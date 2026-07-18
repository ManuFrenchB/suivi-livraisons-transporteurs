#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scraper de statuts de livraison transporteurs (French Bloom) - v2
=================================================================
Lit commandes.csv (numero_commande;transporteur;lien_tracking;date_prevue),
ouvre chaque lien dans un navigateur headless (Playwright), gere les bannieres
cookies, lit le texte rendu (y compris iframes), extrait :
  - le statut brut + le statut normalise (En cours / En retard / Incident / Livré / À vérifier)
  - la DATE DE LIVRAISON REELLE quand le suivi indique "livré"
et ecrit suivi_resultats.csv.

INSTALLATION :
    pip install playwright pandas python-dateutil
    playwright install chromium

UTILISATION :
    python scraper_tracking.py                       # commandes.csv -> suivi_resultats.csv
    python scraper_tracking.py mes_commandes.csv out.csv
    python scraper_tracking.py --headful             # voir le navigateur (debug)
"""

import sys
import csv
import re
import time
import unicodedata
from datetime import date, datetime

from playwright.sync_api import sync_playwright

# ==================================================================
# 1) CLASSIFICATION
# ==================================================================
INCIDENT_KEYWORDS = [
    "incident", "anomalie", "avarie", "refus", "retour", "litige",
    "perdu", "endommag", "souffrance", "non distribu", "echec", "échec",
    "mise en instance", "en instance", "reexpedition", "réexpédition",
]
DELIVERED_KEYWORDS = [
    "livré", "livree", "livrée", "delivered", "remis", "distribué", "distribue",
    "delivery completed", "proof of delivery", "pod",
]


def sans_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s or "")
                   if unicodedata.category(c) != "Mn")


def classifier(statut_brut, date_prevue):
    t = sans_accents(statut_brut or "").lower().strip()
    if not t:
        return "À vérifier"
    if any(sans_accents(k) in t for k in INCIDENT_KEYWORDS):
        return "Incident"
    if any(sans_accents(k) in t for k in DELIVERED_KEYWORDS):
        return "Livré"
    if date_prevue and date_prevue < date.today():
        return "En retard"
    return "En cours"


def jours_retard(statut, date_prevue):
    if statut == "En retard" and date_prevue:
        return (date.today() - date_prevue).days
    return 0


DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%y", "%Y/%m/%d", "%d.%m.%Y")
MOIS_FR = {"janvier": 1, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
           "juillet": 7, "aout": 8, "septembre": 9, "octobre": 10, "novembre": 11, "decembre": 12}


def parse_date(s):
    s = (s or "").strip()
    if not s:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    sl = sans_accents(s).lower().replace(",", " ")
    mois = jour = annee = None
    for p in sl.split():
        if p in MOIS_FR:
            mois = MOIS_FR[p]
        elif p.isdigit():
            n = int(p)
            if n > 31:
                annee = n
            elif jour is None:
                jour = n
    if mois and jour and annee:
        try:
            return date(annee, mois, jour)
        except ValueError:
            return None
    return None


# ==================================================================
# 2) EXTRACTION
# ==================================================================
STATUS_PATTERNS = [
    r"(livr[ée]s?(?:\s+(?:et\s+)?sign[ée])?)",
    r"(delivered)",
    r"(en cours de livraison)",
    r"(out for delivery)",
    r"(en cours d[e']acheminement)",
    r"(en transit|in transit)",
    r"(exp[ée]di[ée]|shipped)",
    r"(pris en charge|picked up)",
    r"(en pr[ée]paration)",
    r"(anomalie[^.\n]*)",
    r"(incident[^.\n]*)",
    r"(retour[^.\n]*)",
    r"(en souffrance)",
    r"(mise en instance|en instance)",
]

# Dates au format numerique ou francais en toutes lettres
DATE_REGEX = re.compile(
    r"(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})"
    r"|(\d{1,2}\s+(?:janvier|f[ée]vrier|mars|avril|mai|juin|juillet|ao[ûu]t|septembre|octobre|novembre|d[ée]cembre)\s+\d{4})"
    r"|((?:janvier|f[ée]vrier|mars|avril|mai|juin|juillet|ao[ûu]t|septembre|octobre|novembre|d[ée]cembre)\s+\d{1,2},?\s+\d{4})",
    re.IGNORECASE)

CONSENT_SELECTORS = [
    "#onetrust-accept-btn-handler",
    "button#truste-consent-button",
    "button[aria-label*='accept' i]",
    "button:has-text('Tout accepter')",
    "button:has-text('Accepter')",
    "button:has-text('J\\'accepte')",
    "button:has-text('Accept all')",
    "button:has-text('Accept All')",
    "button:has-text('I agree')",
]


def dismiss_consent(page):
    for sel in CONSENT_SELECTORS:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click(timeout=2000)
                page.wait_for_timeout(500)
                return
        except Exception:
            continue


def texte_complet(page):
    """Texte du document principal + de toutes les iframes."""
    parts = []
    try:
        parts.append(page.inner_text("body"))
    except Exception:
        pass
    for fr in page.frames:
        try:
            parts.append(fr.inner_text("body"))
        except Exception:
            continue
    return " ".join(" ".join(p.split()) for p in parts if p)


def extraire_statut(text):
    for pat in STATUS_PATTERNS:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return ""


def extraire_date_livraison(text, statut_norm):
    """Cherche une date proche d'un mot-cle de livraison (uniquement si Livré)."""
    if statut_norm != "Livré":
        return ""
    low = sans_accents(text).lower()
    positions = []
    for kw in ["livr", "delivered", "remis", "distribu", "proof of delivery"]:
        i = low.find(kw)
        if i >= 0:
            positions.append(i)
    dates = [(m.start(), m.group(0)) for m in DATE_REGEX.finditer(text)]
    if not dates:
        return ""
    if positions:
        kwpos = min(positions)
        best = min(dates, key=lambda d: abs(d[0] - kwpos))
        d = parse_date(best[1])
        return d.isoformat() if d else ""
    # sinon, date la plus recente trouvee
    parsed = [parse_date(x[1]) for x in dates]
    parsed = [p for p in parsed if p]
    return max(parsed).isoformat() if parsed else ""


def scraper_une_commande(page, lien, tentative=0):
    if not lien or not lien.startswith("http"):
        return "", ""
    try:
        page.goto(lien, wait_until="domcontentloaded", timeout=45000)
        dismiss_consent(page)
        try:
            page.wait_for_load_state("networkidle", timeout=8000)
        except Exception:
            pass
        page.wait_for_timeout(3500)
        text = texte_complet(page)
        brut = extraire_statut(text)
        statut_norm = classifier(brut, None)
        date_liv = extraire_date_livraison(text, statut_norm)
        return brut, date_liv
    except Exception as e:
        if tentative < 1:
            page.wait_for_timeout(2000)
            return scraper_une_commande(page, lien, tentative + 1)
        return "[ERREUR: " + type(e).__name__ + "]", ""


# ==================================================================
# 3) BOUCLE PRINCIPALE
# ==================================================================
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    headful = "--headful" in sys.argv
    entree = args[0] if len(args) > 0 else "commandes.csv"
    sortie = args[1] if len(args) > 1 else "suivi_resultats.csv"

    commandes = []
    with open(entree, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f, delimiter=";"):
            commandes.append(row)
    print(str(len(commandes)) + " commandes a traiter depuis " + entree)

    resultats = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not headful)
        context = browser.new_context(locale="fr-FR", user_agent=UA)
        page = context.new_page()

        for i, cmd in enumerate(commandes, 1):
            num = cmd.get("numero_commande", "").strip()
            transp = cmd.get("transporteur", "").strip().upper()
            lien = cmd.get("lien_tracking", "").strip()
            d_prev = parse_date(cmd.get("date_prevue", ""))

            print("[" + str(i) + "/" + str(len(commandes)) + "] " + num + " (" + transp + ") ...", end=" ", flush=True)
            brut, date_liv = scraper_une_commande(page, lien)
            statut = classifier(brut, d_prev)
            retard = jours_retard(statut, d_prev)
            # coherence : si on a une date de livraison, c'est Livré
            if date_liv and statut != "Livré":
                statut = "Livré"
                retard = 0
            print((brut or "—") + " -> " + statut + (" (" + date_liv + ")" if date_liv else ""))

            resultats.append({
                "numero_commande": num,
                "transporteur": transp,
                "lien_tracking": lien,
                "date_prevue": d_prev.isoformat() if d_prev else "",
                "statut_brut": brut,
                "statut_normalise": statut,
                "jours_retard": retard,
                "date_livraison_reelle": date_liv,
                "date_extraction": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })
            time.sleep(0.4)

        browser.close()

    champs = ["numero_commande", "transporteur", "lien_tracking", "date_prevue",
              "statut_brut", "statut_normalise", "jours_retard",
              "date_livraison_reelle", "date_extraction"]
    with open(sortie, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=champs, delimiter=";")
        w.writeheader()
        w.writerows(resultats)

    from collections import Counter
    c = Counter(r["statut_normalise"] for r in resultats)
    nd = sum(1 for r in resultats if r["date_livraison_reelle"])
    print("\n=== RECAP ===")
    for k in ["En cours", "En retard", "Incident", "Livré", "À vérifier"]:
        print("  " + k.ljust(12) + ": " + str(c.get(k, 0)))
    print("  dates de livraison recuperees : " + str(nd))
    print("\nResultats ecrits dans " + sortie)


if __name__ == "__main__":
    main()
