#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scraper de statuts de livraison transporteurs (French Bloom) - v4
=================================================================
Lit commandes.csv (numero_commande;transporteur;lien_tracking;date_prevue),
ouvre chaque lien (Playwright headless), ferme les bannieres cookies, lit le
texte rendu (iframes compris) et en deduit :
  - le statut normalise (En cours / En retard / Incident / Livré / À vérifier)
  - la DATE DE LIVRAISON REELLE quand le colis est livre.

Regles robustes calees sur des pages reelles UPS / DB Schenker (JOY) / DPD :
  - "Livré" seulement sur un vrai signal ("Livré le <date>", "Livré <date>",
    "expedition livree", "delivered"...), PAS sur un simple libelle d'etape
    "Colis livré" toujours affiche par les bandeaux de progression ;
  - la date de livraison est ancree sur ce signal (pas la date d'expedition) ;
  - "Incident" sur des phrases precises (evite le menu "Retours") ;
  - pages anti-robot (CAPTCHA) marquees sans tentative de contournement.

INSTALLATION :
    pip install playwright pandas python-dateutil
    playwright install chromium

UTILISATION :
    python scraper_tracking.py                       # commandes.csv -> suivi_resultats.csv
    python scraper_tracking.py mes_commandes.csv out.csv
    python scraper_tracking.py --headful
"""

import os
import sys
import csv
import re
import time
import unicodedata
from datetime import date, datetime

from playwright.sync_api import sync_playwright


def proxy_config():
    """Proxy residentiel via variables d'env (GitHub Secrets). None si absent."""
    server = os.environ.get("PROXY_SERVER", "").strip()
    if not server:
        return None
    cfg = {"server": server}
    if os.environ.get("PROXY_USERNAME"):
        cfg["username"] = os.environ["PROXY_USERNAME"]
    if os.environ.get("PROXY_PASSWORD"):
        cfg["password"] = os.environ["PROXY_PASSWORD"]
    return cfg

# ==================================================================
# 1) OUTILS DATES
# ==================================================================
MOIS_FR = {"janvier": 1, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
           "juillet": 7, "aout": 8, "septembre": 9, "octobre": 10, "novembre": 11, "decembre": 12}
MOIS_RE = "(?:janvier|fevrier|mars|avril|mai|juin|juillet|aout|septembre|octobre|novembre|decembre)"
DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%y", "%Y/%m/%d", "%d.%m.%Y")

DATE_REGEX = re.compile(
    r"(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})"
    r"|(\d{1,2}\s+" + MOIS_RE + r"\s+\d{4})"
    r"|(" + MOIS_RE + r"\s+\d{1,2},?\s+\d{4})"
    r"|(\d{1,2}\s+" + MOIS_RE + r")"
    r"|(" + MOIS_RE + r"\s+\d{1,2})",
    re.IGNORECASE)


def sans_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s or "")
                   if unicodedata.category(c) != "Mn")


def parse_date(s, auj=None):
    """Numerique, francais complet, ou francais SANS annee (annee deduite)."""
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
    if not (mois and jour):
        return None
    if annee is None:
        auj = auj or date.today()
        try:
            d = date(auj.year, mois, jour)
        except ValueError:
            return None
        if d > auj:
            try:
                d = date(auj.year - 1, mois, jour)
            except ValueError:
                return None
        return d
    try:
        return date(annee, mois, jour)
    except ValueError:
        return None


# ==================================================================
# 2) ANALYSE DU TEXTE DE LA PAGE
# ==================================================================
# Incident : phrases precises (le bare "retour" matcherait le menu "Retours")
INCIDENT_PATS = [
    r"\bincident\b", r"\banomalie\b", r"\bavarie\b", r"\brefus", r"\blitige\b",
    r"\ben souffrance\b", r"\bmise en instance\b", r"\ben instance\b",
    r"\bperdu\b", r"\bendommag", r"\bnon distribu", r"\bechec de livraison\b",
    r"\breexpedition\b", r"retour[^.]{0,20}(expediteur|instruction)", r"colis retourne",
]
# Livré : signaux forts (pas le simple libelle d'etape "Colis livré")
DELIVERED_PATS = [
    r"\blivree?s?\s+le\b",
    r"\blivree?s?\s+\d{1,2}\s*[/.\-]",
    r"\blivree?s?\s+conforme",            # VCV/Teliway : "Livré conforme ( 20/07/2026 )"
    r"\blivree?s?\s+cfm\s+conforme",       # SPD/veolog : "LIV Livré CFM Conforme 21/07/2026"
    r"\blivree?s?\s*\(\s*\d{1,2}[/.]",     # "Livré ( 20/07/2026 )"
    r"\blivraison\s+effectuee",
    r"\blivree?s?\s+au\s+destinataire",
    r"\best\s+livree?s?\b",
    r"expedition[^.]{0,40}livree",
    r"\bdelivered\b",
    r"proof of delivery",
    r"\blivree?s?\s+check",
    r"\blivree?s?\s+a\b\s+[a-z0-9]",
]
INTRANSIT_PATS = [
    r"en cours de livraison", r"out for delivery", r"mise en livraison", r"en livraison\b",
    r"en cours d.acheminement", r"en transit", r"in transit",
    r"pris en charge", r"picked up", r"\bexpedie", r"en preparation",
    r"colis remis", r"\barrivage\b",
]
CAPTCHA_MARKERS = [
    "verification robot", "je ne suis pas un robot", "recaptcha", "captcha",
    "verify you are human", "are you a robot", "hcaptcha", "cf-challenge",
    "controle de securite",
]


def est_captcha(text):
    low = sans_accents(text).lower()
    return any(m in low for m in CAPTCHA_MARKERS)


def _delivered_pos(low):
    best = -1
    for pat in DELIVERED_PATS:
        m = re.search(pat, low)
        if m and (best < 0 or m.start() < best):
            best = m.start()
    return best


def _date_livraison(low, anchor):
    dates = [(m.start(), m.group(0)) for m in DATE_REGEX.finditer(low)]
    if not dates:
        return ""
    if anchor >= 0:
        apres = [d for d in dates if d[0] >= anchor - 5]
        cand = apres if apres else dates
        best = min(cand, key=lambda d: abs(d[0] - anchor))
    else:
        best = dates[-1]
    d = parse_date(best[1])
    return d.isoformat() if d else ""


def analyser(text, date_prevue):
    """Retourne (statut_normalise, statut_brut, date_livraison_reelle)."""
    low = sans_accents(text).lower()
    for pat in INCIDENT_PATS:
        m = re.search(pat, low)
        if m:
            return "Incident", m.group(0)[:60], ""
    dp = _delivered_pos(low)
    if dp >= 0:
        return "Livré", "Livré", _date_livraison(low, dp)
    for pat in INTRANSIT_PATS:
        m = re.search(pat, low)
        if m:
            return "En cours", m.group(0), ""
    if date_prevue and date_prevue < date.today():
        return "En retard", "", ""
    return "À vérifier", "", ""


def jours_retard(statut, date_prevue):
    if statut == "En retard" and date_prevue:
        return (date.today() - date_prevue).days
    return 0


# ==================================================================
# 3) NAVIGATEUR
# ==================================================================
CONSENT_SELECTORS = [
    "#onetrust-accept-btn-handler",
    "button#truste-consent-button",
    "button[aria-label*='accept' i]",
    "button:has-text('Tout accepter')",
    "button:has-text('Accepter')",
    "button:has-text('J\\'accepte')",
    "button:has-text('Accept all')",
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


def charger_page(page, lien, tentative=0):
    """Retourne le texte de la page, ou un marqueur [ERREUR]/[CAPTCHA]."""
    try:
        page.goto(lien, wait_until="domcontentloaded", timeout=60000)
        dismiss_consent(page)
        try:
            page.wait_for_load_state("networkidle", timeout=8000)
        except Exception:
            pass
        page.wait_for_timeout(3500)
        text = texte_complet(page)
        if est_captcha(text):
            return "[CAPTCHA]"
        return text
    except Exception as e:
        if tentative < 1:
            page.wait_for_timeout(2000)
            return charger_page(page, lien, tentative + 1)
        return "[ERREUR: " + type(e).__name__ + "]"


# ==================================================================
# 4) BOUCLE PRINCIPALE
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
        prx = proxy_config()
        if prx:
            print("Proxy residentiel actif : " + prx["server"])
        context = browser.new_context(locale="fr-FR", user_agent=UA, proxy=prx)
        page = context.new_page()

        for i, cmd in enumerate(commandes, 1):
            num = cmd.get("numero_commande", "").strip()
            transp = cmd.get("transporteur", "").strip().upper()
            lien = cmd.get("lien_tracking", "").strip()
            d_prev = parse_date(cmd.get("date_prevue", ""))

            print("[" + str(i) + "/" + str(len(commandes)) + "] " + num + " (" + transp + ") ...", end=" ", flush=True)

            if not lien or not lien.startswith("http"):
                statut, brut, date_liv = ("En retard" if (d_prev and d_prev < date.today()) else "À vérifier"), "", ""
            else:
                text = charger_page(page, lien)
                if text in ("[CAPTCHA]",) or text.startswith("[ERREUR"):
                    brut = text
                    statut = "En retard" if (d_prev and d_prev < date.today()) else "À vérifier"
                    date_liv = ""
                else:
                    statut, brut, date_liv = analyser(text, d_prev)

            retard = jours_retard(statut, d_prev)
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
