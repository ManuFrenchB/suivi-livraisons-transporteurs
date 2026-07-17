#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scraper de statuts de livraison transporteurs (French Bloom)
============================================================
Lit un CSV de commandes exporté de l'ERP/TMS, ouvre chaque lien de tracking
dans un navigateur headless (Playwright, exécute le JavaScript), extrait le
statut brut, le classe (En cours / En retard / Incident / Livré / À vérifier)
et écrit un CSV de résultats qui alimente le dashboard.

Conçu pour tourner sous Windows / environnement Microsoft, en local.

------------------------------------------------------------------
INSTALLATION (une seule fois)
------------------------------------------------------------------
    pip install playwright pandas python-dateutil
    playwright install chromium

------------------------------------------------------------------
FICHIER D'ENTREE : commandes.csv (séparateur ;)
------------------------------------------------------------------
    numero_commande;transporteur;lien_tracking;date_prevue
    CMD-26754;SPD;https://www.speed-distribution.fr/veolog?code=...&tn=FR26754;2026-07-15
    CMD-JOY01;JOY;https://eschenker.dbschenker.com/app/tracking-public/?...;2026-07-12

  - transporteur : code court (HEP, JOY, VCV, SPD, DHL)
  - date_prevue  : date de livraison prévue (AAAA-MM-JJ). Sert à calculer le retard.
                   Laisser vide si inconnue.

------------------------------------------------------------------
UTILISATION
------------------------------------------------------------------
    python scraper_tracking.py                       # entrée=commandes.csv, sortie=suivi_resultats.csv
    python scraper_tracking.py mes_commandes.csv resultats.csv
    python scraper_tracking.py --headful             # voir le navigateur (debug)
"""

import sys
import csv
import re
import time
from datetime import date, datetime

from playwright.sync_api import sync_playwright

# ==================================================================
# 1) REGLES DE CLASSIFICATION
# ==================================================================
INCIDENT_KEYWORDS = [
    "incident", "anomalie", "avarie", "refus", "retour", "litige",
    "perdu", "endommag", "souffrance", "non distribu", "echec", "échec",
]
DELIVERED_KEYWORDS = [
    "livré", "livree", "livrée", "delivered", "remis", "distribué", "distribue",
]


def classifier(statut_brut: str, date_prevue: date | None) -> str:
    """Retourne le statut normalisé à partir du texte brut + date prévue."""
    t = (statut_brut or "").lower().strip()
    if not t or t.startswith("[erreur") or "erreur" in t:
        return "À vérifier"
    if any(k in t for k in INCIDENT_KEYWORDS):
        return "Incident"
    if any(k in t for k in DELIVERED_KEYWORDS):
        return "Livré"
    if date_prevue and date_prevue < date.today():
        return "En retard"
    return "En cours"


def jours_retard(statut: str, date_prevue: date | None) -> int:
    if statut == "En retard" and date_prevue:
        return (date.today() - date_prevue).days
    return 0


def parse_date(s: str) -> date | None:
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


# ==================================================================
# 2) EXTRACTION PAR TRANSPORTEUR
# ==================================================================
# Pour chaque transporteur : un sélecteur CSS à attendre/lire si tu le connais,
# sinon on retombe sur l'extraction générique (tout le texte visible + mots-clés).
#
# >>> A AFFINER : ouvre une page de tracking dans Chrome, clic droit sur le libellé
#     du statut > Inspecter, et copie le sélecteur CSS dans "selector" ci-dessous.
#     Tu peux aussi m'envoyer le HTML rendu d'une page et je te donne le sélecteur exact.

CARRIERS = {
    "HEP": {"nom": "Heppner",           "selector": None, "wait": 4000},
    "JOY": {"nom": "DB Schenker / DSV",  "selector": None, "wait": 5000},
    "VCV": {"nom": "Teliway VCV",        "selector": None, "wait": 4000},
    "SPD": {"nom": "Speed Distribution", "selector": None, "wait": 4000},
    "DHL": {"nom": "DHL",                "selector": None, "wait": 4000},
}

# Libellés de statut courants à repérer dans le texte de la page (ordre = priorité)
STATUS_PATTERNS = [
    r"(colis livr[ée]|livr[ée]s?(?:\s+(?:et\s+)?sign[ée])?)",
    r"(shipment delivered|delivered|remis au destinataire)",
    r"(en cours de livraison|out for delivery|en livraison)",
    r"(en cours d[e']acheminement)",
    r"(en cours de transport|en transit|in transit)",
    r"(exp[ée]di[ée]|shipped|d[ée]part)",
    r"(enlev[ée]|collect[ée]|pris en charge|picked up)",
    r"(arriv[ée]e?(?:\s+au\s+d[ée]p[ôo]t)?|arrived)",
    r"(d[ée]douanement|customs)",
    r"(en pr[ée]paration|commande enregistr[ée])",
    r"(anomalie[^.\n]*)",
    r"(incident[^.\n]*)",
    r"(retour[^.\n]*)",
    r"(en souffrance|mise en instance|en instance)",
]


def extraire_statut_brut(page, carrier_cfg: dict) -> str:
    """Extrait le libellé de statut de la page rendue."""
    # 1) Sélecteur dédié si fourni
    sel = carrier_cfg.get("selector")
    if sel:
        try:
            el = page.query_selector(sel)
            if el:
                txt = (el.inner_text() or "").strip()
                if txt:
                    return " ".join(txt.split())
        except Exception:
            pass

    # 2) Fallback générique : chercher un libellé de statut dans tout le texte
    try:
        body = page.inner_text("body")
    except Exception:
        body = ""
    body_flat = " ".join(body.split())
    for pat in STATUS_PATTERNS:
        m = re.search(pat, body_flat, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # 3) Rien trouvé
    return ""


def normaliser_url(lien: str) -> str:
    """Corrige les liens Teliway du type '...VEO-FBL&sReference=xxx' sans '?'.
    Sans '?', le paramètre n'est pas transmis à la page -> aucun suivi affiché."""
    if lien and "?" not in lien and "&" in lien:
        lien = lien.replace("&", "?", 1)
    return lien


def scraper_une_commande(page, transporteur: str, lien: str,
                         numero: str = "", dump_dir: str | None = None) -> str:
    cfg = CARRIERS.get(transporteur, {"selector": None, "wait": 4000})
    if not lien or not lien.startswith("http"):
        return ""  # lien vide -> À vérifier
    lien = normaliser_url(lien)
    brut = ""
    try:
        page.goto(lien, wait_until="domcontentloaded", timeout=60000)
        # laisser le JS charger le suivi
        page.wait_for_timeout(cfg.get("wait", 4000))
        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass
        # si un sélecteur est connu, l'attendre explicitement
        if cfg.get("selector"):
            try:
                page.wait_for_selector(cfg["selector"], timeout=8000)
            except Exception:
                pass
        brut = extraire_statut_brut(page, cfg)
    except Exception as e:
        brut = f"[ERREUR: {type(e).__name__}]"

    # DIAGNOSTIC : sauvegarder le HTML des pages sans statut lisible / en erreur
    if dump_dir and (not brut or brut.startswith("[ERREUR")):
        try:
            import os
            os.makedirs(dump_dir, exist_ok=True)
            fname = f"{numero or 'page'}_{transporteur}.html"
            with open(os.path.join(dump_dir, fname), "w", encoding="utf-8") as fh:
                fh.write(page.content())
        except Exception:
            pass
    return brut


# ==================================================================
# 3) BOUCLE PRINCIPALE
# ==================================================================
def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    headful = "--headful" in sys.argv
    entree = args[0] if len(args) > 0 else "commandes.csv"
    sortie = args[1] if len(args) > 1 else "suivi_resultats.csv"

    # Lecture des commandes
    commandes = []
    with open(entree, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            commandes.append(row)
    print(f"{len(commandes)} commandes à traiter depuis {entree}")

    resultats = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not headful)
        context = browser.new_context(locale="fr-FR")
        page = context.new_page()

        for i, cmd in enumerate(commandes, 1):
            num = cmd.get("numero_commande", "").strip()
            transp = cmd.get("transporteur", "").strip().upper()
            lien = cmd.get("lien_tracking", "").strip()
            d_prev = parse_date(cmd.get("date_prevue", ""))

            print(f"[{i}/{len(commandes)}] {num} ({transp}) ...", end=" ", flush=True)
            brut = scraper_une_commande(page, transp, lien, numero=num, dump_dir="debug_html")
            statut = classifier(brut, d_prev)
            retard = jours_retard(statut, d_prev)
            print(f"{brut or '—'} -> {statut}")

            resultats.append({
                "numero_commande": num,
                "transporteur": transp,
                "lien_tracking": lien,
                "date_prevue": d_prev.isoformat() if d_prev else "",
                "statut_brut": brut,
                "statut_normalise": statut,
                "jours_retard": retard,
                "date_extraction": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })
            time.sleep(0.5)  # politesse : ne pas marteler les serveurs

        browser.close()

    # Ecriture des résultats
    champs = ["numero_commande", "transporteur", "lien_tracking", "date_prevue",
              "statut_brut", "statut_normalise", "jours_retard", "date_extraction"]
    with open(sortie, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=champs, delimiter=";")
        w.writeheader()
        w.writerows(resultats)

    # Petit récap console
    from collections import Counter
    c = Counter(r["statut_normalise"] for r in resultats)
    print("\n=== RECAP ===")
    for k in ["En cours", "En retard", "Incident", "Livré", "À vérifier"]:
        print(f"  {k:12}: {c.get(k, 0)}")
    print(f"\nRésultats écrits dans {sortie}")


if __name__ == "__main__":
    main()
