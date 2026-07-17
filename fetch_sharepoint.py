#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetch_sharepoint.py - Recupere le CSV Bonx depuis SharePoint (French Bloom)

Telecharge le dernier « Monitoring Expeditions.csv » depuis SharePoint via
Microsoft Graph, puis le convertit au format attendu par scraper_tracking.py.

Ecrit DEUX fichiers :
  - monitoring_expeditions.csv : copie brute du fichier SharePoint
  - commandes.csv              : format scraper (numero_commande;transporteur;lien_tracking;date_prevue)

Seules les commandes NON encore livrees (date_livraison_reelle vide) passent au scraper.
Une commande peut arriver en 2 fois (lien puis MAJ EDI) : on garde la version avec URL.

Variables d'environnement :
  MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET   (obligatoires ; alias AZURE_* acceptes)
  SP_HOSTNAME   defaut: frenchbloom75.sharepoint.com
  SP_SITE_PATH  defaut: /sites/Exportlogistique
  SP_FILE_NAME  defaut: Monitoring Expeditions.csv
  SP_DRIVE_ID   optionnel

Dependances : pip install msal requests
"""

import os
import sys
import csv
import io
import unicodedata
import requests
import msal

GRAPH = "https://graph.microsoft.com/v1.0"

# Alias de colonnes tolerees (le vrai fichier utilise numero_commande / pays)
COMMANDE = ("numero_commande", "fnumero_commande")
LIEN_BONX = ("lien_bonx",)
TRANSPORTEUR = ("transporteur",)
NUMERO_SUIVI = ("numero_suivi",)
DATE_SOUH = ("date_livraison_souhaitee",)
DATE_SOUH_P1 = ("date_souhaitee_plus_1j_ouvre",)
DATE_REELLE = ("date_livraison_reelle",)


def val(row, keys):
    for k in keys:
        v = row.get(k)
        if v is not None and str(v).strip() != "":
            return str(v).strip()
    return ""


def env_any(*names, required=False, default=None):
    for n in names:
        v = os.environ.get(n)
        if v:
            return v
    if required:
        raise KeyError("Variable manquante : " + " / ".join(names))
    return default


def env_default(name, default):
    v = os.environ.get(name)
    return v if (v and v.strip()) else default


def sans_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s or "")
                   if unicodedata.category(c) != "Mn")


def get_token():
    tenant = env_any("MS_TENANT_ID", "AZURE_TENANT_ID", required=True)
    client_id = env_any("MS_CLIENT_ID", "AZURE_CLIENT_ID", required=True)
    secret = env_any("MS_CLIENT_SECRET", "AZURE_CLIENT_SECRET", required=True)
    app = msal.ConfidentialClientApplication(
        client_id=client_id, client_credential=secret,
        authority="https://login.microsoftonline.com/" + tenant)
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        raise RuntimeError("Auth Graph echouee : " + str(res.get("error_description", res)))
    return res["access_token"]


def gget(url, token):
    r = requests.get(url, headers={"Authorization": "Bearer " + token}, timeout=60)
    r.raise_for_status()
    return r


def resolve_site_id(token, hostname, site_path):
    return gget(GRAPH + "/sites/" + hostname + ":" + site_path, token).json()["id"]


def find_file(token, site_id, drive_id, file_name):
    if drive_id:
        drives = [drive_id]
    else:
        r = gget(GRAPH + "/sites/" + site_id + "/drives", token)
        drives = [d["id"] for d in r.json().get("value", [])]

    stem = file_name.rsplit(".", 1)[0]
    stem_norm = sans_accents(stem).lower()
    q = sans_accents(stem)
    best = None
    for d in drives:
        try:
            r = gget(GRAPH + "/drives/" + d + "/root/search(q='" + q + "')", token)
        except requests.HTTPError:
            continue
        for item in r.json().get("value", []):
            name = item.get("name", "")
            name_norm = sans_accents(name).lower()
            if item.get("file") and name_norm.endswith(".csv") and name_norm.startswith(stem_norm):
                lm = item.get("lastModifiedDateTime", "")
                if best is None or lm > best[0]:
                    best = (lm, d, item["id"], name)
    if best is None:
        raise FileNotFoundError("Fichier « " + file_name + " » introuvable sur le site.")
    print("Fichier trouve : " + best[3] + " (modifie le " + best[0] + ")")
    return best[1], best[2]


def download_csv(token, drive_id, item_id):
    r = gget(GRAPH + "/drives/" + drive_id + "/items/" + item_id + "/content", token)
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return r.content.decode(enc)
        except UnicodeDecodeError:
            continue
    return r.content.decode("utf-8", errors="replace")


def detect_delim(text):
    head = text.splitlines()[0] if text else ""
    counts = {",": head.count(","), ";": head.count(";"), "\t": head.count("\t")}
    return max(counts, key=counts.get)


def to_commandes(rows):
    par_cmd = {}
    for r in rows:
        if val(r, DATE_REELLE):
            continue  # deja livree
        cmd = val(r, COMMANDE)
        if not cmd:
            continue
        suivi = val(r, NUMERO_SUIVI)
        lien = suivi if suivi.lower().startswith(("http://", "https://")) else ""
        cur = {
            "numero_commande": cmd,
            "transporteur": val(r, TRANSPORTEUR),
            "lien_tracking": lien,
            "date_prevue": val(r, DATE_SOUH_P1) or val(r, DATE_SOUH),
        }
        prev = par_cmd.get(cmd)
        if prev is None or (not prev["lien_tracking"] and lien):
            par_cmd[cmd] = cur
    return list(par_cmd.values())


def main():
    token = get_token()
    hostname = env_default("SP_HOSTNAME", "frenchbloom75.sharepoint.com")
    site_path = env_default("SP_SITE_PATH", "/sites/Exportlogistique")
    file_name = env_default("SP_FILE_NAME", "Monitoring Expéditions.csv")
    drive_id = env_default("SP_DRIVE_ID", "") or None

    print("Site : " + hostname + site_path + " | Fichier : " + file_name)
    site_id = resolve_site_id(token, hostname, site_path)
    drive_id, item_id = find_file(token, site_id, drive_id, file_name)
    text = download_csv(token, drive_id, item_id)

    with open("monitoring_expeditions.csv", "w", encoding="utf-8-sig", newline="") as f:
        f.write(text)

    delim = detect_delim(text)
    rows = [{(k or "").strip(): v for k, v in r.items()}
            for r in csv.DictReader(io.StringIO(text), delimiter=delim)]
    print(str(len(rows)) + " lignes lues depuis SharePoint (separateur '" + repr(delim) + "')")

    commandes = to_commandes(rows)
    with open("commandes.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["numero_commande", "transporteur", "lien_tracking", "date_prevue"], delimiter=";")
        w.writeheader()
        w.writerows(commandes)

    sans_lien = sum(1 for c in commandes if not c["lien_tracking"])
    print(str(len(commandes)) + " commandes a suivre (" + str(sans_lien) + " sans lien exploitable)")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERREUR fetch_sharepoint : " + str(e), file=sys.stderr)
        sys.exit(1)
