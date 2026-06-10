# -*- coding: utf-8 -*-
"""
Serveur Flask - Proxy d'export Power BI
Permet aux utilisateurs anonymes de télécharger des rapports PDF/PPTX
en s'authentifiant via un compte de service Azure AD (Client Credentials)
"""

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from msal import ConfidentialClientApplication
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import time
import os
import tempfile
import logging
from datetime import datetime

# ============================================================
#  CONFIGURATION
# ============================================================

CLIENT_ID     = "49b261db-6807-4852-ad9a-e40e1a4c3826"
TENANT_ID     = "b00276c9-2b6f-45fa-bc61-07a02cbf646e"
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "n6x8Q~GHps-jmsoLc7Tg3eM3.Kt4oNY4tnFHqbNL")

REPORT_ID     = "54fe5aa9-31b7-4eee-bcb8-2e09fbce3451"
WORKSPACE_ID  = "41d4b6d3-34d9-41d3-a051-9f91192cc26a"

FILTER_TABLE  = "Dim_Bookmaker"
FILTER_COLUMN = "Bookmaker"

MAX_WAIT_TIME  = 600
CHECK_INTERVAL = 5

ALLOWED_ORGANISMES = [
    "Winamax",
    "Unibet",
]

CORS_ORIGINS = "*"

# ============================================================
#  CONSTANTES
# ============================================================

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES    = ["https://analysis.windows.net/powerbi/api/.default"]

CONTENT_TYPES = {
    "PDF":  "application/pdf",
    "PPTX": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}
EXTENSIONS = {"PDF": "pdf", "PPTX": "pptx"}

# ============================================================
#  INITIALISATION
# ============================================================

app = Flask(__name__)
CORS(app, origins=CORS_ORIGINS)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
log = logging.getLogger(__name__)

_msal_app = ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=AUTHORITY,
)

# ============================================================
#  FONCTIONS
# ============================================================

def get_access_token():
    result = _msal_app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" in result:
        return result["access_token"]
    log.error("Échec auth Azure AD : %s", result.get("error_description"))
    return None


def make_session():
    session = requests.Session()
    retry = Retry(
        total=3,
        backoff_factor=2,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "POST"],
    )
    session.mount("https://", HTTPAdapter(max_retries=retry))
    return session


def base_url():
    return f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/reports/{REPORT_ID}"


def launch_export(token, fmt, organisme=None):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    body = {"format": fmt}
    if organisme:
        body["powerBIReportConfiguration"] = {
            "reportLevelFilters": [
                {"filter": f"{FILTER_TABLE}/{FILTER_COLUMN} eq '{organisme}'"}
            ]
        }
    resp = make_session().post(f"{base_url()}/ExportTo", headers=headers, json=body, timeout=30)
    if resp.status_code not in (200, 202):
        log.error("Erreur lancement export (%s) : %s", resp.status_code, resp.text[:300])
        return None, headers
    return resp.json()["id"], headers


def wait_for_export(export_id, headers):
    status_url = f"{base_url()}/exports/{export_id}"
    session = make_session()
    elapsed = 0
    while elapsed < MAX_WAIT_TIME:
        time.sleep(CHECK_INTERVAL)
        elapsed += CHECK_INTERVAL
        try:
            resp = session.get(status_url, headers=headers, timeout=30)
        except Exception as e:
            log.warning("Erreur vérification statut : %s", e)
            continue
        if resp.status_code == 202:
            log.info("Export en cours… (%ds)", elapsed)
            continue
        if resp.status_code != 200:
            log.error("Statut inattendu : %s", resp.status_code)
            return False
        status = resp.json().get("status")
        if status == "Succeeded":
            log.info("Export terminé en %ds", elapsed)
            return True
        if status == "Failed":
            log.error("Export échoué : %s", resp.json().get("error"))
            return False
    log.error("Timeout export après %ds", MAX_WAIT_TIME)
    return False


def download_export(export_id, headers, dest_path):
    file_url = f"{base_url()}/exports/{export_id}/file"
    with make_session().get(file_url, headers=headers, stream=True, timeout=(30, 300)) as resp:
        if resp.status_code != 200:
            log.error("Erreur téléchargement (%s)", resp.status_code)
            return False
        with open(dest_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
    return True


def sanitize(value):
    for c in '<>:"/\\|?*\'':
        value = value.replace(c, "_")
    return value.strip()

# ============================================================
#  ROUTES
# ============================================================

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "timestamp": datetime.utcnow().isoformat()})


@app.route("/export", methods=["GET"])
def export():
    fmt       = request.args.get("format", "PDF").upper()
    organisme = request.args.get("organisme", "").strip() or None

    if fmt not in ("PDF", "PPTX"):
        return jsonify({"error": "Format invalide. Utiliser PDF ou PPTX."}), 400

    if organisme and ALLOWED_ORGANISMES and organisme not in ALLOWED_ORGANISMES:
        return jsonify({"error": f"Organisme non autorisé : {organisme}"}), 403

    log.info("Demande export | format=%s | organisme=%s", fmt, organisme or "aucun")

    token = get_access_token()
    if not token:
        return jsonify({"error": "Échec authentification Azure AD."}), 500

    export_id, headers = launch_export(token, fmt, organisme)
    if not export_id:
        return jsonify({"error": "Impossible de lancer l'export Power BI."}), 500

    if not wait_for_export(export_id, headers):
        return jsonify({"error": "L'export Power BI a échoué ou a expiré."}), 500

    suffix = f".{EXTENSIONS[fmt]}"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp_path = tmp.name

    try:
        if not download_export(export_id, headers, tmp_path):
            return jsonify({"error": "Erreur lors du téléchargement du fichier."}), 500

        date_str = datetime.now().strftime("%Y-%m-%d")
        org_part = f"_{sanitize(organisme)}" if organisme else ""
        filename = f"PowerBI{org_part}_{date_str}.{EXTENSIONS[fmt]}"

        log.info("Envoi fichier : %s (%d octets)", filename, os.path.getsize(tmp_path))

        return send_file(
            tmp_path,
            mimetype=CONTENT_TYPES[fmt],
            as_attachment=True,
            download_name=filename,
        )
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass


# ============================================================
#  POINT D'ENTRÉE
# ============================================================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n🚀 Serveur proxy Power BI démarré sur le port {port}")
    app.run(host="0.0.0.0", port=port, debug=False)
