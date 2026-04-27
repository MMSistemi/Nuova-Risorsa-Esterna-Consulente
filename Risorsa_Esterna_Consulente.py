import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io
import unicodedata
import re

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    cfg = pd.read_excel(io.BytesIO(data), sheet_name="Consulente", engine="openpyxl")

    grp_df = (
        cfg[cfg["Section"] == "InserimentoGruppi"][["Key/App", "Label/Gruppi/Value"]]
        .rename(columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"})
    )
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"].astype(str)))

    def_df = (
        cfg[cfg["Section"] == "Defaults"][["Key/App", "Label/Gruppi/Value"]]
        .rename(columns={"Key/App": "key", "Label/Gruppi/Value": "value"})
    )
    defaults = dict(zip(def_df["key"], def_df["value"].astype(str)))
    return gruppi, defaults

# ------------------------------------------------------------
# Utility functions
# ------------------------------------------------------------
def auto_quote(fields, quotechar='"', predicate=lambda s: ' ' in s):
    out = []
    for f in fields:
        s = str(f)
        if predicate(s):
            out.append(f'{quotechar}{s}{quotechar}')
        else:
            out.append(s)
    return out

def normalize_name(s: str) -> str:
    nfkd = unicodedata.normalize('NFKD', s)
    ascii_str = nfkd.encode('ASCII', 'ignore').decode()
    return ascii_str.replace(' ', '').replace("'", '').lower()

def formatta_data(data: str) -> str:
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data

# ------------------------------------------------------------
# Generazione SAMAccountName
# ------------------------------------------------------------
def genera_samaccountname(nome, cognome, secondo_nome="", secondo_cognome="", esterno=False):
    n, sn = normalize_name(nome), normalize_name(secondo_nome)
    c, sc = normalize_name(cognome), normalize_name(secondo_cognome)
    suffix = ".ext" if esterno else ""
    limit = 16 if esterno else 20

    cand1 = f"{n}{sn}.{c}{sc}"
    if len(cand1) <= limit:
        return cand1 + suffix

    cand2 = f"{n[:1]}{sn[:1]}.{c}{sc}"
    if len(cand2) <= limit:
        return cand2 + suffix

    base = f"{n[:1]}{sn[:1]}.{c}"
    return base[:limit] + suffix

# ------------------------------------------------------------
# Display name
# ------------------------------------------------------------
def build_full_name(cognome, secondo_cognome, nome, secondo_nome, esterno=False):
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    return " ".join(parts) + (" (esterno)" if esterno else "")

# ------------------------------------------------------------
# Header CSV
# ------------------------------------------------------------
HEADER_USER = [
    "sAMAccountName","Creation","OU","Name","DisplayName","cn","GivenName","Surname",
    "employeeNumber","employeeID","department","Description","passwordNeverExpired",
    "ExpireDate","userprincipalname","mail","mobile","RimozioneGruppo","InserimentoGruppo",
    "disable","moveToOU","telephoneNumber","company"
]

HEADER_COMP = [
    "Computer","OU","add_mail","remove_mail","add_mobile","remove_mobile",
    "add_userprincipalname","remove_userprincipalname","disable","moveToOU"
]

# ------------------------------------------------------------
# App
# ------------------------------------------------------------
st.set_page_config(page_title="1.3 Risorsa Esterna: Consulente")
st.title("1.3 Risorsa Esterna: Consulente")

config_file = st.file_uploader("Carica il file di configurazione (config.xlsx)", type=["xlsx"])
if not config_file:
    st.warning("Caricare il file di configurazione.")
    st.stop()

try:
    gruppi, defaults = load_config_from_bytes(config_file.read())
except Exception as e:
    st.error(f"Errore lettura configurazione: {e}")
    st.stop()

# Defaults
ou_value = defaults.get("ou_default", "")
department_default = defaults.get("department_default", "")
description_default = defaults.get("description_default", "<PC>")
company = defaults.get("company_default", "")
inserimento_base = gruppi.get("esterna_consulente", "")
inserimento_noemail = gruppi.get("esterna_consulente_No_email", "")

o365_groups_raw = defaults.get("o365_groups", "").strip()
o365_std = defaults.get("grp_o365_standard", "").strip()
o365_team = defaults.get("grp_o365_teams", "").strip()
o365_cop = defaults.get("grp_o365_copilot", "").strip()

# ✅ NUOVO DEFAULT FOORBAN
grp_foorban = defaults.get("grp_foorban", "").strip()

# ------------------------------------------------------------
# Input UI
# ------------------------------------------------------------
cognome = st.text_input("Cognome").strip().capitalize()
secondo_cognome = st.text_input("Secondo Cognome").strip().capitalize()
nome = st.text_input("Nome").strip().capitalize()
secondo_nome = st.text_input("Secondo Nome").strip().capitalize()
cf = st.text_input("Codice Fiscale").strip()
telefono = st.text_input("Mobile").replace(" ", "")
description = st.text_input("PC", description_default).strip()
exp_date = st.text_input(
    "Data di Fine (gg-mm-aaaa)",
    defaults.get("expire_default", "30-06-2025")
).strip()

email_flag = st.radio("Email Consip necessaria?", ["Sì", "No"]) == "Sì"

if not email_flag:
    custom_email = st.text_input("Email Personalizzata").strip()

if email_flag:
    profil_flag = st.checkbox("Profilazione SM?")
    sm_lines = st.text_area("SM su quali va profilato").splitlines() if profil_flag else []
else:
    profil_flag = False
    sm_lines = []


# ------------------------------------------------------------
# Template Mail
# ------------------------------------------------------------
if email_flag and st.button("Template per Posta Elettronica"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    upn = f"{sAM}@consip.it"

    st.markdown("Ciao.  \nRichiedo cortesemente la definizione di una casella di posta come sottoindicato.")

    st.markdown(f"""
| Campo | Valore |
|------|--------|
| Tipo Utenza | Remota |
| Utenza | {sAM} |
| Alias | {sAM} |
| Display name | {cn} |
| Common name | {cn} |
| e-mail | {upn} |
| e-mail secondaria | {upn} |
""")

    st.markdown("Inviare batch di notifica migrazione mail a: imac@consip.it")

    # ✅ BLOCCO FOORBAN
    if grp_foorban:
        st.markdown(f"""
Aggiungere utenza al gruppo Azure:

**{grp_foorban}**
""")

    if profil_flag and sm_lines:
        st.markdown("Profilare su SM:")
        for sm in sm_lines:
            if sm.strip():
                st.markdown(f"- {sm}")

    st.markdown("Grazie  \nSaluti")


# ------------------------------------------------------------
# CSV
# ------------------------------------------------------------
if st.button("Genera CSV Consulente"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, True)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, True)
    exp_fmt = formatta_data(exp_date)

    upn = f"{sAM}@consip.it"
    mail = upn if email_flag else (custom_email or upn)
    mobile = f"+39 {telefono}" if telefono else ""

    inser_grp = inserimento_base if email_flag else inserimento_noemail

    row_user = [
        sAM, "SI", ou_value, cn, cn, cn,
        f"{nome} {secondo_nome}".strip(),
        f"{cognome} {secondo_cognome}".strip(),
        cf, "", department_default, description,
        "No", exp_fmt, upn, mail, mobile,
        "", "", "", "", "", company
    ]

    row_comp = [
        description or "", "", f"{sAM}@consip.it", "",
        mobile, "", cn, "", "", ""
    ]

    profile_groups = []
    if o365_groups_raw:
        for t in re.split(r"[;,]", o365_groups_raw):
            if t.strip():
                profile_groups.append(t.strip())

    if inser_grp:
        profile_groups.append(inser_grp.strip())

    profile_row = [""] * len(HEADER_USER)
    profile_row[0] = sAM
    profile_row[18] = ";".join(profile_groups)

    def make_csv(header, row):
        buf = io.StringIO()
        w = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar="\\")
        w.writerow(header)
        w.writerow(auto_quote(row))
        buf.seek(0)
        return buf.getvalue()

    st.download_button("📥 CSV Utente", make_csv(HEADER_USER, row_user), f"{sAM}_utente.csv")
    st.download_button("📥 CSV Computer", make_csv(HEADER_COMP, row_comp), f"{sAM}_computer.csv")
    st.download_button("📥 CSV Profilazione", make_csv(HEADER_USER, profile_row), f"{sAM}_profilazione.csv")

    st.success(f"✅ CSV generati per {sAM}")
