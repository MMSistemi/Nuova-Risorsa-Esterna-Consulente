import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io
import unicodedata
import re
import zipfile


# ------------------------------------------------------------
# Caricamento configurazione da Excel
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):

    workbook = pd.read_excel(
        io.BytesIO(data),
        sheet_name=None,
        engine="openpyxl"
    )

    # Supporta sia "Consulente" sia un eventuale nome alternativo
    if "Consulente" in workbook:
        cfg = workbook["Consulente"]
    elif "Risorsa Esterna" in workbook:
        cfg = workbook["Risorsa Esterna"]
    else:
        cfg = pd.DataFrame()

    # --------------------------------------------------------
    # Normalizzazione colonne
    # --------------------------------------------------------
    if not cfg.empty:

        cfg = cfg.rename(
            columns={
                column: str(column).strip()
                for column in cfg.columns
            }
        )

        # Uniforma Tipo/TIPO
        if "Tipo" in cfg.columns and "TIPO" not in cfg.columns:
            cfg = cfg.rename(
                columns={
                    "Tipo": "TIPO"
                }
            )

        for column in [
            "Section",
            "Key/App",
            "Label/Gruppi/Value",
            "TIPO"
        ]:
            if column in cfg.columns:
                cfg[column] = (
                    cfg[column]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

    required_columns = {
        "Section",
        "Key/App",
        "Label/Gruppi/Value"
    }

    if (
        not cfg.empty
        and required_columns.issubset(cfg.columns)
    ):

        # ----------------------------------------------------
        # OU
        # ----------------------------------------------------
        ou_df = cfg[
            cfg["Section"] == "OU"
        ][
            [
                "Key/App",
                "Label/Gruppi/Value"
            ]
        ].rename(
            columns={
                "Key/App": "key",
                "Label/Gruppi/Value": "label"
            }
        )

        ou_options = dict(
            zip(
                ou_df["key"],
                ou_df["label"]
            )
        )

        # ----------------------------------------------------
        # Inserimento gruppi
        # ----------------------------------------------------
        gruppi_df = cfg[
            cfg["Section"] == "InserimentoGruppi"
        ][
            [
                "Key/App",
                "Label/Gruppi/Value"
            ]
        ].rename(
            columns={
                "Key/App": "app",
                "Label/Gruppi/Value": "gruppi"
            }
        )

        gruppi = dict(
            zip(
                gruppi_df["app"],
                gruppi_df["gruppi"]
            )
        )

        # ----------------------------------------------------
        # Defaults
        # ----------------------------------------------------
        defaults_df = cfg[
            cfg["Section"] == "Defaults"
        ][
            [
                "Key/App",
                "Label/Gruppi/Value"
            ]
        ].rename(
            columns={
                "Key/App": "key",
                "Label/Gruppi/Value": "value"
            }
        )

        defaults = dict(
            zip(
                defaults_df["key"],
                defaults_df["value"]
            )
        )

    else:
        ou_options = {}
        gruppi = {}
        defaults = {}

    # --------------------------------------------------------
    # Estrazione righe Defaults con TIPO = PEL
    # Mantiene l'ordine del file Excel
    # --------------------------------------------------------
    pel_items = []

    if (
        not cfg.empty
        and required_columns.issubset(cfg.columns)
        and "TIPO" in cfg.columns
    ):

        pel_df = cfg[
            (cfg["Section"] == "Defaults")
            & (cfg["TIPO"].str.upper() == "PEL")
        ][
            [
                "Key/App",
                "Label/Gruppi/Value"
            ]
        ]

        for _, row in pel_df.iterrows():

            key = str(
                row["Key/App"]
            ).strip()

            value = str(
                row["Label/Gruppi/Value"]
            ).strip()

            if value:
                pel_items.append(
                    (key, value)
                )

    return (
        ou_options,
        gruppi,
        defaults,
        pel_items
    )


# ------------------------------------------------------------
# Configurazione pagina
# ------------------------------------------------------------
st.set_page_config(
    page_title="1.3 Risorsa Esterna: Consulente",
    layout="wide"
)

st.title(
    "1.3 Risorsa Esterna: Consulente"
)


# ------------------------------------------------------------
# Caricamento configurazione
# ------------------------------------------------------------
config_file = st.file_uploader(
    "Carica il file di configurazione (config.xlsx)",
    type=["xlsx"],
    help=(
        "Il file deve contenere il foglio "
        "'Consulente' con le colonne Section, "
        "Key/App, Label/Gruppi/Value e TIPO."
    )
)

if not config_file:

    st.warning(
        "Caricare il file di configurazione "
        "per continuare."
    )

    st.stop()


(
    ou_options,
    gruppi,
    defaults,
    pel_items
) = load_config_from_bytes(
    config_file.read()
)


# ------------------------------------------------------------
# Funzioni di gestione configurazione
# ------------------------------------------------------------
def split_values(value):

    return [
        item.strip()
        for item in re.split(
            r"[;,]",
            str(value)
        )
        if item.strip()
    ]


def unique_in_order(values):

    result = []

    for value in values:

        value = str(value).strip()

        if value and value not in result:
            result.append(value)

    return result


# ------------------------------------------------------------
# Gruppi O365 letti dalle chiavi grp_o365_*
# ------------------------------------------------------------
o365_groups = []

for key, value in defaults.items():

    key = str(key).strip()

    if key.startswith("grp_o365_") and value:

        for group in split_values(value):

            # Corregge automaticamente "365 ..." in "O365 ..."
            if group.startswith("365 "):
                group = "O" + group

            if group not in o365_groups:
                o365_groups.append(group)


# Compatibilità con la vecchia chiave o365_groups
old_o365_groups = defaults.get(
    "o365_groups",
    ""
)

if old_o365_groups:

    for group in split_values(
        old_o365_groups
    ):

        if group.startswith("365 "):
            group = "O" + group

        if group not in o365_groups:
            o365_groups.append(group)


# ------------------------------------------------------------
# Funzioni di utilità
# ------------------------------------------------------------
def normalize_name(value: str) -> str:

    normalized = unicodedata.normalize(
        "NFKD",
        value
    )

    ascii_value = (
        normalized
        .encode(
            "ASCII",
            "ignore"
        )
        .decode()
    )

    return (
        ascii_value
        .replace(" ", "")
        .replace("'", "")
        .lower()
    )


def formatta_data(data: str) -> str:

    data = str(data).strip()

    for formato in [
        "%d-%m-%Y",
        "%d/%m/%Y"
    ]:
        try:

            data_convertita = datetime.strptime(
                data,
                formato
            )

            # Mantiene la logica precedente:
            # scadenza impostata al giorno successivo
            data_convertita += timedelta(
                days=1
            )

            return data_convertita.strftime(
                "%m/%d/%Y 00:00"
            )

        except ValueError:
            continue

    return data


def genera_samaccountname(
    nome,
    cognome,
    secondo_nome="",
    secondo_cognome="",
    esterno=False
):

    nome_normalizzato = normalize_name(
        nome
    )

    secondo_nome_normalizzato = normalize_name(
        secondo_nome
    )

    cognome_normalizzato = normalize_name(
        cognome
    )

    secondo_cognome_normalizzato = normalize_name(
        secondo_cognome
    )

    suffix = ".ext" if esterno else ""

    # Per gli esterni la parte prima di .ext
    # deve essere al massimo di 16 caratteri
    limit = 16 if esterno else 20

    candidate_1 = (
        f"{nome_normalizzato}"
        f"{secondo_nome_normalizzato}."
        f"{cognome_normalizzato}"
        f"{secondo_cognome_normalizzato}"
    )

    if len(candidate_1) <= limit:
        return candidate_1 + suffix

    candidate_2 = (
        f"{nome_normalizzato[:1]}"
        f"{secondo_nome_normalizzato[:1]}."
        f"{cognome_normalizzato}"
        f"{secondo_cognome_normalizzato}"
    )

    if len(candidate_2) <= limit:
        return candidate_2 + suffix

    base = (
        f"{nome_normalizzato[:1]}"
        f"{secondo_nome_normalizzato[:1]}."
        f"{cognome_normalizzato}"
    )

    return base[:limit] + suffix


def build_full_name(
    cognome,
    secondo_cognome,
    nome,
    secondo_nome,
    esterno=False
):

    parts = [
        part
        for part in [
            cognome,
            secondo_cognome,
            nome,
            secondo_nome
        ]
        if part
    ]

    full_name = " ".join(
        parts
    )

    if esterno:
        full_name += " (esterno)"

    return full_name


def build_name_without_external_label(
    cognome,
    secondo_cognome,
    nome,
    secondo_nome
):

    return " ".join(
        part
        for part in [
            cognome,
            secondo_cognome,
            nome,
            secondo_nome
        ]
        if part
    )


# ------------------------------------------------------------
# Costruzione gruppi Entra ID
# ------------------------------------------------------------
def build_entraid_groups(
    pel_items,
    email_consip
):

    """
    Regole:
    - dl_standard e dl_vip non sono gruppi Entra ID;
    - pillole viene escluso;
    - vengono mantenuti ordine e valori del config;
    - i gruppi vengono generati solo se è richiesta
      la casella Email Consip.
    """

    if not email_consip:
        return []

    entra_groups = []

    for key, value in pel_items:

        key = str(key).strip()
        value = str(value).strip()

        if not value:
            continue

        if key in (
            "dl_standard",
            "dl_vip",
            "pillole"
        ):
            continue

        entra_groups.extend(
            split_values(value)
        )

    return unique_in_order(
        entra_groups
    )


# ------------------------------------------------------------
# Costruzione CSV
# ------------------------------------------------------------
def auto_quote(
    fields,
    quotechar='"',
    predicate=lambda value: " " in value
):

    """
    Racchiude tra doppi apici ogni valore
    contenente almeno uno spazio.
    """

    output = []

    for field in fields:

        value = str(field)

        if predicate(value):

            output.append(
                f'{quotechar}{value}{quotechar}'
            )

        else:
            output.append(
                value
            )

    return output


def prepara_csv_standard(
    row,
    header
):

    """
    Mantiene la costruzione precedente:
    ogni campo contenente almeno uno spazio
    viene racchiuso tra doppi apici.
    """

    buffer = io.StringIO()

    writer = csv.writer(
        buffer,
        quoting=csv.QUOTE_NONE,
        escapechar="\\",
        lineterminator="\n"
    )

    quoted_row = auto_quote(
        row,
        quotechar='"',
        predicate=lambda value: " " in value
    )

    writer.writerow(
        header
    )

    writer.writerow(
        quoted_row
    )

    return buffer.getvalue()


# ------------------------------------------------------------
# Header CSV
# ------------------------------------------------------------
HEADER_UTENTE = [
    "sAMAccountName",
    "Creation",
    "OU",
    "Name",
    "DisplayName",
    "cn",
    "GivenName",
    "Surname",
    "employeeNumber",
    "employeeID",
    "department",
    "Description",
    "passwordNeverExpired",
    "ExpireDate",
    "userprincipalname",
    "mail",
    "mobile",
    "RimozioneGruppo",
    "InserimentoGruppo",
    "disable",
    "moveToOU",
    "telephoneNumber",
    "company"
]


HEADER_COMPUTER = [
    "Computer",
    "OU",
    "add_mail",
    "remove_mail",
    "add_mobile",
    "remove_mobile",
    "add_userprincipalname",
    "remove_userprincipalname",
    "disable",
    "moveToOU"
]


# ------------------------------------------------------------
# Valori di configurazione
# ------------------------------------------------------------
department_default = defaults.get(
    "department_default",
    ""
)

description_default = defaults.get(
    "description_default",
    "<PC>"
)

company = defaults.get(
    "company_default",
    ""
)

expire_default = defaults.get(
    "expire_default",
    ""
)


# ------------------------------------------------------------
# Modulo inserimento consulente
# ------------------------------------------------------------
st.subheader(
    "Modulo Inserimento Nuovo Consulente"
)


colonna_1, colonna_2 = st.columns(2)


with colonna_1:

    cognome = (
        st.text_input(
            "Cognome"
        )
        .strip()
        .title()
    )

    secondo_cognome = (
        st.text_input(
            "Secondo Cognome"
        )
        .strip()
        .title()
    )

    nome = (
        st.text_input(
            "Nome"
        )
        .strip()
        .title()
    )

    secondo_nome = (
        st.text_input(
            "Secondo Nome"
        )
        .strip()
        .title()
    )

    codice_fiscale = st.text_input(
        "Codice Fiscale"
    ).strip()

    numero_telefono = (
        st.text_input(
            "Mobile (+39 già inserito)"
        )
        .replace(" ", "")
        .strip()
    )


with colonna_2:

    description = st.text_input(
        "PC",
        description_default
    ).strip()

    manager = st.text_input(
        "Manager / Referente"
    ).strip()

    data_inizio = st.text_input(
        "Data inizio (gg/mm/aaaa)"
    ).strip()

    data_fine = st.text_input(
        "Data di fine (gg/mm/aaaa)",
        expire_default
    ).strip()

    email_flag = (
        st.radio(
            "Email Consip necessaria?",
            [
                "Sì",
                "No"
            ],
            horizontal=True
        )
        == "Sì"
    )

    if not email_flag:

        custom_email = st.text_input(
            "Email personalizzata"
        ).strip()

    else:
        custom_email = ""


# ------------------------------------------------------------
# Tipologia utente e OU
# ------------------------------------------------------------
ou_values = (
    list(
        ou_options.values()
    )
    if ou_options
    else []
)

default_ou = defaults.get(
    "ou_default",
    ou_values[0]
    if ou_values
    else ""
)

index_default = 0

if default_ou in ou_values:

    index_default = ou_values.index(
        default_ou
    )


if ou_values:

    label_ou = st.selectbox(
        "Tipologia Utente / OU",
        ou_values,
        index=index_default
    )

    selected_key = list(
        ou_options.keys()
    )[
        ou_values.index(
            label_ou
        )
    ]

    ou_value = ou_options.get(
        selected_key,
        ""
    )

else:

    ou_value = st.text_input(
        "OU",
        default_ou
    ).strip()

    selected_key = ""


# ------------------------------------------------------------
# Profilazione SM
# ------------------------------------------------------------
profilazione_sm = (
    st.checkbox(
        "Profilazione SM?"
    )
    if email_flag
    else False
)

if profilazione_sm:

    sm_lines = [
        line.strip()
        for line in st.text_area(
            "SM su cui profilare il consulente, una per riga"
        ).splitlines()
        if line.strip()
    ]

else:

    sm_lines = []


# ------------------------------------------------------------
# Generazione flusso
# ------------------------------------------------------------
if st.button(
    "Genera CSV Consulente",
    type="primary"
):

    # --------------------------------------------------------
    # Validazioni
    # --------------------------------------------------------
    if not cognome or not nome:

        st.error(
            "Inserire almeno Cognome e Nome."
        )

        st.stop()

    if not manager:

        st.error(
            "Inserire il Manager / Referente."
        )

        st.stop()

    if not data_inizio:

        st.error(
            "Inserire la Data inizio."
        )

        st.stop()

    if not email_flag and not custom_email:

        st.error(
            "Se Email Consip non è necessaria, "
            "inserire l'Email personalizzata."
        )

        st.stop()

    # --------------------------------------------------------
    # Dati principali
    # --------------------------------------------------------
    sam = genera_samaccountname(
        nome,
        cognome,
        secondo_nome,
        secondo_cognome,
        True
    )

    cn = build_full_name(
        cognome,
        secondo_cognome,
        nome,
        secondo_nome,
        True
    )

    nome_senza_esterno = build_name_without_external_label(
        cognome,
        secondo_cognome,
        nome,
        secondo_nome
    )

    given_name = (
        f"{nome} {secondo_nome}"
        .strip()
    )

    surname = (
        f"{cognome} {secondo_cognome}"
        .strip()
    )

    upn = f"{sam}@consip.it"

    mail = (
        upn
        if email_flag
        else custom_email
    )

    mobile = (
        f"+39 {numero_telefono}"
        if numero_telefono
        else ""
    )

    expire_date = formatta_data(
        data_fine
    )

    pc_asset = (
        description
        or "<PC>"
    )

    manager_output = (
        manager
        or "<Manager>"
    )

    data_inizio_output = (
        data_inizio
        or "<data inizio>"
    )

    persona_nome_file = nome_senza_esterno

    persona_nome_underscore = "_".join(
        part
        for part in [
            cognome,
            secondo_cognome,
            nome,
            secondo_nome
        ]
        if part
    )

    # ========================================================
    # FILE 1 - UTENZA
    # ========================================================
    row_utente = [
        sam,
        "SI",
        ou_value,
        cn,
        cn,
        cn,
        given_name,
        surname,
        codice_fiscale,
        "",
        department_default,
        pc_asset,
        "No",
        expire_date,
        upn,
        mail,
        mobile,
        "",
        "",
        "",
        "",
        "",
        company
    ]

    csv_utente = prepara_csv_standard(
        row_utente,
        HEADER_UTENTE
    )

    # ========================================================
    # FILE 2 - CASELLA
    # ========================================================
    if email_flag:
    
        casella_lines = [
            "Tipo Utenza --- Remota",
            f"Utenza --- {sam}",
            f"Alias --- {sam}",
            f"Display name --- {cn}",
            f"Common name --- {cn}",
            f"e-mail --- {upn}",
            (
                "e-mail secondaria --- "
                f"{sam}@consipspa.mail.onmicrosoft.com"
            )
        ]
    
        # ----------------------------------------------------
        # Eventuale profilazione SM
        # ----------------------------------------------------
        if profilazione_sm and sm_lines:
    
            casella_lines.append("")
            casella_lines.append("Profilare su SM:")
    
            for sm in sm_lines:
                casella_lines.append(
                    f"- {sm}"
                )
    
        csv_casella = (
            "\n".join(
                casella_lines
            )
            + "\n"
        )
    
    else:
    
        casella_lines = [
            "Casella Consip non richiesta",
            f"Utenza --- {sam}",
            f"Email esterna --- {custom_email}"
        ]
    
        csv_casella = (
            "\n".join(
                casella_lines
            )
            + "\n"
        )
    # ========================================================
    # FILE 3 - PROFILAZIONE GRUPPI AD
    # ========================================================
    if email_flag:

        inserimento_gruppo_value = (
            gruppi.get(
                "esterna_consulente",
                ""
            )
            or ""
        )

    else:

        inserimento_gruppo_value = (
            gruppi.get(
                "esterna_consulente_No_email",
                ""
            )
            or ""
        )

    gruppi_consulente = []

    for group in split_values(
        inserimento_gruppo_value
    ):

        if group.startswith("365 "):
            group = "O" + group

        gruppi_consulente.append(
            group
        )

    gruppi_ad = unique_in_order(
        o365_groups
        + gruppi_consulente
    )

    row_gruppi_ad = [
        ""
    ] * len(
        HEADER_UTENTE
    )

    row_gruppi_ad[0] = sam

    inserimento_gruppo_index = HEADER_UTENTE.index(
        "InserimentoGruppo"
    )

    row_gruppi_ad[
        inserimento_gruppo_index
    ] = ";".join(
        gruppi_ad
    )

    csv_gruppi_ad = prepara_csv_standard(
        row_gruppi_ad,
        HEADER_UTENTE
    )

    # ========================================================
    # FILE 4 - PROFILAZIONE GRUPPI ENTRA ID
    # ========================================================
    gruppi_entraid = build_entraid_groups(
        pel_items,
        email_flag
    )

    entra_lines = [
        (
            f"Per l'utenza {sam} "
            "aggiungere i gruppi EntraID:"
        )
    ]

    if gruppi_entraid:

        entra_lines.extend(
            gruppi_entraid
        )

    else:

        entra_lines.append(
            "Nessun gruppo EntraID previsto."
        )

    if profilazione_sm and sm_lines:

        entra_lines.append(
            ""
        )

        entra_lines.append(
            "Profilare su SM:"
        )

        entra_lines.extend(
            sm_lines
        )

    csv_entraid = (
        "\n".join(
            entra_lines
        )
        + "\n"
    )

    # ========================================================
    # FILE 5 - COMPUTER
    # Inserito nello ZIP ma non nella lista attività punto 02
    # ========================================================
    row_computer = [
        pc_asset,
        "",
        upn,
        "",
        mobile,
        "",
        cn,
        "",
        "",
        ""
    ]

    csv_computer = prepara_csv_standard(
        row_computer,
        HEADER_COMPUTER
    )

    # --------------------------------------------------------
    # Nomi file
    # --------------------------------------------------------
    filename_utente = (
        f"1_{persona_nome_file}_utenza.csv"
    )

    filename_casella = (
        f"2_{persona_nome_file}_casella.csv"
    )

    filename_gruppi_ad = (
        f"3_{persona_nome_file}"
        "_profilazione_Gruppi_AD.csv"
    )

    filename_entraid = (
        f"4_{persona_nome_file}"
        "_profilazione_Gruppi_EntraID.csv"
    )

    filename_computer = (
        f"5_{persona_nome_underscore}"
        "_computer.csv"
    )

    # ========================================================
    # 01 - ASSEGNAZIONE ASSET
    # ========================================================
    subject_assegnazione_asset = (
        "01 - Assegnazione Asset - "
        f"{cn} → GSI"
    )

    message_assegnazione_asset = (
        "Per la nuova risorsa CONSULENTE "
        f"{cn} assegnare asset:\n"
        f"- PC Nome asset {pc_asset} "
        f"Referente: {manager_output}"
    )

    # ========================================================
    # 02 - PROVISIONING UTENTE
    # ========================================================
    subject_provisioning = (
        "02 - Provisioning Utente "
        "(utenza/Casella/Profilazione) - "
        f"{cn} → GIIT"
    )

    message_provisioning = (
        "Ciao,\n"
        "si richiede di elaborare i file contenuti "
        "nella cartella compressa .zip "
        "nel seguente ordine:\n"
        f"1. {filename_utente} per la creazione "
        "di un'utenza On prem su AD\n"
        f"2. {filename_casella} per la creazione "
        "di un'utenza ibrida come indicato\n"
        f"3. {filename_gruppi_ad} per inserimento "
        "dei gruppi AD\n"
        f"4. {filename_entraid} per inserimento "
        "dei gruppi EntraID"
    )

    # ========================================================
    # 03 - PROVISIONING PC
    # Nel testo viene usato il nome senza "(esterno)"
    # ========================================================
    subject_provisioning_pc = (
        "03 - Provisioning PC - "
        f"{cn} → GHDW"
    )

    message_provisioning_pc = (
        "Ciao,\n"
        f"per {nome_senza_esterno} ritirare dal 599 "
        "per CONFIGURAZIONE i seguenti asset:\n"
        f"- PC Nome asset {pc_asset}\n"
        "\n"
        f"la risorsa prende servizio "
        f"{data_inizio_output}"
    )

    # ========================================================
    # 09 - FIRMA ASSET
    # ========================================================
    subject_firma_asset = (
        "09 - Firma Asset - "
        f"{cn} → GHDW"
    )

    message_firma_asset = (
        "Ciao,\n"
        f"si richiede il giorno {data_inizio_output} "
        "di far firmare gli asset ritirati dal 599 "
        f"per la consegna a {cn}, "
        f"Referente {manager_output}."
    )

    # ========================================================
    # GUIDA FLUSSO
    # Contiene esclusivamente 01, 02, 03 e 09
    # ========================================================
    guida_flusso_lines = [
        subject_assegnazione_asset,
        message_assegnazione_asset,
        subject_provisioning,
        message_provisioning,
        subject_provisioning_pc,
        message_provisioning_pc,
        subject_firma_asset,
        message_firma_asset
    ]

    guida_flusso = "\n\n".join(
        guida_flusso_lines
    )

    filename_guida = (
        f"Guida_Flusso_"
        f"{persona_nome_underscore}.txt"
    )

    # ========================================================
    # CREAZIONE BUNDLE ZIP
    # ========================================================
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED
    ) as zip_file:

        zip_file.writestr(
            filename_utente,
            csv_utente
        )

        zip_file.writestr(
            filename_casella,
            csv_casella
        )

        zip_file.writestr(
            filename_gruppi_ad,
            csv_gruppi_ad
        )

        zip_file.writestr(
            filename_entraid,
            csv_entraid
        )

        # Il computer viene incluso nello ZIP,
        # ma non compare nel testo operativo del punto 02
        zip_file.writestr(
            filename_computer,
            csv_computer
        )

        zip_file.writestr(
            filename_guida,
            guida_flusso
        )

    zip_buffer.seek(0)

    bundle_name = (
        f"Provisioning_Consulente_"
        f"{persona_nome_underscore}.zip"
    )

    # ========================================================
    # VISUALIZZAZIONE FLUSSO
    # ========================================================

    # --------------------------------------------------------
    # 01 - Assegnazione Asset
    # --------------------------------------------------------
    st.subheader(
        subject_assegnazione_asset
    )

    st.text(
        message_assegnazione_asset
    )

    st.divider()

    # --------------------------------------------------------
    # 02 - Provisioning Utente
    # --------------------------------------------------------
    st.subheader(
        subject_provisioning
    )

    st.text(
        message_provisioning
    )

    st.markdown(
        "**File collegati al punto 02:**"
    )

    # File 1
    st.markdown(
        f"**1. {filename_utente}**"
    )

    st.dataframe(
        pd.DataFrame(
            [row_utente],
            columns=HEADER_UTENTE
        ),
        use_container_width=True
    )

    st.download_button(
        "Scarica 1 - Utenza",
        data=csv_utente,
        file_name=filename_utente,
        mime="text/csv",
        key="download_utente"
    )

    # File 2
    st.markdown(
        f"**2. {filename_casella}**"
    )

    st.text(
        csv_casella
    )

    st.download_button(
        "Scarica 2 - Casella",
        data=csv_casella,
        file_name=filename_casella,
        mime="text/csv",
        key="download_casella"
    )

    # File 3
    st.markdown(
        f"**3. {filename_gruppi_ad}**"
    )

    st.dataframe(
        pd.DataFrame(
            [row_gruppi_ad],
            columns=HEADER_UTENTE
        ),
        use_container_width=True
    )

    st.download_button(
        "Scarica 3 - Gruppi AD",
        data=csv_gruppi_ad,
        file_name=filename_gruppi_ad,
        mime="text/csv",
        key="download_gruppi_ad"
    )

    # File 4
    st.markdown(
        f"**4. {filename_entraid}**"
    )

    st.text(
        csv_entraid
    )

    st.download_button(
        "Scarica 4 - Gruppi Entra ID",
        data=csv_entraid,
        file_name=filename_entraid,
        mime="text/csv",
        key="download_entraid"
    )

    # File 5 incluso nel bundle, ma non nella lista operativa
    with st.expander(
        "File tecnico aggiuntivo incluso nel bundle ZIP"
    ):

        st.markdown(
            f"**{filename_computer}**"
        )

        st.info(
            "Il file computer viene incluso nel bundle ZIP, "
            "ma non viene indicato nella lista delle attività "
            "da eseguire del punto 02."
        )

        st.dataframe(
            pd.DataFrame(
                [row_computer],
                columns=HEADER_COMPUTER
            ),
            use_container_width=True
        )

        st.download_button(
            "Scarica file Computer",
            data=csv_computer,
            file_name=filename_computer,
            mime="text/csv",
            key="download_computer"
        )

    # Bundle completo
    st.markdown(
        "**Bundle completo del provisioning:**"
    )

    st.download_button(
        "📦 Scarica bundle completo ZIP",
        data=zip_buffer.getvalue(),
        file_name=bundle_name,
        mime="application/zip",
        key="download_bundle"
    )

    st.divider()

    # --------------------------------------------------------
    # 03 - Provisioning PC
    # --------------------------------------------------------
    st.subheader(
        subject_provisioning_pc
    )

    st.text(
        message_provisioning_pc
    )

    st.divider()

    # --------------------------------------------------------
    # 09 - Firma Asset
    # --------------------------------------------------------
    st.subheader(
        subject_firma_asset
    )

    st.text(
        message_firma_asset
    )

    st.success(
        f"Flusso completo generato correttamente per '{sam}'."
    )
