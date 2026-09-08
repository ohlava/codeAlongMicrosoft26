# =========================================================================
# Kopirovani obsahu SharePoint webu (site) - Test_A -> Project02
# =========================================================================
# OPRAVENA / ROBUSTNI VERZE. Klicove zmeny oproti puvodni:
#   1) Doplneny SOURCE_SITE / TARGET_SITE (byly prazdne).
#   2) RequestDigest se AUTOMATICKY OBNOVUJE (plati jen 30 min -> jinak 403).
#   3) Velke soubory (>2 MB) se nahravaji CHUNKED (StartUpload/ContinueUpload/
#      FinishUpload). Male primo pres Files/add.
#   4) Escapovani apostrofu v OData literalech ('' misto ') - jinak spadnou
#      nazvy s apostrofem.
#   5) Varovani u nazvu se znaky % a # (REST endpointy je nepodporuji).
#   6) Jednoduchy retry na 429/503 (throttling) s respektem k Retry-After.
#   7) Volitelne ensure root slozky na cili.
#
# BEZPECNOST: DRY_RUN=True je vychozi -> nic se nezapise. Nejdriv si nech
#             vypsat log, zkontroluj, a teprve pak prepni na False.
# =========================================================================

import re
import json
import time
import uuid
import requests

# ============================ KONFIGURACE ================================

SOURCE_SITE = "https://volkswagengroup.sharepoint.com/sites/Test_A"
TARGET_SITE = "https://volkswagengroup.sharepoint.com/sites/Project02"

SOURCE_COOKIE_FILE = "cookies.txt"
TARGET_COOKIE_FILE = "target_cookies.txt"

PROXY = "http://127.0.0.1:9001"          # <-- dopln svoji proxy (nebo nastav None)
proxies = {"http": PROXY, "https": PROXY} if PROXY else None

DRY_RUN = False                  # True = nic nezapisuje, jen simuluje a loguje
DELETE_TARGET_CONTENT = True    # True = pred kopirovanim smaze OBSAH odpovidajiciho seznamu na cili
PRINT_CONTENTS = True           # True = podrobne vypise obsah kazdeho seznamu na zdroji
COPY_SOURCE_FIELDS = True       # True = vytvori na cili chybejici sloupce ze zdroje (vc. taxonomy)
COPY_SITE_ASSETS = True         # True = zkopiruje SiteAssets/Style Library (obrazky, bannery)
COPY_WELCOME_PAGE = True        # True = nastavi domovskou stranku (WelcomePage) dle zdroje
CREATE_MISSING_LISTS = True     # True = vytvori na cili seznamy/knihovny, ktere existuji jen na zdroji
COPY_VIEWS = True               # True = prevezme zobrazeni (views) seznamu ze zdroje
COPY_NAVIGATION = True          # True = prevezme navigaci (Quick Launch + horni menu) ze zdroje

SMALL_FILE_LIMIT = 2 * 1024 * 1024
CHUNK_SIZE       = 8 * 1024 * 1024
DIGEST_TTL = 1500

# --- KSU trida na slozkach ---
SET_KSU_CLASS = True
KSU_FIELD_VALUE = "5.3"
KSU_TERM_GUID  = "f180d7d0-51f7-4ecb-b85b-8794451fa5fb"   # term 5.3
KSU_TERM_LABEL = "5.3"
KSU_FIELD_TITLE_CANDIDATES = (
    "CSD class", "CSD Class", "CSDclass", "CSD",
    "Trida KSU", "Třída KSU", "KSU Klasse", "KSU-Klasse", "KSU Class", "KSU",
)

SYSTEM_LIBRARY_PATH_SUFFIXES = (
    "/SiteAssets", "/Style Library", "/FormServerTemplates", "/_catalogs", "/_private",
)
SITE_PAGES_PATH_SUFFIX = "/SitePages"
SYSTEM_LIST_TITLES = {
    "User Information List", "Access Requests", "Workflow History",
    "Workflow Tasks", "TaxonomyHiddenList", "Cache Profiles",
    "Long Running Operation Status", "Maintenance Log Library",
}

LOG_FILE = "copy_log.txt"

# ============================ POMOCNE FUNKCE ===============================

_log_lines = []

def log(msg):
    print(msg)
    _log_lines.append(str(msg))

def save_log():
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(_log_lines))

def odata(s):
    return str(s).replace("'", "''")

def has_unsupported_chars(name):
    return "%" in name or "#" in name


def parse_cookie_line(line, cookies_dict):
    line = line.strip().rstrip(",")
    if not line or line.startswith("#"):
        return
    if "; " in line and line.count("=") > 1 and '"' not in line[:2]:
        for part in line.split(";"):
            part = part.strip()
            if "=" in part:
                k, v = part.split("=", 1)
                cookies_dict[k.strip()] = v.strip()
        return
    m = re.match(r'^"([^"]+)"\s*:\s*"(.*)"$', line)
    if m:
        cookies_dict[m.group(1).strip()] = m.group(2).strip()
        return
    if "=" in line:
        k, v = line.split("=", 1)
        cookies_dict[k.strip().strip('"')] = v.strip().strip('"')


def load_cookies(path):
    cookies = {}
    with open(path, encoding="utf-8") as f:
        for raw_line in f:
            parse_cookie_line(raw_line, cookies)
    if "FedAuth" not in cookies or "rtFa" not in cookies:
        raise SystemExit(f"!!! V souboru '{path}' chybi FedAuth nebo rtFa.")
    return cookies


class SPSession:
    def __init__(self, site_url, cookie_file):
        self.site = site_url.rstrip("/")
        self.cookies = load_cookies(cookie_file)
        self.session = requests.Session()
        if proxies:
            self.session.proxies.update(proxies)
        self.session.cookies.update(self.cookies)
        self._digest = None
        self._digest_ts = 0

    def base_headers(self, extra=None):
        h = {
            "Accept": "application/json;odata=verbose",
            "User-Agent": "Mozilla/5.0",
            "Referer": self.site,
            "Origin": "https://volkswagengroup.sharepoint.com",
            "X-FORMS_BASED_AUTH_ACCEPTED": "f",
        }
        if extra:
            h.update(extra)
        return h

    def _request(self, method, endpoint, headers=None, **kwargs):
        url = self.site + endpoint
        for attempt in range(4):
            resp = self.session.request(method, url, headers=self.base_headers(headers), **kwargs)
            if resp.status_code in (429, 503):
                wait = int(resp.headers.get("Retry-After", 5 * (attempt + 1)))
                log(f"      (throttling {resp.status_code}, cekam {wait}s...)")
                time.sleep(wait)
                continue
            return resp
        return resp

    def get(self, endpoint, **kwargs):
        return self._request("GET", endpoint, headers=kwargs.pop("headers", None), **kwargs)

    def post(self, endpoint, extra_headers=None, **kwargs):
        return self._request("POST", endpoint, headers=extra_headers, **kwargs)

    def digest(self):
        now = time.time()
        if self._digest and (now - self._digest_ts) < DIGEST_TTL:
            return self._digest
        r = self._request("POST", "/_api/contextinfo")
        r.raise_for_status()
        self._digest = r.json()["d"]["GetContextWebInformation"]["FormDigestValue"]
        self._digest_ts = now
        return self._digest

    def write_headers(self, method_override=None, extra=None):
        h = {"X-RequestDigest": self.digest()}
        if method_override:
            h["X-HTTP-Method"] = method_override
        if extra:
            h.update(extra)
        return h


_entity_type_cache = {}

def get_entity_type_full_name(sp, list_title):
    key = (sp.site, list_title)
    if key in _entity_type_cache:
        return _entity_type_cache[key]
    r = sp.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')?$select=ListItemEntityTypeFullName")
    et = r.json()["d"]["ListItemEntityTypeFullName"] if r.status_code == 200 else "SP.Data.ListItem"
    _entity_type_cache[key] = et
    return et


# ============================ SEZNAM SEZNAMU =======================

def is_system_library(lst):
    if lst["BaseType"] != 1:
        return False
    root = lst["RootFolder"]["ServerRelativeUrl"]
    return any(root.rstrip("/").endswith(suf) or (suf + "/") in root
               for suf in SYSTEM_LIBRARY_PATH_SUFFIXES)

def is_site_pages_library(lst):
    root = lst["RootFolder"]["ServerRelativeUrl"]
    return root.rstrip("/").endswith(SITE_PAGES_PATH_SUFFIX.strip("/"))

def get_lists(sp):
    r = sp.get("/_api/web/lists"
               "?$select=Title,BaseTemplate,BaseType,Hidden,ItemCount,Description,"
               "EnableFolderCreation,ContentTypesEnabled,RootFolder/ServerRelativeUrl"
               "&$expand=RootFolder")
    r.raise_for_status()
    results = []
    for lst in r.json()["d"]["results"]:
        if lst["Hidden"]:
            continue
        if lst["BaseType"] != 1 and lst["Title"] in SYSTEM_LIST_TITLES:
            continue
        if is_system_library(lst):
            continue
        results.append(lst)
    return results

def get_fields(sp, list_title):
    """Sloupce pro TVORBU schematu (bez base-type poli jako Title)."""
    r = sp.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
               "?$select=Id,Title,InternalName,TypeAsString,ReadOnlyField,Hidden,FromBaseType")
    r.raise_for_status()
    fields = r.json()["d"]["results"]
    return [f for f in fields
            if not f["ReadOnlyField"] and not f["Hidden"]
            and f["InternalName"] not in ("ContentType", "Attachments")
            and not f["FromBaseType"]]


# systemova/technicka pole, ktera se pri kopirovani DAT nikdy neprepisuji
_DATA_SYSTEM_FIELDS = {
    "ContentType", "Attachments", "Author", "Editor", "Created", "Modified",
    "ID", "GUID", "FileLeafRef", "FileRef", "FileDirRef", "Order", "owshiddenversion",
    "_UIVersionString", "_ModerationStatus", "_Level", "AppAuthor", "AppEditor",
    "ComplianceAssetId", "_ComplianceFlags", "_ComplianceTag",
}

def get_data_fields(sp, list_title):
    """Sloupce pro KOPIROVANI DAT - VCETNE base-type poli (napr. Title!),
    ale bez read-only a systemovych technicky poli."""
    r = sp.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
               "?$select=Title,InternalName,TypeAsString,ReadOnlyField,Hidden")
    r.raise_for_status()
    out = []
    for f in r.json()["d"]["results"]:
        if f["ReadOnlyField"] or f["Hidden"]:
            continue
        if f["InternalName"] in _DATA_SYSTEM_FIELDS:
            continue
        out.append(f)
    return out


# ============================ VYTVORENI CHYBEJICICH SEZNAMU ================

def list_exists(tgt, list_title):
    r = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')?$select=Title")
    return r.status_code == 200

def ensure_list_exists(src, tgt, lst):
    """Vytvori seznam/knihovnu na cili dle zdroje, pokud tam neexistuje."""
    if not CREATE_MISSING_LISTS:
        return
    title = lst["Title"]
    if list_exists(tgt, title):
        return
    log(f"  [SEZNAM] '{title}' na cili neexistuje -> vytvarim "
        f"(BaseTemplate={lst['BaseTemplate']})")
    if DRY_RUN:
        return
    body = {
        "__metadata": {"type": "SP.List"},
        "Title": title,
        "BaseTemplate": lst["BaseTemplate"],
        "Description": lst.get("Description", "") or "",
        "ContentTypesEnabled": bool(lst.get("ContentTypesEnabled")),
        "AllowContentTypes": bool(lst.get("ContentTypesEnabled")),
    }
    resp = tgt.post("/_api/web/lists",
                    extra_headers=tgt.write_headers(extra={"Content-Type": "application/json;odata=verbose"}),
                    data=json.dumps(body))
    if resp.status_code not in (200, 201):
        log(f"    !!! chyba vytvareni seznamu: {resp.status_code} {resp.text[:250]}")
        return
    # zapni tvorbu slozek u knihoven, pokud ji ma zdroj
    if lst["BaseType"] == 1 and lst.get("EnableFolderCreation"):
        tgt.post(f"/_api/web/lists/getbytitle('{odata(title)}')",
                 extra_headers=tgt.write_headers(method_override="MERGE",
                                                 extra={"IF-MATCH": "*", "Content-Type": "application/json;odata=verbose"}),
                 data=json.dumps({"__metadata": {"type": "SP.List"}, "EnableFolderCreation": True}))


# ============================ PREVZETI SLOUPCU (SCHEMA) =====================

FIELD_OPTIONS      = 12
NOTE_FIELD_OPTIONS = 12

def _all_source_fields_by_id(src, list_title):
    r = src.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
                "?$select=Id,Title,InternalName,TypeAsString,SchemaXml,Hidden,ReadOnlyField,FromBaseType")
    r.raise_for_status()
    return {str(f["Id"]).lower(): f for f in r.json()["d"]["results"]}

def _field_exists(tgt, list_title, internal):
    r = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
                f"/getbyinternalnameortitle('{odata(internal)}')?$select=InternalName")
    return r.status_code == 200

def _create_field_from_xml(tgt, list_title, schema_xml, options):
    body = {"parameters": {"__metadata": {"type": "SP.XmlSchemaFieldCreationInformation"},
                           "SchemaXml": schema_xml, "Options": options}}
    return tgt.post(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields/CreateFieldAsXml",
                    extra_headers=tgt.write_headers(extra={"Content-Type": "application/json;odata=verbose"}),
                    data=json.dumps(body))

def _get_textfield_guid(schema_xml):
    m = re.search(r"<Name>TextField</Name>\s*<Value[^>]*>\{?([0-9a-fA-F\-]{36})\}?</Value>", schema_xml)
    return m.group(1).lower() if m else None

def copy_list_fields(src, tgt, list_title):
    if not COPY_SOURCE_FIELDS:
        return
    src_fields = get_fields(src, list_title)
    if not src_fields:
        return
    all_by_id = _all_source_fields_by_id(src, list_title)
    log(f"  Prevzeti sloupcu ze zdroje ({len(src_fields)} kandidatu):")
    for f in src_fields:
        internal, ftype, title = f["InternalName"], f["TypeAsString"], f["Title"]
        if _field_exists(tgt, list_title, internal):
            log(f"    [SLOUPEC] '{title}' ({internal}, {ftype}) - uz existuje, preskakuji")
            continue
        src_full = None
        if f.get("Id"):
            src_full = all_by_id.get(str(f["Id"]).lower())
        if not src_full:
            for ff in all_by_id.values():
                if ff.get("InternalName") == internal:
                    src_full = ff
                    break
        schema = src_full["SchemaXml"] if src_full else None
        if not schema:
            log(f"    [SLOUPEC] '{title}' - nelze precist SchemaXml, preskakuji")
            continue
        is_tax = ftype in ("TaxonomyFieldType", "TaxonomyFieldTypeMulti")
        log(f"    [SLOUPEC] vytvarim '{title}' ({internal}, {ftype})" + (" [taxonomy]" if is_tax else ""))
        if DRY_RUN:
            continue
        if is_tax:
            note_guid = _get_textfield_guid(schema)
            note_field = all_by_id.get(note_guid) if note_guid else None
            if note_field and note_field.get("SchemaXml"):
                if not _field_exists(tgt, list_title, note_field["InternalName"]):
                    rn = _create_field_from_xml(tgt, list_title, note_field["SchemaXml"], NOTE_FIELD_OPTIONS)
                    if rn.status_code not in (200, 201):
                        log(f"        !!! chyba Note sloupce: {rn.status_code} {rn.text[:200]}")
            else:
                log(f"        (!) skryty Note sloupec pro '{title}' nenalezen")
            rt = _create_field_from_xml(tgt, list_title, schema, FIELD_OPTIONS)
            if rt.status_code not in (200, 201):
                log(f"        !!! chyba taxonomy sloupce: {rt.status_code} {rt.text[:250]}")
        else:
            r = _create_field_from_xml(tgt, list_title, schema, FIELD_OPTIONS)
            if r.status_code not in (200, 201):
                log(f"        !!! chyba sloupce: {r.status_code} {r.text[:250]}")


# ============================ PREVZETI VIEWS (ZOBRAZENI) ====================

def _get_view_fields(sp, list_title, vtitle):
    """Natahne SEZNAM sloupcu (internal names) zobrazenych ve view.
    ViewFields NELZE ziskat pres $select - je nutny dedikovany endpoint."""
    r = sp.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/views"
               f"/getbytitle('{odata(vtitle)}')/viewfields")
    if r.status_code != 200:
        return []
    d = r.json()["d"]
    # struktura: d.Items.results = ["Title","Ano_x002f_ne",...]
    return d.get("Items", {}).get("results", []) or []


def copy_views(src, tgt, list_title):
    """Prevezme zobrazeni (views) seznamu ze zdroje vc. zobrazenych sloupcu (ViewFields)."""
    if not COPY_VIEWS:
        return
    r = src.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/views"
                "?$select=Title,ViewQuery,RowLimit,DefaultView,Paged,Hidden,PersonalView")
    if r.status_code != 200:
        return
    src_views = [v for v in r.json()["d"]["results"] if not v.get("Hidden") and not v.get("PersonalView")]
    if not src_views:
        return
    rt = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/views?$select=Title")
    existing = set()
    if rt.status_code == 200:
        existing = {v["Title"] for v in rt.json()["d"]["results"]}
    log(f"  Prevzeti views ({len(src_views)}):")
    for v in src_views:
        vtitle = v["Title"]
        # zdrojove zobrazene sloupce (natazene zvlast)
        src_vf = _get_view_fields(src, list_title, vtitle)

        if vtitle in existing:
            log(f"    [VIEW] '{vtitle}' - aktualizuji dotaz + sloupce ({len(src_vf)})")
            if not DRY_RUN:
                _update_view(tgt, list_title, vtitle, v)
                if src_vf:
                    _set_view_fields(tgt, list_title, vtitle, src_vf)
            continue

        log(f"    [VIEW] vytvarim '{vtitle}'" + (" (vychozi)" if v.get("DefaultView") else "")
            + f" - sloupce: {len(src_vf)}")
        if DRY_RUN:
            continue
        body = {"__metadata": {"type": "SP.View"},
                "Title": vtitle,
                "ViewQuery": v.get("ViewQuery", "") or "",
                "RowLimit": v.get("RowLimit", 30),
                "Paged": bool(v.get("Paged", True)),
                "PersonalView": False}
        resp = tgt.post(f"/_api/web/lists/getbytitle('{odata(list_title)}')/views",
                        extra_headers=tgt.write_headers(extra={"Content-Type": "application/json;odata=verbose"}),
                        data=json.dumps(body))
        if resp.status_code not in (200, 201):
            log(f"        !!! chyba vytvareni view: {resp.status_code} {resp.text[:200]}")
            continue
        if src_vf:
            _set_view_fields(tgt, list_title, vtitle, src_vf)

def _update_view(tgt, list_title, vtitle, v):
    body = {"__metadata": {"type": "SP.View"},
            "ViewQuery": v.get("ViewQuery", "") or "",
            "RowLimit": v.get("RowLimit", 30)}
    tgt.post(f"/_api/web/lists/getbytitle('{odata(list_title)}')/views/getbytitle('{odata(vtitle)}')",
             extra_headers=tgt.write_headers(method_override="MERGE",
                                             extra={"IF-MATCH": "*", "Content-Type": "application/json;odata=verbose"}),
             data=json.dumps(body))

def _set_view_fields(tgt, list_title, vtitle, fields):
    """Nastavi zobrazene sloupce view: nejdriv smaze vse, pak prida ze zdroje."""
    base = (f"/_api/web/lists/getbytitle('{odata(list_title)}')/views"
            f"/getbytitle('{odata(vtitle)}')/viewfields")
    tgt.post(base + "/removeallviewfields", extra_headers=tgt.write_headers())
    for fn in fields:
        r = tgt.post(base + f"/addviewfield('{odata(fn)}')", extra_headers=tgt.write_headers())
        if r.status_code not in (200, 204):
            log(f"        (!) sloupec '{fn}' nelze pridat do view: {r.status_code}")


# ============================ NAVIGACE (Quick Launch + horni menu) ==========

def _remap_url(u):
    """Premapuje server-relative i absolutni URL ze zdroje na cil."""
    if not u:
        return u
    sp = SOURCE_SITE.split(".com", 1)[1]     # /sites/Test_A
    tp = TARGET_SITE.split(".com", 1)[1]     # /sites/Project02
    return u.replace(SOURCE_SITE, TARGET_SITE).replace(sp, tp)


def _is_system_nav_url(u):
    """Systemove odkazy (Site contents, Recent, ...) - cil je ma vlastni, nekopirovat."""
    if not u:
        return False
    low = u.lower()
    return "/_layouts/" in low or "viewlsts.aspx" in low

def _get_nav_nodes(sp, which):
    """which = 'quicklaunch' | 'topnavigationbar'. Vrati stromovou strukturu."""
    r = sp.get(f"/_api/web/navigation/{which}?$expand=Children")
    if r.status_code != 200:
        return []
    out = []
    for n in r.json()["d"]["results"]:
        kids = [{"Title": c["Title"], "Url": c["Url"], "IsExternal": c.get("IsExternal", False)}
                for c in (n.get("Children", {}).get("results", []))]
        out.append({"Title": n["Title"], "Url": n["Url"],
                    "IsExternal": n.get("IsExternal", False), "Children": kids})
    return out

def _clear_nav(tgt, which):
    r = tgt.get(f"/_api/web/navigation/{which}?$select=Id")
    if r.status_code != 200:
        return
    for n in r.json()["d"]["results"]:
        tgt.post(f"/_api/web/navigation/{which}/getbyid({n['Id']})",
                 extra_headers=tgt.write_headers(method_override="DELETE", extra={"IF-MATCH": "*"}))

def _add_nav_node(tgt, which, title, url, is_external, parent_id=None):
    # systemove odkazy (_layouts, Site contents) preskoc - cil je ma vlastni
    if _is_system_nav_url(url):
        log(f"        (i) preskakuji systemovy nav uzel '{title}' ({url})")
        return None
    if parent_id is not None:
        endpoint = f"/_api/web/navigation/getnodebyid({parent_id})/children"
    else:
        endpoint = f"/_api/web/navigation/{which}"
    remapped = _remap_url(url)
    # pokud odkaz stale miri na zdrojovy web, oznac jako externi (aby SP neveril, ze je lokalni)
    still_external = bool(is_external) or ("/sites/" in (url or "") and SOURCE_SITE.split(".com",1)[1] in (url or "") and remapped == url)
    body = {"__metadata": {"type": "SP.NavigationNode"},
            "Title": title, "Url": remapped, "IsExternal": still_external}
    try:
        r = tgt.post(endpoint,
                     extra_headers=tgt.write_headers(extra={"Content-Type": "application/json;odata=verbose"}),
                     data=json.dumps(body))
        if r.status_code in (200, 201):
            return r.json()["d"]["Id"]
        log(f"        !!! nav uzel '{title}' preskocen: {r.status_code} {r.text[:120]}")
    except Exception as e:
        log(f"        !!! nav uzel '{title}' vyjimka: {e}")
    return None

def copy_navigation(src, tgt):
    if not COPY_NAVIGATION:
        return
    log(f"\n{'='*70}\nNAVIGACE (Quick Launch + horni menu)\n{'='*70}")
    for which, label in (("quicklaunch", "Quick Launch"), ("topnavigationbar", "Horni menu")):
        nodes = _get_nav_nodes(src, which)
        log(f"  [{label}] nalezeno {len(nodes)} uzlu na zdroji")
        if DRY_RUN:
            for n in nodes:
                log(f"    - {n['Title']} ({n['Url']})")
                for c in n["Children"]:
                    log(f"        - {c['Title']} ({c['Url']})")
            continue
        _clear_nav(tgt, which)
        for n in nodes:
            pid = _add_nav_node(tgt, which, n["Title"], n["Url"], n["IsExternal"])
            for c in n["Children"]:
                _add_nav_node(tgt, which, c["Title"], c["Url"], c["IsExternal"], parent_id=pid)


# ============================ VYPIS OBSAHU ==========

def _print_folder_tree(sp, folder_url, indent=""):
    r = sp.get(f"/_api/web/GetFolderByServerRelativeUrl('{odata(folder_url)}')?$expand=Folders,Files")
    if r.status_code != 200:
        log(f"{indent}!!! nelze nacist slozku '{folder_url}': {r.status_code}")
        return
    data = r.json()["d"]
    for f in data.get("Files", {}).get("results", []):
        log(f"{indent}[SOUBOR] {f['Name']}")
    for sub in data.get("Folders", {}).get("results", []):
        name = sub["Name"]
        if name == "Forms":
            continue
        log(f"{indent}[SLOZKA] {name}/")
        _print_folder_tree(sp, f"{folder_url}/{name}", indent=indent + "    ")

def print_list_contents(sp, lst):
    title = lst["Title"]
    if lst["BaseType"] == 1:
        root = lst["RootFolder"]["ServerRelativeUrl"]
        label = "stranek" if is_site_pages_library(lst) else "knihovny"
        log(f"  Obsah {label} '{title}' (strom):")
        _print_folder_tree(sp, root, indent="    ")
    else:
        r = sp.get(f"/_api/web/lists/getbytitle('{odata(title)}')/items?$select=Id,Title&$top=5000")
        if r.status_code != 200:
            log(f"    !!! nelze nacist polozky '{title}': {r.status_code}")
            return
        items = r.json()["d"]["results"]
        log(f"  Obsah seznamu '{title}' ({len(items)} polozek):")
        for item in items:
            log(f"    [POLOZKA] #{item['Id']:<5} {item.get('Title') or '(bez nazvu)'}")


# ============================ MAZANI OBSAHU =========================

def clear_document_library(tgt, root_folder_url):
    r = tgt.get(f"/_api/web/GetFolderByServerRelativeUrl('{odata(root_folder_url)}')?$expand=Folders,Files")
    if r.status_code != 200:
        log(f"    (cilova slozka neexistuje/chyba: {r.status_code})")
        return
    data = r.json()["d"]
    for f in data.get("Files", {}).get("results", []):
        furl = f["ServerRelativeUrl"]
        log(f"    [MAZAT SOUBOR] {furl}")
        if not DRY_RUN:
            resp = tgt.post(f"/_api/web/GetFileByServerRelativeUrl('{odata(furl)}')",
                            extra_headers=tgt.write_headers(method_override="DELETE", extra={"IF-MATCH": "*"}))
            if resp.status_code not in (200, 204):
                log(f"      !!! chyba mazani souboru: {resp.status_code}")
    for sub in data.get("Folders", {}).get("results", []):
        surl = sub["ServerRelativeUrl"]
        if surl.rstrip("/").endswith("/Forms"):
            continue
        log(f"    [MAZAT SLOZKU] {surl}")
        if not DRY_RUN:
            resp = tgt.post(f"/_api/web/GetFolderByServerRelativeUrl('{odata(surl)}')",
                            extra_headers=tgt.write_headers(method_override="DELETE", extra={"IF-MATCH": "*"}))
            if resp.status_code not in (200, 204):
                log(f"      !!! chyba mazani slozky: {resp.status_code}")

def clear_generic_list(tgt, list_title):
    r = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/items?$select=Id&$top=5000")
    if r.status_code != 200:
        log(f"    (cilovy seznam '{list_title}' neexistuje/chyba: {r.status_code})")
        return
    for item in r.json()["d"]["results"]:
        iid = item["Id"]
        log(f"    [MAZAT POLOZKU] {list_title} #{iid}")
        if not DRY_RUN:
            resp = tgt.post(f"/_api/web/lists/getbytitle('{odata(list_title)}')/items({iid})",
                            extra_headers=tgt.write_headers(method_override="DELETE", extra={"IF-MATCH": "*"}))
            if resp.status_code not in (200, 204):
                log(f"      !!! chyba mazani polozky: {resp.status_code}")


# ============================ KOPIROVANI SOUBORU =========

def ensure_target_folder(tgt, folder_server_relative_url):
    log(f"    [SLOZKA] {folder_server_relative_url}")
    if DRY_RUN:
        return
    resp = tgt.post("/_api/web/folders",
                    extra_headers=tgt.write_headers(extra={"Content-Type": "application/json;odata=verbose"}),
                    data=json.dumps({"__metadata": {"type": "SP.Folder"},
                                     "ServerRelativeUrl": folder_server_relative_url}))
    if resp.status_code not in (200, 201):
        if "already exists" not in resp.text and resp.status_code not in (500,):
            log(f"      !!! chyba vytvareni slozky: {resp.status_code} {resp.text[:200]}")

def _upload_small(tgt, tgt_folder_url, file_name, content):
    endpoint = (f"/_api/web/GetFolderByServerRelativeUrl('{odata(tgt_folder_url)}')"
                f"/Files/add(url='{odata(file_name)}',overwrite=true)")
    resp = tgt.post(endpoint, extra_headers=tgt.write_headers(), data=content)
    if resp.status_code not in (200, 201):
        log(f"      !!! chyba nahrani (small): {resp.status_code} {resp.text[:200]}")

def _upload_large(tgt, tgt_folder_url, file_name, content):
    create = tgt.post(f"/_api/web/GetFolderByServerRelativeUrl('{odata(tgt_folder_url)}')"
                      f"/Files/add(url='{odata(file_name)}',overwrite=true)",
                      extra_headers=tgt.write_headers(), data=b"")
    if create.status_code not in (200, 201):
        log(f"      !!! chyba prazdneho souboru: {create.status_code}")
        return
    fu = odata(f"{tgt_folder_url}/{file_name}")
    upload_id = str(uuid.uuid4())
    total = len(content)
    first = content[:CHUNK_SIZE]
    r = tgt.post(f"/_api/web/GetFileByServerRelativeUrl('{fu}')/StartUpload(uploadId=guid'{upload_id}')",
                 extra_headers=tgt.write_headers(), data=first)
    if r.status_code not in (200, 201):
        log(f"      !!! StartUpload: {r.status_code}")
        return
    offset = len(first)
    while offset < total:
        chunk = content[offset:offset + CHUNK_SIZE]
        op = "FinishUpload" if (offset + len(chunk)) >= total else "ContinueUpload"
        r = tgt.post(f"/_api/web/GetFileByServerRelativeUrl('{fu}')/{op}(uploadId=guid'{upload_id}',fileOffset={offset})",
                     extra_headers=tgt.write_headers(), data=chunk)
        if r.status_code not in (200, 201):
            log(f"      !!! {op}: {r.status_code}")
            return
        offset += len(chunk)
    if total <= CHUNK_SIZE:
        tgt.post(f"/_api/web/GetFileByServerRelativeUrl('{fu}')/FinishUpload(uploadId=guid'{upload_id}',fileOffset={offset})",
                 extra_headers=tgt.write_headers(), data=b"")

def copy_file(src, tgt, src_file_url, tgt_folder_url, file_name):
    log(f"    [SOUBOR] {file_name}  ->  {tgt_folder_url}")
    if has_unsupported_chars(file_name):
        log(f"      (!) PRESKAKUJI - nazev obsahuje '%' nebo '#'.")
        return
    if DRY_RUN:
        return
    r = src.get(f"/_api/web/GetFileByServerRelativeUrl('{odata(src_file_url)}')/$value")
    if r.status_code != 200:
        log(f"      !!! chyba stazeni: {r.status_code}")
        return
    content = r.content
    if len(content) <= SMALL_FILE_LIMIT:
        _upload_small(tgt, tgt_folder_url, file_name, content)
    else:
        log(f"      (velky soubor {len(content)//1024//1024} MB -> chunked)")
        _upload_large(tgt, tgt_folder_url, file_name, content)

def copy_library_recursive(src, tgt, src_folder_url, tgt_folder_url, list_title=None):
    r = src.get(f"/_api/web/GetFolderByServerRelativeUrl('{odata(src_folder_url)}')?$expand=Folders,Files")
    r.raise_for_status()
    data = r.json()["d"]
    for f in data.get("Files", {}).get("results", []):
        copy_file(src, tgt, f["ServerRelativeUrl"], tgt_folder_url, f["Name"])
    for sub in data.get("Folders", {}).get("results", []):
        name = sub["Name"]
        if name == "Forms":
            continue
        new_src = f"{src_folder_url}/{name}"
        new_tgt = f"{tgt_folder_url}/{name}"
        ensure_target_folder(tgt, new_tgt)
        if list_title:
            set_folder_ksu(tgt, list_title, new_tgt)
        copy_library_recursive(src, tgt, new_src, new_tgt, list_title)

def target_root_for(src_root_url):
    sp = SOURCE_SITE.split(".com", 1)[1]
    tp = TARGET_SITE.split(".com", 1)[1]
    return src_root_url.replace(sp, tp)


# ============================ ASSET KNIHOVNY ====

def copy_asset_library(src, tgt, path_suffix):
    r = src.get("/_api/web/lists?$select=Title,BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder")
    if r.status_code != 200:
        log(f"  (!) nelze nacist knihovny pro '{path_suffix}': {r.status_code}")
        return
    src_root = None
    for lst in r.json()["d"]["results"]:
        root = lst["RootFolder"]["ServerRelativeUrl"]
        if root.rstrip("/").endswith(path_suffix.strip("/")):
            src_root = root
            break
    if not src_root:
        log(f"  (i) asset '{path_suffix}' na zdroji neni - preskakuji")
        return
    tgt_root = target_root_for(src_root)
    log(f"\n{'='*70}\nASSET KNIHOVNA: {path_suffix}\n{'='*70}")
    rc = src.get(f"/_api/web/GetFolderByServerRelativeUrl('{odata(src_root)}')?$expand=Folders,Files")
    if rc.status_code != 200:
        log(f"  (!) nelze nacist '{src_root}': {rc.status_code}")
        return
    if DELETE_TARGET_CONTENT:
        log(f"  Mazani obsahu cilove asset knihovny: {tgt_root}")
        clear_document_library(tgt, tgt_root)
    log(f"  Kopirovani assetu: {src_root} -> {tgt_root}")
    copy_library_recursive(src, tgt, src_root, tgt_root, None)


# ============================ MODERNI STRANKY =====

PAGE_FIELDS_TO_COPY = ["Title", "CanvasContent1", "LayoutWebpartsContent",
                       "Description", "PromotedState", "PageLayoutType",
                       "ClientSideApplicationId", "_TopicHeader", "_SPSitePageFlags"]

# GUID aplikace, ktera oznacuje stranku jako MODERNI (client-side). Bez nej
# SharePoint stranku povazuje za klasickou a renderuje ji rozbite.
MODERN_PAGE_APP_ID = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec"

def copy_site_pages_library(src, tgt, lst):
    title = lst["Title"]
    src_root = lst["RootFolder"]["ServerRelativeUrl"]
    tgt_root = target_root_for(src_root)
    if DELETE_TARGET_CONTENT:
        log(f"  Mazani obsahu cilove knihovny stranek: {tgt_root}")
        clear_document_library(tgt, tgt_root)
    r = src.get(f"/_api/web/GetFolderByServerRelativeUrl('{odata(src_root)}')/Files?$select=Name,ServerRelativeUrl")
    r.raise_for_status()
    files = [f for f in r.json()["d"]["results"] if f["Name"].lower().endswith(".aspx")]
    log(f"  Nalezeno {len(files)} stranek ke zkopirovani")
    entity_type = get_entity_type_full_name(src, title)
    for f in files:
        name = f["Name"]
        log(f"    [STRANKA] {name}")
        r_item = src.get(f"/_api/web/GetFileByServerRelativeUrl('{odata(f['ServerRelativeUrl'])}')"
                         "/ListItemAllFields?$select=" + ",".join(PAGE_FIELDS_TO_COPY + ["BannerImageUrl"]))
        if r_item.status_code != 200:
            log(f"      !!! nelze nacist obsah stranky: {r_item.status_code}")
            continue
        item = r_item.json()["d"]
        if DRY_RUN:
            continue
        tgt_file_url = f"{tgt_root}/{name}"
        create_resp = tgt.post(f"/_api/web/GetFolderByServerRelativeUrl('{odata(tgt_root)}')"
                               f"/Files/AddTemplateFile(urlOfFile='{odata(tgt_file_url)}',templateFileType=3)",
                               extra_headers=tgt.write_headers())
        if create_resp.status_code not in (200, 201) and "already exists" not in create_resp.text:
            log(f"      !!! chyba vytvareni stranky: {create_resp.status_code} {create_resp.text[:250]}")
            continue
        r_new = tgt.get(f"/_api/web/GetFileByServerRelativeUrl('{odata(tgt_file_url)}')/ListItemAllFields?$select=Id")
        if r_new.status_code != 200:
            log(f"      !!! nelze najit novou stranku: {r_new.status_code}")
            continue
        new_id = r_new.json()["d"]["Id"]
        values = {"__metadata": {"type": entity_type}}
        for field in PAGE_FIELDS_TO_COPY:
            if item.get(field) is not None:
                val = item[field]
                # v obsahu stranky (canvas/layout) premapuj odkazy /sites/Test_A -> /sites/Project02
                if field in ("CanvasContent1", "LayoutWebpartsContent") and isinstance(val, str):
                    val = _remap_url(val)
                values[field] = val
        # KLICOVE: stranka musi byt oznacena jako MODERNI (client-side), jinak se
        # renderuje rozbite. Nastav natvrdo, i kdyz zdroj vratil null.
        values["ClientSideApplicationId"] = MODERN_PAGE_APP_ID
        if not values.get("PageLayoutType"):
            # Home.aspx -> "Home", ostatni -> "Article"
            values["PageLayoutType"] = "Home" if name.lower() == "home.aspx" else "Article"
        banner = item.get("BannerImageUrl")
        if isinstance(banner, dict) and banner.get("Url"):
            values["BannerImageUrl"] = {"__metadata": {"type": "SP.FieldUrlValue"},
                                        "Url": _remap_url(banner["Url"]), "Description": banner.get("Description", "")}

        def _write_page(vals):
            return tgt.post(f"/_api/web/lists/getbytitle('{odata(title)}')/items({new_id})",
                            extra_headers=tgt.write_headers(method_override="MERGE",
                                                            extra={"IF-MATCH": "*", "Content-Type": "application/json;odata=verbose"}),
                            data=json.dumps(vals))
        upd = _write_page(values)
        if upd.status_code not in (200, 204):
            log(f"      !!! chyba zapisu stranky: {upd.status_code} {upd.text[:300]}")
        expected = values.get("CanvasContent1")
        if expected:
            for attempt in range(2):
                chk = tgt.get(f"/_api/web/lists/getbytitle('{odata(title)}')/items({new_id})?$select=CanvasContent1")
                got = chk.json()["d"].get("CanvasContent1") if chk.status_code == 200 else None
                if got == expected:
                    break
                log(f"      (i) CanvasContent1 se neulozil - opakuji ({attempt+1})")
                _write_page({"__metadata": {"type": entity_type},
                             "CanvasContent1": expected,
                             "LayoutWebpartsContent": values.get("LayoutWebpartsContent", "")})
            else:
                log(f"      !!! CanvasContent1 se nepodarilo ulozit (znamy REST limit).")
        tgt.post(f"/_api/web/GetFileByServerRelativeUrl('{odata(tgt_file_url)}')/CheckIn(comment='Kopie',checkintype=1)",
                 extra_headers=tgt.write_headers())
        tgt.post(f"/_api/web/GetFileByServerRelativeUrl('{odata(tgt_file_url)}')/Publish('Kopie')",
                 extra_headers=tgt.write_headers())
    if COPY_WELCOME_PAGE:
        set_welcome_page(src, tgt)

def set_welcome_page(src, tgt):
    r = src.get("/_api/web/rootfolder?$select=WelcomePage")
    if r.status_code != 200:
        log(f"  (!) nelze precist WelcomePage zdroje: {r.status_code}")
        return
    welcome = r.json()["d"].get("WelcomePage")
    if not welcome:
        log(f"  (i) zdroj nema explicitni WelcomePage")
        return
    log(f"  [DOMOVSKA STRANKA] WelcomePage = '{welcome}'")
    if DRY_RUN:
        return
    resp = tgt.post("/_api/web/rootfolder",
                    extra_headers=tgt.write_headers(method_override="MERGE",
                                                    extra={"IF-MATCH": "*", "Content-Type": "application/json;odata=verbose"}),
                    data=json.dumps({"__metadata": {"type": "SP.Folder"}, "WelcomePage": welcome}))
    if resp.status_code in (401, 403):
        log(f"    (i) WelcomePage nelze nastavit pres REST (chybi opravneni ManageWeb).")
        log(f"        -> Nastav rucne: Nastaveni webu > vyber '{welcome}' jako domovskou stranku,")
        log(f"           nebo v knihovne Site Pages u stranky '...->Make homepage'.")
    elif resp.status_code not in (200, 204):
        log(f"    !!! chyba WelcomePage: {resp.status_code} {resp.text[:250]}")


# ============================ GENERICKE SEZNAMY ======

SKIP_FIELD_TYPES = {"User", "UserMulti", "Lookup", "LookupMulti",
                    "TaxonomyFieldType", "TaxonomyFieldTypeMulti"}

def copy_generic_list_items(src, tgt, list_title):
    fields = get_data_fields(src, list_title)   # VCETNE Title a dalsich base-type poli
    skipped = [f["Title"] for f in fields if f["TypeAsString"] in SKIP_FIELD_TYPES]
    if skipped:
        log(f"    (!) Preskakuji nepodporovane sloupce: {skipped}")
    r = src.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/items?$top=5000")
    r.raise_for_status()
    items = r.json()["d"]["results"]
    log(f"    Nalezeno {len(items)} polozek ke kopirovani")
    entity_type = get_entity_type_full_name(tgt, list_title)   # entity type CILE (kam zapisujeme)
    for item in items:
        values = {"__metadata": {"type": entity_type}}
        for field in fields:
            iname, ftype = field["InternalName"], field["TypeAsString"]
            if ftype in SKIP_FIELD_TYPES:
                continue
            if item.get(iname) is not None:
                values[iname] = item[iname]
        log(f"    [POLOZKA] {list_title} - {values.get('Title', item.get('Id'))}")
        if not DRY_RUN:
            resp = tgt.post(f"/_api/web/lists/getbytitle('{odata(list_title)}')/items",
                            extra_headers=tgt.write_headers(extra={"Content-Type": "application/json;odata=verbose"}),
                            data=json.dumps(values))
            if resp.status_code not in (200, 201):
                log(f"      !!! chyba vytvareni polozky: {resp.status_code} {resp.text[:300]}")


# ============================ KSU ======

_ksu_fields_multi_cache = {}
_ksu_guid_cache = {}
_ksu_dumped = {}
_SRC_SESSION = None

def dump_fields(tgt, list_title):
    r = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
                "?$select=Title,InternalName,TypeAsString,Hidden")
    if r.status_code != 200:
        return
    log(f"      --- Sloupce '{list_title}' ---")
    for f in r.json()["d"]["results"]:
        if f.get("Hidden"):
            continue
        hay = (f["Title"] + f["InternalName"]).lower()
        flag = " [KSU?]" if ("ksu" in hay or "csd" in hay) else ""
        log(f"        {f['Title']:35} | {f['InternalName']:30} | {f['TypeAsString']}{flag}")

def resolve_ksu_fields(tgt, list_title):
    key = (tgt.site, list_title)
    if key in _ksu_fields_multi_cache:
        return _ksu_fields_multi_cache[key]
    found, seen = [], set()
    r = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
                "?$select=Title,InternalName,TypeAsString,Hidden")
    if r.status_code == 200:
        fields = [f for f in r.json()["d"]["results"] if not f.get("Hidden")]
        def _add(f):
            if f["InternalName"] not in seen:
                seen.add(f["InternalName"])
                found.append((f["InternalName"], f["TypeAsString"]))
        for cand in KSU_FIELD_TITLE_CANDIDATES:
            for f in fields:
                if cand.strip().lower() in (f["Title"].strip().lower(), f["InternalName"].strip().lower()):
                    _add(f)
        for f in fields:
            hay = (f["Title"] + f["InternalName"]).lower()
            if "ksu" in hay or "csd" in hay:
                _add(f)
    _ksu_fields_multi_cache[key] = found
    return found

def _label_matches(name, target):
    name, target = name.strip().lower(), target.strip().lower()
    if name == target:
        return True
    first = name.split(" ", 1)[0].split(",", 1)[0].strip()
    return first == target

def _term_labels(t):
    labels = t.get("labels") or []
    names = [str(l.get("name", "")).strip() for l in labels if l.get("name")]
    if not names and t.get("Name"):
        names = [str(t["Name"]).strip()]
    return names

def _harvest_guid_from_items(sp, list_title, internal):
    if sp is None:
        return None
    r = sp.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/items?$select=Id,{internal}&$top=500")
    if r.status_code != 200:
        return None
    for item in r.json()["d"]["results"]:
        val = item.get(internal)
        if isinstance(val, dict) and val.get("TermGuid") and _label_matches(str(val.get("Label", "")), KSU_FIELD_VALUE):
            return val["TermGuid"]
    return None

def _termstore_get(sp, endpoint):
    return sp.get(endpoint, headers={"Accept": "application/json;odata=nometadata"})

def _lookup_guid_from_termstore(tgt, list_title, internal, dump=False):
    r = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
                f"/getbyinternalnameortitle('{odata(internal)}')?$select=TermSetId,SspId,AnchorId")
    if r.status_code != 200:
        return None
    term_set_id = (r.json()["d"].get("TermSetId") or "").strip("{}")
    if not term_set_id or set(term_set_id) <= set("0-"):
        return None
    found = {"guid": None}
    def _walk(url, depth=0):
        while url:
            rr = _termstore_get(tgt, url)
            if rr.status_code != 200:
                return
            data = rr.json()
            for t in data.get("value") or []:
                tid = t.get("id") or t.get("Id")
                names = _term_labels(t)
                if dump:
                    log(f"        {'  '*depth}- {', '.join(names):20} | {tid}")
                if found["guid"] is None and any(_label_matches(n, KSU_FIELD_VALUE) for n in names):
                    found["guid"] = tid
                    if not dump:
                        return
                if t.get("childrenCount", 0) and (dump or found["guid"] is None):
                    _walk(f"/_api/v2.1/termStore/sets/{term_set_id}/terms/{tid}/children", depth + 1)
                    if found["guid"] and not dump:
                        return
            url = data.get("@odata.nextLink")
            if url and "/_api/" in url:
                url = "/_api/" + url.split("/_api/", 1)[1]
    _walk(f"/_api/v2.1/termStore/sets/{term_set_id}/terms")
    return found["guid"]

def get_ksu_term_guid(tgt, list_title, internal):
    if KSU_TERM_GUID:
        return KSU_TERM_GUID
    key = (tgt.site, list_title, internal)
    if key in _ksu_guid_cache:
        return _ksu_guid_cache[key]
    guid = (_harvest_guid_from_items(tgt, list_title, internal)
            or _harvest_guid_from_items(_SRC_SESSION, list_title, internal)
            or _lookup_guid_from_termstore(tgt, list_title, internal))
    _ksu_guid_cache[key] = guid
    return guid

def _set_taxonomy_ksu(tgt, list_title, item_id, internal, guid):
    label = KSU_TERM_LABEL or KSU_FIELD_VALUE
    body = {"formValues": [{"FieldName": internal, "FieldValue": f"{label}|{guid}"}],
            "bNewDocumentUpdate": False}
    resp = tgt.post(f"/_api/web/lists/getbytitle('{odata(list_title)}')/items({item_id})/ValidateUpdateListItem",
                    extra_headers=tgt.write_headers(extra={"Content-Type": "application/json;odata=verbose"}),
                    data=json.dumps(body))
    if resp.status_code not in (200, 201):
        log(f"        !!! ValidateUpdateListItem: {resp.status_code} {resp.text[:250]}")
        return
    try:
        for fv in resp.json()["d"]["ValidateUpdateListItem"]["results"]:
            if fv.get("HasException"):
                log(f"        !!! KSU vyjimka {fv.get('FieldName')}: {fv.get('ErrorMessage')}")
    except Exception:
        pass

def _set_plain_ksu(tgt, list_title, item_id, internal, ftype):
    value = KSU_FIELD_VALUE
    if ftype in ("Number", "Currency"):
        try:
            value = float(KSU_FIELD_VALUE)
        except ValueError:
            pass
    entity_type = get_entity_type_full_name(tgt, list_title)
    body = {"__metadata": {"type": entity_type}, internal: value}
    resp = tgt.post(f"/_api/web/lists/getbytitle('{odata(list_title)}')/items({item_id})",
                    extra_headers=tgt.write_headers(method_override="MERGE",
                                                    extra={"IF-MATCH": "*", "Content-Type": "application/json;odata=verbose"}),
                    data=json.dumps(body))
    if resp.status_code not in (200, 204):
        log(f"        !!! chyba KSU '{internal}': {resp.status_code}")

def set_folder_ksu(tgt, list_title, folder_server_relative_url):
    if not SET_KSU_CLASS:
        return
    ksu_fields = resolve_ksu_fields(tgt, list_title)
    if not ksu_fields:
        if not _ksu_dumped.get((tgt.site, list_title)):
            _ksu_dumped[(tgt.site, list_title)] = True
            log(f"      (!) Zadny KSU/CSD sloupec v '{list_title}'")
            dump_fields(tgt, list_title)
        return
    log(f"      [KSU] {folder_server_relative_url} -> {KSU_FIELD_VALUE} "
        f"(sloupce: {', '.join(i for i, _ in ksu_fields)})")
    if DRY_RUN:
        return
    r = tgt.get(f"/_api/web/GetFolderByServerRelativeUrl('{odata(folder_server_relative_url)}')/ListItemAllFields?$select=Id")
    if r.status_code != 200:
        log(f"        !!! nelze nacist ListItem slozky: {r.status_code}")
        return
    item_id = r.json()["d"]["Id"]
    for internal, ftype in ksu_fields:
        if ftype in ("TaxonomyFieldType", "TaxonomyFieldTypeMulti"):
            guid = get_ksu_term_guid(tgt, list_title, internal)
            if not guid:
                log(f"        (!) '{internal}' Taxonomy, GUID '{KSU_FIELD_VALUE}' neznamy.")
                continue
            _set_taxonomy_ksu(tgt, list_title, item_id, internal, guid)
        else:
            _set_plain_ksu(tgt, list_title, item_id, internal, ftype)


# ============================ HLAVNI BEH ====================================

def main():
    log("=" * 70)
    log(f"START   DRY_RUN={DRY_RUN}   DELETE_TARGET_CONTENT={DELETE_TARGET_CONTENT}")
    log(f"        CREATE_MISSING_LISTS={CREATE_MISSING_LISTS}  COPY_VIEWS={COPY_VIEWS}  "
        f"COPY_NAVIGATION={COPY_NAVIGATION}  COPY_SITE_ASSETS={COPY_SITE_ASSETS}")
    log(f"Zdroj: {SOURCE_SITE}")
    log(f"Cil:   {TARGET_SITE}")
    log("=" * 70)

    src = SPSession(SOURCE_SITE, SOURCE_COOKIE_FILE)
    tgt = SPSession(TARGET_SITE, TARGET_COOKIE_FILE)
    global _SRC_SESSION
    _SRC_SESSION = src

    for name, sp in (("ZDROJ", src), ("CIL", tgt)):
        r = sp.get("/_api/web?$select=Title")
        if r.status_code == 200:
            log(f"[{name}] Pripojeno: {r.json()['d']['Title']}")
        else:
            log(f"[{name}] !!! Pripojeni selhalo: {r.status_code} {r.text[:200]}")
            return

    lists = get_lists(src)
    log(f"\nNalezeno {len(lists)} seznamu/knihoven:\n")
    for lst in lists:
        kind = ("Site Pages" if is_site_pages_library(lst)
                else "Document Library" if lst["BaseType"] == 1 else "Generic List")
        log(f"  - {lst['Title']:30} (BaseType={lst['BaseType']}, ItemCount={lst['ItemCount']}) [{kind}]")

    if PRINT_CONTENTS:
        log(f"\n{'='*70}\nPODROBNY VYPIS OBSAHU (zdroj)\n{'='*70}")
        for lst in lists:
            log(f"\n--- {lst['Title']} ---")
            print_list_contents(src, lst)

    # 0) asset knihovny (obrazky/bannery) - nejdriv
    if COPY_SITE_ASSETS:
        copy_asset_library(src, tgt, "/SiteAssets")
        copy_asset_library(src, tgt, "/Style Library")

    for lst in lists:
        title = lst["Title"]
        log(f"\n{'='*70}\nZPRACOVAVAM: {title}\n{'='*70}")

        # 1) vytvor seznam na cili, pokud chybi
        ensure_list_exists(src, tgt, lst)

        # 2) prevezmi sloupce
        copy_list_fields(src, tgt, title)

        # 3) prevezmi views
        copy_views(src, tgt, title)

        # 4) obsah
        if is_site_pages_library(lst):
            copy_site_pages_library(src, tgt, lst)
        elif lst["BaseType"] == 1:
            src_root = lst["RootFolder"]["ServerRelativeUrl"]
            tgt_root = target_root_for(src_root)
            if DELETE_TARGET_CONTENT:
                log(f"  Mazani obsahu cilove knihovny: {tgt_root}")
                clear_document_library(tgt, tgt_root)
            log(f"  Kopirovani souboru: {src_root} -> {tgt_root}")
            copy_library_recursive(src, tgt, src_root, tgt_root, title)
        else:
            if DELETE_TARGET_CONTENT:
                log(f"  Mazani polozek ciloveho seznamu: {title}")
                clear_generic_list(tgt, title)
            log(f"  Kopirovani polozek seznamu: {title}")
            copy_generic_list_items(src, tgt, title)

    # 5) navigace (Quick Launch + horni menu) - na zaver
    copy_navigation(src, tgt)

    log("\n" + "=" * 70)
    log("HOTOVO" + ("  (DRY RUN - nic se nezapsalo)" if DRY_RUN else ""))
    log("=" * 70)
    save_log()


if __name__ == "__main__":
    main()
