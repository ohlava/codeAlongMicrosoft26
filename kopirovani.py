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

# Bezpecnostni prepinace
DRY_RUN = False                  # True = nic nezapisuje, jen simuluje a loguje
DELETE_TARGET_CONTENT = True    # True = pred kopirovanim smaze OBSAH odpovidajiciho seznamu na cili
PRINT_CONTENTS = True           # True = podrobne vypise obsah kazdeho seznamu na zdroji

# Prah pro chunked upload a velikost chunku
SMALL_FILE_LIMIT = 2 * 1024 * 1024     # <=2 MB -> primy Files/add
CHUNK_SIZE       = 8 * 1024 * 1024     # 8 MB na chunk pro velke soubory

# Digest se preventivne obnovi po tolika sekundach (limit je 1800 s)
DIGEST_TTL = 1500

# Interni (jazykove NEZAVISLE) cesty systemovych knihoven, ktere se nikdy nekopiruji.
SYSTEM_LIBRARY_PATH_SUFFIXES = (
    "/SiteAssets",
    "/Style Library",
    "/FormServerTemplates",
    "/_catalogs",
    "/_private",
)

SITE_PAGES_PATH_SUFFIX = "/SitePages"

SYSTEM_LIST_TITLES = {
    "User Information List", "Access Requests", "Workflow History",
    "Workflow Tasks", "TaxonomyHiddenList", "Cache Profiles",
    "Long Running Operation Status", "Maintenance Log Library",
}

# --- Trida KSU (Klassifizierungssystem fuer Unterlagen) na slozkach ---
# Po vytvoreni kazde slozky na cili se nastavi KSU trida na tuto hodnotu.
SET_KSU_CLASS = True
KSU_FIELD_VALUE = "5.3"
# Pokud je KSU sloupec typu Managed Metadata (Taxonomy), potrebujeme GUID termu.
# GUID termu '5.3' byl zjisten z term-store dumpu (5.3 Car Series and Concept Docs).
# Kdyz je vyplneno, preskoci se automaticke hledani (nejrychlejsi a spolehlive):
KSU_TERM_GUID  = "f180d7d0-51f7-4ecb-b85b-8794451fa5fb"
# Label termu pro ValidateUpdateListItem (staci prefix '5.3'; SP resolvuje dle GUID).
KSU_TERM_LABEL = "5.3"
# Interni nazev sloupce neni napevno (lisi se dle knihovny/jazyka), skript ho
# dohleda podle zobrazovaneho nazvu. Doplneno vice kandidatu (CZ i DE).
KSU_FIELD_TITLE_CANDIDATES = (
    # CSD = Classification System for Documents (EN preklad nemeckeho KSU)
    "CSD class", "CSD Class", "CSDclass", "CSD",
    "Třída KSU", "Trida KSU", "KSU třída", "KSU trida",
    "KSU Klasse", "KSU-Klasse", "KSU Class", "KSU",
)

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
    """Escape apostrofu pro OData literal ('...'). Apostrof se zdvojuje."""
    return str(s).replace("'", "''")


def has_unsupported_chars(name):
    """REST endpointy GetFileByServerRelativeUrl nepodporuji % a #."""
    return "%" in name or "#" in name


def parse_cookie_line(line, cookies_dict):
    """Robustni parser - zvladne 3 formaty zapisu cookies."""
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
        raise SystemExit(f"!!! V souboru '{path}' chybi FedAuth nebo rtFa. Zkontroluj format.")
    return cookies


class SPSession:
    """Obalka nad requests.Session pro jeden SharePoint web (site)."""

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

    # --- HTTP s jednoduchym retry na throttling (429/503) ---
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
        """Ziska (a preventivne obnovuje) X-RequestDigest. Plati jen 1800 s."""
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
    entity_type = r.json()["d"]["ListItemEntityTypeFullName"] if r.status_code == 200 else "SP.Data.ListItem"
    _entity_type_cache[key] = entity_type
    return entity_type


# ============================ ZISKANI SEZNAMU SEZNAMU =======================

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
    r = sp.get(
        "/_api/web/lists"
        "?$select=Title,BaseTemplate,BaseType,Hidden,ItemCount,RootFolder/ServerRelativeUrl"
        "&$expand=RootFolder"
    )
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
    r = sp.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
               "?$select=Title,InternalName,TypeAsString,ReadOnlyField,Hidden,FromBaseType")
    r.raise_for_status()
    fields = r.json()["d"]["results"]
    return [f for f in fields
            if not f["ReadOnlyField"] and not f["Hidden"]
            and f["InternalName"] not in ("ContentType", "Attachments")
            and not f["FromBaseType"]]


# ============================ PODROBNY VYPIS OBSAHU (DIAGNOSTIKA) ==========

def _print_folder_tree(sp, folder_url, indent=""):
    r = sp.get(f"/_api/web/GetFolderByServerRelativeUrl('{odata(folder_url)}')?$expand=Folders,Files")
    if r.status_code != 200:
        log(f"{indent}!!! nelze nacist slozku '{folder_url}': {r.status_code}")
        return
    data = r.json()["d"]
    files = data.get("Files", {}).get("results", [])
    folders = data.get("Folders", {}).get("results", [])

    for f in files:
        log(f"{indent}[SOUBOR] {f['Name']}")

    for sub in folders:
        name = sub["Name"]
        if name == "Forms":
            continue
        log(f"{indent}[SLOZKA] {name}/")
        _print_folder_tree(sp, f"{folder_url}/{name}", indent=indent + "    ")


def print_list_contents(sp, lst):
    title = lst["Title"]
    base_type = lst["BaseType"]

    if base_type == 1:
        root = lst["RootFolder"]["ServerRelativeUrl"]
        label = "stranek" if is_site_pages_library(lst) else "knihovny"
        log(f"  Obsah {label} '{title}' (strom slozek a souboru):")
        _print_folder_tree(sp, root, indent="    ")
    else:
        r = sp.get(f"/_api/web/lists/getbytitle('{odata(title)}')/items?$select=Id,Title&$top=5000")
        if r.status_code != 200:
            log(f"    !!! nelze nacist polozky seznamu '{title}': {r.status_code} {r.text[:200]}")
            return
        items = r.json()["d"]["results"]
        log(f"  Obsah seznamu '{title}' ({len(items)} polozek):")
        if not items:
            log(f"    (seznam je prazdny)")
        for item in items:
            name = item.get("Title") or "(bez nazvu)"
            log(f"    [POLOZKA] #{item['Id']:<5} {name}")


# ============================ MAZANI OBSAHU NA CILI =========================

def clear_document_library(tgt, root_folder_url):
    r = tgt.get(f"/_api/web/GetFolderByServerRelativeUrl('{odata(root_folder_url)}')?$expand=Folders,Files")
    if r.status_code != 200:
        log(f"    (cilova slozka zatim neexistuje nebo chyba cteni: {r.status_code})")
        return

    data = r.json()["d"]
    files = data.get("Files", {}).get("results", [])
    folders = data.get("Folders", {}).get("results", [])

    for f in files:
        furl = f["ServerRelativeUrl"]
        log(f"    [MAZAT SOUBOR] {furl}")
        if not DRY_RUN:
            resp = tgt.post(
                f"/_api/web/GetFileByServerRelativeUrl('{odata(furl)}')",
                extra_headers=tgt.write_headers(method_override="DELETE", extra={"IF-MATCH": "*"}),
            )
            if resp.status_code not in (200, 204):
                log(f"      !!! chyba mazani souboru: {resp.status_code} {resp.text[:200]}")

    for sub in folders:
        surl = sub["ServerRelativeUrl"]
        if surl.rstrip("/").endswith("/Forms"):
            continue
        log(f"    [MAZAT SLOZKU] {surl}")
        if not DRY_RUN:
            resp = tgt.post(
                f"/_api/web/GetFolderByServerRelativeUrl('{odata(surl)}')",
                extra_headers=tgt.write_headers(method_override="DELETE", extra={"IF-MATCH": "*"}),
            )
            if resp.status_code not in (200, 204):
                log(f"      !!! chyba mazani slozky: {resp.status_code} {resp.text[:200]}")


def clear_generic_list(tgt, list_title):
    r = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/items?$select=Id&$top=5000")
    if r.status_code != 200:
        log(f"    (cilovy seznam '{list_title}' zatim neexistuje nebo chyba cteni: {r.status_code})")
        return
    items = r.json()["d"]["results"]
    for item in items:
        item_id = item["Id"]
        log(f"    [MAZAT POLOZKU] {list_title} #{item_id}")
        if not DRY_RUN:
            resp = tgt.post(
                f"/_api/web/lists/getbytitle('{odata(list_title)}')/items({item_id})",
                extra_headers=tgt.write_headers(method_override="DELETE", extra={"IF-MATCH": "*"}),
            )
            if resp.status_code not in (200, 204):
                log(f"      !!! chyba mazani polozky: {resp.status_code} {resp.text[:200]}")


# ============================ KOPIROVANI KNIHOVEN (SOUBORY) =========

# --- KSU: dohledani sloupce + nastaveni na slozce ------------------------

# cache: (site, list_title) -> (internal_name, type_as_string) | (None, None)
_ksu_field_cache = {}
# aby se vypis sloupcu (diagnostika) udelal jen jednou na knihovnu
_ksu_dumped = {}
# cache GUID termu KSU: (site, list_title, internal) -> guid | None
_ksu_guid_cache = {}
# globalni reference na zdrojovou session (pro harvest GUID ze zdroje)
_SRC_SESSION = None

# taxonomy/lookup typy, ktere pres prosty REST MERGE nastavit nejde
_KSU_UNSUPPORTED_TYPES = {"TaxonomyFieldType", "TaxonomyFieldTypeMulti",
                          "Lookup", "LookupMulti"}


def dump_fields(tgt, list_title):
    """DIAGNOSTIKA: vypise vsechny (nesystemove) sloupce knihovny - abys videl,
    jak se KSU sloupec ve skutecnosti jmenuje."""
    r = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
                "?$select=Title,InternalName,TypeAsString,Hidden,ReadOnlyField")
    if r.status_code != 200:
        log(f"      (!) nelze nacist sloupce '{list_title}': {r.status_code} {r.text[:200]}")
        return
    log(f"      --- Sloupce knihovny '{list_title}' (Title | InternalName | Type) ---")
    for f in r.json()["d"]["results"]:
        if f.get("Hidden"):
            continue
        _hay = (f["Title"] + f["InternalName"]).lower()
        flag = " [KSU?]" if ("ksu" in _hay or "csd" in _hay) else ""
        log(f"        {f['Title']:35} | {f['InternalName']:30} | {f['TypeAsString']}{flag}")


def resolve_ksu_field(tgt, list_title):
    """Najde interni nazev a typ KSU sloupce v cilove knihovne.
    1) presna shoda Title/InternalName s kandidaty,
    2) fallback: jakykoli sloupec obsahujici 'ksu' (case-insensitive)."""
    key = (tgt.site, list_title)
    if key in _ksu_field_cache:
        return _ksu_field_cache[key]
    internal, ftype = None, None
    r = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
                "?$select=Title,InternalName,TypeAsString,Hidden")
    if r.status_code == 200:
        fields = [f for f in r.json()["d"]["results"] if not f.get("Hidden")]
        # 1) presna shoda (Title NEBO InternalName) s kandidaty
        for cand in KSU_FIELD_TITLE_CANDIDATES:
            for f in fields:
                if cand.strip().lower() in (f["Title"].strip().lower(),
                                            f["InternalName"].strip().lower()):
                    internal, ftype = f["InternalName"], f["TypeAsString"]
                    break
            if internal:
                break
        # 2) fallback: cokoli s podretezcem 'ksu' nebo 'csd'
        if not internal:
            for f in fields:
                hay = (f["Title"] + f["InternalName"]).lower()
                if "ksu" in hay or "csd" in hay:
                    internal, ftype = f["InternalName"], f["TypeAsString"]
                    log(f"      (i) KSU sloupec nalezen fallbackem: "
                        f"'{f['Title']}' (InternalName '{internal}', typ {ftype})")
                    break
    _ksu_field_cache[key] = (internal, ftype)
    return internal, ftype


# cache seznamu vsech KSU sloupcu: (site, list_title) -> [(internal, ftype), ...]
_ksu_fields_multi_cache = {}


def resolve_ksu_fields(tgt, list_title):
    """Vrati SEZNAM vsech KSU/CSD sloupcu (napr. 'CSD class' i 'csd').
    Kazdy jako (internal_name, type_as_string). Bez duplicit."""
    key = (tgt.site, list_title)
    if key in _ksu_fields_multi_cache:
        return _ksu_fields_multi_cache[key]

    found = []
    seen = set()
    r = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
                "?$select=Title,InternalName,TypeAsString,Hidden")
    if r.status_code == 200:
        fields = [f for f in r.json()["d"]["results"] if not f.get("Hidden")]

        def _add(f):
            if f["InternalName"] not in seen:
                seen.add(f["InternalName"])
                found.append((f["InternalName"], f["TypeAsString"]))

        # 1) presne shody s kandidaty (zachovava poradi kandidatu)
        for cand in KSU_FIELD_TITLE_CANDIDATES:
            for f in fields:
                if cand.strip().lower() in (f["Title"].strip().lower(),
                                            f["InternalName"].strip().lower()):
                    _add(f)
        # 2) fallback: cokoli s 'ksu' nebo 'csd' v nazvu
        for f in fields:
            hay = (f["Title"] + f["InternalName"]).lower()
            if "ksu" in hay or "csd" in hay:
                _add(f)

    _ksu_fields_multi_cache[key] = found
    return found


def _harvest_guid_from_items(sp, list_title, internal):
    """Precte GUID termu z existujici polozky, ktera uz ma KSU = KSU_FIELD_VALUE.
    Taxonomy sloupec vraci pri $select objekt {Label, TermGuid, WssId}."""
    if sp is None:
        return None
    r = sp.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/items"
               f"?$select=Id,{internal}&$top=500")
    if r.status_code != 200:
        return None
    for item in r.json()["d"]["results"]:
        val = item.get(internal)
        if isinstance(val, dict):
            label = str(val.get("Label", "")).strip()
            guid = val.get("TermGuid")
            if guid and _label_matches(label, KSU_FIELD_VALUE):
                return guid
    return None


def _termstore_get(sp, endpoint):
    """GET na v2.1 termStore - MUSI mit Accept: application/json (ne odata=verbose)."""
    return sp.get(endpoint, headers={"Accept": "application/json;odata=nometadata"})


def _term_labels(t):
    """Vytahne vsechny textove labely z term objektu (ruzne verze API)."""
    labels = t.get("labels") or []
    names = [str(l.get("name", "")).strip() for l in labels if l.get("name")]
    if not names and t.get("Name"):
        names = [str(t["Name"]).strip()]
    return names


def _label_matches(name, target):
    """KSU labely maji tvar '5.3 PKW-Serienstandsunterlagen' -> porovnavame
    prvni token (kod tridy) s hledanou hodnotou (napr. '5.3')."""
    name = name.strip().lower()
    target = target.strip().lower()
    if name == target:
        return True
    # prvni "slovo" (kod pred prvni mezerou)
    first_token = name.split(" ", 1)[0].split(",", 1)[0].strip()
    return first_token == target


def _lookup_guid_from_termstore(tgt, list_title, internal, dump=False):
    """Dohleda GUID termu 'KSU_FIELD_VALUE' pres v2.1 term-store.
    dump=True -> vypise vsechny termy (label + GUID) pro diagnostiku."""
    # 1) precti TermSetId ze sloupce (tady odata=verbose zustava, je to /_api/web)
    r = tgt.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/fields"
                f"/getbyinternalnameortitle('{odata(internal)}')"
                "?$select=TermSetId,SspId,AnchorId")
    if r.status_code != 200:
        log(f"        (term-store) nelze precist TermSetId sloupce: {r.status_code}")
        return None
    d = r.json()["d"]
    term_set_id = (d.get("TermSetId") or "").strip("{}")
    if not term_set_id or set(term_set_id) <= set("0-"):
        log(f"        (term-store) sloupec nema platny TermSetId (mozna neni taxonomy).")
        return None
    log(f"        (term-store) TermSetId = {term_set_id}")

    found = {"guid": None}

    def _walk(url, depth=0):
        while url:
            rr = _termstore_get(tgt, url)
            if rr.status_code != 200:
                log(f"        (term-store) GET selhal ({rr.status_code}) na {url[:80]}")
                return
            data = rr.json()
            terms = data.get("value") or []
            for t in terms:
                tid = t.get("id") or t.get("Id")
                names = _term_labels(t)
                if dump:
                    log(f"        {'  '*depth}- {', '.join(names):20} | {tid}")
                if found["guid"] is None and any(
                        _label_matches(n, KSU_FIELD_VALUE) for n in names):
                    found["guid"] = tid
                    if not dump:
                        return
                # rekurze do potomku (hierarchie 5 -> 5.3)
                if t.get("childrenCount", 0) and (dump or found["guid"] is None):
                    _walk(f"/_api/v2.1/termStore/sets/{term_set_id}/terms/{tid}/children",
                          depth + 1)
                    if found["guid"] and not dump:
                        return
            # stránkování
            url = data.get("@odata.nextLink")
            if url and "/_api/" in url:
                url = "/_api/" + url.split("/_api/", 1)[1]

    _walk(f"/_api/v2.1/termStore/sets/{term_set_id}/terms")
    return found["guid"]


def get_ksu_term_guid(tgt, list_title, internal):
    """Vrati GUID termu pro KSU_FIELD_VALUE (manual -> harvest -> term-store)."""
    if KSU_TERM_GUID:
        return KSU_TERM_GUID
    key = (tgt.site, list_title, internal)
    if key in _ksu_guid_cache:
        return _ksu_guid_cache[key]
    guid = (_harvest_guid_from_items(tgt, list_title, internal)
            or _harvest_guid_from_items(_SRC_SESSION, list_title, internal)
            or _lookup_guid_from_termstore(tgt, list_title, internal))
    if guid:
        log(f"      (i) GUID termu '{KSU_FIELD_VALUE}' zjisten: {guid}")
    _ksu_guid_cache[key] = guid
    return guid


def _set_taxonomy_ksu(tgt, list_title, item_id, internal, guid):
    """Zapise taxonomy hodnotu pres ValidateUpdateListItem (format Label|GUID)."""
    label = KSU_TERM_LABEL or KSU_FIELD_VALUE
    body = {
        "formValues": [{"FieldName": internal,
                        "FieldValue": f"{label}|{guid}"}],
        "bNewDocumentUpdate": False,
    }
    resp = tgt.post(
        f"/_api/web/lists/getbytitle('{odata(list_title)}')/items({item_id})"
        "/ValidateUpdateListItem",
        extra_headers=tgt.write_headers(extra={"Content-Type": "application/json;odata=verbose"}),
        data=json.dumps(body),
    )
    if resp.status_code not in (200, 201):
        log(f"        !!! chyba ValidateUpdateListItem: {resp.status_code} {resp.text[:250]}")
        return
    # zkontroluj HasException v odpovedi
    try:
        results = resp.json()["d"]["ValidateUpdateListItem"]["results"]
        for fv in results:
            if fv.get("HasException"):
                log(f"        !!! KSU zapis vratil vyjimku u pole {fv.get('FieldName')}: "
                    f"{fv.get('ErrorMessage')}")
    except Exception:
        pass


def _set_plain_ksu(tgt, list_title, item_id, internal, ftype):
    """Zapise KSU na Text/Choice/Number sloupci prostym MERGE."""
    value = KSU_FIELD_VALUE
    if ftype in ("Number", "Currency"):
        try:
            value = float(KSU_FIELD_VALUE)
        except ValueError:
            pass
    entity_type = get_entity_type_full_name(tgt, list_title)
    body = {"__metadata": {"type": entity_type}, internal: value}
    resp = tgt.post(
        f"/_api/web/lists/getbytitle('{odata(list_title)}')/items({item_id})",
        extra_headers=tgt.write_headers(
            method_override="MERGE",
            extra={"IF-MATCH": "*", "Content-Type": "application/json;odata=verbose"},
        ),
        data=json.dumps(body),
    )
    if resp.status_code not in (200, 204):
        log(f"        !!! chyba nastaveni KSU '{internal}': {resp.status_code} {resp.text[:250]}")


def set_folder_ksu(tgt, list_title, folder_server_relative_url):
    """Nastavi KSU tridu (KSU_FIELD_VALUE) na VSECH KSU/CSD sloupcich slozky."""
    if not SET_KSU_CLASS:
        return

    ksu_fields = resolve_ksu_fields(tgt, list_title)
    if not ksu_fields:
        log(f"      (!) Zadny KSU/CSD sloupec v '{list_title}' nenalezen - preskakuji.")
        if not _ksu_dumped.get((tgt.site, list_title)):
            _ksu_dumped[(tgt.site, list_title)] = True
            dump_fields(tgt, list_title)
        return

    # rozdel na taxonomy vs. ostatni; pro taxonomy zjisti GUID (jednou)
    log(f"      [KSU] {folder_server_relative_url} -> {KSU_FIELD_VALUE} "
        f"(sloupce: {', '.join(i for i, _ in ksu_fields)})")
    if DRY_RUN:
        return

    # najdi Id list-item polozky slozky (jednou pro vsechny sloupce)
    r = tgt.get(f"/_api/web/GetFolderByServerRelativeUrl('{odata(folder_server_relative_url)}')"
                "/ListItemAllFields?$select=Id")
    if r.status_code != 200:
        log(f"        !!! nelze nacist ListItem slozky: {r.status_code} {r.text[:200]}")
        return
    item_id = r.json()["d"]["Id"]

    for internal, ftype in ksu_fields:
        if ftype in ("TaxonomyFieldType", "TaxonomyFieldTypeMulti"):
            guid = get_ksu_term_guid(tgt, list_title, internal)
            if not guid:
                log(f"        (!) '{internal}' je Taxonomy, GUID '{KSU_FIELD_VALUE}' "
                    f"neznamy - preskakuji tento sloupec.")
                if not _ksu_dumped.get((tgt.site, "TERMS:" + list_title)):
                    _ksu_dumped[(tgt.site, "TERMS:" + list_title)] = True
                    _lookup_guid_from_termstore(tgt, list_title, internal, dump=True)
                continue
            _set_taxonomy_ksu(tgt, list_title, item_id, internal, guid)
        else:
            _set_plain_ksu(tgt, list_title, item_id, internal, ftype)


def ensure_target_folder(tgt, folder_server_relative_url):
    log(f"    [SLOZKA] {folder_server_relative_url}")
    if DRY_RUN:
        return
    resp = tgt.post(
        "/_api/web/folders",
        extra_headers=tgt.write_headers(extra={"Content-Type": "application/json;odata=verbose"}),
        data=json.dumps({"__metadata": {"type": "SP.Folder"},
                         "ServerRelativeUrl": folder_server_relative_url}),
    )
    if resp.status_code not in (200, 201):
        if "already exists" not in resp.text and resp.status_code not in (500,):
            log(f"      !!! chyba vytvareni slozky: {resp.status_code} {resp.text[:200]}")


def _upload_small(tgt, tgt_folder_url, file_name, content):
    endpoint = (
        f"/_api/web/GetFolderByServerRelativeUrl('{odata(tgt_folder_url)}')"
        f"/Files/add(url='{odata(file_name)}',overwrite=true)"
    )
    resp = tgt.post(endpoint, extra_headers=tgt.write_headers(), data=content)
    if resp.status_code not in (200, 201):
        log(f"      !!! chyba nahrani (small): {resp.status_code} {resp.text[:200]}")


def _upload_large(tgt, tgt_folder_url, file_name, content):
    """Chunked upload pres StartUpload/ContinueUpload/FinishUpload."""
    # 1) vytvor prazdny soubor
    create = tgt.post(
        f"/_api/web/GetFolderByServerRelativeUrl('{odata(tgt_folder_url)}')"
        f"/Files/add(url='{odata(file_name)}',overwrite=true)",
        extra_headers=tgt.write_headers(), data=b"",
    )
    if create.status_code not in (200, 201):
        log(f"      !!! chyba vytvareni prazdneho souboru: {create.status_code} {create.text[:200]}")
        return

    file_url = f"{tgt_folder_url}/{file_name}"
    fu = odata(file_url)
    upload_id = str(uuid.uuid4())
    total = len(content)

    # 2) StartUpload s prvnim chunkem
    first = content[:CHUNK_SIZE]
    r = tgt.post(f"/_api/web/GetFileByServerRelativeUrl('{fu}')/StartUpload(uploadId=guid'{upload_id}')",
                 extra_headers=tgt.write_headers(), data=first)
    if r.status_code not in (200, 201):
        log(f"      !!! StartUpload selhal: {r.status_code} {r.text[:200]}")
        return
    offset = len(first)

    # 3) ContinueUpload / FinishUpload
    while offset < total:
        chunk = content[offset:offset + CHUNK_SIZE]
        is_last = (offset + len(chunk)) >= total
        op = "FinishUpload" if is_last else "ContinueUpload"
        r = tgt.post(f"/_api/web/GetFileByServerRelativeUrl('{fu}')/{op}(uploadId=guid'{upload_id}',fileOffset={offset})",
                     extra_headers=tgt.write_headers(), data=chunk)
        if r.status_code not in (200, 201):
            log(f"      !!! {op} selhal: {r.status_code} {r.text[:200]}")
            return
        offset += len(chunk)

    # 4) pokud se vse veslo do prvniho chunku, je nutne jeste Finish
    if total <= CHUNK_SIZE:
        r = tgt.post(f"/_api/web/GetFileByServerRelativeUrl('{fu}')/FinishUpload(uploadId=guid'{upload_id}',fileOffset={offset})",
                     extra_headers=tgt.write_headers(), data=b"")
        if r.status_code not in (200, 201):
            log(f"      !!! FinishUpload (single-chunk) selhal: {r.status_code} {r.text[:200]}")


def copy_file(src, tgt, src_file_url, tgt_folder_url, file_name):
    log(f"    [SOUBOR] {file_name}  ->  {tgt_folder_url}")
    if has_unsupported_chars(file_name):
        log(f"      (!) PRESKAKUJI - nazev obsahuje '%' nebo '#', ktere REST endpoint nepodporuje.")
        return
    if DRY_RUN:
        return
    r = src.get(f"/_api/web/GetFileByServerRelativeUrl('{odata(src_file_url)}')/$value")
    if r.status_code != 200:
        log(f"      !!! chyba stazeni souboru: {r.status_code}")
        return
    content = r.content

    if len(content) <= SMALL_FILE_LIMIT:
        _upload_small(tgt, tgt_folder_url, file_name, content)
    else:
        log(f"      (velky soubor {len(content)//1024//1024} MB -> chunked upload)")
        _upload_large(tgt, tgt_folder_url, file_name, content)


def copy_library_recursive(src, tgt, src_folder_url, tgt_folder_url, list_title=None):
    r = src.get(f"/_api/web/GetFolderByServerRelativeUrl('{odata(src_folder_url)}')?$expand=Folders,Files")
    r.raise_for_status()
    data = r.json()["d"]
    files = data.get("Files", {}).get("results", [])
    folders = data.get("Folders", {}).get("results", [])

    for f in files:
        copy_file(src, tgt, f["ServerRelativeUrl"], tgt_folder_url, f["Name"])

    for sub in folders:
        name = sub["Name"]
        if name == "Forms":
            continue
        new_src = f"{src_folder_url}/{name}"
        new_tgt = f"{tgt_folder_url}/{name}"
        ensure_target_folder(tgt, new_tgt)
        if list_title:
            set_folder_ksu(tgt, list_title, new_tgt)   # <-- nastav KSU tridu na slozce
        copy_library_recursive(src, tgt, new_src, new_tgt, list_title)


def target_root_for(src_root_url):
    """Premapuje server-relative cestu ze zdrojoveho webu na cilovy web."""
    source_site_path = SOURCE_SITE.split(".com", 1)[1]
    target_site_path = TARGET_SITE.split(".com", 1)[1]
    return src_root_url.replace(source_site_path, target_site_path)


# ============================ KOPIROVANI MODERNICH STRANEK (Site Pages) =====

PAGE_FIELDS_TO_COPY = [
    "Title", "CanvasContent1", "LayoutWebpartsContent",
    "Description", "PromotedState", "PageLayoutType",
]


def copy_site_pages_library(src, tgt, lst):
    title = lst["Title"]
    src_root = lst["RootFolder"]["ServerRelativeUrl"]
    tgt_root = target_root_for(src_root)

    if DELETE_TARGET_CONTENT:
        log(f"  Mazani obsahu cilove knihovny stranek: {tgt_root}")
        clear_document_library(tgt, tgt_root)

    r = src.get(f"/_api/web/GetFolderByServerRelativeUrl('{odata(src_root)}')/Files"
                "?$select=Name,ServerRelativeUrl")
    r.raise_for_status()
    files = [f for f in r.json()["d"]["results"] if f["Name"].lower().endswith(".aspx")]
    log(f"  Nalezeno {len(files)} stranek ke zkopirovani")

    entity_type = get_entity_type_full_name(src, title)

    for f in files:
        name = f["Name"]
        log(f"    [STRANKA] {name}")

        r_item = src.get(
            f"/_api/web/GetFileByServerRelativeUrl('{odata(f['ServerRelativeUrl'])}')"
            "/ListItemAllFields?$select=" + ",".join(PAGE_FIELDS_TO_COPY + ["BannerImageUrl"])
        )
        if r_item.status_code != 200:
            log(f"      !!! nelze nacist obsah stranky: {r_item.status_code} {r_item.text[:200]}")
            continue
        item = r_item.json()["d"]

        if DRY_RUN:
            continue

        tgt_file_url = f"{tgt_root}/{name}"

        create_resp = tgt.post(
            f"/_api/web/GetFolderByServerRelativeUrl('{odata(tgt_root)}')"
            f"/Files/AddTemplateFile(urlOfFile='{odata(tgt_file_url)}',templateFileType=3)",
            extra_headers=tgt.write_headers(),
        )
        if create_resp.status_code not in (200, 201) and "already exists" not in create_resp.text:
            log(f"      !!! chyba vytvareni stranky: {create_resp.status_code} {create_resp.text[:250]}")
            continue

        r_new = tgt.get(f"/_api/web/GetFileByServerRelativeUrl('{odata(tgt_file_url)}')"
                        "/ListItemAllFields?$select=Id")
        if r_new.status_code != 200:
            log(f"      !!! nelze najit nove vytvorenou stranku na cili: {r_new.status_code}")
            continue
        new_id = r_new.json()["d"]["Id"]

        values = {"__metadata": {"type": entity_type}}
        for field in PAGE_FIELDS_TO_COPY:
            val = item.get(field)
            if val is not None:
                values[field] = val

        banner = item.get("BannerImageUrl")
        if isinstance(banner, dict) and banner.get("Url"):
            values["BannerImageUrl"] = {
                "__metadata": {"type": "SP.FieldUrlValue"},
                "Url": banner["Url"],
                "Description": banner.get("Description", ""),
            }

        update_resp = tgt.post(
            f"/_api/web/lists/getbytitle('{odata(title)}')/items({new_id})",
            extra_headers=tgt.write_headers(
                method_override="MERGE",
                extra={"IF-MATCH": "*", "Content-Type": "application/json;odata=verbose"},
            ),
            data=json.dumps(values),
        )
        if update_resp.status_code not in (200, 204):
            log(f"      !!! chyba zapisu obsahu stranky: {update_resp.status_code} {update_resp.text[:300]}")

        tgt.post(
            f"/_api/web/GetFileByServerRelativeUrl('{odata(tgt_file_url)}')"
            "/CheckIn(comment='Kopie ze zdroje',checkintype=1)",
            extra_headers=tgt.write_headers(),
        )
        tgt.post(
            f"/_api/web/GetFileByServerRelativeUrl('{odata(tgt_file_url)}')/Publish('Kopie ze zdroje')",
            extra_headers=tgt.write_headers(),
        )


# ============================ KOPIROVANI GENERICKYCH SEZNAMU (POLOZKY) ======

SKIP_FIELD_TYPES = {"User", "UserMulti", "Lookup", "LookupMulti",
                    "TaxonomyFieldType", "TaxonomyFieldTypeMulti"}


def copy_generic_list_items(src, tgt, list_title):
    fields = get_fields(src, list_title)
    skipped_fields = [f["Title"] for f in fields if f["TypeAsString"] in SKIP_FIELD_TYPES]
    if skipped_fields:
        log(f"    (!) Preskakuji nepodporovane sloupce (User/Lookup/Taxonomy): {skipped_fields}")

    r = src.get(f"/_api/web/lists/getbytitle('{odata(list_title)}')/items?$top=5000")
    r.raise_for_status()
    items = r.json()["d"]["results"]
    log(f"    Nalezeno {len(items)} polozek ke kopirovani")

    entity_type = get_entity_type_full_name(src, list_title)

    for item in items:
        values = {"__metadata": {"type": entity_type}}
        for field in fields:
            iname = field["InternalName"]
            ftype = field["TypeAsString"]
            if ftype in SKIP_FIELD_TYPES:
                continue
            val = item.get(iname)
            if val is None:
                continue
            values[iname] = val

        log(f"    [POLOZKA] {list_title} - {values.get('Title', item.get('Id'))}")
        if not DRY_RUN:
            resp = tgt.post(
                f"/_api/web/lists/getbytitle('{odata(list_title)}')/items",
                extra_headers=tgt.write_headers(extra={"Content-Type": "application/json;odata=verbose"}),
                data=json.dumps(values),
            )
            if resp.status_code not in (200, 201):
                log(f"      !!! chyba vytvareni polozky: {resp.status_code} {resp.text[:300]}")


# ============================ HLAVNI BEH ====================================

def main():
    log("=" * 70)
    log(f"START kopirovani   DRY_RUN={DRY_RUN}   DELETE_TARGET_CONTENT={DELETE_TARGET_CONTENT}")
    log(f"Zdroj: {SOURCE_SITE}")
    log(f"Cil:   {TARGET_SITE}")
    log("=" * 70)

    src = SPSession(SOURCE_SITE, SOURCE_COOKIE_FILE)
    tgt = SPSession(TARGET_SITE, TARGET_COOKIE_FILE)

    global _SRC_SESSION
    _SRC_SESSION = src   # umozni harvest GUID KSU termu i ze zdroje

    for name, sp in (("ZDROJ", src), ("CIL", tgt)):
        r = sp.get("/_api/web?$select=Title")
        if r.status_code == 200:
            log(f"[{name}] Pripojeno: {r.json()['d']['Title']}")
        else:
            log(f"[{name}] !!! Pripojeni selhalo: {r.status_code} {r.text[:200]}")
            return

    lists = get_lists(src)
    log(f"\nNalezeno {len(lists)} seznamu/knihoven ke zpracovani:\n")
    for lst in lists:
        if is_site_pages_library(lst):
            kind = "Site Pages (moderni stranky)"
        elif lst["BaseType"] == 1:
            kind = "Document Library"
        else:
            kind = "Generic List"
        log(f"  - {lst['Title']:30} (BaseType={lst['BaseType']}, BaseTemplate={lst['BaseTemplate']}, ItemCount={lst['ItemCount']}) [{kind}]")

    if PRINT_CONTENTS:
        log(f"\n{'='*70}")
        log("PODROBNY VYPIS OBSAHU (co skript vidi na zdroji)")
        log(f"{'='*70}")
        for lst in lists:
            log(f"\n--- {lst['Title']} ---")
            print_list_contents(src, lst)

    for lst in lists:
        title = lst["Title"]
        log(f"\n{'='*70}\nZPRACOVAVAM: {title}\n{'='*70}")

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

    log("\n" + "=" * 70)
    log("HOTOVO" + ("  (DRY RUN - nic se nezapsalo)" if DRY_RUN else ""))
    log("=" * 70)
    save_log()


if __name__ == "__main__":
    main()
