# SharePoint Site Copy Tool

Python skript pro kopírování obsahu SharePoint Online webu:

- dokumentové knihovny
- seznamy
- vlastní sloupce
- metadata dokumentů
- views
- formátování views
- formátování sloupců
- navigace
- Site Assets
- Style Library
- moderní stránky (Site Pages)
- Home Page

---

# Požadavky

## Python

Ověření verze:

```powershell
python --version
```

Požadováno:

```text
Python 3.14+
```

---

# Příprava pracovního adresáře

Vytvoř složku:

```powershell
mkdir C:\SharePointCopy
cd C:\SharePointCopy
```

Do ní zkopíruj:

```text
kopirovani.py
cookies.txt
target_cookies.txt
```

---

# Vytvoření virtuálního prostředí

```powershell
python -m venv venv
```

Aktivace:

```powershell
.\venv\Scripts\Activate.ps1
```

Pokud PowerShell blokuje spuštění:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Poté:

```powershell
.\venv\Scripts\Activate.ps1
```

Úspěšná aktivace:

```text
(venv)
```

---

# Instalace závislostí

S průchodem přes proxy server ve Škoda auto

```powershell
pip install requests --proxy http://proxy.mb.skoda.vwg:8080
```

Kontrola:

```powershell
pip list
```

Musí obsahovat:

```text
requests
urllib3
charset-normalizer
certifi
idna
```

---

# Získání SharePoint cookies

Přihlas se do SharePointu:

```text
https://volkswagengroup.sharepoint.com/sites/TESTD
```

Otevři Developer Tools:

```text
F12
Application
Cookies
https://volkswagengroup.sharepoint.com
```

Najdi hodnoty:

```text
FedAuth
rtFa
```

---

# Soubor cookies.txt

```text
FedAuth=xxxxxxxxxxxxxxxx
rtFa=xxxxxxxxxxxxxxxx
```

---

# Soubor target_cookies.txt

Pokud používáš stejný účet:

```text
FedAuth=xxxxxxxxxxxxxxxx
rtFa=xxxxxxxxxxxxxxxx
```

---

# Konfigurace zdrojového a cílového webu

Ve skriptu nastav:

```python
SOURCE_SITE = "https://volkswagengroup.sharepoint.com/sites/TESTD"

TARGET_SITE = "https://volkswagengroup.sharepoint.com/sites/TargetTESTD"
```

---

# Proxy (volitelné)

Pokud používáš lokální proxy:

```python
PROXY = "http://127.0.0.1:9001"
```

Pokud ne:

```python
PROXY = None
```

---

# První test

Doporučené nastavení:

```python
DRY_RUN = True
WIPE_TARGET = False
```

Spuštění:

```powershell
python copy_site.py
```

Očekávaný výstup:

```text
[ZDROJ] Pripojeno: Source TEST D
[CIL] Pripojeno: Target TEST D
```

---

# Chybové stavy

## 401 Unauthorized

```text
401 UNAUTHORIZED
```

Příčina:

- neplatné cookies
- expirované cookies

Řešení:

- znovu zkopírovat FedAuth
- znovu zkopírovat rtFa

---

## 403 Forbidden

```text
403 FORBIDDEN
```

Příčina:

- účet nemá potřebná oprávnění

Řešení:

- ověřit oprávnění na SharePoint webu

---

# Diagnostika připojení

Soubor:

```python
import requests

url = "https://volkswagengroup.sharepoint.com/sites/TESTD/_api/web?$select=Title"

cookies = {
    "FedAuth": "...",
    "rtFa": "..."
}

r = requests.get(
    url,
    cookies=cookies,
    headers={
        "Accept": "application/json;odata=verbose"
    }
)

print(r.status_code)
print(r.text[:1000])
```

Spuštění:

```powershell
python test.py
```

Správný výsledek:

```text
200
```

---

# Produkční spuštění

Nastavení:

```python
DRY_RUN = False
WIPE_TARGET = True
```

Spuštění:

```powershell
python copy_site.py
```

---

# Co skript kopíruje

✅ Document Libraries

✅ Složky

✅ Dokumenty

✅ Metadata dokumentů

✅ Vlastní sloupce

✅ Taxonomy sloupce

✅ Views

✅ View Formatting

✅ Column Formatting
