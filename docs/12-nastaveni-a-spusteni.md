# Nastavení a spuštění

Jediný dokument, který je potřeba k používání. Ostatní v `docs/` jsou detaily
k jednotlivým částem.

---

# Část 1 — Nastavení, jednou

## 1.1 Co musí být na počítači

| Co | Jak ověřit |
|----|-----------|
| PowerShell | Start → napsat `PowerShell` |
| Modul `PnP.PowerShell` | `Get-Module PnP.PowerShell -ListAvailable` vypíše verzi |
| Modul `ImportExcel` | `Get-Module ImportExcel -ListAvailable` — jen pro čtení `.xlsx` |

Instalace, když chybí. Nepotřebuje práva správce:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

`ImportExcel` se dá obejít — v Excelu **Soubor → Uložit jako → CSV UTF-8** a
v konfiguraci uvést ten `.csv`.

## 1.2 Stáhnout soubory

```powershell
git clone https://github.com/ohlava/codeAlongMicrosoft26.git
```

```powershell
cd codeAlongMicrosoft26
```

```powershell
git checkout feature/setup-project-site
```

Bez `git` jde repozitář stáhnout na GitHubu tlačítkem **Code → Download ZIP**
a rozbalit.

## 1.3 Vytvořit konfiguraci

```powershell
Copy-Item .\config\settings.example.json .\config\settings.json
```

```powershell
notepad .\config\settings.json
```

Vyplnit:

```json
{
  "clientId": "<dlouhe cislo z puvodniho script.ps1>",
  "sourceSiteUrl": "https://<tenant>.sharepoint.com/sites/<vzorovy-web>",

  "library": "Dokumenty",
  "folderStructureFile": "Folder_Structure.xlsx",

  "fixedMetadata": {
    "RevIMBCS": "5.3 Car Series and Concept Docs"
  },
  "csdClassScope": "DefaultValue",

  "legacyScript": {
    "copyLists": false,
    "copyPages": true,
    "copyDesign": true,
    "copyRegional": true,
    "copyNavigation": true,
    "setOfflineAvailable": false
  }
}
```

| Klíč | Co to je |
|------|----------|
| `clientId` | ID aplikace pro přihlášení. Stejné, jaké měl původní `script.ps1` v proměnné `$ClientId` |
| `sourceSiteUrl` | Vzorový web, ze kterého se kopírují seznamy, kalendáře, stránky a vzhled |
| `library` | Knihovna dokumentů v cílovém webu, kde vznikne struktura složek |
| `folderStructureFile` | Excel se strukturou složek |
| `fixedMetadata` | Metadata na každé vytvořené složce. Klíč je **interní** název sloupce |
| `csdClassScope` | `DefaultValue`, `ExistingFiles`, nebo `Both` — viz 1.5 |
| `legacyScript` | Nastavení pro volitelný krok `TemplateClone` |

**`config/settings.json` je v `.gitignore`**, do repozitáře se nedostane.

## 1.4 Ověřit, že sedí sloupce

Nic to nemění. Vypíše sloupce knihovny s interními názvy a typy, a upozorní na
duplicitní názvy:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\New-FolderStructure.ps1 -SiteUrl "<URL vzoroveho webu>" -Library "Dokumenty" -Path .\Folder_Structure.xlsx -ListFields
```

A jaké hodnoty přijímá CSD Class:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\Set-CsdClass.ps1 -SiteUrl "<URL vzoroveho webu>" -Library "Dokumenty" -Field RevIMBCS -ListTerms
```

Hodnota v `fixedMetadata` musí být jeden z vypsaných termínů. Stačí i jen jeho
číslo, tedy `"5.3"`.

## 1.5 Jak se CSD Class nastavuje

Jsou to **tři různá místa** a každé se nastavuje jinak:

| Kde | Čím | Nastavení |
|-----|-----|-----------|
| na složce | zápis na položku | vždy, z `fixedMetadata` |
| u nově nahrávaných souborů | výchozí hodnota sloupce knihovny | `csdClassScope: DefaultValue` |
| u souborů, které v knihovně už jsou | zápis na každou položku | `csdClassScope: Both` |

Metadata zapsaná na složku se na soubory v ní **nepřenesou** — proto ta výchozí
hodnota sloupce.

`RevIMBCS` (`CSD Class`) je sloupec se spravovanými metadaty a je `ReadOnly`,
protože patří k records managementu. Zapisuje se do něj přes CSOM
(`SetFieldValueByValue`), kterému příznak `ReadOnly` nevadí — postup převzatý
ze `Set-CsdClass.ps1` od Sergiu Nicy. **Odemykat sloupec není potřeba.**

> Zápis do `RevIMBCS` je zásah do klasifikace záznamů. Že to technicky jde,
> neznamená, že se to smí — potvrďte si to se správcem records managementu a
> zapište do [decisions.md](decisions.md).

---

# Část 2 — Spuštění

Mění se jedině adresa cílového webu. Vše ostatní je v konfiguraci.

## 2.1 Podívat se, co se stane

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\Setup-ProjectSite.ps1 -TargetSiteUrl "<URL noveho webu>"
```

Nic to nezmění. Na konci je souhrn, co by proběhlo.

## 2.2 Připravit web

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\Setup-ProjectSite.ps1 -TargetSiteUrl "<URL noveho webu>" -Apply
```

Tohle je ten běžný případ. Proběhne:

1. **Folders** — složky z Excelu, včetně metadat a CSD Class na každé složce
2. **CsdClass** — výchozí hodnota sloupce, aby ji dostaly nově nahrané soubory
3. **Lists** — seznamy ze vzorového webu, bez položek
4. **Events** — kalendáře ze vzorového webu, bez položek
5. **Pages** — stránky, webparty, obrázky, vzhled, regionální nastavení

## 2.3 Varianty

**Včetně obsahu seznamů a kalendářů:**

```powershell
... -TargetSiteUrl "<URL>" -WithData -Apply
```

**Jen část:**

```powershell
... -TargetSiteUrl "<URL>" -Steps Folders -Apply
```

```powershell
... -TargetSiteUrl "<URL>" -Steps Lists,Events -Apply
```

**Naklonovat vzorový web jako celek a pak doplnit složky:**

```powershell
... -TargetSiteUrl "<URL>" -Steps TemplateClone,Folders,CsdClass -Apply
```

**Přenést i navigaci:**

```powershell
... -TargetSiteUrl "<URL>" -Steps Folders,Lists,Events,Pages,CsdClass,Navigation -Apply
```

## 2.4 Přehled kroků

| Krok | Ve výchozí sadě | Náhled bez `-Apply` | Opakovatelné |
|------|-----------------|---------------------|--------------|
| `Folders` | ano | ano, vypíše plán | ano |
| `CsdClass` | ano | ano | ano |
| `Lists` | ano | ne | ano, chybějící doplní |
| `Events` | ano | ne | ano, chybějící doplní |
| `Pages` | ano | ano, vypíše stránky | **ne, přepisuje stránky** |
| `Navigation` | ne | ano, vypíše strom | ano |
| `TemplateClone` | ne | ne | **ne, maže seznamy** |

Pořadí je dané a nezávisí na tom, jak se kroky vypíšou ve `-Steps`.

**Dva kroky nejsou přírůstkové.** `Pages` stránky přepisuje celé, takže zahodí
ruční úpravy na cílových stránkách. `TemplateClone` maže cílové seznamy včetně
dat, než je vytvoří znovu — jen na čerstvě založený web.

---

# Část 3 — Když něco nefunguje

## Struktura souborů

```
codeAlongMicrosoft26\
├─ config\
│   ├─ settings.example.json     šablona v gitu
│   └─ settings.json             vaše nastavení, není v gitu
├─ src\
│   ├─ Setup-ProjectSite.ps1     jediné, co se spouští
│   ├─ New-FolderStructure.ps1   složky a metadata
│   ├─ Set-CsdClass.ps1          CSD Class na knihovně a souborech
│   ├─ Copy-SharePointLists.ps1  seznamy
│   ├─ Copy-SharePointEvents.ps1 kalendáře
│   ├─ Copy-SitePages.ps1        stránky a vzhled
│   ├─ Copy-SiteNavigation.ps1   navigace
│   └─ script.ps1                původní klonovací skript
├─ docs\
├─ Folder_Structure.xlsx         struktura složek, upravujte podle potřeby
└─ export\                       vzniká sám: plány, seznamy, varování
```

Skripty v `src/` se nespouštějí přímo, kromě diagnostiky v 1.4.
`Setup-ProjectSite.ps1` si je volá sám a na začátku ověří, že tam jsou.

## Časté chyby

| Hláška | Řešení |
|--------|--------|
| `Chybí konfigurace: config/settings.json` | Krok 1.3 |
| `V konfiguraci není vyplněné clientId` | Krok 1.3 |
| `Konfiguraci nelze přečíst - není to platný JSON` | Nejčastěji chybějící nebo přebývající čárka |
| `spouštění skriptů je v tomto systému zakázáno` | Používejte `powershell.exe -ExecutionPolicy Bypass -File ...` jako všude v tomto postupu |
| `Modul PnP.PowerShell není nainstalovaný` | Krok 1.1 |
| `Pro čtení souboru .xlsx je potřeba modul ImportExcel` | Krok 1.1, nebo uložit jako CSV |
| `Knihovnu '<X>' jsem nenašel` | Skript vypíše dostupné knihovny i s GUIDy |
| `Název '<X>' odpovídá N sloupcům` | Předat interní název, viz 1.4 |
| `Termín '<X>' v term setu není` | Použít přesný název z `-ListTerms`, viz 1.4 |
| `Term set ... nelze načíst` | Chybí přístup na Term Store |
| Krok skončil `CHYBA:` | Podrobnosti jsou ve výpisu nad souhrnem a v `export/` |

Jeden neúspěšný krok nezastaví ostatní — zapíše se do souhrnu a pokračuje se.

## Co řešení nedělá

- **nezakládá web** — cílový web musí existovat,
- **neřeší oprávnění** — skupiny ani role vůbec,
- **nepřenáší knihovny dokumentů ze vzoru včetně souborů** — struktura složek se
  bere z Excelu; kopírování knihoven s obsahem umí `Copy-PnPDocLibs` ve
  `script.ps1` verze 2.7 (autor Sergiu Nica, zatím na vlastní branchi),
- **nedoplňuje metadata u souborů, které vznikly dřív** — pokud je chcete
  označit, `csdClassScope: Both`.

## Podrobnosti k jednotlivým částem

| Téma | Dokument |
|------|----------|
| Složky z Excelu, metadata, Term Store | [07-struktura-slozek-z-excelu.md](07-struktura-slozek-z-excelu.md) |
| Navigace | [08-kopirovani-navigace.md](08-kopirovani-navigace.md) |
| Kroky a flagy | [09-jeden-skript.md](09-jeden-skript.md) |
| Stránky a vzhled | [10-stranky-a-vzhled.md](10-stranky-a-vzhled.md) |
| Klonování vzoru | [11-klonovani-vzoru.md](11-klonovani-vzoru.md) |
| Rozbor původního script.ps1 | [05-analyza-skriptu.md](05-analyza-skriptu.md) |
