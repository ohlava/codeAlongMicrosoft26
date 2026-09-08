# Struktura složek z Excelu

Postup pro `src/New-FolderStructure.ps1`. Skript vezme tabulku se strukturou
složek a vytvoří ji v knihovně dokumentů na SharePointu, včetně metadat.

## Jak má tabulka vypadat

Jeden řádek = jedna složka. Cesta se skládá ze sloupců `Level1`, `Level2`, … —
použijí se všechny neprázdné zleva. Ostatní sloupce jsou metadata.

| Level1 | Level2 | Level3 | Made in/responsible | English translation | German translation |
|--------|--------|--------|---------------------|---------------------|--------------------|
| `01_Organisation` | | | A1 | Organisation | Organisation |
| `01_Organisation` | `04_RASI` | | A3 | RASI | RASI |
| `04_Meetings` | `07_Groups` | `05_Zones` | A5 | Zones | Zonen |

Z toho vznikne:

```
01_Organisation
  04_RASI
04_Meetings
  07_Groups
    05_Zones
```

Sloupců `Level` může být libovolný počet, ne jen čtyři — skript si vezme všechny,
které se jmenují `Level<číslo>`. Nadřazené složky nemusí mít vlastní řádek,
doplní se samy.

Ověřeno proti `Folder_Structure.xlsx`: 36 řádků, 36 složek, hloubka 4 úrovně.

## Kam se zapíšou metadata

Výchozí mapování je nastavené na `Folder_Structure.xlsx`:

| Sloupec v tabulce | Sloupec v SharePointu | Typ |
|-------------------|------------------------|-----|
| `Made in/responsible` | `Responsible` | Text |
| `English translation` | `NameEnglish` | Text |
| `German translation` | `NameGerman` | Text |

Chybějící sloupce si skript v knihovně sám vytvoří a přidá do výchozího
zobrazení. Metadata se zapisují **na složku** — v SharePointu je složka položkou
seznamu, takže se na ní dají držet hodnoty sloupců stejně jako na dokumentu.

Jiné mapování jde předat přes `-MetadataMap`, metadata se dají vypnout přes
`-SkipMetadata`.

### Konstantní hodnota na každou složku

Hodnota, která je pro celý běh stejná a v tabulce není, se předá přes
`-FixedMetadata`. Typicky klasifikace celé sady dokumentů:

```powershell
-FixedMetadata @{ "CSD Class" = "5.3 Car Series and Concept Docs" }
```

**Název sloupce s mezerou musí být v uvozovkách.** Bez nich PowerShell hlásí
chybu, protože `CSD` a `Class` bere jako dvě věci před `=`.

Zapíše se na **každou** vytvořenou složku ve všech úrovních. Klíč je interní
název sloupce, hodnota text. Sloupců se dá předat víc:

```powershell
-FixedMetadata @{ "CSD Class" = "5.3 Car Series and Concept Docs"; ProjectId = "EOZ-2026-014" }
```

Sloupec, který v knihovně chybí, se vytvoří jako Text a přidá do výchozího
zobrazení — stejně jako sloupce z `-MetadataMap`. Když stejný sloupec plní
i tabulka, vyhrává hodnota z tabulky, protože je konkrétnější.

### Displejový a interní název

Sloupec má v SharePointu dva názvy: ten, který je vidět (`CSD Class`), a interní,
kterým se k němu přistupuje z API. U sloupce s mezerou v názvu se liší —
interní bývá `CSD_x0020_Class`.

**Skript si to přeloží sám**, takže se zadává ten název, který je vidět
v knihovně. Funguje i interní, kdyby ho někdo znal. Krok „Sloupce pro metadata"
vypíše, co našel:

```
  = Responsible už existuje
  = CSD Class (interně CSD_x0020_Class) už existuje
  + [dry-run] vytvořil bych 'Name (English)' (interně NameEnglish, Text)
```

Nově zakládané sloupce dostanou interní název bez mezer a diakritiky
(`CSD Class` -> `CSDClass`), aby v API nekončily jako nečitelné `_x0020_`.
Displejový název zůstane tak, jak byl zadaný.

Sloupec, který se nepodaří najít ani založit, se u zápisu metadat přeskočí
s varováním — zbytek metadat se zapíše.

### Když je v knihovně víc sloupců se stejným názvem

Knihovna může obsahovat několik sloupců, které se **jmenují stejně**, ale mají
různý interní název a typ — typicky když jeden pochází z content typu, druhý
založil někdo ručně a třetí přišel se šablonou.

Displejový název pak neurčuje sloupec jednoznačně. Skript v takovém případě
**nic nezapíše a nic nezaloží**, jen vypíše kandidáty:

```
WARNING: Název 'CSD Class' odpovídá 3 sloupcům:
         RevIMBCS (TaxonomyFieldType), CSD_x0020_Class (Text), CSDClass (Text).
         Zadejte místo něj interní název toho správného.
```

Řešení je předat **interní** název toho správného:

```powershell
-FixedMetadata @{ "CSD_x0020_Class" = "5.3 Car Series and Concept Docs" }
```

Interní název určuje sloupec jednoznačně, takže duplicitní displejové názvy
přestanou vadit — ale je pořád lepší je uklidit, jinak uživatel v knihovně vidí
tři stejně pojmenované sloupce a neví, do kterého psát.

### Odkud se duplicity berou

Starší verze skriptu hledala existující sloupec **jen podle interního názvu**.
Ten ale nikdy neobsahuje mezeru: `Add-PnPField -InternalName "CSD Class"`
založí sloupec s interním názvem `CSD_x0020_Class`. Při dalším běhu skript
hledal `CSD Class`, nenašel ho a založil další — SharePoint mu přidal číslo:
`CSD_x0020_Class0`, `CSD_x0020_Class1`. Tři běhy tedy udělaly tři sloupce.

Aktuální verze hledá i podle displejového názvu a při nejednoznačnosti **nic
nezakládá**. Duplicity, které už vznikly, ale sama neuklidí — mazání zůstává
na člověku.

Před smazáním se vyplatí zjistit, kde jsou data:

```powershell
Get-PnPListItem -List "<knihovna>" -Fields "FileLeafRef","CSD_x0020_Class","CSD_x0020_Class0","CSD_x0020_Class1" |
    ForEach-Object { $_.FieldValues } |
    Select-Object FileLeafRef, CSD_x0020_Class, CSD_x0020_Class0, CSD_x0020_Class1
```

a přebytečné pak odstranit:

```powershell
Remove-PnPField -List "<knihovna>" -Identity "CSD_x0020_Class0" -Force
Remove-PnPField -List "<knihovna>" -Identity "CSD_x0020_Class1" -Force
```

Co v knihovně skutečně je, ukáže diagnostický režim:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\New-FolderStructure.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<web>" -Library "Shared Documents" -Path .\Folder_Structure.xlsx -ListFields
```

Nic nezapisuje. Vypíše všechny sloupce s interním názvem a typem, upozorní na
duplicitní názvy a uloží úplný seznam do `export/library-fields.csv`.

### Sloupce se spravovanými metadaty (Term Store)

Sloupec typu **TaxonomyFieldType** neukládá text, ale odkaz na termín z Term
Store. Zapsat do něj řetězec nejde — SharePoint hodnotu zahodí a PnP jen varuje:

```
WARNING: Unable to find the specified term. Skipping values for field 'RevIMBCS'
```

Skript to teď řeší sám: u sloupce se spravovanými metadaty si ze `SchemaXml`
přečte, ke kterému **term setu** je připojený, dohledá v něm termín podle názvu
a zapíše jeho GUID. Zadává se tedy normální název termínu:

```powershell
-FixedMetadata @{ "RevIMBCS" = "5.3 Car Series and Concept Docs" }
```

GUID termínu jde předat taky, pokud ho znáte — pozná se podle formátu a použije
se přímo.

#### Nejdřív si termíny vypsat

Než se něco zapíše, ověřte, že termín v term setu skutečně je a jak přesně se
jmenuje. Diagnostický režim nic nezapisuje:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\New-FolderStructure.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<web>" -Library "Dokumenty" -Path .\Folder_Structure.xlsx -ListTerms RevIMBCS
```

Vypíše skupinu, term set, všechny termíny s GUIDy a uloží je do
`export/terms-RevIMBCS.csv`. Když termín podle názvu nenajde při zápisu, napoví
podobné názvy — překlepy a jiná varianta zápisu jsou nejčastější příčina.

Term set se dohledává průchodem skupin v Term Store, protože ze samotného
`TermSetId` skupinu zjistit nelze. Vyžaduje to **právo čtení na Term Store**;
bez něj se termíny nenačtou a skript to řekne.

#### Records management

Pozor, pokud je ten sloupec součástí řešení pro správu záznamů — poznáte to
podle sousedních sloupců s prefixem `RevIM*`, `LegalHold` nebo
`Popisek pro uchovávání informací`. Takový sloupec bývá `ReadOnly` právě proto,
že ho má plnit jen to řešení.

Odemknout ho jde přes `-UnlockReadOnlyFields`, ale je to zásah do klasifikace
záznamů, ne do popisků projektu. **Potvrďte to předem se správcem Term Store
nebo records managementu** a rozhodnutí si zapište do
[decisions.md](decisions.md) — jinak se špatně dohledává, kdo klasifikaci
přepsal a proč.

## Spuštění

Potřebné předpoklady a odkud vzít `ClientId` jsou v
[06-jak-spustit-export.md](06-jak-spustit-export.md) — jsou stejné.

### Krok 1 — Nejdřív dry-run

**Výchozí režim nic nezapisuje.** Vypíše plán a uloží ho do
`export/folder-plan.csv`.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\New-FolderStructure.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<web>" -Library "Shared Documents" -Path .\Folder_Structure.xlsx
```

Výstup vypadá takto — `+` je „vytvořím", `=` je „už existuje, nechávám":

```
  + 01_Organisation   [NameEnglish=Organisation; NameGerman=Organisation; Responsible=A1]
  +   04_RASI   [NameEnglish=RASI; NameGerman=RASI; Responsible=A3]
  = 02_Timeplan   [NameEnglish=Timeplan; NameGerman=Terminplan; Responsible=A3]

  vytvořit: 35     už existuje: 1
```

### Krok 2 — Teprve pak ostře

Stejný příkaz s `-Apply` na konci:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\New-FolderStructure.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<web>" -Library "Shared Documents" -Path .\Folder_Structure.xlsx -Apply
```

S konstantním sloupcem `CSD Class`:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\New-FolderStructure.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<web>" -Library "Shared Documents" -Path .\Folder_Structure.xlsx -FixedMetadata @{ "CSD Class" = "5.3 Car Series and Concept Docs" } -Apply
```

## Když nemáte modul ImportExcel

Čtení `.xlsx` potřebuje modul `ImportExcel`:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

Nebo se dá obejít úplně: v Excelu **Soubor → Uložit jako → CSV UTF-8** a skriptu
předat ten `.csv`. Funguje to bez instalace čehokoli. Oddělovač si skript
detekuje sám (české Excely ukládají středníkem).

## Vlastnosti, na které je dobré se spolehnout

**Dá se pouštět opakovaně.** Existující složky přeskočí, chybějící doplní.
Když se do tabulky přidá nová složka, stačí skript spustit znovu — přidá jen ji.

**Nikdy nic nemaže.** Složku, která je v knihovně ale ne v tabulce, nechá být
a ani ji nehlásí jako chybu. Odstranění je vždy na člověku.

**Jedna chyba nezastaví zbytek.** Problémová složka se zapíše do
`export/folder-warnings.txt` a pokračuje se dál.

**Rozbité řádky přeskočí.** Řádek, kde je zaplněná hlubší úroveň, ale nadřazená
je prázdná, se přeskočí s varováním. Stejně tak název s nepovolenými znaky
(`" * : < > ? / \ |`), s mezerou na konci nebo končící tečkou.

## Když se metadata nezapíšou a nic nehlásí chybu

SharePoint umí zápis tiše zahodit. `Set-PnPListItem` projde bez chyby, ale
hodnota se neuloží. Příčin je několik a liší se tím, kde se odemykají.

Nejdřív diagnostika — nic nezapisuje:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\New-FolderStructure.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<web>" -Library "Shared Documents" -Path .\Folder_Structure.xlsx -ListFields
```

Vypíše u každého sloupce příznaky `ReadOnlyField`, `Sealed` a `Hidden` a zvlášť
upozorní na ty, do kterých zapisovat nelze.

### 1. Sloupec je ReadOnly

Nejčastější případ. Skript to teď pozná a odmítne zápis s vysvětlením místo
tichého selhání. Odemknout na dobu zápisu a hned vrátit zpět umí sám:

```powershell
-UnlockReadOnlyFields
```

Sloupce, které byly zamčené, odemkne, zapíše metadata a **na konci je vrátí zpět
na ReadOnly** — i když zápis mezitím selže, protože vrácení je ve `finally`
bloku. Kdyby se vrácení nepodařilo, napíše to jako varování začínající `POZOR:`.
Takové hlášení nepřehlédněte a sloupec vraťte ručně.

Vyžaduje právo měnit sloupce knihovny (vlastník webu).

### 2. Sloupec je Sealed

Sealed sloupec pochází z content typu a na úrovni knihovny se změnit nedá.
Skript na to upozorní, ale sám s tím nic neudělá — mění se u content typu
v Content Type Gallery, což je zásah do sdílené definice, ne do jednoho webu.

### 3. Celá webová kolekce je uzamčená

Pak neprojde žádný zápis, ani vytvoření složky. Skript stav přečte a upozorní.

**Odemčení z tohoto skriptu nejde** — je na to potřeba role SharePoint
Administrator a připojení do admin centra, tedy na jinou URL:

```powershell
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -Interactive -ClientId "<guid>"

# stav
Get-PnPTenantSite -Identity "https://<tenant>.sharepoint.com/sites/<web>" -Detailed | Select-Object Url, LockState

# odemknout
Set-PnPTenantSite -Identity "https://<tenant>.sharepoint.com/sites/<web>" -LockState Unlock

# ... spustit skript ...

# vrátit zpět
Set-PnPTenantSite -Identity "https://<tenant>.sharepoint.com/sites/<web>" -LockState ReadOnly
```

> Než web odemknete, zjistěte **proč** je zamčený. Read-only bývá výsledek
> archivace nebo retenční politiky, a v takovém případě je uzamčení záměr, ne
> překážka — odemčení by šlo proti governance a je to rozhodnutí správce
> tenantu, ne toho, kdo spouští skript. Změnu stavu si někam zapište, aby se
> nezapomnělo vrátit.

### 4. Knihovna vyžaduje Check-out

Při `ForceCheckout` je potřeba položku před editací vyzvednout. Skript na to
upozorní; nastavení se vypíná ve verzování knihovny.

## Když to nejde

| Hláška | Co to znamená | Řešení |
|--------|---------------|--------|
| `Pro čtení souboru .xlsx je potřeba modul ImportExcel` | Chybí modul | Nainstalovat, nebo uložit jako CSV — viz výše |
| `Knihovnu '<název>' jsem nenašel` | Špatný `-Library` | Skript vypíše seznam dostupných knihoven i s GUIDy, jeden z nich použijte |
| `V tabulce nejsou sloupce Level1, Level2, ...` | Jiná hlavička | Skript vypíše nalezené sloupce; přejmenovat na `Level1`, `Level2`, … |
| `Chybí ClientId` | Nezadané ClientId | Viz [06-jak-spustit-export.md](06-jak-spustit-export.md) |
| `'<název>' není knihovna dokumentů` (varování) | Cíl je seznam, ne knihovna | Zkontrolovat `-Library`, běh přesto pokračuje |
| `Metadata pro '<cesta>' nelze zapsat` | Sloupec neexistuje nebo má nekompatibilní typ | Zkontrolovat `-MetadataMap`, `-FixedMetadata` a typy sloupců v knihovně |
| Hodnota z `-FixedMetadata` se nezapsala | Sloupec je Choice bez té možnosti, nebo Managed Metadata | Přidat hodnotu mezi možnosti sloupce, u Managed Metadata předat GUID termínu |
| `A positional parameter cannot be found` u `-FixedMetadata` | Název sloupce s mezerou není v uvozovkách | `@{ "CSD Class" = "..." }` |
| `Sloupec '<název>' nelze jednoznačně určit, přeskakuji ho` | Víc sloupců se stejným názvem, nebo název neexistuje | Spustit s `-ListFields` a předat interní název |
| `Název '<X>' odpovídá N sloupcům` | Duplicitní displejové názvy v knihovně | Předat interní název toho správného |
| `Unable to find the specified term. Skipping values for field '<X>'` | Starší verze skriptu; textová hodnota v taxonomy sloupci | Aktuální verze termín dohledá sama, stačí předat název |
| `Termín '<X>' v term setu není` | Jiný název termínu | `-ListTerms <sloupec>` a použít přesný název ze seznamu |
| `Term set <guid> se v Term Store nepodařilo najít` | Chybí přístup na Term Store | Vyžádat právo čtení na Term Store |
| `Termín '<X>' je v term setu víckrát` | Stejný název na víc místech | Předat GUID toho správného |
| `<X> nelze určit, nový nezakládám` | Nejednoznačný název, duplikát by to zhoršil | Předat interní název |
| `Sloupec '<X>' je ReadOnly, zápis by se zahodil` | Sloupec je uzamčený | `-UnlockReadOnlyFields`, viz výše |
| `Sloupec '<X>' je Sealed` | Sloupec je z content typu | Změnit u content typu, ne na knihovně |
| `Celá webová kolekce je v režimu ReadOnly` | Site lock | Odemčení přes admin centrum, viz výše |
| `POZOR: sloupec '<X>' se nepodařilo vrátit na ReadOnly` | Selhalo zamčení zpět | **Vrátit ručně** přes `Set-PnPField -Values @{ReadOnlyField=$true}` |
| `Knihovna vyžaduje Check-out` | ForceCheckout | Vypnout v nastavení verzování knihovny |

## Kam to vede dál

Excel je artefakt byznysu, ale není to verzovatelná šablona — v pull requestu
z něj není vidět, co se změnilo. Přirozený další krok je z tabulky vygenerovat
YAML v `templates/` a provisioning pak řídit z něj, přičemž Excel zůstane
vstupním formulářem pro zadavatele. Návrh schématu:
[templates/project-site/project.example.yml](../templates/project-site/project.example.yml).

Skript zatím **nebyl otestován proti tenantu** — logika skládání cest a metadat
je ověřená simulací proti `Folder_Structure.xlsx`, ale samotné volání PnP
cmdletů ne. Proto ten dry-run.
