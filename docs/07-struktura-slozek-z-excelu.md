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

> Pokud je `CSD` v knihovně už založený jako **Choice**, musí `5.3 Car Series and
> Concept Docs` být jednou z jeho možností, jinak SharePoint zápis odmítne
> (nebo hodnotu zahodí, podle nastavení „Allow fill-in choices"). Pro sloupec
> typu **Managed Metadata** takhle text zapsat nelze — tam by bylo potřeba
> předat GUID termínu.

> Všechny tři sloupce jsou zatím typu Text. Pokud se `Responsible` (A1–A5) má
> vybírat ze seznamu hodnot, patří tam typ Choice, a pokud jde o útvary nebo
> osoby, patří to do Term Store, respektive na typ Person. To je otevřená otázka
> na byznys — viz [01-otazky-pro-byznys.md](01-otazky-pro-byznys.md), blok B.

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
| `Sloupec '<název>' v knihovně neexistuje, přeskakuji ho` | Název nesouhlasí s displejovým ani interním | Zkontrolovat podle výpisu v kroku „Sloupce pro metadata" |

## Kam to vede dál

Excel je artefakt byznysu, ale není to verzovatelná šablona — v pull requestu
z něj není vidět, co se změnilo. Přirozený další krok je z tabulky vygenerovat
YAML v `templates/` a provisioning pak řídit z něj, přičemž Excel zůstane
vstupním formulářem pro zadavatele. Návrh schématu:
[templates/project-site/project.example.yml](../templates/project-site/project.example.yml).

Skript zatím **nebyl otestován proti tenantu** — logika skládání cest a metadat
je ověřená simulací proti `Folder_Structure.xlsx`, ale samotné volání PnP
cmdletů ne. Proto ten dry-run.
