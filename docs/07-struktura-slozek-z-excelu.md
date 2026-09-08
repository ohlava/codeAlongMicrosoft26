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
| `Metadata pro '<cesta>' nelze zapsat` | Sloupec neexistuje nebo má nekompatibilní typ | Zkontrolovat `-MetadataMap` a typy sloupců v knihovně |

## Kam to vede dál

Excel je artefakt byznysu, ale není to verzovatelná šablona — v pull requestu
z něj není vidět, co se změnilo. Přirozený další krok je z tabulky vygenerovat
YAML v `templates/` a provisioning pak řídit z něj, přičemž Excel zůstane
vstupním formulářem pro zadavatele. Návrh schématu:
[templates/project-site/project.example.yml](../templates/project-site/project.example.yml).

Skript zatím **nebyl otestován proti tenantu** — logika skládání cest a metadat
je ověřená simulací proti `Folder_Structure.xlsx`, ale samotné volání PnP
cmdletů ne. Proto ten dry-run.
