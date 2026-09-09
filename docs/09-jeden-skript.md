# Jeden skript pro přípravu projektového webu

`src/Setup-ProjectSite.ps1` je jediný příkaz, který je potřeba spouštět. Sám si
zjistí, co má dělat, z konfigurace, a zavolá dílčí skripty.

## Nastavení, které se udělá jednou

ClientId a adresa vzorového webu se nezadávají do příkazu — jsou v konfiguraci.

```powershell
Copy-Item .\config\settings.example.json .\config\settings.json
notepad .\config\settings.json
```

Vyplnit:

```json
{
  "clientId": "<sem to dlouhe cislo z vaseho script.ps1, radek 47>",
  "sourceSiteUrl": "https://<tenant>.sharepoint.com/sites/<vzorovy-web>",
  "library": "Dokumenty",
  "folderStructureFile": "Folder_Structure.xlsx",
  "fixedMetadata": {
    "CSD_x0020_Class": "5.3 Car Series and Concept Docs"
  },
  "setDefaultColumnValues": true
}
```

**`config/settings.json` je v `.gitignore`**, takže se nedostane do repozitáře.
Šablona `config/settings.example.json` v něm zůstává, ale jen s vymyšlenými
hodnotami.

ClientId je nejjednodušší vzít ze funkční kopie `script.ps1`, kde je v proměnné
`$ClientId`. Je to stejná hodnota.

## Běžné použití

### Podívat se, co by se stalo

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\Setup-ProjectSite.ps1 -TargetSiteUrl "https://<tenant>.sharepoint.com/sites/<novy-web>"
```

### Připravit web

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\Setup-ProjectSite.ps1 -TargetSiteUrl "https://<tenant>.sharepoint.com/sites/<novy-web>" -Apply
```

Tohle je ten běžný případ. Udělá:

1. **složky** z `Folder_Structure.xlsx` včetně metadat z Excelu a konstantní
   hodnoty CSD Class,
2. **výchozí hodnoty sloupců**, aby CSD Class dostal každý nově nahraný soubor,
3. **seznamy** ze vzorového webu — jen strukturu, bez položek,
4. **kalendáře** (Events) ze vzorového webu — také bez položek,
5. **stránky, obrázky, vzhled a regionální nastavení** ze vzorového webu.

### Včetně obsahu seznamů a kalendářů

```powershell
... -TargetSiteUrl "..." -WithData -Apply
```

`-WithData` přenese i položky seznamů a kalendářů, včetně příloh.

### Jen část

```powershell
... -TargetSiteUrl "..." -Steps Folders -Apply
... -TargetSiteUrl "..." -Steps Lists,Events -Apply
```

Možnosti: `Folders`, `Lists`, `Events`, `Pages`, `DefaultValues`, `Navigation`.
Výchozí je `Folders, Lists, Events, Pages, DefaultValues` — `Navigation` je
potřeba vyžádat výslovně.

Pořadí kroků je dané a nezávisí na tom, jak se vypíšou ve `-Steps`: nejdřív
složky, pak seznamy a kalendáře, teprve pak stránky. Webparty na stránkách totiž
odkazují na seznamy, které musí existovat dřív.

## Metadata na složkách versus na souborech

Tohle jsou dvě různé věci a skript dělá obě:

**Metadata na složce** se zapisují na složku jako na položku seznamu. Vyplní se
z Excelu (`Responsible`, `NameEnglish`, `NameGerman`) plus konstantní CSD Class.

**Výchozí hodnota sloupce v knihovně** (`Set-PnPDefaultColumnValue`) způsobí, že
hodnotu dostane **každý nově nahraný soubor**. Metadata zapsaná na složku se na
soubory v ní samy nepřenesou — proto je potřeba i tenhle druhý krok.

Vypnout ho jde přes `"setDefaultColumnValues": false` v konfiguraci.

Zpětně už nahrané soubory to neovlivní — výchozí hodnota platí pro nově
přidávané.

## Který sloupec CSD Class

V konfiguraci je zatím `CSD_x0020_Class`, což je textový sloupec — ten funguje
bez dalšího zařizování.

Správným cílem je ale `RevIMBCS` (`Třída KSU`), protože hodnoty typu
"5.3 Car Series and Concept Docs" jsou termíny z firemního klasifikačního
schématu, ne text. Ten sloupec je ale `ReadOnly`, protože patří k records
managementu. Přepnutí je pak jen změna v konfiguraci:

```json
"fixedMetadata": { "RevIMBCS": "5.3 Car Series and Concept Docs" }
```

`New-FolderStructure.ps1` si termín v Term Store dohledá sám. Zápis do
`RevIMBCS` ale vyžaduje odemčení sloupce a **souhlas správce records
managementu** — viz [decisions.md](decisions.md) a
[07-struktura-slozek-z-excelu.md](07-struktura-slozek-z-excelu.md).

## Co která část dělá

| Krok | Skript | Náhled bez -Apply |
|------|--------|-------------------|
| Folders | [New-FolderStructure.ps1](../src/New-FolderStructure.ps1) | ano, vypíše plán složek |
| DefaultValues | přímo v Setup | ano, vypíše, co by nastavil |
| Lists | [Copy-SharePointLists.ps1](../src/Copy-SharePointLists.ps1) | **ne**, jen řekne, že by se spustil |
| Events | [Copy-SharePointEvents.ps1](../src/Copy-SharePointEvents.ps1) | **ne**, jen řekne, že by se spustil |
| Pages | [Copy-SitePages.ps1](../src/Copy-SitePages.ps1) | ano, vypíše stránky |
| Navigation | [Copy-SiteNavigation.ps1](../src/Copy-SiteNavigation.ps1) | ano, vypíše strom |

`Copy-SharePointLists.ps1` a `Copy-SharePointEvents.ps1` (autor Sergiu Nica)
režim náhledu nemají. V náhledu se proto jen napíše, co by se spustilo. Samy
existující seznamy nepřepisují — chybějící zakládají a hlásí, co už bylo.

## Vlastnosti, na které je dobré se spolehnout

**Dá se pouštět opakovaně.** Složky, sloupce i seznamy, které už existují, se
přeskočí. Když se do Excelu přidá složka, stačí spustit znovu.

**Jeden neúspěšný krok nezastaví ostatní.** Selže-li třeba přenos kalendářů,
skript to zapíše do souhrnu a pokračuje. Na konci je přehled, co prošlo a co ne.

**Přihlášení jednou.** Dílčí skripty se přihlašují každý sám, ale běží v jednom
procesu PowerShellu, takže se přihlašovací okno normálně objeví jen u prvního.

## Když to nejde

| Hláška | Řešení |
|--------|--------|
| `Chybí konfigurace: config/settings.json` | Zkopírovat šablonu, viz úvod |
| `V konfiguraci není vyplněné clientId` | Doplnit do `config/settings.json` |
| `Kroky ... potřebují vzorový web` | Doplnit `sourceSiteUrl`, nebo použít `-Steps Folders` |
| `Chybí skript ... ve složce src/` | Chybí soubory z repozitáře, stáhnout znovu |
| `Konfiguraci nelze přečíst - není to platný JSON` | Nejčastěji chybějící nebo přebývající čárka |
| Krok skončil `CHYBA:` | Podrobnosti jsou ve výpisu nad souhrnem a v `export/` |

## Co skript nedělá

**Nezakládá web.** Cílový web musí existovat. Zakládání je zatím ruční krok.

**Nepřenáší knihovny dokumentů ze vzorového webu.** Struktura složek se bere
z Excelu. Kopírování knihoven včetně souborů řeší `Copy-PnPDocLibs` ve
`script.ps1` (autor Sergiu Nica) — zatím na vlastní branchi.

**Nepřenáší oprávnění.** Skupiny ani role se neřeší vůbec.

Stránky, webparty a vzhled přenáší krok `Pages` — podrobněji
v [10-stranky-a-vzhled.md](10-stranky-a-vzhled.md). Pozor, **stránky se
přepisují celé**, takže opakované spuštění zahodí ruční úpravy na cílových
stránkách. Jediný krok, který nepracuje přírůstkově.

**Navigace není seznam.** Levá navigace se neukládá jako list, ale jako
navigační struktura webu, takže ji skripty na kopírování seznamů nepřenesou.
Řeší to `Copy-SiteNavigation.ps1` a je proto ve `-Steps` zvlášť.
