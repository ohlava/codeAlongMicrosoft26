# Jak spustit export vzorového webu

Postup pro `src/Export-SiteInventory.ps1`. Napsáno tak, aby to zvládl i někdo,
kdo PowerShell běžně nepoužívá. Skript **pouze čte** — do SharePointu nic
nezapisuje ani nemaže, takže se nedá nic rozbít.

> Skript zatím **nebyl otestovaný** — psal se na macOS, kde PowerShell není.
> Při prvním spuštění počítejte s tím, že něco bude potřeba doladit. Chybové
> hlášky a jejich řešení jsou na konci.

## Co je potřeba mít

| Co | Jak zjistit, že to mám |
|----|------------------------|
| Přístup ke vzorovému webu | Web se otevře v prohlížeči a je vidět obsah |
| PowerShell | Ve Windows menu Start napsat `PowerShell` |
| Modul `PnP.PowerShell` | `Get-Module PnP.PowerShell -ListAvailable` vypíše verzi |
| **ClientId** aplikace | Viz níže — nejčastější důvod, proč to nejde |

### ClientId

Přihlášení k SharePointu z PowerShellu potřebuje registrovanou aplikaci v Entra ID
a její **ClientId** (dlouhé číslo ve tvaru `a1b2c3d4-...`).

**Nejrychlejší cesta: použít stejné ClientId, jaké má stávající kopírovací
skript.** Ve své funkční kopii `script.ps1` je na řádku 47 v proměnné `$ClientId`.
Ta hodnota funguje i tady — je to stejný způsob přihlášení.

Pokud ji nemáte, vyžádejte si ClientId od správce M365. Pro tento skript stačí
oprávnění pro čtení, žádná nová registrace není potřeba.

## Krok 1 — Získat soubory

Ve Windows Průzkumníku otevřete složku, kam se má stáhnout repozitář, a v ní
otevřete PowerShell (v adresním řádku Průzkumníka napsat `powershell` a Enter).

```powershell
git clone https://github.com/ohlava/codeAlongMicrosoft26.git
cd codeAlongMicrosoft26
```

Když `git` není k dispozici, stáhněte repozitář jako ZIP přes tlačítko **Code →
Download ZIP** na GitHubu a rozbalte ho.

## Krok 2 — Nainstalovat modul PnP.PowerShell

Přeskočte, pokud `Get-Module PnP.PowerShell -ListAvailable` už něco vypíše.

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

`-Scope CurrentUser` znamená, že se instaluje jen pro vás a **není potřeba právo
správce**. Na dotaz, jestli důvěřovat repozitáři PSGallery, odpovězte `A`
(Yes to All).

## Krok 3 — Povolit spuštění skriptu

Windows ve výchozím nastavení nespustí skripty stažené z internetu.

**Nejjednodušší je krok 3 přeskočit** a v kroku 4 použít variantu B, která se
vejde do jednoho příkazu a v systému nic nemění.

Kdo chce povolení nastavit pro celé okno, zadá tyto dva příkazy. **Každý zvlášť,
Enter po každém řádku.** Když se slepí do jednoho, PowerShell zahlásí
`a positional argument cannot be found that accepts argument Unblock-File` —
`Set-ExecutionPolicy` už oba své pozicionální parametry dostal pojmenované,
takže `Unblock-File` nemá kam dát.

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

```powershell
Unblock-File .\src\Export-SiteInventory.ps1
```

Platí to **jen pro aktuálně otevřené okno** PowerShellu, nikde nic trvale
nemění. Když je potřeba mít oba příkazy na jednom řádku, oddělte je středníkem:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass; Unblock-File .\src\Export-SiteInventory.ps1
```

Pokud `Set-ExecutionPolicy` zahlásí, že nastavení **přepisuje zásada skupiny**
(Group Policy), nelze to obejít — použijte variantu B v kroku 4.

## Krok 4 — Spustit export

Nahraďte URL adresou vzorového projektového webu a `<guid>` svým ClientId.

### Varianta A — po kroku 3

```powershell
.\src\Export-SiteInventory.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<vzorovy-web>" -ClientId "<guid>"
```

### Varianta B — bez kroku 3, jeden příkaz (doporučeno)

Povolení platí jen pro tento jeden běh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\Export-SiteInventory.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<vzorovy-web>" -ClientId "<guid>"
```

Pokud používáte PowerShell 7, nahraďte `powershell.exe` za `pwsh`.

> Příkazy jsou schválně na jednom dlouhém řádku. Zalomení přes zpětný apostrof
> funguje, ale při kopírování do terminálu se často rozbije.

Otevře se přihlašovací okno prohlížeče — přihlaste se svým firemním účtem,
stejně jako u stávajícího skriptu.

Aby se ClientId nemuselo psát pokaždé, jde ho nastavit do proměnné prostředí
a pak ho v příkazu vynechat. Opět každý příkaz zvlášť:

```powershell
$env:PNP_CLIENT_ID = "<guid>"
```

```powershell
.\src\Export-SiteInventory.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<vzorovy-web>"
```

Export webu s dvaceti seznamy trvá řádově jednotky minut.

## Krok 5 — Co z toho vypadne

Vše ve složce `export/`:

| Soubor | Co v něm je | Otevřít čím |
|--------|-------------|-------------|
| `inventory.csv` | Přehled všech seznamů **a knihoven dokumentů** — název, typ, URL, počet položek, verzování | Excel |
| `inventory.json` | Totéž pro další zpracování | Poznámkový blok |
| `fields/<název>.csv` | Sloupce daného seznamu nebo knihovny — název, typ, povinnost | Excel |
| `folders/<název>.txt` | Strom složek v knihovně dokumentů | Poznámkový blok |
| `site-template.xml` | Úplná PnP šablona webu, jde ji znovu aplikovat | — |
| `site-script.json` | SharePoint Site Script, čitelnější varianta | Poznámkový blok |

**Nejužitečnější je `inventory.csv` a `folders/`.** To je odpověď na otázku, co
vzorový web skutečně obsahuje — a hlavně které knihovny dokumentů, protože ty
stávající kopírovací skript vůbec nepřenáší (viz [05-analyza-skriptu.md](05-analyza-skriptu.md)).

Složka `export/` je v `.gitignore`, takže se sama nedostane do repozitáře. Než
z ní něco commitnete, projděte to — může obsahovat firemní data.

## Když to nejde

| Hláška | Co to znamená | Řešení |
|--------|---------------|--------|
| `a positional argument cannot be found that accepts argument Unblock-File` | Dva příkazy se slepily do jednoho řádku | Zadat každý zvlášť, nebo oddělit středníkem — viz krok 3 |
| `Modul PnP.PowerShell není nainstalovaný` | Chybí modul | Krok 2 |
| `Chybí ClientId aplikace` | Nezadané ClientId | Doplnit `-ClientId`, viz výše |
| `... nelze načíst, protože spouštění skriptů je v tomto systému zakázáno` | Execution policy | Krok 3 |
| `AADSTS65001` nebo `needs admin consent` | Aplikace nemá schválená oprávnění | Vyžádat u správce M365 |
| `Access denied` / `403` | Chybí přístup na web | Vyžádat aspoň čtení na vzorový web |
| `Seznam <název> neexistuje na serveru` / `List does not exist at site with URL` | Lokalizovaný název knihovny nešel dohledat | Opraveno — skript adresuje seznamy GUIDem. Stáhněte si aktuální verzi skriptu |
| `The term 'Get-PnPSiteScriptFromWeb' is not recognized` | Starší verze modulu | Přidat `-SkipSiteScript`, zbytek exportu proběhne |
| `Site Script nelze vyexportovat` (varování) | Jiné názvy parametrů v této verzi PnP | Není blokující, export pokračuje. Ověřit `Get-Help Get-PnPSiteScriptFromWeb -Full` |
| `PnP šablonu nelze vyexportovat` (varování) | Totéž pro PnP šablonu | Není blokující. `Get-Help Get-PnPSiteTemplate -Full` |
| Skript spadne v polovině | Cokoli jiného | Co už je v `export/`, je použitelné. Zbytek jde přeskočit přes `-SkipPnPTemplate -SkipSiteScript` |

Kdyby cokoli z toho zdržovalo, existuje jednodušší varianta: **spustí to někdo
z technického týmu.** Přihlášení je interaktivní, takže stačí, aby ten člověk
dostal právo čtení na vzorový web — nic dalšího od Markéty není potřeba.

## Verze PowerShellu

`PnP.PowerShell` 2.x vyžaduje PowerShell 7. Starší `PnP.PowerShell` 1.x funguje
i ve Windows PowerShellu 5.1, který je ve Windows předinstalovaný. Skript je
napsaný tak, aby fungoval v obou.

Zjištění verzí:

```powershell
$PSVersionTable.PSVersion
Get-Module PnP.PowerShell -ListAvailable | Select-Object Name, Version
```

Pokud `$PSVersionTable.PSVersion` ukáže 5.1 a modul chybí, je jednodušší
nainstalovat starší verzi modulu než nový PowerShell:

```powershell
Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser
```
