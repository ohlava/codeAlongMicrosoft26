# Kopírování navigace mezi weby

Postup pro `src/Copy-SiteNavigation.ps1`. Přenese navigaci ze vzorového webu na
jiný web, včetně zanoření.

> Skript **nebyl otestován proti tenantu** — na vývojovém stroji není PowerShell.
> Proto ten dry-run.

## Co přenáší

- **QuickLaunch** (levá navigace) a **TopNavigationBar** (horní), nebo jen jednu
  přes `-Location`. Výchozí je obojí.
- Libovolnou hloubku zanoření (strop `-MaxDepth`, výchozí 10). Původní
  `Copy-Navigation` ve `script.ps1` zvládala pevně tři úrovně.
- Příznak externího odkazu.

Odkazy míříci **do zdrojového webu** se přepíšou na cílový web:

```
/sites/Project01/Lists/Tasks   ->   /sites/Project03/Lists/Tasks
```

Odkazy **mimo zdrojový web** se nechávají být — typicky vedou na intranet nebo
do jiné aplikace a přepsat je by je rozbilo.

## Spuštění

### Krok 1 — Dry-run

Nic nemění. Vypíše zdrojovou navigaci jako strom, kolik položek už je v cíli
a kolik odkazů se bude přepisovat. Plán uloží do `export/navigation-plan.csv`.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\Copy-SiteNavigation.ps1 -SourceSiteUrl "https://<tenant>.sharepoint.com/sites/<vzorovy>" -TargetSiteUrl "https://<tenant>.sharepoint.com/sites/<novy>"
```

V CSV je u každé položky `SourceUrl`, `TargetUrl` a `Rewritten` — projděte si
řádky, kde `Rewritten` je `False`, jestli tam opravdu mají zůstat původní odkazy.

### Krok 2 — Přenést

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\Copy-SiteNavigation.ps1 -SourceSiteUrl "https://<tenant>.sharepoint.com/sites/<vzorovy>" -TargetSiteUrl "https://<tenant>.sharepoint.com/sites/<novy>" -Apply
```

Otevřou se **dvě přihlašovací okna** — skript se přihlašuje k oběma webům zvlášť.

## Dva režimy

### Merge (výchozí)

Doplní jen to, co v cíli chybí. Položku se stejným názvem nepřidá znovu, ale
zanoří se do ní a doplní, co chybí uvnitř. Dá se tedy pouštět opakovaně a
navigaci to nezduplikuje.

Ve výstupu je `+` přidáno, `=` už existuje, `.` přeskočeno.

### Replace

```powershell
-Mode Replace -Apply
```

Nejdřív **smaže celou cílovou navigaci** a postaví ji znovu. Je to nevratné a
smaže i položky, které si tam někdo přidal ručně. Používejte jen na čerstvě
založený web.

## Co se nekopíruje

Odkazy, které si SharePoint zakládá sám — `Dokumenty`, `Poznámkový blok`,
`Nedávno použité`, `Stránky webu` — v české, německé i anglické variantě. Jinak
by v cíli vznikly duplicity vedle těch, které tam SharePoint už dal.

Seznam jde přepsat přes `-SkipTitles`:

```powershell
-SkipTitles @("Dokumenty", "Documents")
```

Prázdné pole (`-SkipTitles @()`) zkopíruje vše.

## Co se opravilo proti Copy-Navigation ve script.ps1

Původní funkce funguje, ale má tři vlastnosti, které se sem nepřenesly:

1. **Mazala vždycky.** `Remove-PnPNavigationNode -Force` na všechny cílové
   položky proběhlo při každém spuštění. Tady je to jen `-Mode Replace`.
2. **Pevně tři úrovně zanoření**, napsané jako tři vnořené smyčky. Tady je to
   rekurze bez omezení počtu úrovní.
3. **Chyba v přepisu odkazů.** Volala
   `$Url.replace([Regex]::Escape($SourcePath), [Regex]::Escape($TargetPath))`.
   `[String]::Replace` je ale obyčejná záměna podřetězce, ne regulární výraz —
   `[Regex]::Escape` tam nepatří. U cesty jako `/sites/Project-01` nebo
   `/sites/P.01` by se escapováním změnil hledaný řetězec a záměna by přestala
   fungovat. Tady se používá porovnání prefixu bez ohledu na velikost písmen.

## Když to nejde

| Hláška | Co to znamená | Řešení |
|--------|---------------|--------|
| `Zdrojový a cílový web jsou tentýž` | Stejná URL v obou parametrech | Zkontrolovat `-SourceSiteUrl` a `-TargetSiteUrl` |
| `Položku '<X>' nelze přidat` | Neplatná URL po přepisu, nebo chybí práva | Zkontrolovat `TargetUrl` v `navigation-plan.csv` |
| `Podřízené položky '<X>' nelze přečíst` | Nedostatek práv na zdrojovém webu | Vyžádat aspoň čtení |
| `U '<X>' se nepodařilo zjistit Id` | `Add-PnPNavigationNode` nevrátil objekt | Položka vznikla, ale potomci ne — spustit znovu, merge je doplní |
| Navigace se nepřenesla vůbec | Zdroj má prázdnou navigaci | Ověřit `-Location`; hub navigace se dědí a tady není |

## Na co si dát pozor

**Hub navigace se nepřenáší.** Když web dědí navigaci z hub site, není to
navigace webu a tímto skriptem se nekopíruje — spravuje se na hubu.

**Megamenu a horní navigace u komunikačních webů.** `TopNavigationBar` se
u některých typů webů nepoužívá; pak zůstane prázdná a skript to napíše.

**Odkazy na stránky, které v cíli ještě nejsou.** Navigaci lze přenést dřív než
obsah — odkaz pak vede na neexistující stránku, dokud ji tam něco nedoplní.
Proto navigaci přenášejte jako poslední krok po knihovnách a stránkách.
