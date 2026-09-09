# Stránky, obrázky a vzhled webu

Postup pro `src/Copy-SitePages.ps1`. Přenese ze vzorového webu to, co ostatní
skripty neumí: stránky s webparty, obrázky na nich, domovskou stránku, vzhled
a regionální nastavení.

Normálně se spouští jako krok `Pages` v
[Setup-ProjectSite.ps1](09-jeden-skript.md), samostatně jen když je potřeba
přenést pouze tohle.

> Skript **nebyl otestován proti tenantu**. Vychází z funkcí, které v původním
> `script.ps1` fungují, ale je to přepis, ne totožný kód.

## Odkud to je

Původní `script.ps1` tohle umí ve funkcích `GetAllPages`, `CopyWebparts2`,
`AddImages`, `Copy-Regionalsettings` a v části `Copy-Navigation`, která přenáší
`WebSettings`. Postup zůstal stejný, protože je osvědčený:

```powershell
Export-PnPPage -Identity Home.aspx -Out export.xml   # ze vzoru
Invoke-PnPSiteTemplate -Path export.xml              # do cíle
```

Proti původnímu skriptu je tady navíc režim náhledu, jedna chybná stránka
nezastaví zbytek a nic se nemaže.

Nepřevzal jsem funkci `CopyWebparts` (bez dvojky) s mapováním německých názvů
webpartů — v `script.ps1` se nikde nevolá, nahradila ji `CopyWebparts2`, která
mapování nepotřebuje, protože jde přes šablonu.

## Co přenáší

| Krok | Co to udělá |
|------|-------------|
| `Pages` | Všechny stránky z knihovny stránek včetně webpartů a jejich nastavení |
| `Images` | Obrázky vložené do webpartů — stáhne ze vzoru a nahraje do cíle |
| `HomePage` | Nastaví stejnou domovskou stránku jako na vzoru |
| `Design` | `WebSettings` (téma, logo) a `HeaderLayout` |
| `Regional` | Regionální nastavení — první den týdne, pracovní hodiny, čísla týdnů |

Výchozí je vše. Jen část přes `-Steps`:

```powershell
-Steps Pages,Images
```

Konkrétní stránky přes `-Pages`:

```powershell
-Pages Home.aspx,Reporting.aspx
```

## Spuštění

Náhled — nic nemění, vypíše které stránky by vytvořil a které přepsal:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\Copy-SitePages.ps1 -SourceSiteUrl "https://<tenant>.sharepoint.com/sites/<vzorovy>" -TargetSiteUrl "https://<tenant>.sharepoint.com/sites/<novy>"
```

Pak to samé s `-Apply`.

## Důležité: stránky se přepisují celé

Tady se od ostatních skriptů chová jinak. Složky, sloupce ani seznamy nikdy
nepřepisujeme — existující se přeskočí. **U stránek to nejde**:
`Invoke-PnPSiteTemplate` stránku nahradí celou.

Znamená to, že opakované spuštění zahodí ruční úpravy, které někdo na cílové
stránce udělal. Náhled proto u každé stránky říká, jestli by ji vytvořil, nebo
přepsal, a když už v cíli stránky jsou, upozorní na to červeně.

Na čerstvě založeném webu je to bez rizika. Na běžícím projektu si to nejdřív
projděte v náhledu.

## Na co si dát pozor

**Obrázky mimo vzorový web se nepřenášejí.** Obrázek uložený ve zdrojovém webu
se stáhne a nahraje do cíle na stejnou relativní cestu. Odkaz na stock fotku
nebo na jiný web zůstane, jak je — přepsat by ho rozbilo.

**Odkazy ve webpartech.** Webpart, který ukazuje na konkrétní seznam, si nese
jeho ID. Po přenesení může v cíli ukazovat na nic, dokud tam seznam nevznikne.
Proto stránky přenášejte **po** seznamech a knihovnách — v
`Setup-ProjectSite.ps1` je krok `Pages` schválně až za `Lists` a `Events`.

**Knihovna stránek se jmenuje podle jazyka.** Skript zkouší `Site Pages`,
`Websiteseiten` i `Stránky webu`, a když ani to nesedí, hledá ji podle typu
knihovny (BaseTemplate 119).

## Když to nejde

| Hláška | Co to znamená | Řešení |
|--------|---------------|--------|
| `Knihovnu se stránkami se nepodařilo najít` | Jiný jazyk nebo typ webu | Ověřit, že vzor má moderní stránky |
| `Stránku '<X>' nelze přenést` | Šablona stránky se neaplikovala | Zkusit `-Pages <X>` samostatně, podrobnost je ve varování |
| `Obrázek '<X>' nelze přenést` | Cílová složka neexistuje | Není blokující, obrázek doplnit ručně |
| `Regionální nastavení nelze přenést` | Handler `RegionalSettings` selhal | Původní `script.ps1` slučuje jen vybrané atributy — viz níže |
| `Vzhled webu nelze přenést` | Jiný typ webu nebo chybí práva | Zkontrolovat, že cíl je stejný typ webu jako vzor |

### Když selže přenos regionálního nastavení

Tady se od `script.ps1` liším. Ten nebere šablonu ze vzoru celou, ale vytáhne si
šablonu z cíle a přepíše v ní jen konkrétní atributy: `ShowWeeks`,
`FirstDayOfWeek`, `WorkDays`, `FirstWeekOfYear`, `WorkDayEndHour`,
`WorkDayStartHour`. Mělo to nejspíš důvod — aplikovat cizí `RegionalSettings`
včetně jazyka a časové zóny může selhat.

Tady se pro jednoduchost přenáší celý handler. Když to nefunguje, použijte na
tenhle jeden krok `script.ps1` s `$IsCopyRegionalSettings = $true` a ostatní
příznaky vypnuté, nebo řekněte a doplním sem stejné selektivní slučování.

## Co pořád zůstává v script.ps1

Ani po tomhle není `script.ps1` zbytečný. Umí věci, které jinde nemáme:

- **kopírování knihoven dokumentů včetně souborů** — `Copy-PnPDocLibs` ve verzi
  2.7 (autor Sergiu Nica),
- **dávkové zpracování více cílových webů** ze `Sites.xml` (`$CopyFromList`),
- **kopírování mezi webem a podwebem** (`$CopyFromMainToSubSite`),
- **nastavení offline dostupnosti** seznamů (`$SetOfflineAvailable`).
