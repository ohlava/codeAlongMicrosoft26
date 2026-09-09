# Analýza současného skriptu (script.ps1, v2.6)

Rozbor toho, co dnes používá byznys. Je to nejcennější vstup, který máme — je to
existující specifikace požadavků, jen napsaná v PowerShellu.

## Co to je

**Klonovací skript, ne šablonovací.** Nepracuje s definicí uloženou v gitu.
Připojí se ke dvěma existujícím webům — vzorovému (`$SourcePath`) a cílovému
(`$TargetPath`) — a přenese obsah z prvního do druhého přes PnP PowerShell
s interaktivním přihlášením.

**Cílový web už musí existovat.** Skript ho nezakládá, jen naplňuje. Zakládání
webu je dnes tedy pořád ruční krok.

Konfigurace je pevně napsaná na začátku souboru. Před každým použitím se ručně
přepíše `$TargetPath`. Volitelně umí dávku webů ze `Sites.xml` (`$CopyFromList`).

## Co kopíruje

| Oblast | Jak | Funkce |
|--------|-----|--------|
| Seznamy (listy) | Smaže cílový, vytvoří nový jako `GenericList` | `Copy-SPOLists` |
| Sloupce | Z `SchemaXml` přes `Add-PnPFieldFromXml`; Lookup zvlášť s remapováním | `Copy-SPOLists` |
| Závislosti Lookup polí | Rekurzivně zjistí, které listy musí vzniknout dřív | `Create-ListAndDependencies` |
| Položky seznamů | Včetně příloh, People pickeru (`New-PnPUser`) a Taxonomy hodnot | `Copy-SPOListItems`, `Copy-SPOAttachments` |
| Zobrazení (views) | Včetně `CustomFormatter`, `ViewQuery`, `RowLimit`, default view | `Copy-SPOViews` |
| Stránky a webparty | `Export-PnPPage` → `Invoke-PnPSiteTemplate` | `CopyWebparts2`, `GetAllPages` |
| Obrázky na stránkách | Stáhne a znovu nahraje Banner a Image webparty | `AddImages` |
| Navigace (QuickLaunch) | Smaže cílovou, znovu postaví do 3 úrovní | `Copy-Navigation` |
| Vzhled webu | `Get-PnPSiteTemplate -Handlers WebSettings` + `HeaderLayout` | `Copy-Navigation` |
| Regionální nastavení | Sloučí zdrojové atributy do cílové šablony | `Copy-Regionalsettings` |
| Domovská stránka | `Set-PnPHomePage` | `GetAllPages` |
| Offline dostupnost | `ExcludeFromOfflineClient` podle `$SetOfflineAvailable` | `Copy-SPOLists` |

## Co nekopíruje — a to je pro nás nejdůležitější

**Knihovny dokumentů se nekopírují vůbec.** Filtr na řádku 430 je:

```powershell
$_.RootFolder.ServerRelativeUrl -like "$SourcePath/lists/*"
```

Knihovna dokumentů má root folder `/sites/Project01/Dokumente`, nikoli
`/sites/Project01/Lists/...`. Filtrem tedy neprojde. Zpracují se jen generické
seznamy.

Dál chybí:

- **oprávnění** — žádná práce se skupinami ani rolemi,
- **content types**,
- **term store / spravovaná metadata** — Taxonomy hodnoty se kopírují jako
  `TermGuid`, ale samotné termíny musí v cíli existovat,
- **vytvoření webu**,
- **cokoli k životnímu cyklu a archivaci.**

Odtud plyne nejdůležitější otázka na Markétu: **jak dnes vznikají knihovny
dokumentů?** Ruční klikání, jiný skript, nebo se cílový web zakládá jako kopie
existujícího webu ze SharePointu, takže knihovny už v něm jsou?

## Rizika, na která je dobré upozornit

**Skript maže.** `Remove-PnPList -Force` na řádku 221 smaže cílový seznam včetně
dat, než ho vytvoří znovu. Totéž pro views, stránky a navigační uzly. Proti
běžícímu projektu je to nevratná ztráta obsahu. Je to nástroj na jedno použití
proti prázdnému webu, i když to z něj není vidět.

Proto v naší architektuře stojí zásada „nikdy nemazat" a dry-run jako výchozí
režim ([03-architektura.md](03-architektura.md)).

**Není idempotentní.** Druhé spuštění nedoplní chybějící — přepíše všechno.
Nedá se použít na doaplikování změny šablony do běžících projektů.

**Není verzovaný.** Hlavička souboru žádá „Bitte bei Änderungen am Code diese
kurz dokumentieren", tedy ruční changelog v komentáři. Přesně tuhle bolest má
use case vyřešit přesunem do gitu.

**Ruční konfigurace.** URL cíle se přepisuje v kódu. Snadná záměna cílového webu
v kombinaci s mazáním je ošklivá kombinace.

## Konkrétní chyby, které jsem v kódu našel

Pokud se skript bude přepisovat, tyto stojí za pozornost. Neověřoval jsem je
spuštěním, jde o čtení kódu.

1. **Remapování Lookup hodnot nefunguje.** `$global:Dataa` se plní záznamy
   `@{Title; Id; Items = @()}` (řádek 461), ale do `Items` nikdy nic nepřidá.
   `Copy-SPOListItems` přitom hledá páry `OldId`/`NewId` právě v `$CheckList.Items`
   (řádky 154–160). `$NewLookupIDs` proto zůstane prázdné a hodnoty Lookup polí
   se u zkopírovaných položek ztratí.

2. **`$CheckUser` se přepisuje uvnitř smyčky.** Na řádku 105 se načte seznam všech
   uživatelů, ale řádek 124 (`$CheckUser = $CheckUser | Where-Object Email -eq $Email`)
   ho zúží na jednoho. Od druhé iterace se tedy hledá v jednoprvkovém seznamu a
   `New-PnPUser` se volá zbytečně nebo se přiřadí špatné `LookupId`.

3. **Řádky 2, 5 a 6 nejsou zakomentované.** `Bitte beim Änderungen...`,
   `add Banner-copy function` a `Config` se PowerShell pokusí spustit jako
   příkazy a vypíše chyby. `Clear-Host` na řádku 968 je pak smaže z obrazovky,
   takže si toho nikdo nevšimne. Chybí `#`.

4. **Mrtvý kód.** `$CorrectedSchemaXml` (řádek 282) se spočítá, ale
   `Add-PnPFieldFromXml` dostane neupravený `$ColumnSchemaXml.OuterXml`. Oprava
   `&amp;` se tedy neaplikuje. Funkce `CopyWebparts` (řádky 492–634) se nikde
   nevolá — nahradila ji `CopyWebparts2`.

5. **`$CopyCount` je zavádějící.** Používá se jako `$counter -lt $CopyCount`, takže
   zpracuje `$CopyCount - 1` položek. Komentář to přiznává, ale je to past.

6. **Skrytá podmínka na vzorový web.** Řádek 444 vyžaduje, aby se název seznamu
   rovnal jeho zobrazovanému názvu, jinak vypíše chybu (`exit` je zakomentovaný,
   takže pokračuje dál). Vzorový web tedy musí být pojmenovaný podle nepsaného
   pravidla.

7. **`Start-Sleep -s 3` na každý seznam** je obcházení throttlingu. U webu
   s 20 seznamy to je minuta čekání navíc.

## Co si z toho vzít pro use case

Skript je funkční důkaz, že tenhle přístup jde, a zároveň ukázka, proč
nestačí — chybí mu knihovny dokumentů, oprávnění, idempotence, verzování a
životní cyklus. To je přesně rozsah našeho zadání.

Praktický důsledek pro plán: **nezahazovat ho.** Obsahuje vyřešené detaily, které
by nás jinak zdržely — mapování německých názvů webpartů na jejich typy
(`Einbetten` → `ContentEmbed`, `Dokumentbibliothek` → `MyDocuments`), obcházení
`CustomFormatter` s `< `, dvoufázový pokus o přidání navigačního uzlu.
Tyhle věci někdo vyladil provozem a stojí za převzetí.
