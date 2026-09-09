# Rozhodnutí

Append-only log. Nikdy nepřepisujte starší záznam — pokud se rozhodnutí změní,
připište nové s odkazem na to překonané. Formát: datum, kdo, co, proč.

Kategorie: `TECH` (technologická volba), `SCOPE` (rozsah), `BIZ` (potvrzeno
byznysem), `ENV` (prostředí a přístupy).

---

## 2026-09-08 | TECH | Předpoklad: PnP PowerShell + YAML + GitHub Actions

Rozhodl: Ondřej Hlava (návrh před akcí, k potvrzení týmem)

Provisioning engine PnP PowerShell, definice projektů ve vlastním YAML schématu,
orchestrace GitHub Actions.

Proč: PnP má nejlepší pokrytí SharePointu; vlastní YAML je čitelné pro byznys, což
je hlavní přínos verzování šablon v gitu; GitHub Actions je explicitně v zadání.

Alternativy zvážené a zamítnuté: Site Designs (nestačí na oprávnění a lifecycle),
PnP XML šablony přímo v gitu (nečitelný diff v PR).

---


## 2026-09-08 | SCOPE | Zjištění z rozboru současného skriptu

Zdroj: [05-analyza-skriptu.md](05-analyza-skriptu.md)

Skript, který byznys dnes používá, je klonovač webu, ne šablonovač. Cílový web
musí existovat předem a kopírují se do něj jen generické seznamy — filtr
`RootFolder.ServerRelativeUrl -like "$SourcePath/lists/*"` vylučuje knihovny
dokumentů. Nekopíruje oprávnění, content types, ani nic k životnímu cyklu.
Není idempotentní a před kopírováním cílové seznamy maže.

Důsledky pro rozsah:

- Knihovny dokumentů jsou v zadání zmíněné jako první, ale současné řešení je
  vůbec neřeší. Zjistit u byznysu, jak dnes vznikají.
- Mapování německých názvů webpartů a obcházení `CustomFormatter` ze skriptu
  přebrat — je to vyladěné provozem.
- Naše zásada „nikdy nemazat" a dry-run je reakce na to, že současný skript maže
  cílové seznamy včetně dat.

---

## 2026-09-08 | TECH | Prvním krokem je read-only export vzorového webu

Rozhodl: Ondřej Hlava (návrh)

`src/Export-SiteInventory.ps1` vyexportuje strukturu existujícího webu
(inventář seznamů a knihoven, sloupce, strom složek, PnP šablona, Site Script).
Pouze čte.

Proč: ruční popis struktury je vždy neúplný a export je zároveň startovní bod pro
naši šablonu. Read-only skript také nepotřebuje tenant admina, jen práva
vlastníka webu, takže se dá pustit první den bez čekání na schvalování.

Stav: napsáno, **neotestováno** — na vývojovém stroji není PowerShell,
syntaxi a názvy parametrů PnP cmdletů je potřeba ověřit na Windows.

---

## 2026-09-08 | BIZ | V cílové knihovně běží records management (RevIM*)

Zdroj: výpis sloupců knihovny "Dokumenty" přes `-ListFields`.

Knihovna obsahuje sadu sloupců s prefixem `RevIM*`, které patří nějakému řešení
pro správu záznamů: `Třída KSU` (`RevIMBCS`, TaxonomyFieldType, ReadOnly),
`CSD Record Declaration Time`, `Datum smazání`, `Datum spouštěcí události`,
`Poznámky spouštěcí události`, `Vlastník dokumentu`. Dále `LegalHold`
(TaxonomyFieldTypeMulti, Sealed) a `Popisek pro uchovávání informací`
(retention label).

Důsledky:

- Klasifikace dokumentů už v tenantu existuje jako **spravovaná metadata**, ne
  jako volný text. Hodnoty typu "5.3 Car Series and Concept Docs" jsou nejspíš
  termíny z Term Store, ne řetězce.
- `RevIMBCS` je ReadOnly a zapisuje do něj to řešení, ne my. Plnit ho vlastním
  skriptem je zásah do records managementu a chce to potvrdit od správce.
- Skrytý sloupec `CSD Class_0` (`i0f84bba906045b4af568ee102a52dcb`, Note) je
  textový společník taxonomy pole, což naznačuje, že `RevIMBCS` se dřív
  jmenoval "CSD Class".

Otevřená otázka na byznys: má se "CSD Class" plnit jako **prostý text** vlastním
sloupcem, nebo jde o **existující klasifikaci** v Term Store, kterou má nastavovat
records management? Rozhoduje to, jestli náš skript metadata zapisuje, nebo se
jich nemá dotýkat.

---

## 2026-09-08 | BIZ | "CSD Class" je termín z firemního klasifikačního schématu

Potvrdil: Ondřej Hlava po konzultaci s byznysem.

Hodnoty typu "5.3 Car Series and Concept Docs" jsou termíny z interního
klasifikačního schématu v Term Store, ne volný text. Správným cílem je proto
sloupec se spravovanými metadaty `Třída KSU` (`RevIMBCS`), nikoli textový
sloupec.

Důsledky:

- Textové sloupce `CSD`, `CSD_x0020_Class`, `CSD_x0020_Class0` a
  `CSD_x0020_Class1` v knihovně "Dokumenty" jsou omyl a patří ke smazání.
- `src/New-FolderStructure.ps1` umí dohledat termín podle názvu v term setu, ke
  kterému je sloupec připojený, a zapsat jeho GUID. Diagnostika: `-ListTerms`.
- `RevIMBCS` je `ReadOnly`, protože patří k records managementu. Zápis do něj
  vyžaduje `-UnlockReadOnlyFields` a **předchozí souhlas správce** records
  managementu. Zatím nepotvrzeno.
- Pro provisioning to znamená, že šablona projektu musí odkazovat na termíny
  Term Store, ne nést textové hodnoty. To je argument pro Content Type Gallery
  a sdílené sloupce, viz [02-technologie.md](02-technologie.md), vrstva 5.

---


## 2026-09-09 | TECH | Jeden vstupní bod, dílčí skripty zůstávají oddělené

Rozhodl: Ondřej Hlava

`src/Setup-ProjectSite.ps1` je jediný příkaz, který spouští byznys. Nevolá
vlastní logiku — obaluje `New-FolderStructure.ps1`,
`Copy-SharePointLists.ps1`, `Copy-SharePointEvents.ps1` a
`Copy-SiteNavigation.ps1`. Nastavení, které se nemění (ClientId, vzorový web,
knihovna, hodnota CSD Class), je v `config/settings.json`, který je v
`.gitignore`; šablona `config/settings.example.json` je verzovaná.

Proč obal a ne přepis do jednoho souboru: `Copy-SharePointLists.ps1` a
`Copy-SharePointEvents.ps1` mají po ~1300 řádcích, fungují a mají vlastního
autora. Slití do jednoho skriptu by zahodilo odladěné detaily a rozbilo
vlastnictví kódu podle modulů.

Důsledek pro rozsah: dílčí skripty si každý drží svého vlastníka a rozhraní
mezi nimi je volání s parametry, ne společný stav.

---

## 2026-09-09 | TECH | CSD Class se nastavuje dvakrát, na složku i jako výchozí hodnota

Rozhodl: Ondřej Hlava

Metadata zapsaná na složku platí pro složku, ne pro soubory v ní. Aby hodnotu
dostal každý nově nahraný soubor, nastavuje `Setup-ProjectSite.ps1` navíc
výchozí hodnotu sloupce v knihovně přes `Set-PnPDefaultColumnValue` (krok
`DefaultValues`).

Nedotýká se souborů, které už v knihovně jsou — výchozí hodnota platí pro nově
přidávané. Doplnění metadat u existujících souborů zatím není řešené.

---

## 2026-09-09 | TECH | Do RevIMBCS se zapisuje přes CSOM, odemykání není potřeba

Rozhodl: Ondřej Hlava, na základě `Set-CsdClass.ps1` od Sergiu Nicy
(branch `Set-CSD-Class`).

Sloupec `Třída KSU` / `CSD Class` (`RevIMBCS`) je `TaxonomyFieldType` označený
`ReadOnly`. `Set-PnPListItem` do něj hodnotu zahodí, ale CSOM projde:

    $value = New-Object Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue
    $value.Label = <nazev terminu>; $value.TermGuid = <guid>; $value.WssId = -1
    $field.SetFieldValueByValue($item, $value)
    $item.Update()
    # po vsech zapisech: Invoke-PnPQuery

Přebráno do `New-FolderStructure.ps1` (metadata na složkách) a
`Set-CsdClass.ps1` (výchozí hodnota sloupce a existující soubory).

Termíny se čtou přímo z `TermSetId` sloupce přes `TaxonomySession`, takže se
nemusí procházet skupiny Term Store — nahradilo to původní hledání skupinou po
skupině.

Důsledky:

- `-UnlockReadOnlyFields` je potřeba už jen pro obyčejné ReadOnly sloupce, ne
  pro spravovaná metadata.
- Přepnutí cíle z textového `CSD_x0020_Class` na `RevIMBCS` je jen změna
  v `config/settings.json`.
- **Zůstává nepotvrzené**, jestli se do `RevIMBCS` smí zapisovat. Že to
  technicky jde, není souhlas správce records managementu.

---

## 2026-09-09 | TECH | script.ps1 přesunut do src/ a převeden na parametry

Rozhodl: Ondřej Hlava, se souhlasem vlastníka use casu.

`script.ps1` byl v korenu repozitáře a konfiguroval se editací hlavičky. Je
teď v `src/script.ps1` jako verze 2.8 s parametry (`-SiteDomain`, `-SourcePath`,
`-TargetPath`, `-ClientId`, `-CopyCount`, příznaky `-IsCopy*`), takže ho
`Setup-ProjectSite.ps1` volá jako krok `TemplateClone` bez generování upravené
kopie.

Původní hodnoty zůstaly jako výchozí, takže spuštění bez parametrů se chová
jako dřív. Navíc je vypnutý `Clear-Host`, který mazal výpis volajícího skriptu,
a zakomentované řádky 2, 5 a 6, které nebyly komentáře a PowerShell je zkoušel
spustit jako příkazy.

Riziko: branche `add-library-copy-function-with-folder-structure` a
`copy-listContent` mění `script.ps1` v korenu, takže při merge vznikne konflikt
kvůli přesunu. Řešit spolu se Sergiuem, verze 2.7 s `Copy-PnPDocLibs` má být
zachovaná.

---

## 2026-09-09 | TECH | Termíny se párují podle čísla na začátku názvu

Zdroj: běh proti webu TESTE. Term set sloupce `RevIMBCS` má 74 termínů, ale
`5.3 Car Series and Concept Docs` mezi nimi nebyl - text za číslem se
v Term Store liší formulací.

Pořadí hledání je teď: GUID, přesný název, **shoda čísla na začátku**, jeden
název je začátkem druhého. Číslo je v klasifikačním schématu stabilní, text ne.
Když se hodnota najde jinak než přesnou shodou, skript to vypíše.

Předchozí implementace porovnávala číslo termínu s celým zadaným řetězcem,
takže `5.3` proti `5.3 Car Series and Concept Docs` nikdy nesedlo.

---

## 2026-09-09 | TECH | Vizuál webu se skládá z několika nezávislých přenosů

Zdroj: běh proti webu TESTE - stránky se přenesly, ale web nevypadal jako vzor.

Zjištění:

- **Barevné téma není v handleru `WebSettings`.** Nastavuje se zvlášť přes
  `Set-PnPWebTheme`. Vlastní téma musí být registrované v tenantu.
- **`HeaderEmphasis`, `MegaMenuEnabled` a `QuickLaunchEnabled`** také nejsou
  ve `WebSettings` a jsou dobře vidět. Doplněno do kroku `Pages`.
- **Vzorový web plní navigaci ze seznamu** `Navigation`, dlaždice ze seznamu
  `Hyperlinks`. Bez `-WithData` vzniknou prázdné seznamy a stránka vypadá
  poloprázdná. Skript na to teď upozorní.
- **Krok `Navigation` nebyl ve výchozí sadě**, takže se levá navigace webu
  nepřenášela. Doplněn do výchozí sady.

---

## 2026-09-09 | BIZ | Sloupec RevIMBCS se v knihovně popisuje jako "CSD Class"

Zadal: vlastník use casu.

V cílové knihovně se sloupec `RevIMBCS` zobrazuje jako `Třída KSU`. Má se
jmenovat `CSD Class`. Řeší se v `config/settings.json`:

    "fieldTitles": { "RevIMBCS": "CSD Class" }

Přejmenování platí jen pro tu knihovnu - v Term Store ani na jiných webech se
nic nemění.

---

## 2026-09-09 | TECH | Názvy PnP cmdletů se ověřují za běhu, ne předpokládají

Zdroj: druhý běh proti webu TESTE.

`Set-PnPDefaultColumnValue` a `Get-PnPWebTheme` v nainstalované verzi
PnP.PowerShell neexistují. Skripty teď hledají cmdlet přes
`Get-Command -ErrorAction SilentlyContinue` a zkoušejí známé varianty názvu
(`Set-PnPDefaultColumnValues` i jednotné číslo). Když žádná není, řeknou, co
udělat ručně, místo aby krok spadl.

Motiv webu se navíc dá zadat v konfiguraci klíčem `themeName`, protože ne každá
verze PnP ho umí ze vzoru přečíst.

---

## 2026-09-09 | TECH | Krok Files kopíruje knihovny dokumentů se soubory

Zdroj: běh proti webu TESTE - stránky se přenesly, ale odkazy na nich nikam
nevedly a soubory neměly ikony.

Příčina: stránky odkazují na soubory ve vzorových knihovnách. Bez nich zůstanou
mrtvé odkazy.

`src/Copy-DocumentLibraries.ps1` vychází z `Copy-PnPDocLibs` ve `script.ps1`
verze 2.7 (autor Sergiu Nica). Navíc má náhled, přeskakuje soubory, které v cíli
už jsou, a má strop na velikost souboru.

Je ve výchozí sadě a běží **před** krokem `Pages`. Výchozí knihovnu `Dokumenty`
vynechává, protože tu plní struktura z Excelu; dá se vyžádat klíčem `libraries`.

---

## 2026-09-09 | TECH | Navigace se kopíruje celá, seznam vynechávaných je prázdný

Zdroj: běh proti webu TESTE - položka "Documents" ze vzoru se přeskočila.

Vynechávala se jako "výchozí odkaz, který si SharePoint zakládá sám", jenže ve
vzoru míří na vlastní pohled knihovny (`Forms/Project view.aspx`) a je to
záměrný odkaz. Duplicitám brání už samo slučování podle názvu, takže seznam
`SkipTitles` je teď prázdný a kopíruje se všechno.

---

<!-- Nové záznamy připisujte sem, nejnovější dolů. -->
