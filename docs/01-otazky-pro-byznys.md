# Otázky pro byznys (Markéta a kolegové z EOZ)

Princip: **neptáme se na technologie, ptáme se na proces.** Ona zná bolest, my
známe nástroje. Nejrychlejší cesta k požadavkům není otevřená otázka
("co byste chtěla?"), ale konkrétní návrh, který může odmítnout
("takhle by to vypadalo — co je špatně?"). Proto máme připravený straw man:
[templates/project-site/project.example.yml](../templates/project-site/project.example.yml).

## Blok A — Jak to funguje dnes (15 min, nejdřív jen posloucháme)

1. Můžete nám ukázat **dva reálné projektové weby** — jeden vzorový a jeden,
   který se zvrhl? Ten druhý nám řekne víc.
2. Kdo dnes web fyzicky zakládá? Vy, IT, admin projektu?
3. Kolik projektů se zakládá — za měsíc, za rok?
4. Jak dlouho jedno založení trvá od žádosti k použitelnému webu?
5. Co z toho je čekání na někoho jiného a co skutečná práce?
6. Existuje dnes nějaká napsaná šablona, checklist nebo Wordový postup?
   (I zastaralý — je to nejcennější vstup, který máme.)
7. Co se nejčastěji zapomene nebo udělá špatně?

## Blok B — Co má web obsahovat

8. Projděme knihovny: které knihovny dokumentů má mít každý projekt?
   (Nabídneme náš návrh a necháme škrtat a doplňovat.)
9. Které složky/struktury vevnitř jsou povinné?
10. Jaká **metadata** se u dokumentů vyplňují? (Číslo projektu, fáze, útvar,
    dodavatel, typ dokumentu, důvěrnost…) Které z nich jsou povinné?
11. Existují už schválené číselníky pro tyto hodnoty, nebo si je každý vymýšlí?
12. Má být na webu i něco jiného než dokumenty — úkoly, seznam rizik, kalendář,
    zápisy z porad, kontakty, přehledová stránka?
13. Je potřeba propojení s Teams (kanál k projektu), Planner, Loop?
14. Liší se struktura podle typu projektu? Kolik různých typů projektů je?
    (Toto rozhodne, jestli děláme jednu šablonu, nebo šablonu s variantami.)

## Blok C — Oprávnění (nejcitlivější část)

15. Jaké **role** na projektu existují? (Vedoucí projektu, člen týmu, čtenář,
    externí dodavatel, controlling, management…)
16. Co která role smí — čtení, editace, mazání, správa oprávnění?
17. Kdo rozhoduje, že někdo dostane přístup? Vedoucí projektu, nebo IT?
18. Přistupují k webu i **externisté** mimo Škodu? Přes co dnes?
19. Existují dokumenty, které nesmí vidět celý projektový tým?
    (Rozpočet, personálie, smlouvy.)
20. Jsou už v Entra ID / AD **skupiny** odpovídající projektovým týmům, nebo se
    lidé přidávají po jednom?

## Blok D — Fáze projektu a životní cyklus

21. Jaké **fáze** projekt prochází? Prosíme o oficiální názvy, které se používají.
22. Co se **konkrétně mění** při přechodu do další fáze? Přijmeme i banality:
    - přidá se nová knihovna,
    - předchozí fáze se uzamkne pro editaci,
    - změní se, kdo je vlastník,
    - přidá se schvalovatel,
    - něco se musí vygenerovat nebo nahlásit.
23. Kdo přechod fáze **schvaluje**? Má o tom někdo dostat notifikaci?
24. Vrací se projekt někdy do předchozí fáze? (Rozhoduje, jestli musí být
    přechody obousměrné.)
25. Kde je dnes uložená "pravda" o tom, v jaké fázi projekt je? V SAPu, v Excelu,
    v hlavě vedoucího projektu? **Existuje systém, ze kterého bychom to mohli
    číst automaticky?**

## Blok E — Ukončení a archivace

26. Co znamená "projekt skončil"? Kdo to vyhlásí?
27. Jak dlouho se musí dokumenty držet? Je na to interní směrnice nebo zákonná
    lhůta? (Typicky se liší podle typu dokumentu.)
28. Po archivaci: má být web **neviditelný**, nebo dohledatelný a čitelný?
29. Musí být archiv prohledávatelný? Kdo do něj smí?
30. Smí se něco po uplynutí lhůty smazat automaticky, nebo to vždy někdo potvrdí?

## Blok F — Rozsah pro hackathon

31. Kdyby z celého řešení fungovala **jediná věc**, která to je?
32. Co je největší zdroj frustrace — pomalost, chaos ve struktuře, chybějící
    dokumenty, nebo oprávnění?
33. Komu budeme na konci hackathonu výsledek předvádět a co ten člověk chce vidět?
34. Existuje šance, že se to po hackathonu skutečně nasadí? Kdo by to vlastnil?

## Blok G — Prostředí a přístupy (spíš na IT než na Markétu, ale zeptat se hned)

Toto je **kritická cesta**. Bez přístupu se první den nedá dělat nic než návrh.

35. Máme k dispozici **testovací tenant**, nebo produkci? Kdo je tam admin?
36. Dostaneme **SharePoint Administrator** roli, nebo aspoň možnost registrovat
    aplikaci v Entra ID?
37. Existuje **hub site** pro projekty, pod který se mají weby řadit?
38. Je povolené PowerShell modul `PnP.PowerShell` / CLI for Microsoft 365 proti
    tomu tenantu? (V řadě korporací je app registration s `Sites.FullControl.All`
    zablokovaná — je potřeba vědět hned.)
39. Máme GitHub organizaci s povolenými **GitHub Actions** proti tenantu, nebo
    poběží automatizace v Azure (Functions / Automation) a GitHub bude jen
    úložiště šablon?
40. Jsou v tenantu **sensitivity labels** povinné pro nové weby?

## Menu možností — co jí můžeme nabídnout

Netechnický zadavatel nedokáže chtít věc, o které neví, že existuje. Toto jí
ukážeme jako výběr, ne jako plán.

- **Formulář na založení projektu** — vyplní se název, číslo, vedoucí, typ; web
  se vytvoří sám za pár minut.
- **Schvalovací krok** — před vytvořením webu to někdo odklikne.
- **Automatické čtení fáze z nadřazeného systému** — nikdo nemusí nic překlikávat.
- **Přehledový dashboard** — všechny projekty, jejich fáze, vlastník, velikost,
  datum poslední aktivity.
- **Kontrola shody (drift detection)** — pravidelný report, které weby se
  odchýlily od šablony (někdo smazal knihovnu, přidal si oprávnění).
- **Verzování šablony** — změnu struktury schvaluje pull request, je vidět, kdo a
  proč co změnil, a jde to vrátit.
- **Hromadné doaplikování** — nová povinná knihovna se přidá do všech 200
  běžících projektů jedním spuštěním.
- **Automatická archivace** — po ukončení projektu se web uzamkne, odejde z
  navigace a čeká na uplynutí retenční lhůty.
- **Notifikace** — vedoucí projektu dostane zprávu, když se blíží konec fáze nebo
  když web nikdo 90 dní nepoužil.

## Jak schůzku vést

- Bloky A a B na začátek, dokud je energie. Blok G paralelně poslat na IT hned
  ráno — má nejdelší dobu odezvy.
- Vždy si vyžádat **reálný příklad**, ne popis pravidla. "Ukažte mi projekt Kodiaq
  facelift" je lepší než "jak vypadá projektový web".
- Vše, co zazní, zapsat do [docs/decisions.md](decisions.md) jako rozhodnutí
  s datem. Na hackathonu se lidé po dvou dnech nepamatují, co se domluvilo.
- Otázku, na kterou nezná odpověď, nezahazovat — je to zjištění samo o sobě
  (znamená, že proces není definovaný, a to je součást výsledku).
