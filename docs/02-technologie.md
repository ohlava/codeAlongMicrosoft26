# Technologické možnosti a doporučení

Přehled toho, co na tuto úlohu v ekosystému M365 existuje, s tím, co bych vybral
pro hackathon. Vše před nasazením ověřit proti aktuální dokumentaci a proti
politikám tenantu — verze a licenční podmínky se u M365 mění často.

## Vrstva 1 — Čím web vytvořit a nastavit

| Nástroj | K čemu | Silné stránky | Slabiny |
|---------|--------|---------------|---------|
| **PnP PowerShell** (`PnP.PowerShell`) | Provisioning, konfigurace, export/import šablon | Nejsilnější pokrytí SharePointu, `Get-PnPSiteTemplate` umí web vyexportovat | PowerShell 7, vyžaduje vlastní Entra ID app registraci (verze 2.x) |
| **CLI for Microsoft 365** (`m365`) | Totéž, cross-platform | Dobré do CI, JSON výstupy, skriptovatelné z čehokoli | Menší pokrytí exotických nastavení |
| **Microsoft Graph API** | Skupiny, Teams, členství, soubory | Stabilní, dobře dokumentované, jazykově neutrální | Vytvoření samostatné site collection přes Graph není plně pokryté |
| **SharePoint REST / CSOM** | Cokoli, co výše uvedené neumí | Úplné pokrytí | Verbózní, nízkoúrovňové |
| **Site Designs + Site Scripts** | Nativní SharePoint šablony (JSON) | Nulová infrastruktura, aplikují se při zakládání webu | Omezená sada akcí a jejich počet; na oprávnění a lifecycle nestačí |
| **PnP Provisioning Templates** (XML) | Deklarativní šablona webu | Umí exportovat existující web a znovu ho aplikovat | XML nečitelné pro byznys, špatně se review-uje v PR |

**Doporučení:** PnP PowerShell jako provisioning engine. Alternativa `m365` CLI,
pokud tým nechce PowerShell. Nedoporučuji stavět to na Site Designs — narazíte na
limity v půlce zadání (oprávnění, archivace).

## Vrstva 2 — Kde žije definice projektu

Klíčové designové rozhodnutí. Tři varianty:

**A) PnP XML šablona přímo v gitu.** Nejméně kódu, ale diff v PR je nečitelný a
Markéta si to sama neotevře.

**B) Vlastní YAML/JSON schéma + převodník na PnP volání.** Více práce, ale
šablona je čitelná, dá se review-ovat v PR a byznys ji chápe. Zvládne varianty
podle typu projektu a fází.

**C) Hybrid.** Vzorový web se jednorázově vyexportuje přes `Get-PnPSiteTemplate`
jako startovní bod, z něj se ručně destiluje YAML schéma.

**Doporučení: C jako bootstrap, dál B.** Čitelná šablona je pro tuto úlohu půlka
hodnoty — celý smysl "šablony v GitHubu" je, že je nad ní vidět změna a schválení.
Straw man schématu: [templates/project-site/project.example.yml](../templates/project-site/project.example.yml).

## Vrstva 3 — Co to spouští

| Varianta | Kdy použít |
|----------|------------|
| **GitHub Actions** | Šablony i automatizace v gitu, spouštění na merge nebo `workflow_dispatch`. Přesně odpovídá zadání |
| **Azure Functions** | Reakce na události, delší běhy, korporátní preference "běží to v Azure" |
| **Azure Automation / Runbooks** | Naplánované úlohy, drift detection, archivační dávky |
| **Power Automate** | Formulář a schvalování, hezké pro byznys demo, špatné pro logiku |

**Doporučení pro hackathon:** GitHub Actions s `workflow_dispatch` (ruční
spuštění s parametry) i `on: push` pro šablony. Je to nejkratší cesta k funkční
demonstraci a zadání explicitně chce GitHub.

Pro produkci realisticky vyjde kombinace: formulář v Power Apps / SharePoint listu
→ zápis požadavku → GitHub Actions nebo Azure Function jako vykonavatel.

## Vrstva 4 — Autentizace (nejčastější místo, kde hackathon uvázne)

Provisioning potřebuje app-only přístup s vysokými právy (typicky
`Sites.FullControl.All`, `Group.ReadWrite.All`). To se v korporaci neschvaluje za
hodinu. Řešit **první den**.

Možnosti od nejlepší:

1. **Workload identity federation (OIDC)** mezi GitHub Actions a Entra ID —
   žádný secret v repozitáři. Ideální cíl.
2. **Certifikát** v GitHub Secrets — standardní PnP postup, funguje spolehlivě.
3. **Client secret** — pro hackathon akceptovatelné, do produkce ne.
4. **Interaktivní přihlášení** (`Connect-PnPOnline -Interactive`) — fallback,
   pokud app registraci nedostaneme. Demo pak jede z notebooku, ne z CI.

Připravit varianty 3 i 4, aby demo nezáviselo na schválení.

## Vrstva 5 — Metadata

- **Term Store (Managed Metadata)** — sdílené číselníky napříč weby (útvary,
  typy dokumentů). Správné řešení, vyžaduje práva na Term Store.
- **Content Type Gallery / Content Type Hub** — centrální content types
  publikované do webů. Správné řešení pro "jednotná metadata".
- **Lokální site columns na každém webu** — rychlé, ale za rok nekonzistentní.

Pro hackathon lokální sloupce stačí, ale v prezentaci pojmenovat, že produkční
verze patří do Content Type Gallery. Ukazuje to, že tým rozumí governance.

## Vrstva 6 — Oprávnění

- **M365 group-connected team site** — členství řeší M365 skupina (Owners /
  Members). Jednodušší, ale jen dvě úrovně a vždy s tím vzniká Teams/Planner.
- **Communication site + SharePoint skupiny** — víc kontroly nad rolemi,
  bez M365 skupiny.
- **Entra ID security groups mapované do SharePoint skupin** — nejlepší pro
  korporaci: členství se spravuje v jednom místě, ne na každém webu.
- **Sensitivity labels (container labels)** — mohou být v tenantu povinné a
  vynucují nastavení externího sdílení.

Pro projektové weby s rolemi (vedoucí / člen / čtenář / externista) vede
kombinace Entra ID skupin. Otázky 15–20 v [01-otazky-pro-byznys.md](01-otazky-pro-byznys.md)
mají za cíl zjistit, kolik rolí skutečně existuje.

## Vrstva 7 — Lifecycle a archivace

**Přechod fáze** = znovu-aplikování šablony s jiným parametrem fáze. Prakticky:
změna metadat webu, přidání knihovny, uzamčení předchozí knihovny pro editaci,
změna navigace, notifikace.

**Archivace** — možnosti:

- `Set-PnPTenantSite -LockState ReadOnly` — web zůstane čitelný, nikdo needituje.
  Nejrychlejší a pro demo dostatečné.
- **Retention labels / policies (Microsoft Purview)** — správné řešení pro
  zákonné lhůty, drží dokumenty i po smazání webu.
- **Microsoft 365 Archive** — vyvedení obsahu do levnějšího úložiště s možností
  reaktivace. Zpoplatněno zvlášť, licenční model ověřit.
- Odebrání z hub navigace, přeznačení metadaty, přesun do jiného hubu.

Pro hackathon: read-only lock + retention label + odebrání z navigace. Popsat
M365 Archive jako produkční variantu.

## Co ještě stojí za zvážení

- **Drift detection** — porovnání reálného webu proti šabloně a report odchylek.
  Málo práce, silné demo, protože to je bolest, kterou nikdo neřeší.
- **Dry-run / plán změn** — vypsat, co by se provedlo, bez provedení. Nutné, aby
  si to někdo vůbec troufl pustit na produkci.
- **Rollback** — u SharePointu částečně nemožný (smazanou knihovnu vrátíte
  z koše, smazaná metadata jsou pryč). Radši navrhnout režim, kdy skript nikdy
  nemaže, jen přidává a uzamyká.
