# Use case: Automatic SharePoint Sites (tým 7, EOZ)

## Zadání (jak přišlo)

Automatizace vytváření a správy projektových SharePointových webů pomocí šablon
uložených a verzovaných v GitHubu. Místo manuálního zakládání projektových
portálů má řešení automaticky vytvořit standardizovaný web včetně:

- knihoven dokumentů,
- metadat,
- nastavení,
- oprávnění.

Součástí je automatizace životního cyklu projektu — přechody mezi fázemi
projektu a archivace po ukončení.

Cíle: zjednodušit správu portálů, zajistit jednotnou strukturu napříč projekty,
omezit manuální práci, využít GitHub pro verzování šablon i skriptů.

## Co to znamená přeloženo do techniky

Zadání se rozpadá na čtyři nezávislé části. Každá se dá demonstrovat samostatně,
což je pro hackathon důležité (viz [04-plan-hackathonu.md](04-plan-hackathonu.md)).

| # | Část | Co to je technicky |
|---|------|--------------------|
| 1 | **Definice šablony** | Deklarativní popis webu (YAML/JSON/PnP XML) v gitu, code review přes pull request |
| 2 | **Provisioning** | Skript/pipeline, která z definice vytvoří web v M365 tenantu |
| 3 | **Lifecycle** | Idempotentní přeaplikování definice při změně fáze projektu (jiná metadata, oprávnění, navigace) |
| 4 | **Archivace** | Read-only lock, retention/label, přesun z aktivní navigace, případně M365 Archive |

Klíčové slovo je **idempotence**: stejná operace jde spustit opakovaně a web
dokonverguje do stavu popsaného v šabloně. Bez toho části 3 a 4 nefungují —
"změna fáze" je jen další aplikace šablony s jiným vstupem.

## Co je na tom skutečně obtížné

Vytvořit web je nejjednodušší část. Skutečná složitost je v:

- **Oprávnění** — kdo smí založit projekt, kdo je vlastník, jak se řeší odchod
  člena z firmy, dědičnost z Entra ID skupin vs. SharePoint skupiny.
- **Metadata** — sdílené sloupce a content types napříč projekty (Term Store,
  Content Type Gallery) vs. lokální kopie na každém webu. Sdílené je správné a
  výrazně těžší.
- **Governance** — jmenné konvence URL, hub site, sensitivity labels, kvóty.
- **Zpětná kompatibilita** — co se stane s 200 existujícími weby, když se šablona
  ve verzi 2 změní.

Na tohle se ptáme byznysu, ne technologie. Viz
[01-otazky-pro-byznys.md](01-otazky-pro-byznys.md).
