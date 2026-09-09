# Automatic SharePoint Sites — tým 7 (EOZ)

Automatizované zakládání a správa projektových SharePointových webů ze šablon
verzovaných v tomto repozitáři, včetně přechodů mezi fázemi projektu a archivace.

Hackathon Code Along with Microsoft 2026, tým 7.

## Rychlý start

Připravit nový projektový web jedním příkazem:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\Setup-ProjectSite.ps1 -TargetSiteUrl "https://<tenant>.sharepoint.com/sites/<novy-web>" -Apply
```

Před prvním použitím je potřeba vytvořit `config/settings.json` — postup
v [docs/09-jeden-skript.md](docs/09-jeden-skript.md).

## Kde začít

| Dokument | Obsah |
|----------|-------|
| [docs/00-use-case.md](docs/00-use-case.md) | Zadání a jeho rozbor do technických částí |
| [docs/01-otazky-pro-byznys.md](docs/01-otazky-pro-byznys.md) | Otázky pro zadavatele + menu možností, které mu můžeme nabídnout |
| [docs/02-technologie.md](docs/02-technologie.md) | Přehled možností v M365 a doporučení |
| [docs/03-architektura.md](docs/03-architektura.md) | Návrh komponent a zásady |
| [docs/04-plan-hackathonu.md](docs/04-plan-hackathonu.md) | Fáze práce, demo scénář, rizika |
| [docs/05-analyza-skriptu.md](docs/05-analyza-skriptu.md) | Rozbor skriptu, který byznys používá dnes — co umí, co ne, kde má chyby |
| [docs/06-jak-spustit-export.md](docs/06-jak-spustit-export.md) | Postup spuštění exportu krok za krokem, i pro netechnického uživatele |
| [docs/08-kopirovani-navigace.md](docs/08-kopirovani-navigace.md) | Přenesení navigace ze vzorového webu na jiný |
| [docs/09-jeden-skript.md](docs/09-jeden-skript.md) | **Hlavní postup pro byznys** — jeden vstupní bod, konfigurace, flagy |
| [docs/10-stranky-a-vzhled.md](docs/10-stranky-a-vzhled.md) | Přenesení stránek, webpartů, obrázků a vzhledu ze vzoru |
| [docs/11-klonovani-vzoru.md](docs/11-klonovani-vzoru.md) | Krok TemplateClone — spuštění původního script.ps1 |\n| [docs/decisions.md](docs/decisions.md) | Log rozhodnutí — zdroj pravdy |
| [docs/workstreams.md](docs/workstreams.md) | Kdo na čem pracuje |
| [CLAUDE.md](CLAUDE.md) | Pravidla pro práci s AI agenty v tomto repozitáři |

Straw man definice projektu k diskusi s byznysem:
[templates/project-site/project.example.yml](templates/project-site/project.example.yml)

## Stav

Přípravná fáze před akcí. Zatím jen dokumentace a návrh — žádný funkční kód.

## Pravidla, na které je dobré myslet hned

- Do repozitáře nepatří tajemství ani firemní data. Podrobněji v [CLAUDE.md](CLAUDE.md).
- Práce vždy na vlastní branch, nikdy přímo na `main`.
- Skripty zapisující do SharePointu mají výchozí režim dry-run.
