# Přispívání do projektu

Tento dokument je jednoduchý a je určen i pro kolegy, kteří nepracují pravidelně s Gitem nebo technickým vývojem.

## 1. Jak pracovat s větvemi

Používáme jednoduchou konvenci pojmenování větví:

- `feature/nazev-funkce` – pro novou funkci nebo novou část projektu
- `fix/nazev-opravy` – pro opravu chyby

Příklady:

- `feature/frontloading-overview`
- `feature/ssp-template-setup`
- `fix/changelog-typo`

Doporučení:

- Větev vytvářejte z aktuální hlavní větve (`main`).
- Po dokončení práce ji sloučte přes pull request.
- Nepracujte dlouho jen na jedné větvi bez pravidelných aktualizací.

## 2. Jak psát commit zprávy

Commit zpráva by měla mít formát:

`typ(oblast): popis`

Příklady:

- `feat(doc): přidána demo struktura changelogu`
- `fix(template): opraveno pojmenování sekce archivace`
- `docs(contributing): doplňena pravidla pro práci s větvemi`

Typy zpráv:

- `feat` – nová funkce
- `fix` – oprava chyby
- `docs` – změna dokumentace
- `refactor` – úprava kódu bez změny funkce
- `chore` – drobné úklidové změny

## 3. Sémantické verzování šablon

V projektu používáme semantické verzování podle pravidla:

- `MAJOR` – velká změna, která rozbíjí kompatibilitu nebo mění zásadní pravidla
- `MINOR` – přidání nové funkce nebo rozšíření bez rozbití starého chování
- `PATCH` – drobná oprava, malá úprava nebo oprava chyby

Příklady:

- `1.0.0` – první stabilní verze projektu
- `1.1.0` – přidána nová funkce, například Frontloading
- `1.2.0` – přidán nový modul, například SSP
- `1.3.0` – rozšíření pro archivaci
- `1.3.1` – drobná oprava v archivní části

### Kdy zvýšit verzi

- Pokud přidáváte novou funkci nebo nový modul: `MINOR`
- Pokud měníte už existující chování a je to zásadní: `MAJOR`
- Pokud opravujete chybu nebo drobnou nepřesnost: `PATCH`

## 4. Jak označit verzi pomocí Git tagu

Po dokončení verze se označí tagem:

`git tag v1.2.0`

A následně se tag odešle do vzdáleného repozitáře:

`git push origin v1.2.0`

Příklady:

- `v1.0.0` – první oficiální verze
- `v1.1.0` – rozšíření frontloadingu
- `v1.2.0` – přidání SSP
- `v1.3.0` – archivace

## 5. Praktický workflow

1. Vytvořte novou větev z `main`.
2. Upravte obsah a doplňte změny.
3. Commitněte podle konvence `typ(oblast): popis`.
4. Ověřte, zda je změna v souladu s pravidly projektu.
5. Vytvořte pull request.
6. Po schválení sloučte do `main`.
7. Pokud jde o novou verzi, označte ji tagem a aktualizujte changelog.

## 6. Demo pravidla pro tým

Toto je pouze demo příklad, který má pomoci i méně technickým kolegům pochopit běžný pracovní postup. V reálném projektu je vždy nutné dodržovat interní pravidla týmu a bezpečnostní požadavky.

## 7. Malý tip

Pokud nejste jistí, kdy zvýšit verzi, řiďte se jednoduchým pravidlem:

- nová funkce = `MINOR`
- oprava = `PATCH`
- radikální změna = `MAJOR`

Díky tomu je práce s verzemi jednoduchá a přehledná pro všechny.
