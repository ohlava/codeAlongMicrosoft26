# Kontext projektu pro AI agenty

Tento soubor je sdílený kontext pro všechny členy týmu i jejich agenty.
Když se něco dozvíte, co by agent měl vědět příště, dopište to sem — ne do chatu.

## Co stavíme

Automatizované zakládání a správu projektových SharePointových webů pro EOZ
(Škoda Auto) ze šablon verzovaných v tomto repozitáři. Součástí je životní cyklus
projektu (přechody fází) a archivace po ukončení.

Zadání a rozbor: [docs/00-use-case.md](docs/00-use-case.md).
Architektura a zásady: [docs/03-architektura.md](docs/03-architektura.md).
Rozhodnutí učiněná v průběhu: [docs/decisions.md](docs/decisions.md).

## Nepřekročitelná pravidla

1. **Do repozitáře nepatří tajemství ani firemní data.** Žádné client secrets,
   certifikáty, tokeny, connection stringy. Žádné konkrétní adresy tenantu,
   e-maily, jména zaměstnanců, čísla reálných projektů ani exporty dat.
   V ukázkách používejte `example.com` a smyšlené hodnoty.
2. **Skript, který zapisuje do SharePointu, má výchozí režim dry-run.**
   Ostrá aplikace jen s explicitním přepínačem.
3. **Nikdy negenerujte kód, který maže knihovny, sloupce, weby nebo oprávnění.**
   Engine přidává, upravuje a uzamyká. Mazání zůstává člověku.
4. **Všechny operace musí být idempotentní.** Před zápisem zjistit stav a
   doplnit jen chybějící.
5. **Nespouštějte nic proti produkčnímu tenantu**, dokud to není v
   [docs/decisions.md](docs/decisions.md) výslovně povolené.

## Technologické volby

Než začnete kódovat, zkontrolujte [docs/decisions.md](docs/decisions.md) — volby
se během hackathonu mění a decisions.md je zdroj pravdy, ne tento odstavec.

Aktuální předpoklad: PnP PowerShell jako provisioning engine, definice projektů
ve YAML validovaném JSON Schematem, orchestrace přes GitHub Actions.

## Konvence

- Definice projektů: `templates/`, validované proti `schema/`.
- Kód: `src/`, rozdělený na moduly podle domény (weby, knihovny, metadata,
  oprávnění, lifecycle, archivace). Modul nemá vědět o ostatních.
- Testy: `tests/`, Pester. Nový modul bez testu se nemerguje.
- Dokumentace ke rozhodnutí patří do `docs/`, nikoli do komentářů v kódu.
- Komentáře v kódu vysvětlují *proč*, ne *co*. Bez metadat typu autor a datum.

## Spolupráce s ostatními agenty v týmu

Repozitář je sdílený a pracuje v něm několik lidí s agenty současně. Aby se
nepřepisovali:

- **Vždy pracujte na vlastní branch**, nikdy přímo na `main`.
  Konvence: `<jmeno>/<oblast>`, například `hlava/permissions`.
- **Před začátkem práce `git pull --rebase origin main`.**
- **Malé PR, často.** Velký PR na hackathonu nikdo nezreviduje.
- **Jeden modul = jeden vlastník v daný okamžik.** Kdo na čem pracuje, je
  v [docs/workstreams.md](docs/workstreams.md).
- **Neupravujte cizí modul.** Když potřebujete změnu, dopište požadavek do PR
  popisu nebo do workstreams.md.
- **Rozhraní modulů se dohodnou předem** a mění se jen po dohodě, protože na
  nich staví ostatní.
- **Nepřepisujte `docs/decisions.md`** — jen připisujte na konec.

## Definice hotového

- Funguje to opakovaně (druhé spuštění nic nerozbije).
- Existuje dry-run výstup, který lidsky popíše, co se stane.
- Je na to test nebo aspoň zdokumentovaný ruční postup ověření.
- Chování je popsané v `docs/`.
