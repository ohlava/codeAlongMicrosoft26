# Plán hackathonu

Délka a přesný program akce nejsou zatím známé — plán je napsaný jako fáze,
přeškálujte na skutečný počet dní.

## Fáze 0 — Ještě před akcí (teď)

- [x] Repozitář, dokumentace, otázky na byznys, straw man šablony
- [ ] Zjistit, jaký tenant dostaneme a kdo je tam admin (**kritická cesta**)
- [ ] Zjistit, jestli dostaneme Entra ID app registraci
- [ ] Ověřit, že `PnP.PowerShell` funguje z notebooků členů týmu
- [ ] Domluvit s Markétou 45 min na začátku prvního dne
- [ ] Vyžádat si předem export nebo screenshoty dvou reálných projektových webů

## Fáze 1 — Discovery (první ráno, max 2 hodiny)

Rozhovor podle [01-otazky-pro-byznys.md](01-otazky-pro-byznys.md). Výstup:

- seznam knihoven a metadat, který zadavatel potvrdil,
- seznam rolí a co smí,
- seznam fází a co se při přechodu mění,
- jedna věta: "úspěch hackathonu znamená, že …".

Tvrdý časový strop. Discovery se dá dělat tři dny a nemít co ukázat.

## Fáze 2 — Vertikální řez (nejdřív celá cesta, ne dokonalé části)

Cíl: jeden příkaz vytvoří jeden web s jednou knihovnou a jedním sloupcem. Ošklivě,
ale od konce ke konci a z pipeline, ne z notebooku.

Tohle odemkne všechna překvapení (auth, práva, URL konvence) v době, kdy je ještě
čas je řešit. Nedělejte to naopak.

## Fáze 3 — Rozšíření do šířky

Paralelizovatelné mezi členy týmu, protože moduly jsou oddělené:

- knihovny, složky, verzování
- metadata a content types
- oprávnění a skupiny
- validace šablony proti JSON Schema
- dry-run výstup

## Fáze 4 — Lifecycle a archivace

Přechod fáze a archivace. Tohle je část, kterou ostatní týmy typicky nemají
hotovou, a v zadání je explicitně. Zaslouží si čas.

## Fáze 5 — Demo (rezervovat poslední půlden, opravdu)

Scénář, který funguje bez internetu a bez improvizace:

1. Markéta otevře pull request, kde se do šablony přidává povinná knihovna.
   Vidí čitelný diff a schválí.
2. Merge spustí pipeline.
3. Vytvoří se nový projektový web se vším nastavením. Ukázat živě.
4. Spustí se přechod do další fáze — přibude knihovna, předchozí se uzamkne.
5. Spustí se archivace — web je read-only, mimo navigaci, s retenční značkou.
6. Drift report ukáže web, na kterém někdo ručně smazal knihovnu.

Připravit **záložní nahrané video** celého scénáře. Tenant během dema padne
častěji, než se čeká.

## Rizika a co s nimi

| Riziko | Dopad | Reakce |
|--------|-------|--------|
| Nedostaneme app registraci / admin práva | Blokující | Interaktivní přihlášení, demo z notebooku, fallback na osobní dev tenant |
| Discovery se protáhne | Nic hotového | Tvrdý strop 2 hodiny, dál pracujeme na předpokladech a označíme je |
| Zadavatel chce příliš mnoho | Nic dokončeného | Vertikální řez první, rozsah řezaný po fázích |
| Tenant nebo síť padne při demu | Ztracená prezentace | Nahrané video |
| Tým si rozdělí práci a nejde to složit | Poslední den v panice | Rozhraní modulů dohodnout před dělením, integrovat každý den |
| Šablona nečitelná pro byznys | Ztracená hlavní pointa | YAML, ne PnP XML |
