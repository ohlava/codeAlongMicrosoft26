# Návrh architektury (draft před schůzkou s byznysem)

## Tok

```
  Požadavek na projekt                Definice v gitu
  (formulář / SP list / ruční)        templates/project-site/*.yml
            |                                   |
            +----------------+------------------+
                             v
                    Orchestrátor (GitHub Actions)
                      - validace schématu
                      - dry-run: co se změní
                      - aplikace
                             |
                             v
                    Provisioning engine (PnP PowerShell)
                             |
        +--------------------+--------------------+
        v                    v                    v
   SharePoint web       Entra ID skupiny      Purview labels
   knihovny, sloupce    členství, role        retence, archivace
                             |
                             v
                    Audit log + report stavu
```

## Komponenty a co má která na starost

| Komponenta | Odpovědnost |
|------------|-------------|
| `templates/` | Deklarativní definice — struktura, metadata, role, fáze. Verzované, měněné přes PR |
| `schema/` | JSON Schema pro validaci šablon. Chyba se pozná v PR, ne až na produkci |
| `src/` | Provisioning engine — čte definici, čte aktuální stav webu, dopočítá rozdíl, aplikuje |
| `.github/workflows/` | Spouštěče: validace na PR, provisioning na `workflow_dispatch`, drift report na plán |
| `docs/` | Rozhodnutí, otázky, návrh |

## Zásady, na kterých bych trval

**Idempotence.** Každá operace zjistí aktuální stav a doplní chybějící. Druhé
spuštění nesmí nic rozbít ani zduplikovat. Bez toho nefungují fáze ani hromadné
doaplikování šablony.

**Nikdy nemazat.** Engine přidává, upravuje nastavení a uzamyká. Mazání knihovny
nebo sloupce nechá na člověku. U SharePointu je návrat obtížný a jeden špatný
běh proti produkci by projekt zabil.

**Dry-run jako výchozí režim.** Ostrá aplikace vyžaduje explicitní přepínač.
Tohle je jediný důvod, proč by správce tenantu takový nástroj kdy pustil.

**Šablona je čitelná pro byznys.** Když zadavatel nedokáže v PR poznat, co se
mění, přišli jsme o hlavní přínos verzování.

**Stav webu je odvozený z definice, ne naopak.** Zdrojem pravdy je git. Ruční
změna na webu je drift a má se reportovat.

## Navrhovaná struktura repozitáře

```
templates/
  project-site/
    project.example.yml       # straw man pro diskusi
    types/                    # varianty podle typu projektu
schema/
  project-site.schema.json
src/
  Provision-ProjectSite.ps1   # hlavní vstupní bod
  modules/
    Sites.psm1                # vytvoření a nastavení webu
    Libraries.psm1            # knihovny, složky, verzování
    Metadata.psm1             # site columns, content types
    Permissions.psm1          # skupiny, role
    Lifecycle.psm1            # přechody fází
    Archive.psm1              # uzamčení, retence
tests/
  Provision.Tests.ps1         # Pester, proti mocku nebo dev tenantu
.github/workflows/
  validate-templates.yml
  provision-site.yml
  drift-report.yml
docs/
```

Struktura je návrh — pokud se tým rozhodne pro `m365` CLI nebo Python
orchestraci, vrstvy zůstávají stejné, mění se jen obsah `src/`.
