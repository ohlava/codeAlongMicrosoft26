# Kdo na čem pracuje

Aby si agenti a lidé nepřepisovali stejné soubory. Aktualizujte, když si berete
nebo odkládáte práci.

| Oblast | Soubory | Vlastník | Branch | Stav |
|--------|---------|----------|--------|------|
| Discovery s byznysem | `docs/01-*`, `docs/decisions.md` | — | — | čeká na akci |
| Prostředí a přístupy | `docs/decisions.md` (ENV) | — | — | **kritická cesta** |
| Export vzorového webu | `src/Export-SiteInventory.ps1` | — | — | draft, neotestováno |
| **Vstupní bod pro byznys** | `src/Setup-ProjectSite.ps1`, `config/` | Ondřej Hlava | `feature/setup-project-site` | draft, neotestováno |
| Seznamy a kalendáře | `src/Copy-SharePoint{Lists,Events}.ps1` | Sergiu Nica | `copy-listContent` | funguje |
| CSD Class / spravovaná metadata | `src/Set-CsdClass.ps1` | Sergiu Nica, Ondřej Hlava | `feature/setup-project-site` | draft |
| Původní klonovací skript | `src/script.ps1` (v2.8, parametry) | Ondřej Hlava, Sergiu Nica | `feature/setup-project-site` | přesunuto |
| Struktura složek z Excelu | `src/New-FolderStructure.ps1` | Ondřej Hlava | — | funguje, testováno |
| Stránky, vzhled, regionální nastavení | `src/Copy-SitePages.ps1` | Ondřej Hlava | `feature/setup-project-site` | draft, neotestováno |
| Kopírování navigace | `src/Copy-SiteNavigation.ps1` | Ondřej Hlava | — | draft, neotestováno |
| Kopírování knihoven ze vzoru | `script.ps1` (`Copy-PnPDocLibs`) | Sergiu Nica | `add-library-copy-function-with-folder-structure` | rozpracováno |
| Schéma šablony | `templates/`, `schema/` | — | — | draft |
| Vytvoření webu | `src/modules/Sites.psm1` | — | — | nezačato |
| Knihovny a složky | `src/modules/Libraries.psm1` | — | — | nezačato |
| Metadata | `src/modules/Metadata.psm1` | — | — | nezačato |
| Oprávnění | `src/modules/Permissions.psm1` | — | — | nezačato |
| Lifecycle / fáze | `src/modules/Lifecycle.psm1` | — | — | nezačato |
| Archivace | `src/modules/Archive.psm1` | — | — | nezačato |
| CI / pipeline | `.github/workflows/` | — | — | skeleton |
| Demo a prezentace | `docs/04-plan-hackathonu.md` | — | — | nezačato |
