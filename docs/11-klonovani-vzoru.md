# Klonování vzorového webu původním script.ps1

Krok `TemplateClone` ve [Setup-ProjectSite.ps1](09-jeden-skript.md) spustí
původní `script.ps1`, který naklonuje vzorový web jako celek.

**Není ve výchozí sadě kroků** a je potřeba ho vyžádat výslovně. Důvody jsou
níže.

## Kdy to použít

Když chcete nový web postavit "jako ten vzorový" jedním tahem a teprve pak
doplnit strukturu z Excelu:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\Setup-ProjectSite.ps1 -TargetSiteUrl "https://<tenant>.sharepoint.com/sites/<novy-web>" -Steps TemplateClone,Folders,DefaultValues -Apply
```

Pořadí je dané: nejdřív klonování vzoru, pak složky z Excelu a metadata.

Kroky `Lists` a `Events` v tomhle případě většinou nechcete — `script.ps1` už
seznamy zpracovává sám a dělaly by se dvakrát.

## Proč není ve výchozí sadě

**Maže.** `script.ps1` na řádku 221 volá `Remove-PnPList -Force` na cílový
seznam, než ho vytvoří znovu. Smaže ho i s daty. Použitelné je to jen na
čerstvě založený web — na běžícím projektu je to nevratná ztráta obsahu.

**Není přírůstkový.** Druhé spuštění nedoplní chybějící, ale přepíše všechno.
Nehodí se na doaplikování změn.

**Ztrácí hodnoty Lookup polí.** Remapování starých ID na nová v něm podle čtení
kódu nefunguje — viz [05-analyza-skriptu.md](05-analyza-skriptu.md), nález 1.
Seznamy přenesené `Copy-SharePointLists.ps1` tuhle vadu nemají.

Proto je výchozí cesta modulární: `Folders, Lists, Events, Pages,
DefaultValues`. `TemplateClone` je pro případ, kdy chcete rychle celý klon.

## Jak se konfiguruje

`script.ps1` nemá parametry — konfigurace je napsaná v jeho hlavičce. Aby se
nemusel upravovat (je společný a byznys ho spouští i ručně), `Setup-ProjectSite`
si z něj **vygeneruje kopii** s doplněnými hodnotami do
`export/script.generated.ps1` a spustí tu. Originál zůstane nedotčený.

Nahrazují se tyto řádky:

| Řádek v `script.ps1` | Čím se nahradí |
|----------------------|----------------|
| `$SiteDomain` | Doména z `-TargetSiteUrl` |
| `$SourcePath` | Cesta ze `sourceSiteUrl` v konfiguraci |
| `$TargetPath` | Cesta z `-TargetSiteUrl` |
| `$ClientId` | `clientId` z konfigurace |
| `$CopyCount` | 1000, nebo 1 když `copyLists: false` |
| `$IsCopyPages`, `$IsCopyTemplateDesign`, `$IsCopyRegionalSettings`, `$IsCopyNavigation` | Podle `legacyScript` v konfiguraci |
| `$SetOfflineAvailable` | `ja` / `nein` podle `setOfflineAvailable` |
| `Clear-Host` | Zakomentuje se, jinak by smazal výpis předchozích kroků |
| Řádky 2, 5 a 6 | Zakomentují se — nejsou to komentáře, PowerShell je zkouší spustit jako příkazy |

Generovaná kopie obsahuje ClientId a adresy tenantu, proto vzniká v `export/`,
který je v `.gitignore`.

### Vypnutí kopírování seznamů

```json
"legacyScript": { "copyLists": false }
```

`script.ps1` na to nemá vypínač, takže se to obchází přes `$CopyCount = 1`.
Podmínka ve skriptu je `$counter -lt $CopyCount` a `$counter` začíná na 1, takže
se nezpracuje žádný seznam. Stránky, navigace a vzhled se přenesou i tak,
protože se volají až za smyčkou přes seznamy.

Je to trik, ne podporovaný způsob. Kdyby se ve `script.ps1` změnila ta
podmínka, přestane fungovat — pak se to pozná tak, že se seznamy začnou
kopírovat, i když mají být vypnuté.

## Náhled

Bez `-Apply` se kopie **vygeneruje a vypíšou se doplněné hodnoty**, ale nespustí
se. Režim náhledu `script.ps1` nemá.

Vyplatí se v tom výpisu zkontrolovat `$SourcePath` a `$TargetPath` — záměna
zdroje a cíle je v kombinaci s mazáním nejhorší možná chyba.

## Co dělá script.ps1 a nikdo jiný

I když se `TemplateClone` nepoužije, `script.ps1` zůstává v repozitáři, protože
umí věci, které jinde nemáme:

- **knihovny dokumentů včetně souborů** — `Copy-PnPDocLibs` ve verzi 2.7
  (autor Sergiu Nica, zatím na vlastní branchi, na této není),
- **dávkové zpracování více cílových webů** ze `Sites.xml` (`$CopyFromList`),
- **kopírování mezi webem a podwebem** (`$CopyFromMainToSubSite`),
- **offline dostupnost seznamů** (`$SetOfflineAvailable`).
