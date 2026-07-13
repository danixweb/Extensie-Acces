# Schema jurnalului de progres (`.accdb-progress/`)

Document canonic — referențiat de `vba-progres`, `vba-analiza`, `vba-include-tratare-erori` și `vba-corectare-erori`. Orice schimbare de format se face aici, o singură dată, nu în fiecare SKILL.md.

## Locație și ciclu de viață

`.accdb-progress/`, sibling al `.accdb`-ului (lângă `.accdb-ai/`, `.accdb-errors/`, `.accdb-backups/`) — local, gitignored, persistent (nu se șterge la închiderea bazei, spre deosebire de `.accdb-ai/`).

```
.accdb-progress/
  INDEX.md
  Modules/<nume-obiect>.md
  Forms/<nume-obiect>.md
  Reports/<nume-obiect>.md
  Queries/<nume-obiect>.md
  Tables/<nume-obiect>.md
```

Numele folderelor urmează exact `Category`/`CATEGORIES` din `access-explorer/src/model.ts` (`Modules`, `Forms`, `Reports`, `Queries`, `Tables`) — aceeași convenție ca `.accdb-ai/<Category>/`.

**Regulă de bază: se creează fișier per-obiect DOAR pentru obiecte efectiv atinse.** Nu se pre-generează unul gol per obiect din inventarul complet al bazei (ar fi zgomot inutil pentru o bază cu sute/mii de obiecte). Un obiect fără fișier propriu = netratat, punct — nu necesită nicio altă verificare pentru a ști asta.

**Granularitate de rutină doar unde s-a lucrat.** Într-un modul deja atins, rutinele care NU au secțiune proprie în fișierul `.md` sunt implicit netratate — nu le enumera explicit (ar necesita parsarea completă a modulului doar ca să dovedești că nu s-a lucrat la ceva, cost nejustificat).

## `INDEX.md`

Regenerat integral de fiecare dată când un skill face status/actualizare (nu editat incremental manual) — sursa de adevăr pentru inventar e mereu un `list` proaspăt prin `access-bridge.ps1`, combinat cu fișierele per-obiect existente la momentul regenerării.

```markdown
# Jurnal de progres — <DbPath>
Ultima verificare inventar: <timestamp ISO>, via bridge `list` (<total> obiecte)

## Imagine de ansamblu
<paragraf liber, opțional, sintetizat din secțiunile "Scop general" ale obiectelor atinse —
ce face aplicația per ansamblu, cum se leagă obiectele atinse între ele (fluxuri identificate).
Actualizat manual/de skill pe măsură ce apar suficiente obiecte atinse ca să contureze un tipar;
nu forța o narațiune dacă doar 1-2 obiecte izolate au fost atinse.>

## Modules (<N atinse> din <total>)
- [<nume>](Modules/<nume>.md) — <Scop general, pe scurt> · <rezumat foarte scurt al ultimei lucrări>
- ...
- (<M> netratate: <nume1>, <nume2>, ...)

## Forms (<N> din <total>)
...

## Reports (<N> din <total>)
...

## Queries (<N> din <total>)
...

## Tables (<N> din <total>)
...
```

Categoriile fără niciun obiect atins tot apar (cu "(0 din <total>)" și lista completă netratată) — utilizatorul trebuie să vadă imediat ce categorii n-au fost deloc explorate.

## Fișier per obiect (`<Category>/<nume-obiect>.md`)

```markdown
---
category: <Modules|Forms|Reports|Queries|Tables>
object: <nume exact obiect, cum apare in inventarul `list`>
lastUpdated: <YYYY-MM-DD>
---

## Scop general
<Ce rol are acest obiect in aplicatie, per ansamblu — nu doar ce face o rutina anume.>
(Sursă: dedus de Claude din cod — sau: dictat de utilizator <data>.)

## <NumeRutina1>
- **Scop:** <ce face concret aceasta rutina/subrutina — intrare, iesire, efect>
- **Lucrat:** <data>, via `<nume-skill>` <context suplimentar daca aplica, ex. "(eroare id=7186)">
- **Ce s-a făcut:** <descriere concreta a modificarii sau a analizei>
- **Stare:** <aplicat si confirmat | propus, neconfirmat | doar analizat, fara modificare>
- **Antet în cod:** `id=<id-ul din antetul AI-Track>` (doar daca s-a scris efectiv cod — vezi `antet-cod.md`)

## <NumeRutina2>
...

<!-- rutine neatinse din acelasi obiect NU apar aici -->
```

Câmpuri obligatorii per rutină atinsă: `Scop`, `Lucrat`, `Ce s-a făcut`, `Stare`. `Antet în cod` apare doar dacă acea atingere a scris efectiv cod (nu pentru o analiză pur read-only).

Pentru obiecte fără rutine individuale (ex. o `Query` sau un `Table` — nu au "Sub"/"Function") secțiunea per-rutină nu se aplică; se folosește direct "Scop general" + un bloc unic de status:

```markdown
## Modificare
- **Lucrat:** <data>, via `<skill>`
- **Ce s-a făcut:** <ex. "corectat SQL-ul interogarii: JOIN gresit pe cheia externa">
- **Stare:** aplicat și confirmat
```

## Reguli pentru câmpul "Scop"/"Scop general"

- Se completează **prima dată** când un skill atinge obiectul/rutina respectivă — fie din propria înțelegere a codului (Claude citește și rezumă ce face), fie preluat verbatim dacă utilizatorul îl dictează explicit (marchează sursa: "dictat de utilizator").
- **Nu se rescrie automat** la o atingere ulterioară doar pentru că rutina a fost modificată din nou — se actualizează Scop-ul doar dacă modificarea confirmată chiar schimbă ce face rutina (nu la un simplu fix de eroare care nu-i schimbă rolul).
- Se citește de toate skill-urile **înainte** de a propune orice modificare/optimizare ulterioară pe acel obiect — sursă de context despre intenția reală a codului, nu doar despre forma lui sintactică.

## Reguli de regenerare a `INDEX.md`

1. Rulează `list` prin `access-bridge.ps1` — sursa de adevăr pentru "ce există", niciodată un cache vechi.
2. Pentru fiecare categorie, listează obiectele care AU fișier propriu în `.accdb-progress/<Category>/` separat de cele care nu au.
3. Pentru obiectele cu fișier propriu, extrage "Scop general" (prima linie relevantă) + un rezumat de-o-linie al ultimei intrări (cea mai recentă `Lucrat:` din fișier) pentru linia din `INDEX.md`.
4. Nu modifica fișierele per-obiect la regenerarea `INDEX.md` — regenerarea e read-only asupra lor, doar `INDEX.md` se rescrie.
