# Întrebări Gramaticescu (repository public)

Aici stau **doar întrebările** jocului Gramaticescu. Jocul are repository-ul lui, privat.

Când urci o versiune nouă a fișierului **questions.xlsx**, GitHub verifică întrebările, le transformă în JSON și le publică pe GitHub Pages. Aplicația de pe telefon citește de acolo întrebările la pornire, deci corecturile ajung la jucători fără o actualizare în App Store sau Google Play.

Întrebările din acest repository sunt publice: oricine le poate citi. Codul jocului rămâne privat.

## Ce conține

| Fișier | Ce face |
| --- | --- |
| `questions.xlsx` | Întrebările (foaia „Întrebări”) și regulile pe teme (foaia „Reguli”) |
| `scripts/build_questions.py` | Verifică Excelul și generează fișierele pentru aplicație |
| `.github/workflows/publica.yml` | Automatizarea GitHub Actions |
| `data/` | Creat automat: copia JSON a întrebărilor, pentru istoric. Nu-l edita. |

Ce se publică pe `https://NUMELE-TĂU.github.io/gramaticescu-intrebari/`:

- `manifest.json`: versiunea, un fișier foarte mic pe care aplicația îl verifică la pornire;
- `questions.json`: toate întrebările și regulile;
- o pagină simplă de control, cu versiunea publicată și numărul de întrebări pe temă.

---

## Prima publicare (o singură dată)

### 1. Creează repository-ul

1. Pe [github.com](https://github.com): sus în dreapta **+** → **New repository**.
2. **Repository name:** `gramaticescu-intrebari`.
3. Alege **Public**. GitHub Pages e gratuit doar pentru repository-urile publice.
4. **Create repository**.

### 2. Urcă fișierele

1. Apasă **uploading an existing file** (sau **Add file** → **Upload files**).
2. Trage tot conținutul folderului `gramaticescu-intrebari`: `questions.xlsx`, `README.md` și folderele `scripts` și `.github`.
   - Pe Mac, folderul `.github` e ascuns. În Finder apasă **Cmd + Shift + .** ca să-l vezi.
3. **Commit changes**.

Verifică după urcare că vezi în listă folderul `.github`. Dacă lipsește, urmează pasul de rezervă:

<details>
<summary>Pas de rezervă: creează manual fișierul automatizării</summary>

1. **Add file** → **Create new file**.
2. La nume scrie exact: `.github/workflows/publica.yml` (cu barele oblice, GitHub face singur folderele).
3. Copiază în fereastră tot conținutul fișierului `publica.yml` din pachet.
4. **Commit changes**.

</details>

### 3. Pornește GitHub Pages

1. **Settings** (sus) → în meniul din stânga, la „Code, planning, and automation”, **Pages**.
2. La **Build and deployment** → **Source**, alege **GitHub Actions**.

### 4. Publică

1. Tabul **Actions** → în stânga **Publică întrebările** → **Run workflow** → **Run workflow**.
2. După 1–2 minute apare bifa verde.
3. Deschide `https://NUMELE-TĂU.github.io/gramaticescu-intrebari/`: trebuie să vezi pagina de control, cu versiunea și numărul de întrebări.

Adresa asta o pui apoi în jocul din repository-ul privat (vezi README-ul de acolo).

---

## Cum corectezi sau adaugi întrebări

1. Deschide `questions.xlsx` în Excel și fă modificările.
2. Pe GitHub: **Add file** → **Upload files** → trage `questions.xlsx` → **Commit changes**. Fișierul vechi se înlocuiește.
3. Automatizarea pornește singură. După 1–2 minute, jucătorii primesc întrebările noi la următoarea pornire a aplicației.

### Reguli pentru Excel

| Coloană | Ce scrii |
| --- | --- |
| **ID** | Un număr unic. **Nu schimba ID-ul unei întrebări existente.** Pentru o întrebare nouă folosește următorul număr liber. |
| **Temă** | Alege din listă. Temele și regulile lor sunt în foaia „Reguli”. |
| **Întrebarea** | Propoziția, cu exact un spațiu gol scris `___` (trei liniuțe jos). |
| **Răspuns corect** / **Răspuns greșit** | Cele două variante, diferite între ele. |
| **Activă** | `da` sau `nu`. |

- **Corectezi o întrebare:** schimbi textul, dar păstrezi ID-ul. Provocarea zilei și nivelurile rămân aceleași pentru toți.
- **Scoți o întrebare:** pui `nu` la Activă. **Nu șterge rândul**, altfel s-ar strica istoricul din calendarul jucătorilor.
- **Adaugi o întrebare:** un rând nou, cu următorul ID liber. Apare ca nivel nou la capătul hărții imediat, iar în provocarea zilei după 7 zile, ca toți jucătorii să aibă deja lista nouă.
- **Adaugi o temă:** un rând nou în foaia „Reguli”, cu numele temei și regula afișată copilului după o greșeală.

### Dacă ceva nu e în regulă

- În tabul **Actions** apare un **X roșu** în loc de bifă verde.
- Apasă pe rularea respectivă: vezi exact rândul greșit, de exemplu *„Foaia «Întrebări», rândul 12: răspunsul corect și cel greșit sunt identice”*.
- Corectezi Excelul și îl urci din nou. Până atunci jucătorii păstrează lista anterioară, deci nu ajunge nimic stricat la ei.
- **Avertismentele** (de exemplu „variantele sunt lungi”) nu opresc publicarea. Le vezi în rezumatul rulării.

### Istoricul modificărilor

GitHub nu arată ce s-a schimbat într-un fișier Excel. Automatizarea salvează însă după fiecare publicare `data/questions.json`. Deschide-l și apasă **History**: vezi exact ce întrebare s-a modificat, când și cum.

## Rulare locală (opțional)

```
pip install openpyxl
python3 scripts/build_questions.py
```

Fișierele generate apar în folderul `site/`.
