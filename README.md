# Cel voor Cel

Stap voor stap Excel leren voor leerlingen van 9 tot 13 jaar. De leerling leert in de app, oefent in het echte Excel, laadt het bestand op en krijgt meteen feedback. De leerkracht volgt alles op.

## Online zetten (zoals Kangoeroe-Kluis)

1. Maak een nieuwe repository op GitHub en zet alle bestanden van deze map erin.
2. In Netlify: **Add new site → Import an existing project** en kies de repository. De instellingen staan al in `netlify.toml`.
3. In Netlify: **Site configuration → Environment variables** → voeg toe:
   - `LEERKRACHT_WACHTWOORD` = het wachtwoord voor het leerkrachtoverzicht.
4. Klik op **Deploy**. Klaar.

- Leerlingen: `https://<jouw-site>.netlify.app/`
- Leerkracht: `https://<jouw-site>.netlify.app/leerkracht.html`

De opslag (leerlingen, voortgang, opgeladen Excel-bestanden) gebeurt in **Netlify Blobs**. Daar hoef je niets voor in te stellen.

Open je `index.html` zonder Netlify (bv. lokaal), dan start de app in **demomodus**: alles blijft in die browser, het leerkrachtwachtwoord is dan `demo`.

## De 12 stappen

1. Rijen, kolommen en cellen
2. Typen, verbeteren en kolombreedte
3. Cellen opmaken
4. Excel vult zelf aan (vulgreep, reeksen, maanden, dagen, honderdveld)
5. Getallen tonen (procent, euro, decimalen, datum, komma)
6. Je eerste formules (=, + − * /, celadressen, doortrekken, haakjes, foutmeldingen)
7. Functies (SOM, GEMIDDELDE, MIN, MAX, MEDIAAN, AANTAL, AFRONDEN)
8. Project: mijn maandrapport (oefenrapport met gegeven punten, daarna eigen rapport met echte punten)
9. Excel kiest: ALS (juist/jammer, lege cel, ALS in een ALS, boodschap bij een score)
10. Keuzelijsten en kleuren die reageren (gegevensvalidatie, eigen foutmelding, voorwaardelijke opmaak, kleuren in het rapport)
11. Klaar voor een gebruiker (AANTAL.ALS, SPATIES.WISSEN, score in procent, afbeelding, kolom verbergen, rasterlijnen, ontgrendelen en beveiligen)
12. Project: mijn interactief werkblad (eigen werkblad voor de eigen of een andere klas, minstens 10 vragen met rekenen én taal/WO, drie niveaus ★ ★★ ★★★, tussentijds nakijken, testronde door een klasgenoot, inleveren)

Het ontwerp van elke stap staat in de map `docs/`.

## Hoe het werkt

- **Aanmelden**: voornaam, naam, klas en een eigen code (4 tot 8 cijfers). De eerste keer maakt de leerling zelf zijn account. Code vergeten? De leerkracht stelt een nieuwe in.
- **Elke stap** heeft drie delen: *Leren* (uitleg met korte oefeningen in de app, met voorleesknop), *Oefenen in Excel* (met hulplijnen) en *Test* (nieuwe opdrachten, geen hulplijnen, 80% nodig).
- **Ik kan dit al**: een leerling mag meteen naar de test. De volgende stap gaat open als de test gehaald is.
- **Elke leerling krijgt andere opdrachten** (andere cellen, woorden, kleuren). De leerkracht ziet precies welke.
- **Opladen**: twee routes (map OneDrive zichtbaar of niet). De app waarschuwt bij een oud bestand of een onveranderd startbestand.
- **Stap 12** is gedifferentieerd. De leerling kiest een niveau (★ Basis, ★★ Gevorderd, ★★★ Expert) en mag altijd wisselen. Bij *Tussentijds nakijken* laadt hij zo vaak op als hij wil, met tips. Bij *Inleveren* vult hij eerst in wie zijn werkblad testte en welke tip hij kreeg. Het werkblad is goedgekeurd als alles van ★ in orde is; de leerkracht ziet het niveau, de behaalde sterren, de tester en de tip.
- **Vrije opbouw in stap 12**: de app zoekt zelf de vragen (elke ALS-formule die naar een antwoordcel kijkt) en speelt daarna gebruiker: ze vult in gedachten alles juist, fout en leeg in en rekent het hele werkblad door (feedback, score, procent, eindboodschap, kleuren). Oplossingen mogen in een verborgen kolom, op een verborgen blad of in de formule zelf staan.
- **Terugblik**: na elke geslaagde test kiest de leerling met één klik hoe het ging (😊 makkelijk, 🙂 goed te doen, 😓 moeilijk) en mag hij kort schrijven wat nieuw was. Na het eindproject volgt een grotere terugblik (trots op, moeilijkst, volgende keer anders). De leerkracht ziet dit in het overzicht, bij de leerling en in de CSV.
- **Verder werken**: het laatst opgeladen bestand is op elk toestel opnieuw te downloaden.

## Wat de app in een Excel-bestand controleert

Afbeeldingen (zwevend en in een cel), verborgen kolommen, rasterlijnen, geblokkeerde cellen en bladbeveiliging, gegevensvalidatie (keuzelijsten, gehele getallen, foutmelding) en voorwaardelijke opmaak (de app probeert de regels uit met voorbeeldwaarden), formules (de app rekent ze zelf na, ook doorgetrokken formules: juiste adressen, juiste functie, juist bereik, getal getypt in plaats van formule, = vergeten, x in plaats van *), celinhoud, reeksen (ook gekopieerd in plaats van doorgevuld, te ver gesleept, Engelse maandnamen), honderdveld, getalnotatie (procent, euro, aantal decimalen, datum), getal of tekst (komma/punt), bladnaam, bestandsnaam, ingevoegde rijen, twee lijnen in een cel (Alt + Enter), kolombreedte, vet/schuin/onderstreept, letter- en opvulkleur (ook themakleuren), uitlijning, rijhoogte, samengevoegde cellen, lettergrootte en randen. De lezer kan ook al formules, gegevensvalidatie, voorwaardelijke opmaak en bladbeveiliging uitlezen, voor de volgende stappen.

## Bestanden

| Bestand | Wat |
|---|---|
| `index.html`, `js/app.js` | leerling-app |
| `leerkracht.html`, `js/leerkracht.js` | leerkrachtoverzicht |
| `js/modules.js` | **inhoud**: lessen, opdrachten, hulplijnen en controles per stap |
| `js/xlsx-check.js` | leest een .xlsx-bestand (waarden en opmaak) |
| `js/formula.js` | rekent formules na (Nederlands en Engels) |
| `js/xlsx-write.js` | maakt de startbestanden |
| `js/api.js` | praat met de opslag (of demomodus) |
| `netlify/functions/api.mjs` | opslag op Netlify |
| `css/app.css` | uiterlijk |

## Goed om te weten

- Het nakijken gebeurt in de browser van de leerling. De leerkracht krijgt altijd ook het bestand zelf, om na te kijken.
- Een tikfout in naam of klas maakt een tweede account. Verwijder het dubbele account in het leerkrachtoverzicht.
- De kolombreedte-controle is een schatting met wat speling.
