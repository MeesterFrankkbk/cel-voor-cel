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

## Hoe het werkt

- **Aanmelden**: voornaam, naam, klas en een eigen code (4 tot 8 cijfers). De eerste keer maakt de leerling zelf zijn account. Code vergeten? De leerkracht stelt een nieuwe in.
- **Elke stap** heeft drie delen: *Leren* (uitleg met korte oefeningen in de app, met voorleesknop), *Oefenen in Excel* (met hulplijnen) en *Test* (nieuwe opdrachten, geen hulplijnen, 80% nodig).
- **Ik kan dit al**: een leerling mag meteen naar de test. De volgende stap gaat open als de test gehaald is.
- **Elke leerling krijgt andere opdrachten** (andere cellen, woorden, kleuren). De leerkracht ziet precies welke.
- **Opladen**: twee routes (map OneDrive zichtbaar of niet). De app waarschuwt bij een oud bestand of een onveranderd startbestand.
- **Verder werken**: het laatst opgeladen bestand is op elk toestel opnieuw te downloaden.

## Wat de app in een Excel-bestand controleert

Celinhoud, bladnaam, bestandsnaam, ingevoegde rijen, twee lijnen in een cel (Alt + Enter), kolombreedte, vet/schuin/onderstreept, letter- en opvulkleur (ook themakleuren), uitlijning, rijhoogte, samengevoegde cellen, lettergrootte en randen. De lezer kan ook al formules, gegevensvalidatie, voorwaardelijke opmaak en bladbeveiliging uitlezen, voor de volgende stappen.

## Bestanden

| Bestand | Wat |
|---|---|
| `index.html`, `js/app.js` | leerling-app |
| `leerkracht.html`, `js/leerkracht.js` | leerkrachtoverzicht |
| `js/modules.js` | **inhoud**: lessen, opdrachten, hulplijnen en controles per stap |
| `js/xlsx-check.js` | leest een .xlsx-bestand (waarden en opmaak) |
| `js/xlsx-write.js` | maakt de startbestanden |
| `js/api.js` | praat met de opslag (of demomodus) |
| `netlify/functions/api.mjs` | opslag op Netlify |
| `css/app.css` | uiterlijk |

## Goed om te weten

- Het nakijken gebeurt in de browser van de leerling. De leerkracht krijgt altijd ook het bestand zelf, om na te kijken.
- Een tikfout in naam of klas maakt een tweede account. Verwijder het dubbele account in het leerkrachtoverzicht.
- De kolombreedte-controle is een schatting met wat speling.
