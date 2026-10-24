# Ontwerp stap 4 en 5

## Stap 4 – Excel vult zelf aan

**Kern:** Excel herkent een patroon en zet het voort. Je hoeft het patroon maar één keer te tonen.

| Les | Wat de leerling ontdekt | Oefenvorm in de app |
|---|---|---|
| De vulgreep | het vierkantje rechtsonder; slepen kopieert of zet voort | vulgreep-simulator: "ja" kopiëren, 1-2 → 10, 5-10 → 40 |
| Excel zoekt het patroon | twee startwaarden tonen de stap | quiz: wat volgt er? (5, 10 / 20, 18 / Les 1 / één getal) |
| Eén getal is een kopie | één getal slepen = kopiëren; Ctrl ingedrukt of *Opvulopties → Reeks doorvoeren* = tellen | quiz + simulator met Ctrl |
| Lijsten die Excel kent | maanden, dagen, afkortingen; Excel moet in het Nederlands staan | simulator: januari → december, maandag → zondag, Les 1 → Les 6 |
| Een honderdveld in een paar seconden | kolom A (1, 11 …), kolom B (2, 12 …), dan beide samen naar rechts slepen: elke rij telt zelf verder | simulator in drie stappen |

**Opdrachten in Excel** (startbestand met kopjes in rij 1, elke leerling andere getallen):
1. Tellen van 1 tot N (N tussen 30 en 50) in kolom A.
2. De tafel van k (10 getallen) in kolom C.
3. Aftellen van S tot 1 in kolom E.
4. Twaalf maanden in kolom G (test: beginnen bij een andere maand en doorlopen).
5. Twee weken (14 dagen) in kolom I, vanaf een gegeven dag.
6. Een honderdveld 1–100 in K2:T11 (test: 101–200).

**Controles en foutmeldingen** (de app ziet alleen het resultaat, niet het slepen; daarom zijn de reeksen te lang om met de hand te typen):
- leeg / begint in de verkeerde rij ("begint in A3, moet in A2");
- alle cellen hetzelfde getal → "Excel kopieerde het getal: selecteer eerst de twee eerste getallen";
- alle maanden hetzelfde → "Excel herkent 'januari' niet. Staat Excel in het Nederlands?";
- eerste fout: "In A17 verwacht ik 16, maar daar staat 18";
- te ver gesleept: "Na 40 moet je stoppen: A42 moet leeg zijn";
- honderdveld gekanteld (1–10 onder elkaar) krijgt een eigen uitleg.

## Stap 5 – Getallen tonen

**Kern:** Excel bewaart een getal en kiest los daarvan hoe het eruitziet. 0,675 en 68% is hetzelfde getal. Dat is nodig voor het maandrapport.

| Les | Wat de leerling ontdekt | Oefenvorm in de app |
|---|---|---|
| Wat Excel bewaart en wat je ziet | cel en formulebalk verschillen | weergave-simulator: knoppen %, €, meer/minder decimalen |
| Komma of punt? | in België typ je 2,5; getallen staan rechts, tekst links | quiz |
| Procent | 0,75 = 75%; 68% is eigenlijk 0,675 | quiz |
| Datums | 15/3/2027 typen; Excel telt dagen (1 = 1 januari 1900) | quiz |

**Opdrachten in Excel** (startbestand met getallen, elke leerling andere waarden):
1. Toetsscore als procent zonder decimalen.
2. Toetsscore als procent met 1 decimaal.
3. Twee prijzen in euro met 2 decimalen (twee cellen samen selecteren).
4. Een getal met minder decimalen tonen, **zonder** het opnieuw te typen.
5. Een datum typen.
6. Een kommagetal typen (lengte in meter).

**Controles:** getalnotatie (procent, euro, aantal decimalen, datum) én de waarde zelf. Typt een leerling 3,14 in plaats van de notatie aan te passen, dan zegt de app dat het getal zelf 3,14159 moet blijven. Typt hij 1.45, dan wordt dat tekst en legt de app de komma uit.
