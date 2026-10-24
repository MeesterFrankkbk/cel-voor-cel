# Ontwerp stap 12 – Project: mijn interactief werkblad

**Kern:** de leerling maakt een **echt** werkblad voor iemand anders: voor de eigen klas of voor een andere klas (bv. L6 voor L3). Alles uit stap 9 tot 11 komt samen. Het werk is **gedifferentieerd**: de leerling kiest zelf een niveau ★, ★★ of ★★★ en mag altijd een niveau hoger proberen.

## Waarom zo?

- **Minstens 10 vragen** met **rekenen én taal/WO/Frans …** (wens van de leerkracht). De app kan het vak niet lezen, maar wel zien of er antwoorden met een **getal** en antwoorden met een **woord** zijn: minstens 2 van elk.
- **Een vraag met een afbeelding** zit al in niveau ★: het is eenvoudig en maakt het werkblad leuk.
- Niveaus verschillen in **vaardigheden**, niet in hoeveelheid werk: 10 vragen voor iedereen.
- **Eerst goed denken en dan pas bouwen:** de lessen beginnen met *voor wie?* en *een plan op papier*.

## De drie niveaus

| | ★ Basis | ★★ Gevorderd (= ★ +) | ★★★ Expert (= ★★ +) |
|---|---|---|---|
| 1 | titel en een korte uitleg voor de gebruiker | SPATIES.WISSEN: een spatie te veel is geen fout | een eindboodschap die afhangt van de score (ALS) |
| 2 | minstens 10 vragen, elk met een antwoordcel | minstens 2 keuzelijsten, met het juiste antwoord erin | de score kleurt mee (voorwaardelijke opmaak) |
| 3 | rekenen én taal/WO: minstens 2 getallen en 2 woorden als antwoord | de score als procent | een rekenvraag die alleen getallen toelaat, met een eigen foutmelding |
| 4 | feedback per vraag met ALS, ook als de cel leeg is | rasterlijnen uit | een keuzelijst die de keuzes uit een bereik haalt |
| 5 | een score die telt hoeveel antwoorden juist zijn (AANTAL.ALS) | blad beveiligd: antwoordcellen ontgrendeld, de rest geblokkeerd | |
| 6 | juist kleurt groen, jammer kleurt rood | | |
| 7 | een afbeelding bij een vraag | | |
| 8 | de oplossingen verborgen | | |
| inleveren | de antwoordcellen zijn weer leeg | | |

Geslaagd = 80% van de opdrachten van het gekozen niveau (zoals bij alle stappen). Daarnaast toont de app de **sterren** die echt gehaald zijn: ★★ betekent dat alles van ★ en ★★ in orde is.

## Hoe vindt de app de vragen? (vrije opbouw)

Elke leerling bouwt zijn werkblad anders. De app zoekt daarom niet op vaste cellen, maar redeneert zoals een gebruiker:

1. Elke formule met **ALS** is een mogelijke **feedbackcel**.
2. De cellen waarnaar die formule verwijst en waarin geen formule staat, zijn mogelijke **antwoordcellen**.
3. De app **probeert uit**: ze typt (in gedachten) niets, een fout antwoord en de mogelijke oplossingen (de inhoud van de andere cel of een getal/woord in de formule zelf). Geeft één ervan iets anders dan "fout" en "leeg", dan is dat het **juiste antwoord**. Zo werkt het ook met eigen woorden ("goed", "super", "probeer opnieuw") en met de oplossing in een verborgen kolom, op een ander blad of in de formule zelf.
4. Wie is de antwoordcel en wie de oplossing? Voorrang voor: ontgrendeld, niet verborgen, met een keuzelijst, en de cel die "leeg" opmerkt.
5. Heeft de werkmap meerdere bladen met vragen (zoals het oude werkblad met een blad per leerling + "Blanco"), dan kiest de app het blad met de meeste vragen, en bij gelijkspel het blad met de meeste lege antwoordcellen.

Zo herkent de app ook het oude voorbeeldwerkblad (opgave in B, oplossing in de verborgen kolom C, antwoord in D, feedback in E, keuzes in de verborgen kolommen K:N).

## De app speelt gebruiker

Met een **wat-als-rekenaar** vult de app alle antwoordcellen tegelijk in en rekent het hele werkblad door (feedback → score → procent → eindboodschap):

- alles juist, alles fout, alles leeg, de helft juist, telkens één vraag juist;
- de score moet elke vraag meetellen ("je score telt C14 nog niet mee");
- het procent moet 100% zijn als alles juist is ("deel je door het juiste aantal vragen?");
- de eindboodschap moet veranderen tussen alles juist en niets juist;
- een antwoord met een spatie erachter moet juist blijven (★★), behalve bij een keuzelijst.

De **methode is vrij**: telt een leerling met punten (ALS → 1/0 en SOM), dan werkt de score ook. De opdracht noemt AANTAL.ALS als aanbevolen weg.

## Valkuilen die de app herkent

- een getal als tekst in de oplossing ('42): de gebruiker typt 42 en krijgt "jammer";
- SPATIES.WISSEN aan één kant: getallen worden "jammer";
- een lege antwoordcel geeft al "jammer";
- de score telt niet alle vragen mee, of telt ook "jammer";
- het juiste antwoord staat niet in de keuzelijst;
- het juiste antwoord past niet in de eigen getalregel (dan kan niemand het typen!);
- de antwoordkolom verborgen in plaats van de oplossingskolom;
- beveiligd terwijl de antwoordcellen nog geblokkeerd zijn;
- oplossingen in de formule zelf en een beveiligd blad: dan moet bij Bescherming ook **Verborgen** aan, anders leest de gebruiker de oplossing in de formulebalk;
- een functie die de app niet kent: dan zegt ze dat eerlijk.

## Werken in de app: drie tabbladen

1. **Leren** – Voor wie maak je het? · Kies je niveau (de keuze gebeurt hier) · Een plan op papier · Zo kijkt de app na · Testen door een klasgenoot.
2. **Tussentijds nakijken** – zo vaak opladen als je wil, met tips. Je kan hier ook van niveau wisselen.
3. **Inleveren** – eerst een korte checklist: getest door een klasgenoot (naam + de tip die je kreeg), tips verwerkt, antwoordcellen leeg. Daarna oploaden zonder tips. De leerkracht ziet het niveau, de sterren, de tester en de tip.

## Testronde door een klasgenoot (zonder extra organisatie)

De tester zit gewoon aan de laptop van de maker en volgt een **testkaart** (in de les): uitleg lezen, alles juist invullen (100%?), iets fout, iets leeg, een spatie te veel, proberen een oplossing te zien of een formule te wissen, en tot slot één top en één tip. Zo blijft het eenvoudig in de klas, en toch maakt de leerling echt iets voor iemand anders.

## Technisch

- Het niveau zit in de code van de opdrachten (seed `2~abc…`). Zo ziet de leerkracht later precies dezelfde lijst opdrachten als de leerling.
- Geen wijziging aan de server nodig: de testronde wordt bewaard in de voortgang van de leerling.
- Nieuw in de rekenmotor: ISLEEG, ISGETAL, ISTEKST, WAARDE, GELIJK, ALS.VOORWAARDEN, hele kolommen (D:D) en de nieuwe functienamen van Excel 365 (`_xlfn.`).
