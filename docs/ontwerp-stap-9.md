# Ontwerp stap 9 – Excel kiest: ALS

**Kern:** met ALS laat je Excel zelf beslissen wat er in een cel komt. Dat is het hart van elk interactief werkblad: Excel vergelijkt het antwoord met de oplossing en zegt "juist" of "jammer".

`=ALS(voorwaarde; wat als het klopt; wat als het niet klopt)` — drie delen, met een puntkomma ertussen. Tekst staat tussen aanhalingstekens.

## Lessen

| Les | Wat de leerling ontdekt | Oefenvorm |
|---|---|---|
| Excel beslist | de drie delen van ALS; een score verandert en de boodschap verandert mee | mini-rekenblad |
| Vergelijken | = &lt;&gt; &lt; &gt; &lt;= &gt;=; "minstens 5" is >=5 | quiz |
| Juist of jammer? | het antwoord vergelijken met de oplossing: `=ALS(C2=B2;"juist";"jammer")` | mini-rekenblad |
| Eerst kijken of er iets staat | een lege cel geeft meteen "jammer"; oplossing: ALS in een ALS met `C2=""` | mini-rekenblad |
| Tekst vergelijken | hoofdletters tellen niet, een spatie wel; "400 cm" is niet 400 | quiz |

## Opdrachten in Excel (startbestand "Mijn eerste interactief werkblad")

Zes vragen met oplossing en (ingevulde) antwoorden van een proefleerling: twee juist, twee fout, twee leeg.

1. In D4 een ALS die "juist" of "jammer" toont.
2. De formule doortrekken tot D9.
3. De formule verbeteren: is het antwoord leeg, dan "vul in" (ALS in een ALS), en opnieuw doortrekken.
4. Een boodschap bij de score: "Goed gedaan!" vanaf een bepaalde score, anders "Oefen nog wat".

## Controle: de app probeert de formule zelf uit

De app vult in gedachten zelf antwoorden in (het juiste, een fout, niets) en kijkt wat de formule dan toont. Zo wordt de **logica** nagekeken, niet alleen de uitkomst van dat ene moment.

Foutmeldingen voor: geen ALS gebruikt, tekst zonder aanhalingstekens (#NAAM?), juist en jammer omgewisseld, de verkeerde cellen vergeleken, de lege cel vergeten, > in plaats van >= (bij precies de grens).
