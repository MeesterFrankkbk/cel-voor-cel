# Ontwerp stap 6 en 7

## Stap 6 – Je eerste formules

**Kern:** een formule begint met `=` en rekent met **celadressen**, niet met getallen. Verander je een getal, dan rekent Excel vanzelf opnieuw. Trek je een formule door, dan schuiven de adressen mee.

| Les | Wat de leerling ontdekt | Oefenvorm in de app |
|---|---|---|
| Excel als rekenmachine | `=`, `+ - * /` (geen × of :), zonder `=` is het tekst | mini-rekenblad: zelf formules typen |
| Rekenen met celadressen | `=B1*B2` in plaats van `=2,5*4`; adres aanklikken tijdens het typen; getal veranderen → resultaat verandert mee | mini-rekenblad |
| Formules doortrekken | vulgreep op een formule: `=A2*B2` wordt `=A3*B3` … | vulgreep-simulator met formules + formulebalk |
| Voorrang en haakjes | `=2+3*4` is 14, `=(2+3)*4` is 20; gemiddelde van twee getallen | quiz |
| Foutmeldingen | `#DEEL/0!`, `#NAAM?`, `#####` | quiz + mini-rekenblad (fout oplossen) |

**Opdrachten in Excel** (startbestand "Inkopen voor het klasfeest": artikel, prijs per stuk, aantal; budget; aantal leerlingen; elke leerling andere artikels en getallen):
1. Totaal van de eerste rij: prijs × aantal (`=B4*C4`).
2. De formule doortrekken tot de laatste rij.
3. Alle totalen samen optellen met `+` (stap 7 toont daarna hoe het korter kan met SOM).
4. Wat blijft er over van het budget? (`=G3-D11`)
5. Kost per leerling (`=D11/G6`), in euro met 2 decimalen (herhaling stap 5).
6. Het aantal leerlingen veranderen en zien dat alles mee verandert.

**Controles:** de app rekent elke formule zelf na (ook als Excel ze als "gedeelde formule" bewaart na doortrekken). Foutmeldingen voor: getal getypt in plaats van formule, formule met getallen in plaats van adressen, `=` vergeten (tekst), verkeerde rij in de formule, × in plaats van `*`, foute uitkomst.

## Stap 7 – Functies

**Kern:** Excel heeft ingebouwde rekenhulpjes. `=SOM(D4:D9)` is korter en veiliger dan `=D4+D5+…`.

| Les | Wat de leerling ontdekt | Oefenvorm in de app |
|---|---|---|
| Een functie | naam, haakjes, bereik; AutoSom-knop Σ | mini-rekenblad |
| SOM, GEMIDDELDE, MIN, MAX | wat elke functie doet | mini-rekenblad |
| MEDIAAN en AANTAL | het middelste getal; AANTAL telt alleen getallen (een lege cel telt niet mee) | quiz + mini-rekenblad |
| AFRONDEN en de puntkomma | `=AFRONDEN(GEMIDDELDE(B2:B6);1)`, in België `;` tussen de delen | mini-rekenblad |
| Functies in een andere taal | `#NAAM?` bij een tikfout; SOM heet SUM in een Engelse Excel | quiz |

**Opdrachten in Excel** (startbestand "Sportdag: touwtjespringen", 10 leerlingen, één afwezige):
SOM, GEMIDDELDE, AFRONDEN(GEMIDDELDE), MAX, MIN, MEDIAAN, AANTAL.

**Controles:** juiste functie, juist bereik (te kort, te lang, met de kop erbij), uitkomst klopt, het antwoord niet zelf uitgerekend en getypt.
