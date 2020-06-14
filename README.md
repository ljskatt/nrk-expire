# NRK-Expire
Dette prosjektet er inspirert av [NRK-DL](https://github.com/ljskatt/nrk-dl)

Scriptene i prosjektet gjør det mulig å hente ut rapporter utifra NRK på hvilke programmer/serier som er på vei ut (i ferd med å bli slettet fra NRK), slik at du eventuelt kan se på programmene/seriene før de blir borte, eller benytte [NRK-DL](https://github.com/ljskatt/nrk-dl) til å laste ned programmene selv.

## Teknisk info

Scriptene oppretter en mappe til cache av filer `nrk-cache`, denne er nyttig hvis du skal generere begge rapportene, eller tilpasse rapportene, da unngår man å sende mange requests til NRK, samtidig som at det tar kortere tid å generere rapporter.

Om du ønsker å laste ned ferske data, så er det bare å slette `nrk-cache` før du kjører scriptene.

Den lager rapporter utifra programmer/serier som:

- Ikke har gått ut
- Utgår i løpet av 12 måneder fra tidspunktet rapporten ble kjørt (Kan tilpasses)

## Har dere ferdig genererte rapporter?
Ja, det har vi:

- [Kalender rapport](https://ljskatt.no/nrk-expire.ics)
Kan linkes til kalender ved å legge til URL basert kalender: `https://ljskatt.no/nrk-expire.ics`

## Excel
Excel-scriptet henter ut informasjon fra NRK sitt API og legger det i en `.xlsx` (Excel) fil

`.\nrk-expire-excel.ps1`

## Kalender
Ikke i repository enda