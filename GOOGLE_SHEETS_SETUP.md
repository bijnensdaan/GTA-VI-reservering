# Gedeelde opslag via Google Sheets

De app ondersteunt een Google Sheet als gedeelde opslag. De Sheet blijft de bron van waarheid; de loting en controle op overlappende boekingen gebeuren in Apps Script, zodat aanvragen vanaf verschillende apparaten dezelfde uitkomst zien.

## Eenmalige koppeling

1. Maak in Google Drive een lege Google Sheet, bijvoorbeeld **PlayStation reserveringen**.
2. Open **Extensies → Apps Script**. Vervang de inhoud van `Code.gs` door de code uit dit projectbestand en sla op.
3. Kies in de Apps Script-editor de functie `setup` en druk op **Uitvoeren**. Geef de gevraagde Google-toestemming. Dit maakt de tabbladen `Reserveringen` en `Instellingen` aan.
4. Ga naar **Implementeren → Nieuwe implementatie → Web-app**. Stel **Uitvoeren als** in op **Ik** en **Toegang** op **Iedereen**. Implementeer en kopieer de URL die eindigt op `/exec`.
5. Plak die URL in `config.js` als waarde van `window.PS_API_URL`.
6. Publiceer de appbestanden (`index.html`, `style.css`, `app.js` en `config.js`) op GitHub Pages. `Code.gs` is de servercode en hoeft niet op Pages te staan.
7. Open de Pages-link, vul de drie namen in en sla ze op. Daarna gebruikt ieder apparaat met diezelfde Pages-link de gedeelde Sheet.

## Toegang en privacy

Met de instelling **Iedereen** kan iedereen die de web-app-URL kent de reserveringen lezen en acties uitvoeren. Omdat de Pages-link openbaar is en de Apps Script-URL in `config.js` staat, behandel dit als een openbare familieplanner: zet er geen privégegevens in. Deel de Google Sheet zelf alleen met wie hem mag bekijken of beheren. De app gebruikt geen wachtwoorden en heeft geen gebruikersaccounts.

Als je Apps Script later opnieuw implementeert, controleer dan of de `/exec`-URL is veranderd en werk `config.js` zo nodig bij.
