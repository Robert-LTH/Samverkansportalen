# Förbättringsportalen för SharePoint Online

Detta projekt innehåller en SharePoint Framework-webbdel som låter medarbetare föreslå förbättringar och rösta fram vilka förslag som ska prioriteras. Varje användare får fem samtidiga röster som automatiskt återförs när ett förslag markeras som genomfört eller borttaget. Borttagna förslag behålls i historiken och kan sökas fram på samma sätt som aktiva förslag.

## Funktioner

- Lägg till nya förbättringsförslag direkt från webbdelens gränssnitt.
- Se hur många aktiva röster ett förslag har samt hur många röster du själv har kvar.
- Rösta på eller återta din röst från aktiva förslag (status "Föreslagen" eller "Pågående").
- Uppdatera status på förslag till "Genomförd" eller "Avslutad" (borttagen) utan att förlust av historik.
- Sök bland samtliga förslag, inklusive avslutade eller borttagna.
- Listor i SharePoint för förslag och röster provisioneras automatiskt vid första körning.

## Teknisk översikt

- **Plattform:** SharePoint Framework (SPFx) 1.17.4 med React.
- **SharePoint-listor:**
  - `ImprovementSuggestions` för själva förslagen.
  - `SuggestionVotes` för röster kopplade till varje förslag.
- **Bibliotek:** PnPjs används för att kommunicera med SharePoint och hantera listor och data.

## Kom igång

1. **Förberedelser**
   - Installera Node.js 16 (>=16.13.0 och <17). Ett `.nvmrc`-file ingår i repot så att `nvm use` / `nvm install` automatiskt växlar till en kompatibel version.
   - Klona detta repo och öppna projektmappen.
   - Kör `npm install` (projektet levereras med `.npmrc` som aktiverar `legacy-peer-deps`).

2. **Utvecklingsläge**
   - Kör `npm run watch` för att bygga om projektet kontinuerligt medan du utvecklar.
   - Vill du snabbt inspektera det byggda paketet kan du använda `npm run serve`, vilket startar en enkel statisk server för innehållet i `dist/`.

3. **Bygg för distribution**
   ```bash
   npm run clean
   npm run build
   ```
   - Resultatet hamnar i mappen `dist/`. Paketera den tillsammans med övriga SharePoint-artefakter enligt er befintliga releaseprocess.

4. **Lägg till på en sida**
   - Redigera valfri modern sida i SharePoint.
   - Lägg till webbdelens "Förbättringsportalen" och spara sidan.

## Användning

1. **Skapa förslag** – Fyll i titel och beskrivning och klicka på *Spara förslag*.
2. **Rösta** – Klicka på *Rösta* på ett aktivt förslag för att använda en av dina fem röster.
3. **Återta röst** – Klicka på *Återta röst* för att frigöra rösten innan förslaget är klart.
4. **Ändra status** – Använd statuslistan på kortet för att markera förslag som pågående, genomfört eller avslutat. När status är "Genomförd" eller "Avslutad" återlämnas röster automatiskt och förslaget finns kvar i sökningen.

## Vidareutveckling

- Lägg till egna vyer eller kolumner i listorna om ytterligare metadata behövs.
- Komplettera med notifieringar, exempelvis via Power Automate, när nya förslag skapas eller status ändras.
- Anpassa UI:t i `ImprovementPortal.styles.ts` för att matcha organisationens grafiska profil.

## Support

Om något inte fungerar, se till att kontot som kör webbdelens kod har rättigheter att skapa och uppdatera listor på webbplatsen. Kontrollera också att appen har distribuerats korrekt via appkatalogen.
