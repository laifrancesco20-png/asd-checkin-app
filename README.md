# ASD Check-in App

App di check-in per i tornei: lista iscritti, incassi, sincronizzazione tra dispositivi tramite Firebase Realtime Database e chat con Gemini.

## Chiave Gemini

La chiave di Google AI Studio si inserisce in *Impostazioni* e resta salvata solo nel browser di quel dispositivo. Non viene mai scritta nel database.

Le versioni precedenti copiavano la chiave in `/config.json` del database, leggibile da chiunque conosca l'URL. All'avvio l'app cancella quella copia.

## Regole del database

Il file `database.rules.json` contiene le regole da incollare in Firebase console → Realtime Database → Regole:

- `tournament` resta leggibile e scrivibile, perché l'app sincronizza senza login;
- `config` si può solo cancellare;
- tutto il resto è chiuso.

Nota: chi conosce l'URL del database può ancora leggere e modificare `tournament`. Per chiuderlo davvero serve aggiungere un login (Firebase Authentication).
