# Preventivatore Web

App web per preventivi e fatture da prezziario Excel.
Funziona su qualsiasi browser — iPhone, Mac vecchio, PC.

## Avvio locale (per testare)

```bash
pip install -r requirements.txt
python app.py
```
Apri http://localhost:5000

## Struttura

```
preventivatore-web/
├── app.py           ← server Flask (tutto il backend)
├── index.html       ← interfaccia (tutto il frontend)
├── requirements.txt
├── render.yaml      ← config deploy Render.com
└── data/
    ├── prezziario_1.xlsx   ← il tuo file Excel
    ├── prezziario_2.xlsx   ← secondo file (opzionale)
    ├── clienti.json        ← generato automaticamente
    └── impostazioni.json   ← generato automaticamente
```

## Deploy su Render.com (gratis, link pubblico)

1. Crea repo GitHub, carica tutti i file **inclusi gli Excel** nella cartella `data/`
2. Vai su [render.com](https://render.com) → New → Web Service
3. Collega il repo GitHub
4. Render legge `render.yaml` e configura tutto da solo
5. Clicca **Deploy** — in 2 minuti hai un link tipo `https://preventivatore.onrender.com`

### Aggiornare i prezziari in futuro
Sostituisci i file Excel nel repo GitHub → Render fa il redeploy automatico.

## Salvare come app su iPhone

1. Apri il link su Safari iPhone
2. Tocca il tasto **Condividi** (quadrato con freccia su)
3. Scorri → **Aggiungi a schermata Home**
4. Conferma → appare l'icona come un'app normale

## Note
- I dati clienti e impostazioni sono salvati sul server in `data/`
- Su Render.com il piano gratuito "dorme" dopo 15 min di inattività
  (primo accesso dopo pausa: ~30 secondi di attesa)
- Per uso continuativo considera il piano Starter da $7/mese
