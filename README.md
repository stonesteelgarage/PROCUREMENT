# Procurement Intelligence MVP

Dashboard locale Python/Streamlit per:
1. Importare vendor list storiche in un database SQLite locale.
2. Generare una vendor list compilata partendo da un template Excel.
3. Fare scouting base su URL pubblici con supporto OpenAI.
4. Consultare la memoria storica vendor.

## Installazione su Mac

Apri Terminale nella cartella del progetto e lancia:

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Poi apri `secrets.py` e inserisci la tua chiave OpenAI:

```python
OPENAI_API_KEY = "sk-..."
```

## Avvio

```bash
streamlit run app.py
```

## Database

Il database locale viene creato automaticamente:

```text
procurement_intelligence.db
```

## Nota

Questa è una prima versione MVP. Lo scouting automatico web via motore di ricerca dedicato potrà essere aggiunto dopo senza rifare l'architettura.
