import docx
import os
import re

# 1. Carica il documento (assicurati che il file sia nella stessa cartella)
doc_path = 'documents/testotesi.docx'
if not os.path.exists(doc_path):
    # Prova a cercarlo nella root se non è in documents
    doc_path = 'testotesi.docx'

doc = docx.Document(doc_path)
full_text = [p.text for p in doc.paragraphs]
content = "\n".join(full_text)

# 2. Trova il punto esatto dove inizia la Sezione 3
# Cerchiamo "3  Analisi del tracciato" o "3. IL TRACCIATO"
match = re.search(r'(?m)^3\s+Il tracciato della Via Valeria', content, re.IGNORECASE) or \
        re.search(r'(?m)^3\s+IL TRACCIATO', content, re.IGNORECASE)

if match:
    analisi_integrale = content[match.start():]
    
    # Crea la cartella se non esiste
    os.makedirs('docs/frammenti_sparsi', exist_ok=True)
    
    # 3. Salva il file integrale
    with open('docs/frammenti_sparsi/03_analisi_tracciato.md', 'w', encoding='utf-8') as f:
        f.write("# " + analisi_integrale)
    
    print("SUCCESSO: Il file integrale è stato creato in docs/frammenti_sparsi/03_analisi_tracciato.md")
else:
    print("ERRORE: Non ho trovato l'inizio della Sezione 3. Controlla il titolo nel Word.")