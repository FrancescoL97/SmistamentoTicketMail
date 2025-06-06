Suite di Test per `smista_ticket.py`

Contenuto:
----------
La suite di test `test_smista_ticket.py` testa le seguenti funzioni del modulo:

1. ricerca_id(mail)
   - Estrae l'ID del ticket dal subject o dal body di una mail.
   - Copre casi con:
     - ID nel subject
     - ID nel body
     - Nessun ID

2. valutazione_stato(mail)
   - Determina lo stato del ticket in base al subject.
   - Ritorna: "aperto", "chiuso", "ignoto"

3. trova_mail_collegate(mail, ticket_id)
   - Verifica se una mail è collegata a un determinato ID ticket (subject o body).

Struttura file:
---------------
test_smista_ticket.py

- test_ricerca_id_subject()
- test_ricerca_id_body()
- test_ricerca_id_none()
- test_valutazione_stato_chiuso()
- test_valutazione_stato_aperto()
- test_valutazione_stato_ignoto()
- test_trova_mail_collegate_subject()
- test_trova_mail_collegate_body()
- test_trova_mail_collegate_none()

Come eseguire i test:
---------------------
1. Posizionati nella directory del progetto
2. Esegui:
   pytest test_smista_ticket.py

Prerequisiti:
-------------
- Python 3.11
- pytest
- pluggy, iniconfig, packaging, colorama, pygments (già installati)

Obiettivo:
----------
Garantire che le funzioni core del parser ticket siano robuste, affidabili e testabili in isolamento.

