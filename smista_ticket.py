# ───────────────────────────────────────────────────────────────
# Script: smista_ticket.py
# Descrizione: Riconosce mail di ticket da Outlook, ne valuta lo stato
# (aperto/chiuso) e le sposta nelle cartelle corrispondenti. Se un ticket è
# chiuso, sposta automaticamente anche tutte le mail correlate.
# Autore: Francesco Labianca
# Data: 2025-06-04
# ───────────────────────────────────────────────────────────────

import win32com.client
import re
import time

# === Collegamento a Outlook e alle cartelle ===
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # Posta in arrivo

ticket_folder = inbox.Folders["TICKET"]
aperti_folder = ticket_folder.Folders["APERTI"]
chiusi_folder = ticket_folder.Folders["CHIUSI"]

# === Parole chiave per valutazione stato ===
KEYWORDS_APERTO = [
    "opened",
    "comment",
    "commento",
    "presa in carico",
    "take in charge",
    "approved"
]

KEYWORDS_CHIUSO = [
    "closed",
    "resolved",
    "completed"
]

# ───────────────────────────────────────────────────────────────
def ricerca_id(mail):
    '''
    Estrae l'ID del ticket (REQxxxx, INCxxxx o RITMxxxx) dal subject o body dell'email.

    :param mail: Oggetto mail di Outlook (MailItem)
    :return: ID del ticket come stringa (es. "REQ0001234") o None se non trovato
    '''

    ticket_regex = re.compile(r"(REQ\d+|INC\d+|RITM\d+)", re.IGNORECASE)
    subject = mail.Subject
    body = mail.Body

    match_sub = ticket_regex.search(subject)
    if not match_sub:
        match_bod = ticket_regex.search(body)
        if match_bod:
            return match_bod.group(1).upper()
    else:
        return match_sub.group(1).upper()

# ───────────────────────────────────────────────────────────────
def valutazione_stato(mail):
    '''
    Determina lo stato della mail analizzando il subject. Lo stato può essere:
    - "aperto" se contiene parole chiave di apertura/commento
    - "chiuso" se contiene parole chiave di chiusura
    - "ignoto" se non trova nulla

    :param mail: Oggetto mail di Outlook (MailItem)
    :return: Stringa 'aperto', 'chiuso' o 'ignoto'
    '''
    subject = mail.Subject.lower()
    if any(k in subject for k in KEYWORDS_CHIUSO):
        return "chiuso"
    elif any(k in subject for k in KEYWORDS_APERTO):
        return "aperto"
    else:
        return "ignoto"

# ───────────────────────────────────────────────────────────────
def trova_mail_collegate(mail, ticket_id):
    '''
    Verifica se una mail è collegata a un ticket (tramite ID nel subject o body).
    Utile per spostare tutte le mail correlate quando un ticket è chiuso.

    :param mail: Oggetto mail di Outlook (MailItem)
    :param ticket_id: ID del ticket da cercare (es. "REQ0001234")
    :return: True se la mail è collegata, False altrimenti
    '''
    try:
        id_regex = re.compile(ticket_id, re.IGNORECASE)
        return bool(id_regex.search(mail.Subject) or id_regex.search(mail.Body))
    except Exception as e:
        print(f"Errore durante controllo correlazione mail: {e}")
        return False

# ───────────────────────────────────────────────────────────────
# MAIN: smistamento delle mail in base allo stato
# ───────────────────────────────────────────────────────────────
def smista_ticket():
    '''
    Scansiona tutte le email nella cartella 'TICKET' e le smista nelle sottocartelle
    'APERTI' o 'CHIUSI' in base allo stato del ticket rilevato.

    Per ogni mail:
    - Estrae l'ID del ticket (es. REQ123456, INC123456, RITM123456)
    - Determina se è 'aperto' o 'chiuso' in base a parole chiave nel subject
    - Sposta la mail nella cartella corretta ('APERTI' o 'CHIUSI')

    In caso di ticket chiuso:
    - Cerca eventuali altre email collegate già presenti in 'APERTI'
    - Se trovate, le sposta anch'esse in 'CHIUSI'

    :return: La funzione non restituisce nulla. Opera direttamente sulle email di Outlook.
    '''
    tickets_da_valutare = ticket_folder.Items
    ticket_aperti = aperti_folder.Items

    for ticket in tickets_da_valutare:
        id_valutato = ricerca_id(ticket)

        if id_valutato:
            stato = valutazione_stato(ticket)

            if stato == "aperto":
                ticket.Save()
                ticket.Move(aperti_folder)

            elif stato == "chiuso":
                ticket.Save()
                ticket.Move(chiusi_folder)

                # Snapshot della cartella APERTI dopo lo spostamento
                aperti_snapshot = list(ticket_aperti)

                # Cerca e sposta tutte le mail collegate
                for aperti in aperti_snapshot:
                    if trova_mail_collegate(aperti, id_valutato):
                        aperti.Save()
                        aperti.Move(chiusi_folder)

if __name__ == "__main__":
    for i in range(5):
        print(f"\n Lancio {i+1}/5...\n")
        smista_ticket()
        time.sleep(1)  # 1 secondo di pausa tra le esecuzioni