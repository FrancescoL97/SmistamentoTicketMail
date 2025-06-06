import win32com.client, re, datetime, os

### Dichiarazioni
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)
ticket_folder = inbox.Folders["TICKET"]
aperti_folder = ticket_folder.Folders["APERTI"]
chiusi_folder = ticket_folder.Folders["CHIUSI"]
tickets_da_valutare = ticket_folder.Items
ticket_aperti = aperti_folder.Items

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

### Funzioni
def ricerca_id(mail):
    ticket_regex = re.compile(r"(REQ\d+|INC\d+|RITM\d+)", re.IGNORECASE)
    subject = mail.Subject
    body = mail.Body
    match_sub = ticket_regex.search(subject)
    if not match_sub:
        match_bod = ticket_regex.search(body)
        if match_bod:
            id = match_bod.group(1).upper()
            return id
    if match_sub:
        id = match_sub.group(1).upper()
        return id

def valutazione_stato(mail):
    subject = mail.Subject.lower()
    if any(k in subject for k in KEYWORDS_CHIUSO):
        stato = "chiuso"
    elif any(k in subject for k in KEYWORDS_APERTO):
        stato = "aperto"
    else:
        stato = "ignoto"
    return stato

def trova_mail_collegate(mail, ticket_id):
    try:
        id_regex = re.compile(ticket_id, re.IGNORECASE)
        return bool(id_regex.search(mail.Subject) or id_regex.search(mail.Body))
    except Exception as e:
        print(f"Errore durante controllo correlazione mail: {e}")
        return False

### Azioni
for ticket in tickets_da_valutare:
    id_valutato = ricerca_id(ticket)
    if id_valutato:
        if valutazione_stato(ticket) == "aperto":
            ticket.Save()
            ticket.Move(aperti_folder)
        elif valutazione_stato(ticket) == "chiuso":
            ticket.Save()
            ticket.Move(chiusi_folder)
            aperti_snapshot = list(ticket_aperti)
            for aperti in ticket_aperti:
                if trova_mail_collegate(aperti,id_valutato):
                    aperti.Save()
                    aperti.Move(chiusi_folder)