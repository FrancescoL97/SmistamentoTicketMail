
from smista_ticket import ricerca_id, valutazione_stato, trova_mail_collegate

class FakeMail:
    def __init__(self, subject, body):
        self.Subject = subject
        self.Body = body

# === Test ricerca_id ===

def test_ricerca_id_subject():
    mail = FakeMail("Richiesta REQ000123", "")
    assert ricerca_id(mail) == "REQ000123"

def test_ricerca_id_body():
    mail = FakeMail("Nessun ID", "Controlla INC000999 qui.")
    assert ricerca_id(mail) == "INC000999"

def test_ricerca_id_none():
    mail = FakeMail("Niente", "Zero roba utile")
    assert ricerca_id(mail) is None

# === Test valutazione_stato ===

def test_valutazione_stato_chiuso():
    mail = FakeMail("Your ticket REQ000123 has been closed", "")
    assert valutazione_stato(mail) == "chiuso"

def test_valutazione_stato_aperto():
    mail = FakeMail("Incident INC123 opened on your behalf", "")
    assert valutazione_stato(mail) == "aperto"

def test_valutazione_stato_ignoto():
    mail = FakeMail("Something weird", "")
    assert valutazione_stato(mail) == "ignoto"

# === Test trova_mail_collegate ===

def test_trova_mail_collegate_subject():
    mail = FakeMail("Update REQ000777", "")
    assert trova_mail_collegate(mail, "REQ000777") is True

def test_trova_mail_collegate_body():
    mail = FakeMail("Info", "See ticket RITM000888 for details")
    assert trova_mail_collegate(mail, "RITM000888") is True

def test_trova_mail_collegate_none():
    mail = FakeMail("Unrelated", "This is a test")
    assert trova_mail_collegate(mail, "INC999999") is False
