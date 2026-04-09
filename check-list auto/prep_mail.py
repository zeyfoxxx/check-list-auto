
import win32com.client as win32
import os
import time
from datetime import datetime

def prepare_mail(to_list, cc_list, subject, body="", from_address=None):

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # ✅ Envoyer au nom de WAN
    if from_address:
        mail.SentOnBehalfOfName = from_address

    # ✅ Préremplir tout SAUF le HTML
    mail.To = "; ".join(to_list)
    mail.CC = "; ".join(cc_list)
    mail.Subject = subject

    # ✅ Étape 1 : OUVRIR l'email AVANT la modification HTML
    mail.Display()
    time.sleep(0.3)   # laisse Outlook charger la signature

    # ✅ Étape 2 : RÉCUPÉRER le HTMLBody AVEC la signature Outlook
    current_html = mail.HTMLBody

    # ✅ Étape 3 : préparer ton texte au format HTML
    body_html = body.replace("\n", "<br>")

    # ✅ Étape 4 : injecter ton texte AVANT la signature, sans la supprimer
    # On met ton texte juste après l’ouverture du <body>
    if "<body" in current_html.lower():
        # on trouve la fin de <body>
        idx = current_html.lower().find(">") + 1
        final_html = current_html[:idx] + f"<p>{body_html}</p><br>" + current_html[idx:]
    else:
        # fallback ultra-simple
        final_html = f"<p>{body_html}</p><br>" + current_html

    # ✅ Étape 5 : réinjecter le HTML modifié
    mail.HTMLBody = final_html

date_str = datetime.now().strftime("%d/%m/%Y")
titre = "SPID RS OPERE-INV-check-list "+date_str


# ✅ TEST
prepare_mail(
    to_list=["groupe.dsiun.spsn.tbirs@departement13.fr"],
    cc_list=["projetcd13@freepro.com", "sesn.wan@departement13.fr"],
    subject=titre,
    body="Bonjour,\n\nVeuillez trouver ci-dessous la check-list du jour :\n\n\n\nCordialement,",
    from_address="sesn.wan@departement13.fr"
)
