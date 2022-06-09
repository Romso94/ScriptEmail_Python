import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import csv
import pandas as pd


read_file = pd.read_excel(r'Liste_Clients.xlsx')

read_file.to_csv(r'Liste_Clients.csv',index=' ',header=True)

Liste_email=[]
Liste_nom=[]
Liste_ville=[]
Nombre_Email=0

with open('Liste_Clients.csv', encoding='utf-8') as csvFile:
    csvRead = csv.reader(csvFile)

    for row in csvRead:
        if row[4]=='Mail ':
            pass
        else :
            Nombre_Email +=1
           
            Liste_email.append(row[4]) #Numero Colonne Email
            




Email_Send =0

for i in  range (Nombre_Email):
    msg = MIMEMultipart()
    msg['From'] = "XXXXXX@gmail.com" # Changer l'adresse mail et mettre la votre
    msg['To'] = f"{Liste_email[i]}" # Adresse Destinataire ne pas toucher
    password = "XXXXXX"  #Mot de passe de votre compte Google |  Aller dans ->Gérer votre compte Google -> Sécurité ->Connexion a Google 
                                                                       #    -> Mot de passe des applications -> selectionner une application : Messagerie / selectionner un appareil et générer

    msg['Subject']= f"XXXXXX" # l'objet du mail
    body = f"XXXXXXXXXXX"  #Corps du mail
    msg.attach(MIMEText(body,'html'))

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(msg["From"],password)
    server.sendmail(msg["From"],msg["To"],msg.as_string())
    Email_Send +=1
    print(f'Email envoyé a {Liste_email[i]}')
    server.quit()


print(f"Nombre d'email envoyé = {Email_Send}")



