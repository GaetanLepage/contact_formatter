#!/usr/bin/env python3

"""
Script de génération de fichiers vcf ou csv formatés à partir d'un tableau csv

On privilégiera le format vcf qui est compatible toutes plateformes (téléphone, Google, Apple, Outlook...)

Auteur : Gaétan Lepage
Date : Janvier 2018
"""


import os
import sys


"""
    Lit le fichier csv passé en argument et remplit puis retourne un tableau.
"""
def lecture_csv(input_file):
    if (not os.path.exists(input_file)):
        print("Fichier introuvable ou invalide")
        return 0

    les_contacts = []
    with open(input_file, "r") as old:
        les_lignes = old.readlines()
        for ligne in les_lignes:
            ligne = ligne.strip("\n")
            ligne = ligne.split(",")
            prenom = ligne[0]
            nom = ligne[1]
            numero = ligne[2]
            mail = ligne[3]
            les_contacts.append((prenom, nom, numero, mail))
    return les_contacts

"""
    Génère un fichier csv pour google.
"""
def generation_csv_google(les_contacts):
    les_chaines = []
    for contact in les_contacts:
        prenom = contact[0]
        nom = contact[1]
        numero = contact[2]
        mail = contact[3]
        chaine = prenom + " " + nom + "," + prenom + ",," + nom
        chaine += ",,,,,,,,,,,,,,,,,,,,,,,* My Contacts,* Home,"
        chaine += mail + ",Mobile," + numero
        les_chaines.append(chaine)

    with open("export_google.csv", "w") as google:
        print("Name,Given Name,Additional Name,Family Name,Yomi Name,Given Name\
         Yomi,Additional Name Yomi,Family Name Yomi,Name Prefix,Name Suffix,\
         Initials,Nickname,Short Name,Maiden Name,Birthday,Gender,Location,Billing \
         Information,Directory Server,Mileage,Occupation,Hobby,Sensitivity,Priority,\
         Subject,Notes,Group Membership,E-mail 1 - Type,E-mail 1 - Value,Phone 1 \
         - Type,Phone 1 - Value", file=google)
        for chaine in les_chaines:
            print(chaine, file=google)
    
    print("Export au format csv pour google effectué avec succès dans le fichier export_google.csv")

"""
    Génère un fichier csv pour outlook.
"""
def generation_csv_outlook(les_contacts):
    les_chaines = []
    for contact in les_contacts:
        prenom = contact[0]
        nom = contact[1]
        numero = contact[2]
        mail = contact[3]
        chaine = prenom + ",," + nom + ",,,,,,,,,,,," + mail
        chaine += ",,,,,," + numero
        chaine += ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,\
        ,,,,Normal,,Imported 12/7/17 1;My Contacts,"
        les_chaines.append(chaine)

    with open("export_outlook.csv", "w") as outlook:
        print("First Name,Middle Name,Last Name,Title,Suffix,Initials,Web Page,\
        Gender,Birthday,Anniversary,Location,Language,Internet Free Busy,Notes,\
        E-mail Address,E-mail 2 Address,E-mail 3 Address,Primary Phone,Home Phone\
        ,Home Phone 2,Mobile Phone,Pager,Home Fax,Home Address,Home Street,Home \
        Street 2,Home Street 3,Home Address PO Box,Home City,Home State,Home Postal \
        Code,Home Country,Spouse,Children,Manager's Name,Assistant's Name,Referred \
        By,Company Main Phone,Business Phone,Business Phone 2,Business Fax,Assistant's \
        Phone,Company,Job Title,Department,Office Location,Organizational ID Number,\
        Profession,Account,Business Address,Business Street,Business Street 2,Business \
        Street 3,Business Address PO Box,Business City,Business State,Business Postal \
        Code,Business Country,Other Phone,Other Fax,Other Address,Other Street,Other \
        Street 2,Other Street 3,Other Address PO Box,Other City,Other State,Other \
        Postal Code,Other Country,Callback,Car Phone,ISDN,Radio Phone,TTY/TDD Phone,\
        Telex,User 1,User 2,User 3,User 4,Keywords,Mileage,Hobby,Billing Information,\
        Directory Server,Sensitivity,Priority,Private,Categories", file=outlook)
        for chaine in les_chaines:
            print(chaine, file=outlook)
    
    print("Export au format csv pour outlook effectué avec succès dans le fichier export_outlook.csv")

"""
    Génère un fichier vcf.
"""
def generation_vcf(les_contacts):
    with open("export_vcf.vcf", "w") as vcard:
        for contact in les_contacts:
            prenom = contact[0]
            nom = contact[1]
            numero = contact[2]
            mail = contact[3]

            print("BEGIN:VCARD", file=vcard)
            print("VERSION:3.0", file=vcard)
            print("FN:" + prenom + " " + nom, file=vcard)
            print("N:" + nom + ";" + prenom + ";;;", file=vcard)
            print("EMAIL;TYPE=INTERNET;TYPE=HOME:" + mail, file=vcard)
            print("TEL;TYPE=CELL:" + numero, file=vcard)
            print("END:VCARD", file=vcard)
    
    print("Export au format vcf effectué avec succès dans le fichier export_vcf.vcf")

def main():

    help = "\nUtilisation : contact_generator [OPTION] <INPUT_FILE.csv>\n"
    help += "\nOPTION =\n"
    help += "--> -vcf (option par défaut)\n"
    help += "--> -csv_google\n"
    help += "--> -csv_outlook\n"


    print("/!\ Attention, le fichier source doit comporté exactement et dans cet ordre les champs suivants :")
    print("--> prénom")
    print("--> nom")
    print("--> numéro")
    print("--> mail\n")

    if len(sys.argv) not in [2, 3]:
        print(help)
        return

    les_contacts = lecture_csv(sys.argv[-1])

    if (not les_contacts):
        return
    
    if len(sys.argv) == 2:
        generation_vcf(les_contacts)

    if len(sys.argv) == 3:
        if (sys.argv[1] == "vcf"):
            generation_vcf(les_contacts)
        elif (sys.argv[1] == "csv_google"):
            generation_csv_google(les_contacts)
        elif (sys.argv[1] == "csv_outlook"):
            generation_csv_outlook(les_contacts)
        else:
            print("invalid option")
        

main()
