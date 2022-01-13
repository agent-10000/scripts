# BEFORE EXECUTION: adjust "path" according to your operating system
#
# Author: The Sang Nguyen
#
# Inspired by Corey Schafer's Tutorial:
# https://www.youtube.com/watch?v=JRCJ6RtE3xU

"""
This script automates the delivery of corrected homework sheets by
- reading csv-file containing all names and e-mail addresses
- changing into directory with corrected sheets
- reading corrected sheets of the form
    "XX_surname_korrigiert.pdf" or 
    "XX_surname_surname2_korrigiert.pdf" or 
    "XX_surname_surname2_surname3_korrigiert.pdf"
- sending them by e-mail via SMTP

SET UP:
1 - Save this script in a directory together with "Punkteliste.csv" (containing 
all names and e-mail addresses) and a folder with names "BlattXX" containing 
the corrected sheets (of the forms above).
2 - main(): Adjust "path" (according to your OS) and your login data 
"UNI_USER", "UNI_ADDRESS" and "UNI_PW". It is recommended to use environment 
variables instead of clear text.
3 - send_mail(): Adjust e-mail text.
"""

# built-in modules
import os
import smtplib
from email.message import EmailMessage

# third-party modules
import pandas as pd


def send_mail(UNI_USER, UNI_ADDRESS, UNI_PW,
              email, firstname, surname, sheet_no, filename):
    msg = EmailMessage()
    msg["Subject"] = "AGLA1 Blatt " + sheet_no + " Korrektur"
    msg["From"] = UNI_ADDRESS
    msg["To"] = email
    msg.set_content(
        f"Hallo {firstname},\
            \n\nanbei findest Du deinen korrigierten Zettel.\
            \n\nViele Grüße,\
            \nSang\
            \n\nBrought to you by Platypus Inc.\
            \nhttps://github.com/agent-10000"
    )

    with open(filename, "rb") as f:
        file_data = f.read()
        file_name = f.name

    msg.add_attachment(
        file_data,
        maintype="application",  # for PDFs
        subtype="octet-stream",  # for PDFs
        filename=file_name,
    )

    with smtplib.SMTP('email.stud.uni-goettingen.de', 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(UNI_USER, UNI_PW)
        smtp.send_message(msg)

    print("E-Mail to \""+surname+", "+firstname +
          "\" ("+email+") was sent succesfully.")


def main():
    UNI_USER = os.environ.get('UNI_USER')  # "ug-student\firstname.surname"
    UNI_ADDRESS = os.environ.get('UNI_ADDRESS')  # ecampus e-mail address
    UNI_PW = os.environ.get('UNI_PW')  # password

    df = pd.read_csv("Punkteliste.csv", sep=";")

    # FOR TESTING PURPOSES:
    # ---------------------------------------------
    # new_row = {
    #     "Stud.IP Benutzername": "thesang.nguyen",
    #     "Nachname": "Nguyen",
    #     "Vorname": "The Sang",
    # }
    # new_row_2 = {
    #     "Stud.IP Benutzername": "thesang.nguyen",
    #     "Nachname": "Nguyen",
    #     "Vorname": "Doppelgänger",
    # }
    # df = df.append(new_row, ignore_index=True)
    # df = df.append(new_row_2, ignore_index=True)
    # ---------------------------------------------

    # duplicate surnames
    duplicate_names = list(df[df.duplicated("Nachname")]["Nachname"])

    sheet_no = input("Please enter sheet no.: ")
    if len(sheet_no) == 1:
        sheet_no = "0" + sheet_no

    # ADJUST path according to your os
    path = os.getcwd()
    path += r"\Blatt"  # for Windows
    # path += "/Blatt"  # for Linux
    path += sheet_no

    os.chdir(path)  # change into directory of given path

    # get list of filenames of corrected sheets
    corr_sheets = [filename for filename in os.listdir(path)
                   if filename[-14:] == "korrigiert.pdf"]
    names_in_filenames = [filename[3:-15] for filename in corr_sheets]
    print(f"There are {len(corr_sheets)} corrected sheets.")

    for idx, names in enumerate(names_in_filenames):
        print(names)
        # split "names" into names of group members
        for name in names.split(sep="_"):
            if name in duplicate_names:
                options = list(df.loc[df["Nachname"] == name]["Vorname"])

                prompt = f"There are {len(options)} {name}s.\
                    \nDo you mean {options[0]} [0]"
                for k in range(1, len(options)):
                    prompt += f" or {options[k]} [{k}]"
                prompt += "? "
                opt = int(input(prompt))

                # unlike "Nachname" the dataframe index is unique
                df_idx = df.loc[df["Nachname"] == name]\
                    .loc[df["Vorname"] == options[opt]].index[0]
            else:
                df_idx = df.loc[df["Nachname"] == name].index[0]

            dd = df.iloc[df_idx]
            email = dd["Stud.IP Benutzername"] + "@stud.uni-goettingen.de"
            firstname = dd["Vorname"]
            surname = dd["Nachname"]

            # "corr_sheets" and "names_in_files" share same index of same file
            send_mail(UNI_USER, UNI_ADDRESS, UNI_PW,
                      email, firstname, surname, sheet_no, corr_sheets[idx])


if __name__ == "__main__":
    main()
