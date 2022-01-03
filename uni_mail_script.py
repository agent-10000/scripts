# Code by The Sang Nguyen inspired by Corey Schafer's Tutorial:
# https://www.youtube.com/watch?v=JRCJ6RtE3xU
#
# Written for Windows

import os
import smtplib
from email.message import EmailMessage
import pandas as pd


def get_corr_sheets():
    corr_sheets = []
    for file_name in os.listdir(os.getcwd()):
        if file_name[-14:] == "korrigiert.pdf":
            corr_sheets.append(file_name)
    return corr_sheets


def send_mail(UNI_USER, UNI_ADDRESS, UNI_PW, sheet_no, df, df_indeces, filenames):
    for idx, df_idx in enumerate(df_indeces):
        mail = df.iloc[df_idx]["Stud.IP Benutzername"] + \
            "@stud.uni-goettingen.de"
        firstname = df.iloc[df_idx]["Vorname"]

        msg = EmailMessage()
        msg["Subject"] = "AGLA1 Blatt " + sheet_no + " Korrektur"
        msg["From"] = UNI_ADDRESS
        msg["To"] = mail
        msg.set_content(
            f"Hallo {firstname},\
            \n\nanbei findest Du deinen korrigierten Zettel.\
            \n\nViele Grüße,\
            \nSang"
        )

        with open(filenames[idx], "rb") as f:
            file_data = f.read()
            file_name = f.name

        msg.add_attachment(
            file_data,
            maintype="application",
            subtype="octet-stream",
            filename=file_name,
        )

        with smtplib.SMTP('email.stud.uni-goettingen.de', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(UNI_USER, UNI_PW)
            smtp.send_message(msg)

        print("E-Mail to "+firstname+" ("+mail+") was sent succesfully.")


def main():
    UNI_USER = os.environ.get('UNI_USER')
    UNI_ADDRESS = os.environ.get('UNI_ADDRESS')
    UNI_PW = os.environ.get('UNI_PW')

    df = pd.read_csv("Punkteliste.csv", sep=";")
    # ---------------------------------------------
    # FOR TESTING PURPOSES:
    # ---------------------------------------------
    new_row = {
        "Stud.IP Benutzername": "thesang.nguyen",
        "Nachname": "Nguyen",
        "Vorname": "The Sang",
    }
    new_row_2 = {
        "Stud.IP Benutzername": "thesang.nguyen",
        "Nachname": "Nguyen2",
        "Vorname": "The Sang2",
    }
    new_row_3 = {
        "Stud.IP Benutzername": "thesang.nguyen",
        "Nachname": "Nguyen3",
        "Vorname": "The Sang3",
    }
    df = df.append(new_row, ignore_index=True)
    df = df.append(new_row_2, ignore_index=True)
    df = df.append(new_row_3, ignore_index=True)
    # ---------------------------------------------

    duplicate_names = list(df[df.duplicated("Nachname")]["Nachname"])

    sheet_no = input("Homework Sheet No.: ")
    if len(sheet_no) == 1:
        sheet_no = "0" + sheet_no

    path = r"C:\Users\thesa\Desktop\AGLA1_Korrektur\Blatt"  # on windows
    # path = "/mnt/c/Users/thesa/desktop/AGLA1_Korrektur/Blatt"  # on linux

    os.chdir(path + sheet_no)  # change into directory of given path

    df_indeces = []
    corr_sheets = get_corr_sheets()
    print(f"There are {len(corr_sheets)} corrected sheets.")
    filenames = [filename[:-15] for filename in corr_sheets]
    corr_sheets_multiples = []
    for idx, filename in enumerate(filenames):
        fn = filename.split(sep="_")
        for i in range(1, len(fn)):
            corr_sheets_multiples.append(corr_sheets[idx])
            if fn[i] in duplicate_names:
                options = list(df.loc[df["Nachname"] == fn[i]]["Vorname"])
                opt = input(
                    f"There are two {fn[i]}s.\
                    \nDo you mean {options[0]} [0] or {options[1]} [1]? "
                )
                opt = int(opt)
                df_indeces.append(
                    df.loc[df["Vorname"] == options[opt]].index[0])
            else:
                df_indeces.append(df.loc[df["Nachname"] == fn[i]].index[0])

    send_mail(UNI_USER, UNI_ADDRESS, UNI_PW,
              sheet_no, df, df_indeces, corr_sheets_multiples)


if __name__ == "__main__":
    main()
