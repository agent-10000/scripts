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
2 - main(): Adjust login data "UNI_USER", "UNI_ADDRESS" and "UNI_PW". It is 
recommended to use environment variables instead of clear text.
3 - send_mail(): Adjust e-mail text.
"""

# built-in modules
import os
import sys
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
        f"Hallo {firstname},\n\n" +
        "anbei findest Du deinen korrigierten Zettel.\n\n" +
        "Viele Grüße,\n" +
        "Sang"
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

    print(
        f"|    +--- E-Mail to \"{surname}, {firstname}\" was sent succesfully.\n" +
        f"|         ({email})"
    )


def main():
    UNI_USER = os.environ.get('UNI_USER')  # "ug-student\firstname.surname"
    UNI_ADDRESS = os.environ.get('UNI_ADDRESS')  # ecampus e-mail address
    UNI_PW = os.environ.get('UNI_PW')  # password

    df = pd.read_csv("Punkteliste.csv", sep=";")

    # FOR TESTING PURPOSES:
    # ---------------------------------------------
    new_row = {
        "Stud.IP Benutzername": "thesang.nguyen",
        "Nachname": "Nguyen",
        "Vorname": "The Sang",
    }
    # new_row_2 = {
    #     "Stud.IP Benutzername": "thesang.nguyen",
    #     "Nachname": "Nguyen",
    #     "Vorname": "Doppelgänger",
    # }
    # new_row_3 = {
    #     "Stud.IP Benutzername": "thesang.nguyen",
    #     "Nachname": "Nguyen2",
    #     "Vorname": "Doppelgänger2",
    # }
    df = df.append(new_row, ignore_index=True)
    # df = df.append(new_row_2, ignore_index=True)
    # df = df.append(new_row_3, ignore_index=True)
    # ---------------------------------------------

    # duplicate surnames
    duplicate_names = list(df[df.duplicated("Nachname")]["Nachname"])

    path = os.getcwd()
    if sys.platform.startswith('win32'):
        path += r"\Blatt"
    elif sys.platform.startswith('linux'):
        path += "/Blatt"
    elif sys.platform.startswith('darwin'):  # for macOS
        path += "/Blatt"  # not sure if script runs on macOS though LOL
    else:
        raise OSError('this script does not support your OS')

    sheet_no = input("++++ Please enter sheet no.: ")
    if len(sheet_no) == 1:
        sheet_no = "0" + sheet_no

    path += sheet_no

    os.chdir(path)  # change into directory of given path

    # get list of filenames of corrected sheets
    corr_sheets = [filename for filename in os.listdir(path)
                   if filename[-14:] == "korrigiert.pdf"]
    names_in_filenames = [filename[3:-15] for filename in corr_sheets]
    print(f"+--- There are {len(corr_sheets)} corrected sheets.")

    unknown = []  # names not found in database during following search
    for idx, names in enumerate(names_in_filenames):
        # split "names" into names of group members
        names = names.split(sep="_")
        print(f"+--- Handling file {idx+1} from:", *names)
        for name in names:
            dd = df.loc[df["Nachname"] == name]
            if name in duplicate_names:
                options = list(dd["Vorname"])

                prompt = f"++++ There are {len(options)} {name}s.\
                    \n|    Do you mean {options[0]} [0]"
                for k in range(1, len(options)):
                    prompt += f" or {options[k]} [{k}]"
                prompt += "? "
                opt = int(input(prompt))

                # unlike "Nachname" the dataframe index is unique
                dd = dd.loc[df["Vorname"] == options[opt]]

            try:
                df_idx = dd.index[0]
            except IndexError:
                unknown.append(name)
                print(
                    '+--- UNICODE ERROR:' +
                    f'\n|    Name "{name}" not found in database (possibly due to weird Umlaut-formats)' +
                    '\n|    +-- possible SOLUTIONS: retype filename on PC or change name in database'
                )
                continue
            else:
                dd = df.iloc[df_idx]
                email = dd["Stud.IP Benutzername"] + "@stud.uni-goettingen.de"
                firstname = dd["Vorname"]
                surname = dd["Nachname"]

                # "corr_sheets" and "names_in_files" share same index of same file
                send_mail(UNI_USER, UNI_ADDRESS, UNI_PW,
                          email, firstname, surname, sheet_no, corr_sheets[idx])

    if unknown:
        print('+--- ATTENTION: following names were NOT found:', *unknown)


if __name__ == "__main__":
    main()
