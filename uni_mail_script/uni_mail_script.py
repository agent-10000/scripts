# Author: The Sang Nguyen
# Written for the University of Goettingen. For other universities, please adapt the script.

"""
This script automates the delivery of corrected homework sheets by sending them by e-mail via SMTP.

--- SET UP ---:
Step 1: 
    Save this script in a directory together with "Punkteliste.csv" from StudIP. 
    Then create a folder with name "BlattXX" containing the corrected sheets with names of the form
        - "XX_surname_corrected.pdf" (single submissions), or,
        - "XX_surname_surname2_corrected.pdf" (group submissions), etc,
    where XX is the sheet number.
Step 2 - in method "main()": 
    Adjust login data "UNI_USER", "UNI_ADDRESS" and "UNI_PW". 
    (It is recommended to use environment variables instead of clear text.)
Step 3 - in method "send_mail()": 
    Adjust e-mail text.
"""

# built-in modules
import os
import sys
import smtplib
from email.message import EmailMessage

# third-party modules
import pandas as pd


def get_sheet_number():
    """
    Get sheet number from user.
    """
    sheet_number = input("++++ Please enter sheet no.: ")
    return "0" + sheet_number if len(sheet_number) == 1 else sheet_number


def get_path(sheet_number):
    """
    Get path to the folder containing the corrected sheets.
    """
    path = os.getcwd()
    if sys.platform.startswith('win32'):
        path += r"\Blatt"
    elif sys.platform.startswith('linux'):
        path += "/Blatt"
    elif sys.platform.startswith('darwin'):  # don't know if the script runs on macOS LOL
        path += "/Blatt"
    else:
        raise OSError('this script does not support your OS')
    return path + sheet_number


def get_corr_sheets(path, ending):
    """
    Get list of filenames of the corrected sheets.
    """
    return [filename for filename in os.listdir(path) if filename[-len(ending):] == ending]


def get_names(filenames, skip_first=0, skip_last=0):
    """
    Get list of (students') names from filenames.
    """
    return [filename[skip_first:-skip_last] for filename in filenames]


def send_mail(UNI_USER, UNI_ADDRESS, UNI_PW, receiver_email, firstname, surname, sheet_number, filename):
    """
    Send e-mail to student via SMTP. 
    """
    msg = EmailMessage()
    msg["From"] = UNI_ADDRESS
    msg["To"] = receiver_email
    msg["Subject"] = "Blatt " + sheet_number + " Korrektur"
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
        smtp.ehlo()  # identify ourselves to smtp gmail client
        smtp.starttls()  # secure our email with tls encryption
        smtp.ehlo()  # re-identify ourselves as an encrypted connection
        smtp.login(UNI_USER, UNI_PW)
        smtp.send_message(msg)

    print(
        f"|    +--- E-Mail to \"{surname}, {firstname}\" was sent succesfully.\n" +
        f"|         ({receiver_email})"
    )


def main():
    UNI_USER = os.environ.get('UNI_USER')  # "ug-student\firstname.surname"
    UNI_ADDRESS = os.environ.get('UNI_ADDRESS')  # eCampus e-mail address
    UNI_PW = os.environ.get('UNI_PW')  # password

    df = pd.read_csv("Punkteliste.csv", sep=";")

    # FOR TESTING PURPOSES:
    # ---------------------------------------------
    new_row = {
        "Stud.IP Benutzername": "thesang.nguyen",
        "Nachname": "Nguyen",
        "Vorname": "The Sang",
    }
    new_row_2 = {
        "Stud.IP Benutzername": "thesang.nguyen",
        "Nachname": "Nguyen",
        "Vorname": "Doppelgänger",
    }
    new_row_3 = {
        "Stud.IP Benutzername": "thesang.nguyen",
        "Nachname": "Nguyen2",
        "Vorname": "Doppelgänger2",
    }
    df = df.append(new_row, ignore_index=True)
    df = df.append(new_row_2, ignore_index=True)
    df = df.append(new_row_3, ignore_index=True)
    # ---------------------------------------------

    # duplicate surnames
    duplicate_names = list(df[df.duplicated("Nachname")]["Nachname"])

    sheet_number = get_sheet_number()
    path = get_path(sheet_number)
    os.chdir(path)  # change into directory of given path

    beginning = "XX_"
    ending = "_corrected" + ".pdf"

    corr_sheets = get_corr_sheets(path, ending)  # get list of filenames of corrected sheets
    names_in_filenames = get_names(corr_sheets, len(beginning), len(ending))  # get list of names in filenames
    print(f"+--- There are {len(corr_sheets)} corrected sheets.")

    unknown = []  # names not found in database during following search
    for idx, names in enumerate(names_in_filenames):
        names = names.split(sep="_")  # split "names" into names of group members
        print(f"+--- Handling file no. {len(names_in_filenames) - idx} from:", *names)
        for name in names:
            dd = df.loc[df["Nachname"] == name]  # get dataframe row(s) of student(s) with given name
            if name in duplicate_names:  # if students have the same surname
                options = list(dd["Vorname"])  # list of firstnames with same surname

                # let user choose which student to send e-mail to
                prompt = f"++++ There are {len(options)} {name}s.\n" + f"|    Do you mean {options[0]} [0]"
                for k in range(1, len(options)):
                    prompt += f" or {options[k]} [{k}]"
                prompt += "? "
                opt = int(input(prompt))

                dd = dd.loc[df["Vorname"] == options[opt]]  # unlike "Nachname" the dataframe index is unique

            try:
                df_idx = dd.index[0]  # get index of dataframe row of student with given name
            except IndexError:  # if student not found in database
                unknown.append(name)
                print(
                    '+--- UNICODE ERROR:\n' +
                    f'|    Name "{name}" not found in database (possibly due to weird Umlaut-formats)\n' +
                    '|    +-- possible SOLUTIONS: retype filename on PC or change name in database'
                )
                continue
            else:
                dd = df.iloc[df_idx]  # wanted dataframe row

                # "corr_sheets" and "names_in_files" share same index of same file
                send_mail(UNI_USER, UNI_ADDRESS, UNI_PW, dd["Stud.IP Benutzername"] + "@stud.uni-goettingen.de",
                          dd["Vorname"], dd["Nachname"], sheet_number, corr_sheets[idx])

    if unknown:
        print('+--- ATTENTION: following names were NOT found:', *unknown)


if __name__ == "__main__":
    main()
