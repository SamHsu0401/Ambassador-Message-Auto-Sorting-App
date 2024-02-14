import pandas as pd


def spreadsheet():
    # spreadsheet access
    sheetname_input = int(
        input(
            """
    ### Only input number
    1. Aerospace
    2. Biomedical
    3. Chemical
    4. Civil
    5. CE
    6. CIS
    7. CS
    8. Electrical
    9. IIT
    10. Mechanical
    What is the sheet of major you want to access: """
        )
    )

    student_major_list = {
        1: "Aerospace",
        2: "Biomedical",
        3: "Chemical",
        4: "Civil",
        5: "Computer Eng.",
        6: "Computer Information Systems",
        7: "Computer Science",
        8: "Electrical",
        9: "Integrated Information Tech.",
        10: "Mechanical",
    }
    sheetname = student_major_list[sheetname_input]

    filename = "2024 Texting Tracker.xlsx"
    workbook = pd.read_excel(filename, sheet_name=sheetname)
    return workbook


def print_text(preferred_name, user_name, major):
    message = []
    for name in preferred_name:

        text = f"Hi {name}! My name is {user_name}. I am a {major} Major at the University of South Carolina. Congratulations on your acceptance to USC! Do you have any questions about being an engineering student here that I could answer for you?"
        message.append(text)
    return message


def ask_user_info():

    major_list = {
        1: "Aerospace Engineering",
        2: "Biomedical Engineering",
        3: "Chemical Engineering",
        4: "Civil Engineering",
        5: "Computer Engineering",
        6: "Computer Information Systems",
        7: "Computer Science",
        8: "Electrical Engineering",
        9: "Integrated Information Technology",
        10: "Mechanical Engineering",
    }

    name = input("What is your name: ")
    major = int(
        input(
            """
### Only input number
1. Aerospace
2. Biomedical
3. Chemical
4. Civil
5. CE
6. CIS
7. CS
8. Electrical
9. IIT
10. Mechanical
What is your major: """
        )
    )
    major = major_list[major]

    return name, major


def write_to_excel(preferred_name, names_not_sent, phone_numbers):
    name, major = ask_user_info()
    message = print_text(preferred_name, name, major)
    messages_df = pd.DataFrame(
        {
            "Preferred Name": names_not_sent,
            "Phone Number": phone_numbers,
            "Message": message,
        }
    )
    print("***Spreadsheet Created Successfully***")

    output_filename = "Message Output.xlsx"
    messages_df.to_excel(output_filename, index=False)


def ask_num():
    text_num_request = int(input("how many people do you want to send: "))
    return text_num_request


def main():
    workbook = spreadsheet()
    people_to_send = workbook[workbook["Text Sent"].isnull()]
    people_to_send = people_to_send.head(ask_num())
    preferred_name = people_to_send["Preferred Name"].tolist()
    phone_numbers = people_to_send["Mobile Number"].tolist()

    write_to_excel(preferred_name, preferred_name, phone_numbers)


main()
