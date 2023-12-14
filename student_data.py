import datetime
import os
import pandas as pd

current_year = int(str(datetime.datetime.now().year)[2:])

department_capacity = {
    'IT': 60,
    'CSE': 120,
    'ME': 25,
    'ECE': 60,
    'EE': 40
}

def load_existing_records():
    file_path = str(f"student_db_{current_year}.xlsx")

    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
    else:
        df = pd.DataFrame(columns=['Department', 'Name', 'Exam Type', 'Rank', 'Roll'])

    return df

def write_records_to_excel(df):
    df.to_excel(str(f"student_db_{current_year}.xlsx"), index=False)

def addStudent(df):
    try:
        # current_year = int(str(datetime.datetime.now().year)[2:])
        name = input("Enter student name: ").capitalize()
        department = input("Enter Student Department [Example: IT, CSE, ECE, EE, ME] : ").upper()

        if department not in department_capacity:
            raise ValueError("Invalid department. Exit with Status code x2")


        exam_type = input("What type of exam did he crack (JEE) or (WBJEE)? ").upper()

        if exam_type == 'JEE':
            jee_rank = int(input("Enter the JEE Rank of the Student: "))
            wbjee_rank = None  # Define wbjee_rank as None
            status = 0
        elif exam_type == 'WBJEE':
            wbjee_rank = int(input("Enter WBJEE Rank of the Student: "))
            jee_rank = None
            status = 0
        else:
            print("Invalid input. Exit with Status code x1")
            raise ValueError("Invalid exam type")
    
        if not df.empty:
            existing_department_records = df[df['Department'] == department]
            num_existing_students = len(existing_department_records)
            if(num_existing_students >= department_capacity[department]):
                print(str(f"All Seats are Full in {department} for academic year"+str(datetime.datetime.now().year)+"try next Year"))
            else:
                if num_existing_students == 0:
                    roll_number = f"{department}{current_year}01"
                else:
                    roll_number = f"{department}{current_year}{str(num_existing_students + 1).zfill(2)}"

        else:
            roll_number = f"{department}{current_year}01"

        new_record = pd.DataFrame([[department, name, exam_type, jee_rank if exam_type == 'JEE' else wbjee_rank, roll_number]],
                                  columns=['Department', 'Name', 'Exam Type', 'Rank', 'Roll'])

        # df = df._append(new_record, ignore_index=True)
        df = pd.concat([df, new_record], ignore_index=True)
        write_records_to_excel(df)

    except ValueError as e:
        print(f"Error: {e}")
        name, department, jee_rank, wbjee_rank, status = None, None, None, None, 1
        return name, department, jee_rank, wbjee_rank, status  # Return the values explicitly

    return name, department, jee_rank, wbjee_rank, status

def searchStudentByRoll(roll_number, df):
    try:
        student_record = df[df['Roll'] == roll_number].iloc[0]
        return tuple(student_record)
    except IndexError:
        print("Student not found with the given roll number.")
        return None

def deleteStudentByRoll(roll_number, df):
    try:
        df = df[df['Roll'] != roll_number]
        write_records_to_excel(df)
        print(f"Student with roll number {roll_number} deleted successfully.")
    except Exception as e:
        print(f"Error deleting student: {e}")

def displayAllRecords(df):
    try:
        if not df.empty:
            print("\nAll Student Records:")
            print(df.to_string(index=False))
        else:
            print("No student records found.")
    except Exception as e:
        print(f"Error displaying records: {e}")


if __name__ == '__main__':
    df = load_existing_records()

    print("[0] Add New Student Record")
    print("[1] Delete a Specific Student Record")
    print("[2] Search a Student Record")
    print("[3] Display All Student Record")
    print("[4] Exit")

    choice = int(input("\nWhat you Want [only Numerical Input {0,1,2,3,4}]: "))

    if(choice == 0):
        addStudent(df)
    elif(choice == 1):
        delete_roll_number = input("Enter the Roll Number to delete: ").upper()
        deleteStudentByRoll(delete_roll_number, df)
    elif(choice == 2):
        search_roll_number = input("Enter the Roll Number to search: ").upper()
        result = searchStudentByRoll(search_roll_number, df)
        if result is not None:
            print(result)
    elif(choice == 3):
        displayAllRecords(df)
    elif(choice == 4):
        exit()
    else:
        print("Exit with Status code x999 Invalid Input")

