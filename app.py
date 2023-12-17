from flask import Flask, render_template, request
import datetime
import os
import pandas as pd

app = Flask(__name__)

current_year = int(str(datetime.datetime.now().year)[2:])
department_capacity = {'IT': 60, 'CSE': 120, 'ME': 25, 'ECE': 60, 'EE': 40}


def load_existing_records():
    file_path = str(f"student_db_{current_year}.xlsx")
    if os.path.exists(file_path):
        return pd.read_excel(file_path)
    return pd.DataFrame(columns=['Department', 'Name', 'Exam Type', 'Rank', 'Roll'])


def write_records_to_excel(df):
    file_path = str(f"student_db_{current_year}.xlsx")
    df.to_excel(file_path, index=False)


def addStudent(existing_df, name, department, exam_type, jee_rank = None, wbjee_rank = None):
    try:
        new_record = None
        if department not in department_capacity:
            raise ValueError("Invalid department. Exit with Status code x2")

        if exam_type not in ['JEE', 'WBJEE']:
            raise ValueError("Invalid exam type. Exit with Status code x1")

        if not existing_df.empty:
            existing_department_records = existing_df[existing_df['Department'] == department]
            num_existing_students = len(existing_department_records)
            if num_existing_students >= department_capacity[department]:
                raise ValueError(f"All Seats are Full in {department} for the academic year {current_year}. Try next year.")

            if num_existing_students == 0:
                roll_number = f"{department}{current_year}01"
            else:
                roll_number = f"{department}{current_year}{str(num_existing_students + 1).zfill(2)}"
        else:
            roll_number = f"{department}{current_year}01"

        if exam_type == 'JEE':
        # Handle the case when rank data is not provided
            try:
                jee_rank = int(jee_rank) if jee_rank else None
                new_record = pd.DataFrame([[department, name, exam_type, jee_rank, roll_number]],columns=['Department', 'Name', 'Exam Type', 'Rank', 'Roll'])
            except ValueError:
                raise ValueError("Rank should be numeric")
            
        elif exam_type == 'WBJEE':
            try:
                wbjee_rank = int(wbjee_rank) if wbjee_rank else None
                new_record = pd.DataFrame([[department, name, exam_type, wbjee_rank, roll_number]],columns=['Department', 'Name', 'Exam Type', 'Rank', 'Roll'])
            except ValueError:
                raise ValueError("Rank Should be Numeric")
            
        updated_df = pd.concat([existing_df, new_record], ignore_index=True)
        write_records_to_excel(updated_df)

        # Return the added student details
        return name, department, exam_type, jee_rank, wbjee_rank, roll_number, 0

    except ValueError as e:
        print(f"Error: {e}")
        return None, None, None, None, None, None, str(e)


@app.route("/", methods=["GET", "POST"])
def home():
    existing_df = load_existing_records()  # Load existing records at the beginning of each request

    if request.method == "POST":
        name = request.form.get("name").title()
        department = request.form.get("department").upper()
        exam_type = request.form.get("exam_type").upper()

        if exam_type == "JEE":
            try:
                jee_rank = int(request.form.get("jee_rank"))
                wbjee_rank = None
            except ValueError as z:
                error_message = f"JEE Rank should be a numeric value. Cannot accept '{z}'."
                result = (None, None, None, None, None, None, error_message)
                return render_template("index.html", result=result, df=existing_df)

        elif exam_type == "WBJEE":
            try:
                wbjee_rank = int(request.form.get("wbjee_rank"))
                jee_rank = None
            except ValueError as y:
                error_message = f"WBJEE Rank should be a numeric value. Cannot accept '{y}'."
                result = (None, None, None, None, None, None, error_message)
                return render_template("index.html", result=result, df=existing_df)

        else:
            # Handle invalid exam type
            jee_rank = None
            wbjee_rank = None

        result = addStudent(existing_df, name, department, exam_type, jee_rank, wbjee_rank)
        existing_df = load_existing_records()
        return render_template("index.html", result=result, df=existing_df)
    
    return render_template("index.html", df=existing_df)



if __name__ == '__main__':
    app.run(debug=True)
