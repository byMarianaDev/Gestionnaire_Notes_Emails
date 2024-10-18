
# Student Grades Email Automation Script

Welcome my **Student Grades Email Automation Script** project. This script automates the process of generating personalized reports for students based on their grades in Excel and sends them via email. The project leverages multi-threading for efficient email sending and ensures that every student gets their report in a timely manner.

## Key Features
- Automatically generates personalized grade reports for students.
- Sends reports via email with attachments (grade reports).
- Multi-threaded execution for faster processing.
- Easy-to-configure `.env` file for sensitive credentials.

## Installation

To get started, you will need to install the following dependencies. Make sure you have `Python 3.12` or a compatible version installed on your machine.

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/student-grades-automation.git
cd student-grades-automation
```

### 2. Install Required Python Packages

Before running the script, make sure you have installed the necessary Python libraries. You can install them using `pip`:

```bash
pip install -r requirements.txt
```

### 3. Set Up Environment Variables

This script uses environment variables to keep sensitive data such as your email and password secure. Create a `.env` file in the project root with the following variables:

```
EMAIL_ADDRESS=your email
EMAIL_PASSWORD=your password
```

Replace `your email` and `your password` with your actual credentials.

### 4. Excel File Format

You need to provide an Excel file (`.xlsx`) containing the students' information. The file should have **4 columns** as shown below:

| Student Name | Mathematics | Science | History | Email |
|--------------|-------------|---------|---------|-------|
| John Doe     | 85          | 90      | 78      | john.doe@example.com |
| Jane Smith   | 92          | 88      | 81      | jane.smith@example.com |

This file will be processed to generate individualized reports for each student.

### 5. Template Document

The script uses a template Word document (`plantilla_alumnos.docx`) for generating student reports. The template includes placeholders for the following variables:

- `{ALUMNO}`: Student Name
- `{MATEMÁTICAS}`: Mathematics Grade
- `{CIENCIAS}`: Science Grade
- `{HISTORIA}`: History Grade
- `{PROMEDIO}`: Average Grade

The generated reports will be filled with the respective grades for each student and attached to the email.

### How to Run the Script

Once everything is set up, you can run the script by executing the following command:

```bash
python procesamiento_datos.py
```

The script will:
1. Read the students' data from the Excel file.
2. Generate personalized Word documents with each student's grades.
3. Send the documents via email to the corresponding recipients using your email credentials from the `.env` file.

## Example Output

Upon successful execution, the script will log the status of email sending and report generation:

```
Email sent to john.doe@example.com
Email sent to jane.smith@example.com
All emails sent successfully.
```

If any issues arise (e.g., missing files or incorrect credentials), detailed error logs will be saved in the `app.log` file.

## Dependencies

The following Python packages are required:

- `pandas`: For reading the Excel file.
- `openpyxl`: For handling `.xlsx` Excel files.
- `python-docx`: For generating Word documents.
- `dotenv`: For loading environment variables from the `.env` file.
- `smtplib`: For sending emails.
- `threading`: For concurrent email sending.

You can find all required dependencies listed in `requirements.txt`.

## Notes

This project showcases my ability to handle **file I/O operations**, work with **external libraries** for email automation, and implement **multi-threading** for performance optimization. I have also followed best practices by securing sensitive data in environment variables and logging important events for troubleshooting.

Feel free to clone this project, configure your own `.env` file, and run it on your local machine to see the functionality in action.

## Future Improvements

- Implementing a user-friendly GUI for non-technical users.
- Adding support for additional file formats (e.g., CSV).
- Enhancing error handling with retry mechanisms for failed email deliveries.

---

If you have any questions or feedback, feel free to reach out!
