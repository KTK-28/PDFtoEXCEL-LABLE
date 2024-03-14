PDF to Excel Converter Web Application 

For Amazon Seller Who Dispatch their order with "Self-Ship" option & Delivery through Indian Postal Service 

This web application allows users to upload PDF invoices, convert them to Excel format, and generate shipping labels.

Setup Instructions
Prerequisites

    Install Python:
    Download and install Python from the official website.

Installation Steps

    Open Windows PowerShell:
    Right-click on Windows PowerShell and select "Run as administrator".

    Set Execution Policy:
    Run the following command to set the execution policy to unrestricted:

    powershell

Set-ExecutionPolicy Unrestricted

Clone the Project:

bash

git clone https://github.com/your-username/pdf-to-excel.git

Navigate to the Project Directory:

bash

cd pdf-to-excel

Create Virtual Environment:
Create a virtual environment using the following commands:

bash

pip install venv
venv\Scripts\Activate.ps1

Install Dependencies:
Install the required dependencies using:

bash

    pip install -r requirements.txt

Running the Application

    Start the Flask Server:
    Run the following command to start the application:

    bash

    python app.py

    Access the Web Interface:
    Open your web browser and navigate to http://localhost:5000.

Folder Structure

    uploads: Stores uploaded PDF files.
    excels: Stores converted Excel sheets.
    labels: Stores generated shipping labels.

Additional Notes

    Ensure that the uploads, excels, and labels directories have proper read and write permissions.
    You can modify the host and port settings in the app.py file if needed.
    This project is licensed under the MIT License. See the LICENSE file for details.

Troubleshooting

    If you encounter any issues during setup or usage, please refer to the troubleshooting section in this README or seek assistance from the project's contributors.
