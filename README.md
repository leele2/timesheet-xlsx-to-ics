# Shift Calendar Uploader

## Overview
This project provides a web-based tool for uploading an Excel file containing work shifts and converting it into an ICS calendar file. Users can then import the generated ICS file into their preferred calendar applications.

## Features
- Upload an Excel (.xlsx) file containing shift data.
- Search for shifts based on a provided name.
- Generate an ICS calendar file containing the retrieved shifts.
- Download the ICS file for easy import into calendar applications.

## Tech Stack
- **Backend:** Django
- **Frontend:** HTML
- **Libraries Used:**
  - `pytz` for timezone handling
  - `ics` for calendar file generation
  - `pandas` (used in `read_xls` for Excel processing)

## Folder Structure
```
project-root/
│-- shifts/
│   ├── views.py  # Handles file upload and ICS generation
│   ├── utils.py  # Contains helper functions like reading Excel files
│-- api/
│   ├── wsgi.py   # WSGI entry point for deployment
│-- vercel.json   # Deployment configuration for Vercel
│-- requirements.txt  # Dependencies
│-- README.md  # Project documentation
```

## 📤 How to Upload Your File

1. Open the application in your web browser.
2. Select your correctly formatted Excel file (`.xlsx` format).
3. Enter the name you want to search shifts for.
4. Click "Upload" to generate an `.ics` file for your calendar

### 🗓 Table Format Example

Below is an example of the required structure:

|          | Sunday      | Monday      | Tuesday     | Wednesday   | Thursday    | Friday      | Saturday    |
|----------|------------|------------|------------|------------|------------|------------|------------|
| **Date** | 26th Jan   | 27th Jan   | 28th Jan   | 29th Jan   | 30th Jan   | 31st Jan   | 1st Feb    |
| **Shift** | **Employee Name** | **Employee Name** | **Employee Name** | **Employee Name** | **Employee Name** | **Employee Name** | **Employee Name** |
| 9:00-17:00  | Alice       | Bob        | Charlie    | David      | Emma       | Finn       | Grace      |
| 10:00-15:00 | Henry       | Isla       | Jack       | Karen      | Liam       | Mia        | Noah       |
| 10:30-15:30 | Olivia      | Peter      | Quinn      | Rachel     | Sam        | Tina       | Victor     |
| 11:30-16:30 | Wendy       | Xavier     | Yvonne     | Zach       | Amy        | Brian      | Chloe      |

## ✅ Accepted Data Formats

To ensure proper processing, the following data formats must be used:

### **1. Days of the Week**
- Must be written as:  
  `Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday`

### **2. Dates**
- Format: `26th Jan`, `27th Jan`, `28th Jan`  
- The day should be a number with `st`, `nd`, `rd`, or `th` suffix, followed by a three-letter month abbreviation.

### **3. Shift Times**
- Format: `HH:MM-HH:MM`  
- Example: `09:00-17:00`, `10:30-15:30`  
- Must use a 24-hour format.

### **4. Employee Names**
- Each shift slot should contain only one name.
- Example: `Alice`, `Bob`, `Charlie`, etc.

## Installation
### Prerequisites
- Python 3.8+
- pip
- Virtual environment (optional but recommended)

### Setup
1. Clone the repository:
   ```sh
   git clone https://github.com/your-repo/shift-calendar.git
   cd shift-calendar
   ```
2. Create and activate a virtual environment:
   ```sh
   python -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```
3. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```
4. Run the Django server:
   ```sh
   python manage.py runserver
   ```
5. Access the app at `http://127.0.0.1:8000/`.

## Usage
1. Navigate to the web interface.
2. Select an Excel file (.xlsx) that contains shift data.
3. Enter the name of the person whose shifts you want to extract.
4. Click "Upload" to process the file.
5. Download the generated ICS file and import it into your preferred calendar application.

## Deployment
This project is deployed on **Vercel**. Below is the relevant `vercel.json` configuration:
```json
{
  "routes": [
    {
      "src": "/(.*)",
      "dest": "api/wsgi.py"
    }
  ],
  "functions": {
    "api/wsgi.py": {
      "maxDuration": 60
    }
  }
}
```
To deploy manually:
1. Install Vercel CLI:
   ```sh
   npm install -g vercel
   ```
2. Deploy the project:
   ```sh
   vercel
   ```

## Troubleshooting
- **Timeout issues**: Ensure `maxDuration` is set appropriately in `vercel.json`.
- **File upload errors**: Verify that your Excel file follows the expected format.
- **Calendar import issues**: Make sure your calendar app supports `.ics` files.
