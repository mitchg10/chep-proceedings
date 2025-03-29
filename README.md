# Booklet Creator

This script, `booklet_creator.py`, generates a conference booklet in Microsoft Word format (`.docx`) from a CSV file containing session details. The booklet includes session titles, authors, affiliations, abstracts, proposals, and references.

---

## Prerequisites

Before running the script, you need to set up your computer with the required tools and libraries. Follow these steps:

---

### Step 1: Install Python

1. **Check if Python is already installed**:
   - Open the **Terminal** (on Mac) or **Command Prompt** (on Windows).
   - Type: `python3 --version` and press Enter.
   - If you see a version number (e.g., `Python 3.x.x`), Python is installed. If not, proceed to the next step.

2. **Install Python**:
   - Go to the [Python website](https://www.python.org/downloads/).
   - Download and install the latest version of Python for your operating system.
   - During installation, make sure to check the box that says **"Add Python to PATH"** (on Windows).

---

### Step 2: Install Required Libraries

1. Open the **Terminal** (on Mac) or **Command Prompt** (on Windows).
2. Install the required libraries by typing the following command and pressing Enter:
   ```bash
   pip install pandas python-docx
    ```

This will install:
- `pandas`: A powerful data manipulation library for Python.
- `python-docx`: A library for creating and updating Microsoft Word (.docx) files.

---

### Step 3: Get/Prepare your CSV file
1. Make sure your CSV file is named chep_data.csv and contains the following columns:
    - **Submission title**: The title of the session.
    - **Submission authors**: The authors of the session, with superscripts for affiliations.
    - **Affiliations**: The institutions corresponding to the superscripts.
    - **Abstract**: The abstract of the session.
    - **Proposal**: The proposal text.
    - **References**: Any references for the session.
  2. Place the `chep_data.csv` file in the same directory as the `booklet_creator.py` script.

---

### Step 4: Download and Run the Script

#### Step 4.1: Clone the Repository (Using Git)
1. Open the **Terminal** (on Mac) or **Command Prompt** (on Windows).
2. Navigate to the directory where you want to download the script. Use the `cd` command to change directories. For example:
   ```bash
   cd path/to/your/desired/directory
   ```
3. Clone the repository containing the script by typing the following command and pressing Enter:
   ```bash
   git clone https://github.com/mitchg10/chep-proceedings.git
   ```

#### Step 4.2: Download the Script Manually (If Git is Not Installed)
1. Open a web browser and navigate to the GitHub repository URL:
   ```
   https://github.com/mitchg10/chep-proceedings.git
   ```
2. Click the green **Code** button and select **Download ZIP**.
3. Extract the downloaded ZIP file to a directory of your choice.

#### Step 4.3: Navigate to the Script Directory
1. Open the **Terminal** (on Mac) or **Command Prompt** (on Windows).
2. Navigate to the directory containing the script. For example:
   ```bash
   cd path/to/booklet-creator
   ```
3. Use the `ls` (on Mac) or `dir` (on Windows) command to list the files in the directory and ensure the `booklet_creator.py` script is present.

#### Step 4.4: Run the Script
1. Run the script by typing the following command and pressing Enter:
   ```bash
   python3 booklet_creator.py
   ```
2. The script will read the `chep_data.csv` file and generate a Word document named `booklet.docx` in the same directory.

---

### Troubleshooting
If you encounter any issues while running the script, here are some common troubleshooting steps:
1. **Python not found**: If you see an error saying `python3: command not found`, make sure Python is installed and added to your PATH.
2. **Module not found**: If you see an error saying `ModuleNotFoundError: No module named 'pandas'` or `ModuleNotFoundError: No module named 'docx'`, make sure you have installed the required libraries using pip.
3. **CSV file not found**: If you see an error saying `FileNotFoundError: [Errno 2] No such file or directory: 'chep_data.csv'`, make sure the CSV file is in the same directory as the script and is named correctly.
4. **Permission denied**: If you see a permission error, make sure you have the necessary permissions to read the CSV file and write to the directory.
5. **Invalid CSV format**: If the script fails to read the CSV file, check that it is properly formatted and contains the required columns.
6. **Word document not generated**: If the script runs without errors but the Word document is not created, check the script for any print statements or error messages that might indicate what went wrong.

---

### Output
The script will generate a Word document named `booklet.docx` in the same directory as the script. The document will contain:
1. **Session Title**: The title of the session.
2. **Authors and Affiliations**: Authors listed with their institutions.
3. **Abstract**: The abstract of the session.
4. **Proposal**: The proposal text.
5. **References**: Any references provided.

**Note**: You will have to customize the Word document further to match your desired formatting and layout. The script provides a basic structure, but you may want to adjust fonts, styles, and other formatting options in Word after generating the document.
---

### You're all set!
Now you can create your conference booklet using the `booklet_creator.py` script! If you have any questions or need further assistance, feel free to ask.