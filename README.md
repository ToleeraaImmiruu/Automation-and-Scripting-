# Excel Automation with Python (openpyxl)

This project demonstrates how to use **Python** and **openpyxl** to automate common Excel tasks such as:

* Reading Excel files
* Filtering data
* Cleaning data
* Detecting duplicates
* Renaming files
* Moving files
* Creating reports
* Writing and formatting Excel sheets

It is useful for anyone learning **Python automation, data processing, and Excel scripting.**

---

## 🚀 Features

This project includes examples of how to:

* Read Excel workbooks and worksheets
* Loop through rows and columns
* Convert Excel rows into Python dictionaries
* Filter students based on conditions
* Compare strings and numbers correctly
* Detect duplicate records
* Rename multiple files automatically
* Move files based on extension
* Create new Excel files from Python
* Format Excel cells (bold, color, etc.)

---

## 🛠️ Technologies Used

* Python 3
* openpyxl
* os module
* shutil module

---

## 📂 Example Data Used

Sample student data structure:

| Name  | Sex | Age | Dept |
| ----- | --- | --- | ---- |
| Toli  | M   | 23  | Soft |
| Anaan | F   | 21  | IT   |
| Ayyu  | M   | 22  | IT   |
| Simo  | M   | 24  | Soft |

---

## 📌 Key Concepts Demonstrated

### 1. Reading Excel

```python
from openpyxl import load_workbook

wb = load_workbook("students.xlsx")
ws = wb.active
```

### 2. Looping through rows

```python
for row in ws.iter_rows(min_row=2, values_only=True):
    print(row)
```

### 3. Filtering data

```python
for row in ws.iter_rows(min_row=2, values_only=True):
    name, sex, age, dept = row
    age = int(age)

    if sex == "M" and age > 20:
        print(f"{name} qualifies")
```

### 4. Detecting duplicates

```python
seen = set()
for row in ws.iter_rows(min_row=2, values_only=True):
    name = row[0]
    if name in seen:
        print("Duplicate:", name)
    else:
        seen.add(name)
```

### 5. Moving files by type

```python
import os, shutil

folder = r"C:\Users\Toli\Desktop\action"
dest = r"C:\Users\Toli\Desktop\Images"

os.makedirs(dest, exist_ok=True)

for file in os.listdir(folder):
    full_path = os.path.join(folder, file)
    if file.endswith(".jpg"):
        shutil.move(full_path, dest)
```

---

## 📦 Installation

Install required package:

```bash
pip install openpyxl
```

---

## 📖 How to Use

1. Clone the repository
2. Install dependencies
3. Update file paths to match your system
4. Run scripts using Python

---

## 🎯 Purpose of This Project

This project is meant for:

* Learning Python automation
* Practicing Excel processing
* Understanding data cleaning
* Building real-world automation skills

---

## 📬 Contact

If you are improving this project or using it for learning, feel free to contribute!

Happy coding! 🚀
