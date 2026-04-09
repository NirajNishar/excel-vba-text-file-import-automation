# 📂 Excel VBA Text File Import Automation

## 📌 Project Overview

This project is an Excel VBA-based automation tool that imports multiple text files into Excel using a single macro.

Each selected file is automatically loaded into a separate worksheet, eliminating repetitive manual work and improving efficiency in handling structured text data.

---

## 🎯 Objectives

* Automate the process of importing multiple text files into Excel
* Reduce manual copy-paste effort
* Create a separate worksheet for each imported file
* Improve workflow efficiency using VBA macros
* Build a practical Excel automation project for portfolio use

---

## 📊 Dataset / Data Source

The project uses sample monthly sales data stored as text files.

### 📁 Sample Files

* Apr2015Sales.txt
* Dec2015Sales.txt
* Feb2015Sales.txt
* Jan2015Sales.txt
* Mar2015Sales.txt
* May2015Sales.txt
* Nov2015Sales.txt

📍 Location:

```text
data/sample-files/
```

Each file represents monthly sales data and is imported into a separate worksheet.

---

## 🛠️ Tools & Technologies

* Microsoft Excel
* VBA (Visual Basic for Applications)
* Excel Macros
* File Handling in VBA
* Git & GitHub

---

## ⚙️ Key Features

* Select multiple text files at once
* Automatically import each file into a new worksheet
* Batch processing using VBA
* Dynamic worksheet creation
* Sheet naming based on file names
* Simple and beginner-friendly automation

---

## 🔄 How It Works

1. Run the `Import_Text_File` macro
2. Select one or more text files
3. Each file is opened automatically
4. Data is copied from the first sheet
5. A new worksheet is created
6. Data is pasted into the worksheet
7. The sheet is renamed using the file name
8. The source file is closed
9. The process repeats for all selected files

---

## 📸 Screenshots

### 📁 File Selection

![File Selection](screenshots/file-selection.PNG)

---

### 🟡 Before Import

![Before Import](screenshots/before-import.PNG)

---

### 🟢 After Import

![After Import](screenshots/after-import.PNG)

---

## 📂 Project Structure

```text
excel-vba-text-file-import-automation/
├── data/
│   └── sample-files/
│       ├── Apr2015Sales.txt
│       ├── Dec2015Sales.txt
│       ├── Feb2015Sales.txt
│       ├── Jan2015Sales.txt
│       ├── Mar2015Sales.txt
│       ├── May2015Sales.txt
│       └── Nov2015Sales.txt
├── screenshots/
│   ├── file-selection.png
│   ├── before-import.png
│   └── after-import.png
├── vba/
│   └── modules/
│       └── Module1.bas
├── excel-vba-text-file-import-automation.xlsm
└── README.md
```

---

## ▶️ How to Run This Project

1. Clone or download this repository
2. Open the `.xlsm` file in Microsoft Excel
3. Enable macros
4. Run the `Import_Text_File` macro
5. Select the text files
6. The data will be imported automatically

---

## 📈 Key Insights

* VBA can significantly reduce repetitive Excel tasks
* Batch file processing improves efficiency
* Automating file imports saves time and reduces errors
* Even simple VBA projects can solve real-world problems
* Clean structure and documentation improve project presentation

---

## 🧾 Conclusion

This project demonstrates how Excel VBA can be used to automate file import tasks efficiently.

It highlights the practical use of macros in handling multiple files and showcases a beginner-friendly automation workflow suitable for real-world scenarios.

---

## 👨‍💻 Author

**Niraj Nishar**

---

## ⭐ If you found this project useful

Give it a ⭐ on GitHub
