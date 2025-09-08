# 📊 Excel Data Processing Project

## 📌 Overview

This project automates Excel data handling tasks using **Python (pandas & openpyxl)**. It processes two Excel files, performs data transformations, and generates new sheets with useful summaries.

### Key Features

* Reads and processes multiple Excel sheets.
* Generates a **Count Summary** sheet with occurrences of unique values.
* Performs a **VLOOKUP-style join** between two datasets.
* Writes processed results back to the Excel file without modifying original data.

---

## 🛠️ Technologies Used

* **Python 3**
* **pandas** – for data manipulation.
* **openpyxl** – for writing results to Excel.
* **Jupyter Notebook** – for code execution and experimentation.

---

## 📂 Project Structure

```
├── Project_01.ipynb   # Jupyter Notebook with the full code
├── WS1.xlsx           # Input Excel file 1
├── WS2.xlsx           # Input Excel file 2 
├── Output_File_WS2    # Output Excel file (processed and updated)
```

---

## 🚀 How to Run

1. **Clone the Repository**

```bash
 git clone https://github.com/your-username/excel-data-processor.git
 cd excel-data-processor
```

2. **Install Dependencies**

```bash
pip install pandas openpyxl jupyter
```

3. **Run the Notebook**

```bash
jupyter notebook Project_01.ipynb
```

4. **Output**

* `WS2.xlsx` will be updated with:

  * **Count Summary** sheet (unique values & counts)
  * **Joined Sheet** with additional `Days for Settlement` column

---

## 📊 Example Workflow

* Input: `WS1.xlsx` and `WS2.xlsx`
* Process:

  * Count occurrences of `A/C Reference` in `WS2`.
  * Join with `Days for Settlement` from `WS1`.
* Output: Updated `WS2.xlsx` with new sheets.

---

## 🤝 Contributing

Contributions are welcome! Feel free to fork the repo, open issues, or submit pull requests.

---
