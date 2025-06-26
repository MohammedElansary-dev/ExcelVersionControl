# ExcelVersionControl

## 📦 One-Click Excel Backup & Version Control (VBA Macro)

This is a plug-and-play Excel VBA macro that allows users to back up their current workbook in a **single click**, with support for:

* `.xlsm` (macro-enabled backup),
* `.xlsx` (macro-free backup),
* `.csv` (active sheet export),
* automatic timestamping,
* version control via filename,
* and backup run logging.

---

## 🔧 Features

✅ Save backup copies in selected formats: `.xlsm`, `.xlsx`, `.csv`
✅ Automatically names backups with a date-time stamp
✅ Creates a `Backups/` folder in the workbook directory if it doesn’t exist
✅ Maintains a visible sheet called `BackupLog` to track all backup runs
✅ Requires no external tools, libraries, or installation
✅ Fully customizable and beginner-friendly

---

## 💻 How It Works

When the macro `OneClickBackupAndVersionControl` is executed:

1. You are asked which formats to save (`.xlsm`, `.xlsx`, `.csv`).
2. The chosen files are saved in a subfolder called `Backups` located next to the original workbook.
3. Each backup file is named like:

   ```
   WorkbookName_Backup_2024-06-25_18-45-12.xlsm
   ```
4. A log entry is created (or appended) in a sheet called `BackupLog` with:

   * Timestamp
   * Backup file base name
   * Formats saved
   * Windows username
   * Path to the backups

---

## 📂 Folder Structure

```text
📁 YourWorkbookFolder/
├── YourWorkbook.xlsm
├── 📁 Backups/        <--- created automaticly
│   ├── YourWorkbook_Backup_YYYY-MM-DD_HH-MM-SS.xlsm
│   ├── YourWorkbook_Backup_YYYY-MM-DD_HH-MM-SS.xlsx
│   ├── YourWorkbook_Backup_YYYY-MM-DD_HH-MM-SS.csv
└── 📄 BackupLog (Excel sheet inside the workbook)
```

---

## 📌 Installation

1. Open your Excel workbook (`.xlsm` recommended).
2. Press `ALT + F11` to open the **VBA Editor**.
3. Go to `File > Import File...` and import `MakeBackUp.bas`.
4. Press `CTRL + S` to save the workbook.
5. Optionally: Add a button to the ribbon or Quick Access Toolbar to run `OneClickBackupAndVersionControl`.

---

## 🛠️ Requirements

* Excel 2016 or later (with support for `.xlsm`)
* Macro-enabled Excel file (`.xlsm`)
* Macros must be **enabled** for the script to run

---

## 🚀 Usage

You can run the macro manually or assign it to a button:

* Press `ALT + F8`, select `OneClickBackupAndVersionControl`, then click **Run**
* Or assign it to a button from `Developer > Insert > Button`

---

## 🧪 To Customize

Edit the following in the script:

* Backup folder name (`Backups`)
* Date format (`yyyy-mm-dd_HH-mm-ss`)
* Which formats to auto-save by default (skip prompts)

---

## 📋 Example Log Output (in `BackupLog` sheet)

| Timestamp           | FileBaseName                         | Formats    | User    | BackupPath                    |
| ------------------- | ------------------------------------ | ---------- | ------- | ----------------------------- |
| 2024-06-25 18:45:12 | MyWorkbook\_Backup\_20240625\_184512 | .xlsm .csv | JohnDoe | C:\Users\Foo\Documents... |

---

## ❓FAQ

**Q: Will this overwrite existing backups?**
A: No — each backup includes a timestamp, so all versions are preserved.

**Q: What if the `BackupLog` sheet doesn’t exist?**
A: The script will create it automatically on first run.

**Q: Can I hide the `BackupLog` sheet?**
A: Yes, but make sure it's not `VeryHidden` unless you're okay with editing it manually later.

---

## 📄 License

MIT License — free to use, modify, and redistribute.

---

## 🙌 Credits

Made by Mohamed El-ansary — inspired by the need to version Excel workbooks without version control software.

---
