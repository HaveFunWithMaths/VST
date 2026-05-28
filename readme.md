# Bank Account & Statement Processor

A Python-based utility to automate the mapping and processing of raw bank statements into structured, monthly bank ledger templates. This tool parses daily transactions from an input statement, sorts them chronologically, and populates monthly sheets with automated formulas for opening balances, inflows, outflows, closing balances, and totals.

---

## 📁 Repository Structure

Your workspace contains the following core files:

*   **`process_statement.py`**: The core Python script that handles parsing the raw statement, grouping transactions by month/day, writing data, applying styles, and maintaining formulas.
*   **`statementTemplate.xlsx`**: A clean, professionally-styled template for the input bank statement.
*   **`readme.md`**: This guide.

> [!NOTE]
> To run the script, the directory must also contain the Bank Account Ledger template (`OutputTemplate.xlsx` or `OutputTemplate_Backup.xlsx`). This template houses the target sheets (`Apr 2025` through `Mar 2026`) with pre-configured header/footer styles and formatting rules.

---

## 🛠️ Prerequisites & Installation

To run this tool, you need:
1.  **Python 3.6+** installed on your system.
2.  Dependencies specified in `requirements.txt`.

Install the required dependencies via your terminal or command prompt:
```bash
pip install -r requirements.txt
```

---

## 🚀 How to Use the Tool

Follow these simple steps to process your bank statements:

### Step 1: Prepare the Input Bank Statement (`statement.xlsx`)
1.  Open the provided **`statementTemplate.xlsx`** for reference.
2.  Create a copy of this file and rename it to **`statement.xlsx`** in the same folder as the script.
3.  Fill or paste your raw bank transactions into `statement.xlsx`. Ensure the following columns exist in row 1:
    *   **Value Date**: The date of the transaction (e.g., `YYYY-MM-DD` or other standard Excel date formats).
    *   **Description**: Description or particulars of the transaction.
    *   **Chq / Ref No.**: Cheque or transaction reference number. Use `-` if not applicable.
    *   **Amount**: The absolute value of the transaction.
    *   **Dr / Cr**: Indicate whether it's a Debit (`DR` for outflows/payments) or Credit (`CR` for inflows/receipts).
    *   **Balance**: The running balance of the bank account after this transaction.

### Step 2: Prepare the Ledger Template (`OutputTemplate.xlsx`)
1.  Ensure you have either **`OutputTemplate_Backup.xlsx`** or **`OutputTemplate.xlsx`** in the script's directory.
2.  The script will automatically detect `OutputTemplate_Backup.xlsx` if present (to protect your clean template from being overwritten) and save the populated data to a new `OutputTemplate.xlsx`.

### Step 3: Run the Processing Script
Open a terminal in the folder containing `process_statement.py` and run:
```bash
python process_statement.py
```

### Step 4: Review the Output
1.  Open the newly generated or modified **`OutputTemplate.xlsx`**.
2.  Navigate through the tabs (`Apr 2025`, `May 2025`, etc.).
3.  Each sheet is filled chronologically day-by-day.
4.  All transaction details are neatly formatted.
5.  Opening/closing balances and the **Tally** rows at the bottom are calculated dynamically using Excel formulas (e.g., `SUBTOTAL`), ensuring your sheets remain interactive.

---

## ⚙️ How It Works (Under the Hood)

1.  **Dynamic Initial Balance**: The script automatically calculates the starting balance of the fiscal year (April 1st, 2025) by looking at the first transaction's amount, direction, and running balance.
2.  **Chronological Sorting**: Transactions are sorted by date and sequence, ensuring order is preserved even for multiple transactions on the same day.
3.  **Cross-Month Linking**: The opening balance of each sheet is formula-linked to the closing balance of the preceding month's **Tally** row.
4.  **Style Cloning**: All rows, fonts, fills, alignments, borders, and number formats are dynamically cloned from row 4 of the template, preserving the original design.
5.  **Dynamic Tally Row**: Generates a dynamic "Tally" row for each month that sums Inflows and Outflows using `=SUBTOTAL(109, ...)` and updates sheet tables accordingly.

---

## 🔍 Troubleshooting & FAQs

> [!WARNING]
> **Error: Statement file not found...**
> Ensure your raw statement is saved exactly as `statement.xlsx` in the same directory as the script.

> [!IMPORTANT]
> **Warning: could not parse date in statement row...**
> If you see this warning in the terminal, check the date column in `statement.xlsx` for that row. Dates should be in a standard format (e.g., `YYYY-MM-DD` or numeric Excel dates).

> [!TIP]
> **Running for a new financial year?**
> Update the sheet names inside `OutputTemplate.xlsx` and the calendar mapping inside `process_statement.py` (lines 173-178) to match the new year.

---

## 🌐 Public Hosting via Vercel

Since the web application is built as a single-page static site (HTML, CSS, and client-side JavaScript), you can host it publicly on **Vercel** for free.

### Step 1: Install Vercel CLI (Optional)
If you want to deploy from the command line:
1. Install Vercel globally:
   ```bash
   npm install -g vercel
   ```
2. Run the deployment command in the project folder:
   ```bash
   vercel
   ```
3. Follow the CLI prompts to log in and set up your project.

### Step 2: Deploy via GitHub (Recommended)
For automatic continuous deployments:
1. Push this project folder to a repository on **GitHub**, **GitLab**, or **Bitbucket**.
2. Go to [Vercel.com](https://vercel.com/) and log in with your Git provider.
3. Click **Add New** -> **Project**.
4. Import your repository.
5. Vercel will automatically detect the static project and the `vercel.json` settings. Click **Deploy**.
6. Within a minute, your web application will be live at a public URL (e.g., `https://your-project-name.vercel.app`).
