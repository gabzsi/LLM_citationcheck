# LLM Citation Check

> Python tools for analyzing citation outputs from Large Language Models.

---

## 📦 Requirements

Install these packages before running any scripts:

```bash
pip install python-docx openpyxl pandas requests
```

---

## 🛠️ Scripts

### 📄 `citation_lookup.py`
Checks the validity of citations found in Word documents.

**Expected citation format:**
```
ISSN(s) | journal | author | volume | issue | page | year | type | key | DOI
```

**How it works:**
- Prompts you to select Word documents to analyze
- Calls the **CrossRef API** to validate citations
- Outputs one Excel sheet with multiple columns
- Categorizes citations by LLM engine and topic (extracted from filename)

**Example filename:** `chatgpt_perovskite.docx`

> ⚠️ Requires an internet connection.
>
> ✅ Human review of results is strongly advised before running other scripts.

---

### 📊 `citation_analysis.py`
Performs analysis on the Excel sheet created by `citation_lookup.py`.

- Handles multiple Excel files and merges them into a Combined Analysis sheet

**Analyses include:**
- Invalidity rates
- Publisher preference
- Publication years
- Top 5 journals (all citations & valid only)
- Publisher preference by topic

---

### 🔓 `paywallcheck.py`
Checks whether citations are open-access or behind a paywall, using the **Unpaywall API**.

- Prompts you to select the analysis Excel file
- Only runs on **valid citations** (uses DOI to query the API)

> ⚠️ Requires an internet connection.

---

### 🔍 `duplicatefind.py`
Finds DOIs that appear more than once — across different LLMs or across multiple prompts from the same LLM.

---

## ⚠️ Notes

- These scripts were written for a specific analysis workflow and may require adjustments for other use cases.
- The Word files used in the original analysis are also included in this repository.
