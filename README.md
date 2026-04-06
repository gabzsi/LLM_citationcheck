# LLM_citationcheck
Repository contains python codes that were used to analyze citation outputs from LLMs

The Python scripts are dependent on python-docx, openpyxl, pandas and requests packages. Installation of these is required before running the scripts.

citation_lookup.py checks validity of citations from Word files.
The format of citation is expected to be: "ISSN(s)|journal|author|volume|issue|page|year|type|key|DOI".
User is prompted to select Word documents to analyze. After analysis, one Excel sheet is created with multiple columns. Citations are categorized based on LLM engine and topic extracted from word filename.
Example filename: chatgpt_perovskite.docx
Note that execution of this script requires internet connection as CrossRef API is called.
Human revision of analysis is strongly advised prior to running the other Python scripts.
citation_analysis.py performs analysis on the as-created Excel sheet. It can also handle multiple Excel files and will create a Combined Analysis sheet.
Analyses performed: invalidity, publisher preference, publication years, top 5 journals (all citation, valid citations only), publisher preference within topics.
paywallcheck.py checks whether citation is open-access or requires subscription using Unpaywall API.
User is prompted to select the analysis Excel file. Paywall check is performed by passing DOI to API, therefore analysis is only performed on valid citations.
Note that execution of this script requires internet connection as Unpaywall API is called.
duplicatefind.py looks for DOIs that appear more than once within a given run accross different LLMs or accross multiple prompts from the same LLM.
The scripts were written specifically for our analysis and are not guaranteed to work without issues in any case. The World files we analyzed can be also found in this repository.
