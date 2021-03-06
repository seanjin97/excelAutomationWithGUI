Template code to check all values in an excel file and populate another excel file with the correct values.

How to use:

1. Clone repository and install type "pip install -r requirements.txt" in CLI
2. Generate the UOB branch codes using pdf_to_csv/pdf_to_csv.ipynb if not already done
3. Run app.pyw
4. Locate file to check, file to populate and the UOB branch code reference files in the GUI.

   ![GUI](screenshots/GUI.jpg)

5. Change the fields in the GUI to fit your needs.
6. Click save
7. Click Check

- Before check

![Before Check](screenshots/check_before.jpg)

- After check

![After Check](screenshots/check_after.jpg)

8. Click Populate. Only rows marked as TRUE will be copied and populated. Rows marked as FALSE will be ignored.

- Before populate

![Before populate](screenshots/populate_before.jpg)

- After populate

![After populate](screenshots/populate_after.jpg)
