# CSV to XLSX and PDF to XLSX Conversion Tool

>**Note**:<br>
>1- The CSV file needs to be in .csv format. If your file is in .numbers format, you must export it to CSV first.<br>
>2- There are online tools available for free that can convert .numbers to .csv.

## Features
1- CSV to XLSX: Update an Excel file with data from a CSV file.<br>
2- PDF to XLSX: Extract data from a PDF file and update the Excel file accordingly.

## Instructions for Windows Users:
### Pre-Generated Executables:
* Executables (Installer.exe and Application.exe) are already provided.

* If you need to regenerate the executables:<br>
Open Command Prompt in this folder and type:
```bash
python generate_executables.py
```

### Steps to run program:
1- Run Installer.exe to install all required libraries.<br>
2- Run Application.exe to start the program.<br>
3- Select the desired operation: CSV to XLSX or PDF to XLSX.<br>
4- Provide the required file paths:<br>
&emsp; **For CSV to XLSX**: CSV file path and Excel file path.<br>
&emsp; **For PDF to XLSX**: PDF file path and Excel file path.<br>
5- Click Run.<br>
6- Once processing is complete, the updated Excel file will be saved in the same folder as the original Excel file.<br>

### Additional Modifications:
#### PDF to XLSX Matching Enhancements:
&emsp; 1- Data extracted from the PDF is matched to the correct worksheet based on the year and category in the Excel file.<br>
&emsp; 2- Sheet names are matched dynamically, so they can contain variations of "RESPITE" or "PERSONAL CARE" (e.g., "Respite 2024" or "personal 2025").<br>
&emsp; 3- Data matching includes patient names, dates of service, and charges.<br>

#### Automatic Cleanup:
&emsp; After processing, the temporary file pdf_output.csv is automatically deleted to avoid clutter.<br>

#### Fixed Window Size:
&emsp; The application window is now set to a fixed size to prevent resizing.<br>

#### Application Test
![CSV to XLSX Option](https://github.com/omarmustafa130/csv-data-to-xlsx/blob/main/Assets/Application_test.png)
![PDF to XLSX Option](https://github.com/omarmustafa130/csv-data-to-xlsx/blob/main/Assets/Application_test2.png)



#### Enjoy :)
<b>OM<br><b>

