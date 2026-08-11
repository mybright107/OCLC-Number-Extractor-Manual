# OCLC-Number-Extractor-Manual
After the OCLC Data Sync job is completed, you may see “Records not imported upon no match.” Since the downloaded Binary (MARC) file appears empty, my workaround is to download the XML file and use this tool to extract the OCLC numbers.


**Background**

This tool is created to extract OCLC numbers from the XML file downloaded from Alma.
After the OCLC Data Sync import jobs are completed, you may see some **Records Not Imported** section, specifically **“Records not imported upon no match”** records. This means that our holdings are set in OCLC, but we don't have the records in Alma.
In this case, I delete the holdings in OCLC.
We can download the file in either **XML** or **Binary (MARC)** format. However, I noticed that the downloaded MARC file does not contain any content once you open the file. My workaround is to download the **XML file** and then use this tool to extract the OCLC numbers from the file.

**Purpose**

The OCLC Number Extractor reads a MARC XML file and creates an Excel workbook containing OCLC numbers found in the records.
The workbook separates records into online and physical resource worksheets.

**Running the tool**

1. Open the OCLC_Extractor folder.
2. Double-click OCLC_Extractor.exe.
3. In the file-selection window, locate the XML file you downloaded from Alma. Extract the zip folder. You'll see file 0, file 1, and more depending on the file size.
4. Select the XML file you need to process
5. Click Open. Wait for the completion message.
6. Click OK.

The tool saves the Excel workbook in the same folder as the selected XML file.
For example, if the input file is: file_0.xml
the output will be: file_0_oclc_numbers.xlsx
If that workbook already exists, the tool creates a new version such as: file_0_oclc_numbers_2.xlsx
Existing workbooks are not overwritten.

**Excel workbook contents**

The workbook contains three worksheets:
**1. All records** — every record with a recognized OCLC number
**2. Online**— records whose MARC 300 field contains the phrase “online resource”
**3. Physical** — records whose MARC 300 field does not contain the phrase “online resource”

Each worksheet contains these columns:
**1. OCLC_Number** — the numeric portion of the OCLC identifier
**2. Full_035_Field** — the extracted identifier, such as (OCoLC)123456
**3. Record_Number** — the record’s sequential position in the XML file
**4. Resource_Type** — Online or Physical

**Important interpretation note**
For this tool, “Physical” means that the phrase “online resource” was not found in the record’s MARC 300 field.
This is a processing category, not a complete bibliographic format determination. Review questionable records before using the category for deletion or other irreversible cleanup.

**Troubleshooting**
Wait approximately 10 seconds and try once more.

If it still does not open:
Confirm that the entire application folder was extracted from the ZIP file.
Do not run the application from inside the ZIP file.

