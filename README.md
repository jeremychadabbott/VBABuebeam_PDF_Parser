# VBABuebeam_PDF_Parser
This tool iterates over PDFs in a folder, opens them, converts and parses the content for data entry into financial software.

It was developed to replicate a paperwork flow, where we receive PDFs in our "Fax" file, and processing begins at that point.

These .bas files target specific types of PDFs we commonly receive in our office. They parse the data for entry into our financial software. Different vendors have different PDF formats, requiring different algorithm approaches for each vendor's PDFs.

To use these .bas files in your project, load them into an XLSM module and connect a command to them. I maintain them here for version control.

Some of the .bas modules will launch Chrome to verify that the data found on the PDF is valid. Website data, if different from that found on the PDF, will override the PDF data. This is because sometimes PDF legibility is poor, and the website will be more accurate.
