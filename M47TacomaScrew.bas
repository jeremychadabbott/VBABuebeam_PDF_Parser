Attribute VB_Name = "M47TacomaScrew"
Sub TacomaScrew(path)

Dim InvoiceDate As String
Dim i As Long
Dim URL As String
Dim IE As Object
Dim objElement As Object
Dim objCollection As Object
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Temp")
Dim rawData As String
Dim splitData() As String
Dim j As Integer
Dim itemDesc As String, itemQuantity As String, itemCost As String
Dim lineItemStart As Boolean

start:
Found = 0

' Scan parent folder path for Tacoma Screw PDF
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(path)
For Each objFile In objFolder.Files
    fname = objFile.Name
    'MsgBox fname
    fpath = objFile.path
    ' Reference JMJ1437437
    If UCase(fname) Like "*JMJ[0-9][0-9][0-9][0-9][0-9][0-9]*" Then
        Found = 1
    End If '
If Found = 1 Then Exit For
Next objFile

' If no Tacoma Screw PDF was found, exit sub
If Found = 0 Then Exit Sub

Call FormatTempSheet
Call Convert_PDF_to_Excel(fname, fpath, XLSpath, XLSname, emailmessage)



'delete bad characters on thisworkbook.temp sheet that will trip the code
For xoffset = 1 To 300
    'MsgBox xoffset
    If Left(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0).Formula, 1) = "=" Then
        MsgBox "detected formula embeded at offset " & xoffset & Chr(13) & ws.Range("BA2").Offset(xoffset, 0).Formula
        Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0) = ""
        'ws.Range("BA2").Offset(xoffset, 0) = Replace(ws.Range("BA2").Offset(xoffset, yoffset), "&", "")
    End If
Next xoffset


' Set Variables
tempsheetoffset = 0
InvoiceDate = ""
Possibleerror = 0
TotalInvoice = ""
VendorInvoiceNo = ""

' Algorythm Scape Macro Data
For xoffset = 0 To 50 'Gather Macro Information
                
    'Assign variable "Line" with the data
    Line = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0)
                   
    'PDFtype
    If Line = "Remit To TACOMA SCREW PRODUCTS INC" And PDFtype = "" Then PDFtype = "Invoice"
    'If Line = "INVOICE - ORIGINAL" And PDFtype = "" Then PDFtype = "Invoice"
    'If Line = "CREDIT MEMO" And PDFtpye = "" Then PDFtype = "Invoice"
                   
    'INVOICE NO ref 260058666-00
    If VendorInvoiceNo = "" Then
        If UCase(Line) Like "*[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]-[0-9][0-9]*" Then
            ' Find the start position of the invoice number pattern
            startPos = InStr(Line, Mid(Line, InStr(Line, "-") - 9, 12))
            ' Extract the invoice number assuming it's of fixed length
            VendorInvoiceNo = Mid(Line, startPos, 12)
            MsgBox "Tacome Screw InvoiceNO=:" & VendorInvoiceNo & ":"
        End If
    End If
                       
    ' INVOICE Date
    If InvoiceDate = "" Then
        If UCase(Line) Like "*[0-9][0-9]/[0-9][0-9]/2[2-9]*" Then
            ' Find the start position of the date pattern
            startPos = InStr(Line, Mid(Line, InStr(Line, "/2") - 5, 10))
            ' Extract the date assuming it's in the format MM/DD/YYYY
            InvoiceDate = Mid(Line, startPos, 8)
            ' Error check
            If Not InvoiceDate Like "*/*/*" Then InvoiceDate = ""
            MsgBox "InvoiceDate=:" & InvoiceDate & ":"
        End If
    End If
                   
    ' DECO PO
    If DecoPO = "" Then
        If UCase(Line) Like "*NET30*" Then
            ' Find the start position of "Net30"
            startPos = InStr(UCase(Line), "NET30")
            If startPos > 0 Then
                ' Adjust the starting position to the end of "Net30" and skip extra spaces
                startPos = startPos + Len("NET30")
                Do While Mid(Line, startPos, 1) = " "
                    startPos = startPos + 1
                Loop
                
                ' Extract the substring that follows "Net30"
                endPos = InStr(startPos, Line & " ", " ")
                TargetPO = Trim(Mid(Line, startPos, endPos - startPos))
                
                ' Check if PO is valid
                Call CheckPONumber(TargetPO, Found)
                ' Found variable Index:
                ' Found = 0 TargetPO does not conform, refuse to process
                ' Found = 1 TargetPO is OK to Process
                ' Found = 2 TargetPO is Subcontract, Send PDF to Fax File
                ' Found = 3 TargetPO is SHOP, Send PDF to Fax File
                If Found = 1 And TargetPO <> "" Then
                    DecoPO = TargetPO
                    MsgBox "decoPO=:" & DecoPO & ":"
                End If
            End If
        End If
    End If

     
    'ORDER DATE - Tacoma Screw invoices do not have seperate order dats and invoice dates
    If InvoiceDate <> "" Then OrderDate = InvoiceDate

    ' TOTAL $ ref <Qty Shipped Total Merchandise Total  $ 52.16>
    If VendorTotalInvoice = "" Then
        ' Find the position of the keyword "Merchandise Total  $"
        startPos = InStr(1, UCase(Line), "MERCHANDISE TOTAL  $")
    
        ' Check if the keyword is found
        If startPos > 0 Then
            ' Adjust the starting position to the end of the keyword
            startPos = startPos + Len("MERCHANDISE TOTAL  $")
            
            ' Skip any spaces following the keyword
            Do While Mid(Line, startPos, 1) = " "
                startPos = startPos + 1
            Loop
            
            ' Find the end position which is the next space or end of the string
            endPos = InStr(startPos, Line & " ", " ")
            
            ' Extract the substring
            VendorTotalInvoice = Trim(Mid(Line, startPos, endPos - startPos))
            
            ' Remove any character that is not a number or a decimal
            Dim cleanInvoice As String
            Dim char As String
            cleanInvoice = ""
            For i = 1 To Len(VendorTotalInvoice)
                char = Mid(VendorTotalInvoice, i, 1)
                If IsNumeric(char) Or char = "." Then
                    cleanInvoice = cleanInvoice & char
                End If
            Next i
            
            ' Assign cleaned value back to VendorTotalInvoice
            VendorTotalInvoice = cleanInvoice
            
            ' Display the extracted value
            MsgBox "VendorTotalInvoice=:" & VendorTotalInvoice & ":"
        End If
    End If


                    
Next xoffset
            
MsgBox "finished macro data scrape"


'Get Line Items
'Scan for each product line items and invoice total
Set ws = Workbooks(XLSname).Sheets(1)
    ' Loop through all rows in column A to search for line items
    i = 1
    lineItemStart = False
    Do While ws.Cells(i, 1).Value <> ""
        rawData = ws.Cells(i, 1).Value
        
        ' Find product line items
        If InStr(1, rawData, "***EXT***") > 0 Then
            ' Extract item quantity (assumed to be numeric and before the price)
            itemQuantity = ""
            itemCost = ""
            itemDesc = ""
            splitData = Split(rawData, "  ")  ' Split the data using double spaces as the delimiter
            
            For j = 0 To UBound(splitData)
                'If splitData(j) <> "" Then MsgBox "j:" & j & Chr(13) & splitData(j)
                If IsNumeric(splitData(j)) And itemQuantity = "" And j > 2 Then
                    itemQuantity = splitData(j)
                ElseIf IsNumeric(Replace(splitData(j), ".", "")) And j > 60 And itemCost = "" Then
                    itemCost = splitData(j)
                ElseIf Len(splitData(j)) > 20 And itemDesc = "" Then
                    itemDesc = splitData(j)
                    itemDesc = Replace(itemDesc, "***EXT***", "")
                    For Repeat = 1 To 4
                        If Left(itemDesc, 1) = vlbf Or Left(itemDesc, 1) = " " Then itemDesc = Right(itemDesc, Len(itemDesc) - 1)
                    Next Repeat
                    Dim position As Integer
                    position = InStr(itemDesc, Chr(10))
                    If position > 0 And position < 20 Then
                        itemDesc = Mid(itemDesc, position + 1)
                    End If
                    itemDesc = Replace(itemDesc, vlbf, "")
                    itemDesc = Replace(itemDesc, Chr(13), "")
                    itemDesc = Replace(itemDesc, Chr(10), "")
                    'MsgBox "itemdesc :" & itemDesc & ":"
  
                End If
            Next j

            MsgBox "Desc:" & itemDesc & Chr(13) & "Cost:" & itemCost & Chr(13) _
                & "Qty:" & itemQuantity
            
            ' Output the extracted information
             'Write the Data to the Sheet
            ThisWorkbook.Sheets("Temp").Range("P2").Offset(itemline, 0) = itemDesc
            ThisWorkbook.Sheets("Temp").Range("Q2").Offset(itemline, 0) = Unit
            ThisWorkbook.Sheets("Temp").Range("R2").Offset(itemline, 0) = itemQuantity
            ThisWorkbook.Sheets("Temp").Range("S2").Offset(itemline, 0) = itemCost 'unitprice
            ThisWorkbook.Sheets("Temp").Range("T2").Offset(itemline, 0) = linetotal
            ThisWorkbook.Sheets("Temp").Range("A2").Offset(itemline, 0) = DecoPO
            ThisWorkbook.Sheets("Temp").Range("B2").Offset(itemline, 0) = OrderDate
            ThisWorkbook.Sheets("Temp").Range("C2").Offset(itemline, 0) = "267"
            ThisWorkbook.Sheets("Temp").Range("AH2").Offset(itemline, 0) = Tax
            ThisWorkbook.Sheets("Temp").Range("H2").Offset(itemline, 0) = VendorInvoiceNo
            ThisWorkbook.Sheets("Temp").Range("N2").Offset(itemline, 0) = VendorTotalInvoice
            ThisWorkbook.Sheets("Temp").Range("J2").Offset(itemline, 0) = InvoiceDate
            itemline = itemline + 1
        End If
        
        'find invoice total
        If InStr(1, rawData, "Balance Due") > 0 Then
            splitData = Split(rawData, "$")  ' Split the data using double spaces as the delimiter
            For j = 0 To UBound(splitData)
                VendorTotalInvoice = splitData(j)
                VendorTotalInvoice = Replace(VendorTotalInvoice, " ", "")
                VendorTotalInvoice = Replace(VendorTotalInvoice, vbLf, "")
                VendorTotalInvoice = Replace(VendorTotalInvoice, Chr(13), "")
                VendorTotalInvoice = Replace(VendorTotalInvoice, Chr(10), "")
                'MsgBox "VendorTotalInvoice" & Chr(13) & ":" & VendorTotalInvoice & ":"
                If IsNumeric(VendorTotalInvoice) Then
                    MsgBox "Found VendorTotalInvoice" & Chr(13) & ":" & VendorTotalInvoice & ":"
                    ThisWorkbook.Sheets("Temp").Range("N2") = VendorTotalInvoice
                End If
            Next j
        End If
        
        i = i + 1
    Loop


'confirm write vendor to temp sheet
ThisWorkbook.Sheets("Temp").Range("C2") = "267"

MsgBox "Freeze"
MsgBox "Freeze"
MsgBox "Freeze"
MsgBox "Freeze"


' check for common error fixes
Call SelfHealTempPage

' Check if PO conforms before bothering to Enter
TargetPO = ThisWorkbook.Sheets("Temp").Range("A2")
Call CheckPONumber(TargetPO, Found)

' Found variable Index:
' Found = 0 TargetPO does not conform, refuse to process
' Found = 1 TargetPO is OK to Process
' Found = 2 TargetPO is Subcontract, Send PDF to Fax File
' Found = 3 TargetPO is SHOP, Send PDF to Fax File

' Check if missing Vendor Invoice Number (crtitical error)
If VendorInvoiceNo = "" Then MsgBox "Failed to Pickup Tacome Screw Invoice No! Will not be able to input in Sage"

docno:

'If we have a Vendor Invoice Number, do a webscrape
If VendorInvoiceNo <> "" Then
    Call TacomaScrewWebscrape(VendorInvoiceNo)
    'Exit Sub
End If

'Messgae user if webscrape wasn't done
MsgBox "Tacoma Screw module failed to go to webscrape"


' Insert Code here if Found = 2 TargertPO is a subcontract
If Found = 2 Then MsgBox "Tacoma screw PO number is a subcontract number, FREEZE"
  
' If there were no Errors deteted during data scrape, proceed to write to Sage
If Possibleerror < 1 Then

    ' Open Sage
    Call ClickOnSage
    xoffset = 0
    emailmessage = "Tacoma Screw Invoice"
    
    ' Enter Invoice
    Call SageEnterINVOICEfromTEMP(xoffset, emailmessage, fpath)
    If emailmessage = "Temp Sheet Total Error" Then Exit Sub
    
    'Build new File name
    TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
    TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
    If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
    If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
    pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
    & "INVOICE " & AddZdero & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
    & TotalInvoiceAmount & ").pdf"
    pdfoption1 = Replace(pdfoption1, "$", "")
      
    ' If new file name doesn't already exist, rename and move it
    If Dir("\\server2\Faxes\TACOMA SCREW - 267" & "\" & pdfoption1) = "" Then _
    Name fpath As "\\server2\Faxes\TACOMA SCREW - 267\" & pdfoption1
    
    ' Minimize Sage
    SetCursorPos 1083, 11
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Application.Wait (Now + TimeValue("00:00:01"))
    
Else
    'Notify User errors were found during datascrape
    MsgBox "Did not enter " & TargetPO & Chr(13) & "PossibleErrors =" & Possibleerror _
        & Chr(13) & "Found =" & Found & Chr(13) & "Moving to Fax File as;" & Chr(13) _
        & TargetPO & " Tacome Screw"
                   
    'Rename and move PDF as .1 if PDF.name already exists
    If Dir("\\server2\Faxes\" & TargetPO & " Tacoma Screw.pdf") <> "" Then TargetPO = TargetPO & ".1"

    'Rename and Move PDF to fax file
    If Dir(fpath) <> "" And Dir("\\server2\Faxes\" & TargetPO & " Tacoma Screw.pdf") = "" Then Name fpath As "\\server2\Faxes\" & TargetPO & " Tacoma Screw.pdf"
                
    'Delete original PD from parent folder
    If Dir(fpath) <> "" And UCase(path) Like "*ATTACH*" Or emailmessage Like "*already been entered*" And Dir(fpath) <> "" Then Kill (fpath)

End If

'Loop back through parent folder and see if there are any more Tacoma Screw PDF's to process
GoTo start:

End Sub

Sub TacomaScrewWebscrape(VendorInvoiceNo)





End Sub

