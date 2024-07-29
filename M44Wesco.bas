Attribute VB_Name = "M44Wesco"

Sub IfWesco(path, XLSname, XLSpath, fpath, emailmessage)
'--------------------------------------------------------------
'                       WESCO
'--------------------------------------------------------------
    Dim InvoiceDate As String
    Dim i As Long
    Dim URL As String
    Dim IE As Object
    Dim objElement As Object
    Dim objCollection As Object
    Dim try As Integer
    Dim ws As Worksheet
    Dim vendoritemno As String
    Dim fso As Object
    Dim sourcePath As String
    Dim destinationPath As String
    
    Set ws = ThisWorkbook.Sheets("Temp")
    ws.Range("A2:CA200").ClearContents


start:
    fname = ""
    Found = 0

' Scan the parent folder for WESCO PDF's
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(path)
    For Each objFile In objFolder.Files
        fname = objFile.Name
        'MsgBox fname
        fpath = objFile.path
        ' Reference S010203436-0002_26163
        If UCase(fname) Like "*WESCO*" Then
            Found = 1
        End If '
    If Found = 1 Then Exit For
    Next objFile
    
'If No Wesco PDF's foiund, exit sub
    If Found = 0 Then Exit Sub

    Call FormatTempSheet

    
    If fname = "" Then MsgBox "WARNING Fname is nothing"

    
    If Dir(fpath) <> "" Then
     ThisWorkbook.FollowHyperlink Address:=fpath
    End If
        
    Application.Wait (Now + TimeValue("00:00:05"))
        
wait_time = Pages

'Convert to Text

'Set curser on page
    Application.Wait (Now + TimeValue("00:00:01"))
    SetCursorPos 500, 500
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo

' Save-as, Using hot keys
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "+^s", True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "{tab}", True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "t", True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "~", True
    Application.Wait (Now + TimeValue("00:00:05"))
     
    TXTpath = Left(fpath, Len(fpath) - 3) & "txt"
    HTMname = Left(fname, Len(fname) - 3) & "txt"

 
'Copy Notepad Info
    Application.Wait (Now + TimeValue("00:00:03"))
    Application.SendKeys "^a", True
    Set Clipboard = New MSForms.DataObject
    Application.CutCopyMode = False
    Clipboard.Clear
    Application.SendKeys ("^c")
    Sleep 150
    Clipboard.GetFromClipboard
        
' Close Notepad
    Application.SendKeys "%{F4}" 'Close Notepad
    Application.Wait (Now + TimeValue("00:00:02"))
    
'Delete Notepad file
    If Dir(TXTpath) <> "" Then Kill TXTpath

' close bluebeam
    Application.SendKeys "%{F4}" 'Close Notepad
    Application.Wait (Now + TimeValue("00:00:02"))

' Get the data from the clipboard
    clipboardData = Clipboard.GetText

' Split the data into lines
    Lines = Split(clipboardData, vbCrLf)

' Process each line
    x = 0
    y = 0
    For i = LBound(Lines) To UBound(Lines)
        'MsgBox Lines(i)
        ws.Range("BA2").Offset(x, y) = Lines(i)
        x = x + 1
        ' Do something with the line, for example:
    Next i

'-------------------------------------------------------------------------------
'                              PROCESS DATA
'-------------------------------------------------------------------------------

GoTo skip_Save_Product_links:
' Get product numbers
    Dim myRange As Range
    Dim cell As Range
    Dim myString As String
    Dim partialString As String
    Vendor = "Wesco"
    Set myRange = ThisWorkbook.Sheets("Temp").Range("BA2:BA1000") 'change to your desired range
    totalProducts = 0
    For Each cell In myRange
        'Ref Wesco UPC
        If cell.Text Like "*[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]*" Then
            Product_Link = cell.Text
            Product_Link = Replace(Product_Link, " ", "")
            If Product_Link Like "[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]" Then
                Call Save_Product_Links(Product_Link, Vendor, workbook_open_status)
                totalProducts = totalProducts + 1
            End If
            
        End If
    Next cell
    Dim wbname As String
    On Error Resume Next
    wbname = "Product_links.xlsx"
    If Workbooks(wbname).IsOpen Then
        'MsgBox "The workbook is open."
        Workbooks(wbname).Save
        Workbooks(wbname).Close
    Else
        'MsgBox "The workbook is not open."
    End If
    On Error GoTo 0

skip_Save_Product_links:

'preclean data of bad characters that will trip the code
    For xoffset = 1 To 300
        'MsgBox xoffset
        If Left(ws.Range("BA2").Offset(xoffset, 0).Formula, 1) = "=" Then
            
            MsgBox "detected formula embeded at offset " & xoffset & Chr(13) & ws.Range("BA2").Offset(xoffset, 0).Formula
            ws.Range("BA2").Offset(xoffset, 0) = ""
            'ws.Range("BA2").Offset(xoffset, 0) = Replace(ws.Range("BA2").Offset(xoffset, yoffset), "&", "")
        End If
    Next xoffset
    
'Process and get Macro Data
    tempsheetoffset = 0
    InvoiceDate = ""
    Possibleerror = 0
    TotalInvoice = ""
    VendorInvoiceNo = ""
    For xoffset = 1 To 1000
    
       Line = ws.Range("BA2").Offset(xoffset, yoffset)
       Line = Replace(Line, vbLf, "")
       
       'PDFtype
       If Line = "ORDER #" And PDFtype = "" Then PDFtype = "Order"
       If Line = "INVOICE - ORIGINAL" And PDFtype = "" Then PDFtype = "Invoice"
       If Line = "CREDIT MEMO" And PDFtpye = "" Then PDFtype = "Invoice"
       
        'INVOICE NO
        If UCase(Line) Like "INVOICE NUMBER" And VendorInvoiceNo = "" Then '
           VendorInvoiceNo = ws.Range("BA2").Offset(xoffset + 1, yoffset)
           'MsgBox "WESCO InvoiceNO=:" & VendorInvoiceNO & ":"
           'Special exception for Invoice 972186 which is already taken by hardware sales
           If VendorInvoiceNo = "972186" Then VendorInvoiceNo = "972186W"
           
           
        End If
       
       'INVOICE Date
        If UCase(Line) Like "INVOICE DATE" And InvoiceDate = "" Then '
           InvoiceDate = ws.Range("BA2").Offset(xoffset + 1, yoffset)
           If Not InvoiceDate Like "*/*/*" Then InvoiceDate = ""
           'MsgBox "InvoiceDate=:" & InvoiceDate & ":"
        End If
       
       'DECO PO
       If UCase(Line) Like "*CUSTOMER ORDER NUMBER*" Then '
        DecoPO = ws.Range("BA2").Offset(xoffset + 1, yoffset)
            DecoPO = Replace(DecoPO, vbLf, "")
           DecoPO = Replace(DecoPO, "PO", "")
           DecoPO = Replace(DecoPO, "#", "")
           DecoPO = Replace(DecoPO, " ", "")
           'MsgBox "decoPO=:" & DecoPO & ":"
       End If
       
       'ORDER DATE
       If Line Like "*DATE ORDERED*" Then '
           OrderDate = ws.Range("BA2").Offset(xoffset + 1, yoffset)
           If Not OrderDate Like "*/*/*" Then OrderDate = ""
           MsgBox "OrderDate=:" & OrderDate & ":"
       End If

       'DELIVERY METHOD
       'ws.Range("AE2") = ""
       
        'TOTAL $
        If UCase(Line) Like "*TOTAL >*" Then '
           'MsgBox "Freeze"
            Found = 0
            VendorTotalInvoice = Right(UCase(Line), 9)
            VendorTotalInvoice = Replace(VendorTotalInvoice, " ", "")
            VendorTotalInvoice = Replace(VendorTotalInvoice, ">", "")
            VendorTotalInvoice = Replace(VendorTotalInvoice, "T", "")
            VendorTotalInvoice = Replace(VendorTotalInvoice, "O", "")
            VendorTotalInvoice = Replace(VendorTotalInvoice, "A", "")
            VendorTotalInvoice = Replace(VendorTotalInvoice, "L", "")
            'VendorTotalInvoice = Replace(VendorTotalInvoice, " ", "")
            'MsgBox "WESCO VendorTotalInvoice=:" & VendorTotalInvoice & ":"
       End If
       
       'Discount
       'Ref "WITHIN 10 DAYS - NET 30 DAYS > 4.42 TOTAL > 221.03"
       If UCase(Line) Like "*WITHIN 10 DAYS*" Then
            Discount = UCase(Line)
            Discount = Replace(Discount, "WITHIN 10 DAYS - NET 30 DAYS", "")
            For Repeat = 1 To 10
                If Right(Discount, 1) = ">" Then Exit For
                Discount = Left(Discount, Len(Discount) - 1)
            Next Repeat
            Discount = Replace(Discount, ">", "")
            Discount = Replace(Discount, "TOTAL", "")
            Discount = Replace(Discount, " ", "")
            
            'MsgBox "WESCO Discount->" & Discount
       End If
    Next xoffset
     
    'MsgBox "done scraping macro data, now for line items...."
                 
     If VendorTotalInvoice = "" And PDFtype = "Invoice" Or VendorTotalInvoice = 0 And PDFtype = "Invoice" Then MsgBox "Failed to scrape invoice total"
     'MsgBox PDFtype
     
              
'Detect how many line items there are
    For Detect = 0 To 1000
        If ws.Range("BA2").Offset(Detect, 0) = "ID" And _
        ws.Range("BA2").Offset(Detect + 1, 0) = "NUMBER" Then
            'MsgBox "Made it here"
            For LineItems = 1 To 100
                If ws.Range("BA2").Offset(Detect + 1 + LineItems, 0) Like "[0-9][0-9][0-9][0-9][0-9][0-9]*" Then
                'Do nothing
                Else
                    Exit For
                End If
            Next LineItems
            Exit For
        End If
    Next Detect
    
    LineItems = LineItems - 1
    If LineItems <= 0 Then

    If emailmessage <> "busy" Then MsgBox "Failed to detect qty of line items on Wesco Sheet" & Chr(13) & "emailmessage:" & emailmessage
        Exit Sub
    End If
    'If LineItems > 0 Then MsgBox "Found " & LineItems & " items on Wesco Sheet"
            
'Get line item product info
For itemline = 1 To LineItems

'ITEM DESCRIPTION Range(P2)
    itemdescription = ""
    For Detect = 0 To 200 'Detect how many line items there are
        If ws.Range("BA2").Offset(Detect, 0) = "CATALOG NUMBER" And _
        ws.Range("BA2").Offset(Detect + 1, 0) = "AND DESCRIPTION" Then
        'How to skip over short lines
            adjLine = 0
            For Repeat = 0 To 100
                itemdescription = ws.Range("BA2").Offset(Detect + 2 + Repeat)
                'MsgBox "unfiltered line is->" & itemdescription
                'WESCO invoice line EXCLUSIONS
                If Len(itemdescription) > 10 And Not UCase(itemdescription) Like "*TUAN*" _
                And Not UCase(itemdescription) Like "*PH 425*" And _
                Not UCase(itemdescription) Like "*DEL*[0-9]AM*" And _
                Not UCase(itemdescription) Like "*JORDAN*" And _
                Not UCase(itemdescription) Like "*MONDAY*" And Not UCase(itemdescription) Like "*TWIC*CARD*" And _
                Not UCase(itemdescription) Like "*TUESDAY*" And Not UCase(itemdescription) Like "*SOLD*IN*CARTON*" And _
                Not UCase(itemdescription) Like "*WEDNESD*" And Not UCase(itemdescription) Like "WITH BASE" And _
                Not UCase(itemdescription) Like "*THURSDAY*" And Not UCase(itemdescription) Like "*THIS*ORDER*" And _
                Not UCase(itemdescription) Like "*FRIDAY*" And Not UCase(itemdescription) Like "*DRIVE*WA*" And _
                Not UCase(itemdescription) Like "*DAVE*" And Not UCase(itemdescription) Like "*GRAVEL*" And _
                Not UCase(itemdescription) Like "*NEEDED*" And Not UCase(itemdescription) Like "*STORAGE*" And _
                Not UCase(itemdescription) Like "*POSSIBLE*" And Not UCase(itemdescription) Like "*BETWEEN*" And _
                Not UCase(itemdescription) Like "*DELIVER*" And Not UCase(itemdescription) Like "*JOHN*" And _
                Not UCase(itemdescription) Like "*MEET*AT*" And Not UCase(itemdescription) Like "*XCN=*" And _
                Not UCase(itemdescription) Like "*TUAN*" And Not UCase(itemdescription) Like "*XDC=*" And _
                Not UCase(itemdescription) Like "*DALE*" And Not UCase(itemdescription) Like "*MIKE*" And _
                Not UCase(itemdescription) Like "*DEDUCT*" And Not UCase(itemdescription) Like "*ACCEPTANCE*" And _
                Not UCase(itemdescription) Like "*WITHIN*" And Not UCase(itemdescription) Like "*TERMS*" And _
                Not UCase(itemdescription) Like "*ACCOUNT*" And Not UCase(itemdescription) Like "*MONSTER*" And _
                Not UCase(itemdescription) Like "*TRK:*" And Not UCase(itemdescription) Like "*RENTON*" And _
                Not UCase(itemdescription) Like "*PLEASE*" And Not UCase(itemdescription) Like "*WWW.*" And _
                Not UCase(itemdescription) Like "*JOSH*" And Not UCase(itemdescription) Like "*HANGER*" And _
                Not UCase(itemdescription) Like "*ON*SITE*" And Not UCase(itemdescription) Like "*BOEING*" And _
                Not UCase(itemdescription) Like "*CALL *" _
                And UCase(itemdescription) <> "************************" Then adjLine = adjLine + 1
                If adjLine = itemline Then Exit For
            
            Next Repeat
            itemdescription = ws.Range("BA2").Offset(Detect + 2 + Repeat)
            Exit For
        End If
        If itemdescription <> "" Then Exit For
    Next Detect
    
'Error Check
    If itemdescription = "" Then
        For Repeat = 1 To 10
            MsgBox "Error, failed to pick up item description. Possibly check if excluding comma from acceptable descriptions has made it so less descriptions were picked up than part numbers"
        Next Repeat
    End If
    'MsgBox "Item " & ItemLine & " Description->" & itemdescription
    



'QUANTITY (R2)
    Quantity = ""
    For Detect = 0 To 200 'Detect how many line items there are
        If ws.Range("BA2").Offset(Detect, 0) = "QUANTITY" And _
        ws.Range("BA2").Offset(Detect + 1, 0) = "SHIPPED" Then
            Quantity = ws.Range("BA2").Offset(Detect + 1 + itemline)
            Exit For
        End If
    Next Detect

    'error check
    If Quantity = 0 Then
        'MsgBox "On WESCO invoice, failed to scrape quantity-> will cause imperfections when entering invoice. Try running again?"
    End If
    
'UNIT Range (Q2)
    Unit = ""
    For Detect = 0 To 400
        If ws.Range("BA2").Offset(Detect, 0) = "UOM" Then
            Found = 1
            Unit = ws.Range("BA2").Offset(Detect + itemline)
            Exit For
        End If
    Next Detect
    'Error check
    If Unit = "" Then
        If Unit = "" Then MsgBox "Failed to pick up Unit of measurement, so price per unit may not be accurate / try re-running?"
    End If
    'MsgBox "Unit->" & Unit & Chr(13) & "Found->" & Found
    'Found = 0


'UNIT PRICE (S2)
    unitprice = ""
    For Detect = 0 To 400
        If ws.Range("BA2").Offset(Detect, 0) = "UNIT" And _
        ws.Range("BA2").Offset(Detect + 1, 0) = "PRICE" Then
            unitprice = ws.Range("BA2").Offset(Detect + 1 + itemline)
            Exit For
        End If
    Next Detect
    If Unit = "C" Then
        unitprice = unitprice / 100
        Unit = "Ea"
    End If
    If Unit = "M" Then
        unitprice = unitprice / 1000
        Unit = "Ea"
    End If
    
    'MsgBox "UnitPrice->" & UnitPrice
    
'LINE TOTAL (T2)
    linetotal = ""
    For Detect = 0 To 400
        If ws.Range("BA2").Offset(Detect, 0) = "EXTENSION" Then
            linetotal = ws.Range("BA2").Offset(Detect + itemline)
            Exit For
        End If
    Next Detect
    'MsgBox "LineTotal->" & LineTotal
    
'POPULATE DATA TO THE SHEET
    ws.Range("P1").Offset(itemline, 0) = itemdescription
    ws.Range("Q1").Offset(itemline, 0) = Unit
    ws.Range("R1").Offset(itemline, 0) = Quantity
    ws.Range("S1").Offset(itemline, 0) = unitprice
    ws.Range("T1").Offset(itemline, 0) = linetotal
    ws.Range("A1").Offset(itemline, 0) = DecoPO
    ws.Range("B1").Offset(itemline, 0) = OrderDate
    ws.Range("C1").Offset(itemline, 0) = "430"
    ws.Range("AH1").Offset(itemline, 0) = Tax
    ws.Range("H1").Offset(itemline, 0) = VendorInvoiceNo
    ws.Range("N1").Offset(itemline, 0) = VendorTotalInvoice
    ws.Range("J1").Offset(itemline, 0) = InvoiceDate
Next itemline
        
ThisWorkbook.Sheets("Temp").Range("C2") = "430"
Call SelfHealTempPage
'Check if PO conforms before bothering to Enter
TargetPO = ThisWorkbook.Sheets("Temp").Range("A2")
Call CheckPONumber(TargetPO, Found)
            
' Found variable Index:
' Found = 0 TargetPO does not conform, refuse to process
' Found = 1 TargetPO is OK to Process
' Found = 2 TargetPO is Subcontract, Send PDF to Fax File
' Found = 3 TargetPO is SHOP, Send PDF to Fax File

If VendorInvoiceNo = "" Then
    MsgBox "Failed to Pickup WESCO Invoice No! Will not be able to input in Sage"
Else
    'Call WescoWebscrape(VendorInvoiceNo, Fail, fpath, TargetPO, docno, PDFtype, shipping, path, fname, Discount)
End If
docno:      'Insert Webscraper here: but, haven't done this yet because scnas are as high quality as info on web page and
            'also webpage addresses not conducive to scraping
            
            
            If Found = 2 Then
                'adjust for the fact that WESCO invoices sometimes start with a ZERO!!
                AddZero = ""
                If UCase(ThisWorkbook.Sheets("Temp").Range("C2")) Like "*WESCO*" Then
                    InvoiceNO = ThisWorkbook.Sheets("Temp").Range("H2").Offset(xoffset, 0)
                    'MsgBox Len(InvoiceNo)
                    If Len(InvoiceNO) < 6 Then
                        Application.SendKeys ("0")
                        AddZero = "0"
                    End If
                End If
                
                TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
                TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
                If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
                If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
                pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
                & "INVOICE " & AddZdero & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
                & TotalInvoiceAmount & ").pdf"
                pdfoption1 = "Subcontract " & Replace(pdfoption1, "$", "")
                
                If Dir("\\server2\Faxes\" & pdfoption1) = "" Then
                Name fpath As "\\server2\Faxes\" & pdfoption1
                updatelog = pdfoption1
                Call logupdate(updatelog)
                Application.Wait (Now + TimeValue("00:00:06"))
                End If
                
                SetCursorPos 1083, 11 '--------------------------'Sage Minimize
                Call Mouse_left_button_press
                Call Mouse_left_button_Letgo
                Application.Wait (Now + TimeValue("00:00:01"))
                
                'MsgBox "Sent TargetPO that matches Subcontract number to Fax Folder->" & TargetPO
                Application.Wait (Now + TimeValue("00:00:03"))
                Exit Sub
            
            End If
            
            
            
            If Possibleerror < 1 Then
                Call ClickOnSage
                xoffset = 0
                emailmessage = "Wesco Invoice"
                Call SageEnterINVOICEfromTEMP(xoffset, emailmessage, fpath)
                If emailmessage = "Temp Sheet Total Error" Then Exit Sub
                'rename and move file
                
                'adjust for the fact that WESCO invoices sometimes start with a ZERO!!
                AddZero = ""
                If UCase(ThisWorkbook.Sheets("Temp").Range("C2")) Like "*WESCO*" Then
                    InvoiceNO = ThisWorkbook.Sheets("Temp").Range("H2").Offset(xoffset, 0)
                    'MsgBox Len(InvoiceNo)
                    If Len(InvoiceNO) < 6 Then
                        Application.SendKeys ("0")
                        AddZero = "0"
                    End If
                End If
                
                TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
                TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
                If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
                If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
                pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
                & "INVOICE " & AddZdero & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
                & TotalInvoiceAmount & ").pdf"
                pdfoption1 = Replace(pdfoption1, "$", "")
                
                ' Move or Kill PDF
                If Dir(fpath) <> "" Then
                    If Dir("\\server2\Faxes\WESCO - 430\" & pdfoption1) = "" Then
                        Name fpath As "\\server2\Faxes\WESCO - 430\" & pdfoption1
                        updatelog = pdfoption1
                        Call logupdate(updatelog)
                    Else
                        Kill fpath
                    End If
                End If
                
                SetCursorPos 1083, 11 '--------------------------'Sage Minimize
                Call Mouse_left_button_press
                Call Mouse_left_button_Letgo
                Application.Wait (Now + TimeValue("00:00:01"))
            Else
                    MsgBox "Did not enter " & TargetPO & Chr(13) & "PossibleErrors =" & Possibleerror _
                    & Chr(13) & "Found =" & Found & Chr(13) & "Moving to Fax File as;" & Chr(13) _
                    & TargetPO & " Wesco"
                    
                    If Dir("\\server2\Faxes\" & TargetPO & " Wesco.pdf") <> "" Then TargetPO = TargetPO & ".1"

                    If Dir(fpath) <> "" And Dir("\\server2\Faxes\" & TargetPO & " Wesco.pdf") = "" Then
                        Name fpath As "\\server2\Faxes\" & TargetPO & " Wesco.pdf"
                        Application.Wait (Now + TimeValue("00:00:06"))
                    End If
                
                    If Dir(fpath) <> "" And UCase(path) Like "*ATTACH*" Or emailmessage Like "*already been entered*" And Dir(fpath) <> "" Then Kill (fpath)
            End If
GoTo start:

End Sub

Sub WescoWebscrape(VendorInvoiceNo, Fail, fpath, TargetPO, docno, PDFtype, shipping, path, fname, Discount, emailmessage)
Dim ws As Worksheet
Dim vendoritemno As String
Set ws = ThisWorkbook.Sheets("Temp")

Exit Sub

' Prepare this workbooks temp sheet
'Call FormatTempSheet

'Set Target URL, Reference format: https://www.platt.com/Order.aspx?itemid=2D22679&CustNum=36850
TargetURL = "https://buy.wesco.com/accountoverview/orders"

' Open Chrome
Call OpenChrome(TargetURL)
try = 0

copy_orders_page_data:
    
' COPY->PASTE webpage Data
Application.CutCopyMode = False
Application.SendKeys ("^a"), True
Sleep (250)
For Repeat = 1 To 3
    Application.SendKeys ("^c"), True
    Sleep 500
Next Repeat
ws.Paste Destination:=ws.Range("BA2")
Sleep 500
ws.DrawingObjects.Delete
ws.Range("BA2:CA300").UnMerge
Sleep 500

'Read Paste Data and check if at login page
Found = 0
For x = 0 To 20
    For y = 0 To 10
        If UCase(ws.Range("BA2").Offset(x, y)) Like "*JEREMY*" Then Found = 1
    Next y
    If Found = 1 Then Exit For
Next x

If Found = 0 Then
    ' MsgBox "We're at the login page"
    ' tab 16 times to hit "Login" with pre-populated data
    MsgBox "Freeze Need to login!"
    MsgBox "Freeze Need to login!"
    MsgBox "Freeze Need to login!"
    MsgBox "Freeze Need to login!"
    MsgBox "Freeze Need to login!"
    
End If
    
    
'Check if at ORDER PAGE
Found = 0
For x = 0 To 100
    For y = 0 To 26
        If UCase(ws.Range("BA2").Offset(x, y)) Like "*" & VendorInvoiceNo & "*" Then Found = 1
    Next y
    If Found = 1 Then Exit For
Next x
    
If Found = 0 Then
        try = try + 1
        If try < 2 Then
            Application.SendKeys ("%d")
            Application.Wait (Now + TimeValue("00:00:01"))
            Application.SendKeys (TargetURL)
            Application.Wait (Now + TimeValue("00:00:01"))
            Application.SendKeys ("~")
            For Wait = 1 To 10
                Application.Wait (Now + TimeValue("00:00:01"))
            Next Wait
            GoTo copy_orders_page_data:
        Else
            MsgBox "can't find " & VendorInvoiceNo & " on this page!!"
            MsgBox "can't find " & VendorInvoiceNo & " on this page!!"
            MsgBox "can't find " & VendorInvoiceNo & " on this page!!"
            MsgBox "can't find " & VendorInvoiceNo & " on this page!!"
            
            'close chrome
            Application.SendKeys ("^w")
            Sleep 500
            Application.SendKeys ("^w")
            Sleep 500
            Exit Sub
        End If
End If


' use find function to fiund and click into specific order
Application.SendKeys ("^f")
Application.Wait (Now + TimeValue("00:00:01"))
Application.SendKeys (VednorInvoiceNo)
Application.Wait (Now + TimeValue("00:00:01"))
Application.SendKeys ("^~")
For Wait = 1 To 5
    Application.Wait (Now + TimeValue("00:00:01"))
Next Wait





' Copy order page data
' COPY->PASTE webpage Data
ws.Range("BA2:CA300").UnMerge
ws.Range("BA2:CA300").Clear
Application.CutCopyMode = False
Application.SendKeys ("^a"), True
Sleep (250)
For Repeat = 1 To 3
    Application.SendKeys ("^c"), True
    Sleep 500
Next Repeat
ws.Paste Destination:=ws.Range("BA2")
Sleep 500
ws.DrawingObjects.Delete
ws.Range("BA2:CA300").UnMerge
Sleep 500

' Error check if we actually copied from an orders Page

MsgBox "Freeze, just finished copying WESCO order from internet to temp sheet!!!!"

MsgBox "Freeze, just finished copying order to temp sheet"

MsgBox "Freeze, just finished copying order to temp sheet"

MsgBox "Freeze, just finished copying order to temp sheet"


' Get product numbers

Dim myRange As Range
Dim cell As Range
Dim myString As String
Dim partialString As String
Set myRange = ThisWorkbook.Sheets("Temp").Range("BA2:BC200") 'change to your desired range
partialString = "Item #" 'change to your desired partial string
For Each cell In myRange
    'If cell.Text <> "" Then MsgBox cell.Text
    If InStr(1, cell.Text, partialString, vbTextCompare) > 0 Then 'check if cell contains partial string
        Product_Link = cell.Text 'copy cell contents to variable
        Product_Link = Replace(Product_Link, "Item #", "")
        Product_Link = Replace(Product_Link, " ", "")
        'MsgBox Product_Link
        'Do something with myString
        Vendor = "Platt"
        Call Save_Product_Links(Product_Link, Vendor, workbook_open_status)
    End If
Next cell
Workbooks("Product_links.xlsx").Save
Workbooks("Product_links.xlsx").Close
    
    
    
    
' Reset variables
tempsheetoffset = 0
InvoiceDate = ""
Possibleerror = 0
TotalInvoice = ""
yoffset = 0

' Scrape copied wepage data
For xoffset = 1 To 500 'Gather Macro Information
       'MsgBox "Freeze"
       Line = ws.Range("BA2").Offset(xoffset, 0)
       Line = Replace(Line, vbLf, "")

    If Line = "CREDIT MEMO" And PDFtpye = "" Then
        PDFtype = "Invoice"
        MsgBox "PDFtype_>" & PDFtype
    End If

       'INVOICE Date
        If Line Like "*Placed on*" And PDFtype = "Invoice" And InvoiceDate = "" Then  '
           InvoiceDate = Line
           'If Not InvoiceDate Like "*/*/*" Then InvoiceDate = ""
           InvoiceDate = Replace(InvoiceDate, "Placed on", "")
            'InvoiceDate = Replace(InvoiceDate, " ", "")
           InvoiceDate = Format(InvoiceDate, "MM/DD/YYYY")
           InvoiceDate = Format(Date, "MM/DD/YYYY")
           'MsgBox "InvoiceDate=:" & InvoiceDate & ":"
       End If
       
       'ORDER DATE
       If Line Like "*Placed on*" And PDFtype = "Order" Then '
           OrderDate = Line
           'If Not OrderDate Like "*/*/*" Then OrderDate = ""
           OrderDate = Replace(OrderDate, "Placed on", "")
           OrderDate = Format(OrderDate, "MM/DD/YYYY")
           'MsgBox "OrderDate=:" & OrderDate & ":"
       End If

       'DELIVERY METHOD
       'ws.Range("AE2") = ""
       
        'TOTAL $
        If UCase(Line) Like "TOTAL" Then 'ws.Range("BA2").Offset(xoffset, 1) Like "*[$]*[0-9].[0-9]*" Then '
           'MsgBox "Freeze"
            If ws.Range("BA2").Offset(xoffset, 1) <> "" And _
                ws.Range("BA2").Offset(xoffset, 1) Like "*[0-9]*" Then _
                VendorTotalInvoice = CDbl(ws.Range("BA2").Offset(xoffset, yoffset + 1))
           'MsgBox "Platt VendorTotalInvoice=:" & VendorTotalInvoice & ":"
       End If
        
       'DECO PO
       If Line = "PO:" Then '
        DecoPO = ws.Range("BA2").Offset(xoffset + 1, yoffset)
        DecoPO = Replace(DecoPO, vbLf, "")
           DecoPO = Replace(DecoPO, "PO", "")
           DecoPO = Replace(DecoPO, "#", "")
           DecoPO = Replace(DecoPO, " ", "")
           'MsgBox "decoPO=:" & DecoPO & ":"
       End If
       
        'Vendor Invoice #
        vendorInvoice = docno
        
       'TAX
       If UCase(Line) Like "*SALES*" And Not UCase(Line) Like "*PLATT*" Then '
            For Repeat = 1 To 10
                If ws.Range("BA2").Offset(xoffset, yoffset + Repeat) <> "" Then
                    Tax = ws.Range("BA2").Offset(xoffset, yoffset + Repeat)
                End If
            Next Repeat
       End If
       
        If UCase(Line) Like "*HANDLING*" And Not UCase(Line) Like "*PLATT*" Then '
            For Repeat = 1 To 10
                If ws.Range("BA2").Offset(xoffset, yoffset + Repeat) <> "" Then
                    handling = ws.Range("BA2").Offset(xoffset, yoffset + Repeat)
                End If
            Next Repeat
       End If
       
Next xoffset

'MsgBox "done, Target PO = " & DecoPO
'MsgBox "Freeze"
'MsgBox "Freeze"
'MsgBox "Freeze"
'MsgBox "Freeze"
            
 ' Message user if failed to fetch Invoice Total
 If VendorTotalInvoice = "" And PDFtype = "Invoice" Or VendorTotalInvoice = 0 And PDFtype = "Invoice" Then MsgBox "Failed to scrape invoice total"
 
            
 ' Get Line Items
 For xoffset = 0 To 500 'Now get Line Items
     yoffset = 0
     'For Yoffset = 0 To 20
         Line = ws.Range("BA2").Offset(xoffset, 0)
         Line = Replace(Line, vbLf, "")
         
          'ITEM DESCRIPTION
         If Line Like "*Item*#*" Then
 
             itemDesc = ws.Range("BA2").Offset(xoffset - 1, 0)
             itemDesc = Replace(itemDesc, "Item # ", "")
             itemDesc = Replace(itemDesc, vbLf, "")
             itemDesc = Replace(itemDesc, "+", "")
             If Left(itemDesc, 1) = " " Then itemDesc = Right(itemDesc, Len(itemDesc - 1))
             If Left(itemDesc, 1) = "," Then itemDesc = Right(itemDesc, Len(itemDesc - 1))
             If itemDesc Like "*" & Chr(173) & "*" Then itemDesc = Replace(itemDesc, Chr(173), "")
             If Len(itemDesc) > 60 Then ItemDec = Left(itemDesc, 60)
             If itemDesc Like "1000*" Then MsgBox "WESCO WEBSCRAPER ERROR, description seems to be quantity" & Chr(13) & "Description->" & itemDesc
             If itemDesc = "" Then Possibleerror = Possibleerror + 1
             'MsgBox "ItemDesc=:" & ItemDesc

            ' Vendor Item No. (vendoritemno)
            MsgBox "Freeze to get vendor item #"
            MsgBox "Freeze to get vendor item #"
            MsgBox "Freeze to get vendor item #"
            MsgBox "Freeze to get vendor item #"
            MsgBox "Freeze to get vendor item #"
            MsgBox "Freeze to get vendor item #"
            
                
             'UNIT Type
             For y = 0 To 6
                 If ws.Range("BA2").Offset(xoffset + y, yoffset) Like "*Price:*" Then Exit For
             Next y
             If y > 6 Then MsgBox "Found line item but couldn't find the unit!" & Chr(13) & "excel row ->" & xoffset + 2 & _
                  "might be one of those cases where the units are on the line below the item. fixed by resetting and going again"
             
             Unit = Right(ws.Range("BA2").Offset(xoffset + y, yoffset), 2)
             'MsgBox unit
             Unit = Replace(Unit, "Price:", "")
             Unit = Replace(Unit, vbLf, "")
             Unit = Replace(Unit, vbLf, "")
             Unit = Replace(Unit, "(100 EA)", "")
             Unit = Replace(Unit, "(100 FT)", "")
             Unit = Replace(Unit, " ", "")
             
             If Unit Like "*E*E*" Then Possibleerror = Possibleerror + 1 'idicates that rows are combined in excel conversion
             For re = 0 To 15
                 If Left(Unit, 1) Like "[0-9]" Or Left(Unit, 1) Like "/" Or Left(Unit, 1) Like "." _
                 Or Left(Unit, 1) Like "$" Then Unit = Right(Unit, Len(Unit) - 1)
             Next re
             If Unit = "FT" Then Unit = "EA"
             If Unit Like "*C*" Then Unit = "C"
             If Unit Like "*M*" Then Unit = "M"
            
             
             'MsgBox "Unit=:" & unit & ":"
              
             'QUANTITY ORDERED OR SHIPPED
             For y = 0 To 6
                 If ws.Range("BA2").Offset(xoffset + y, 0) Like "*Order Qty:*" And _
                     PDFtype = "Order" Then Exit For
                 If ws.Range("BA2").Offset(xoffset + y, 0) Like "*Ship Qty:*" And _
                     PDFtype = "Invoice" Then Exit For
             Next y
             
             Quantity = ws.Range("BA2").Offset(xoffset + y, 0)
             Quantity = Replace(Quantity, "Order Qty:", "")
             Quantity = Replace(Quantity, "Ship Qty:", "")
             Quantity = Replace(Quantity, vbLf, "")
             Quantity = Replace(Quantity, " ", "")
             Negative = 0
             For re = 0 To 15
                 If UCase(Right(Quantity, 1)) = "-" Then Negative = 1
                 If UCase(Right(Quantity, 1)) Like "[A-Z]" Then Quantity = Left(Quantity, Len(Quantity) - 1)
             Next re
             If Negative = 1 Then Quantity = Quantity * -1
             Quantity = Replace(Quantity, " ", "")
             'MsgBox "Quantity=:" & Quantity & ":"
             

             'UNIT PRICE
             For y = 0 To 6
                 If ws.Range("BA2").Offset(xoffset + y, 0) Like "*Price:*" Then Exit For
             Next y
             unitprice = ws.Range("BA2").Offset(xoffset + y, 0)
             'MsgBox unitprice
             unitprice = Replace(unitprice, vbLf, "")
             unitprice = Replace(unitprice, "Price:", "")
             unitprice = Replace(unitprice, "(100 FT)", "")
             unitprice = Replace(unitprice, " ", "")
             unitprice = Replace(unitprice, "$", "")
             For re = 0 To Len(unitprice)
                 If Not Right(unitprice, 1) Like "[0-9]" Then unitprice = Left(unitprice, Len(unitprice) - 1)
             Next re
             
             ' Unit Price
             If Unit = "C" Then
                 If unitprice = "" Then unitprice = 0
                 unitprice = unitprice / 100
                 Unit = "EA"
             End If
             If Unit = "M" Then
                 unitprice = unitprice / 1000
                 Unit = "Ea"
             End If
             If unitprice = "" Then unitprice = "0"
             If unitprice Like "*[A-Z]*" Then unitprice = "0"
             'MsgBox "UnitPrice=:" & unitprice & ":"
             
             'LINE TOTAL
             If UCase(Quantity) Like "*[A-Z]*" Then Quantity = 1
             ws.Range("S2").Offset(tempsheetoffset, 0) = unitprice
             If SHIP = 0 Then lineprice = unitprice * Quantity
             'MsgBox "LinePrice=:" & LinePrice & ":"
             
             'Write Data to thisworkbok temp sheet
             ws.Range("P2").Offset(tempsheetoffset, 0) = itemDesc
             ws.Range("Q2").Offset(tempsheetoffset, 0) = Unit
             ws.Range("R2").Offset(tempsheetoffset, 0) = Quantity
             ws.Range("T2").Offset(tempsheetoffset, 0) = lineprice
             ws.Range("A2").Offset(tempsheetoffset, 0) = DecoPO
             ws.Range("B2").Offset(tempsheetoffset, 0) = OrderDate
             ws.Range("C2").Offset(tempsheetoffset, 0) = "234"
             ws.Range("AH2").Offset(tempsheetoffset, 0) = Tax
             ws.Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
             ws.Range("N2").Offset(tempsheetoffset, 0) = VendorTotalInvoice
             ws.Range("O2").Offset(tempsheetoffset, 0).NumberFormat = "@"
             ws.Range("O2").Offset(tempsheetoffset, 0) = vendoritemno
             ws.Range("J2").Offset(tempsheetoffset, 0) = InvoiceDate
             tempsheetoffset = tempsheetoffset + 1
             'Cross Check
             If itemDesc = Quantity Then _
                 MsgBox "Wesco Webscrape Transcription Error, Description " & itemDesc & " is Equal to Quantity " & Quantity & " on line " & xoffset _
                 & Chr(13) & "temp sheet line " & tempsheetoffset
         End If
     'Next yOffset
 Next xoffset
 
' If a discount was found, insert it now
ThisWorkbook.Sheets("Temp").Range("AF2") = Discount
 
'If there was shipping costs, insert it now at the end of the line items
If shipping <> "" Then
        ws.Range("S2").Offset(tempsheetoffset, 0) = shipping
        ws.Range("A2").Offset(tempsheetoffset, 0) = DecoPO
        ws.Range("B2").Offset(tempsheetoffset, 0) = OrderDate
        ws.Range("C2").Offset(tempsheetoffset, 0) = "234"
        ws.Range("AH2").Offset(tempsheetoffset, 0) = Tax
        ws.Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
        ws.Range("N2").Offset(tempsheetoffset, 0) = VendorTotalInvoice
        ws.Range("P2").Offset(tempsheetoffset, 0) = "Shipping and Handling"
        ws.Range("R2").Offset(tempsheetoffset, 0) = 1
        ws.Range("J2").Offset(tempsheetoffset, 0) = InvoiceDate
End If
            
                
'MsgBox "Done scraping Web data"
Call SelfHealTempPage






End Sub
