Attribute VB_Name = "M42Stoneway"

Sub IfStonewayOrdAck(fname, fpath, XLSpath, XLSname, path, emailmessage) 'STONEWAY STONEWAY STONEWAY
Dim InvoiceDate As String
Dim i As Long
Dim URL As String
Dim try As Integer
'path = source where pdf's reside. Attachements folder or backup folder for submittal scrape
'MsgBox "Arrived at stoneway " & Chr(13) & fname

Call FormatTempSheet

Found = 0



    Call FormatTempSheet
    Call Convert_PDF_to_Excel(fname, fpath, XLSpath, XLSname, emailmessage)
'Verify it's a stoneway Invoice
    Found = 0
    For x = 0 To 100
        For y = 0 To 20
            
            If UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(x, y)) Like "*STONEWAY*" Then Found = 1
            If UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(x, y)) Like "*NORTH*COAST*" Then Found = 2
        Next y
    If Found <> 0 Then Exit For
    Next x

'MsgBox "In Stoneway loop"
    If Found = 0 Then
        'kill xls
        Workbooks(XLSname).Close SaveChanges:=False
        Exit Sub
    End If

    If Found = 2 Then
        'MsgBox "Found North Coast Document"
        Call IfNorthCoast(path, XLSname, XLSpath, fpath, fname, emailmessage)
        Exit Sub
    End If


'Scrape Macro Data
    For xoffset = 0 To 200 'Gather Macro Information
        For y = 0 To 6
            Line = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, y)
            
            ' Document number
            If Line Like "S[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].[0-9][0-9][0-9]*" Then
                docno = Line
                docno = Replace(docno, " ", "")
                'MsgBox "Stoneway DocumentNumber->" & DocNo
            End If
            
            ' Invoice Date / order date
            If UCase(Line) Like "* DATE*" Then '
                OrderDate = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 1, y)
                'MsgBox Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, y) & Chr(13) & "Date=:" & OrderDate & ":"
                If UCase(Line) Like "*INVOIC*DAT*" And InvoiceDate = "" Then
                    InvoiceDate = OrderDate
                    'MsgBox "Located Invoice Date->" & InvoiceDate
                End If
            End If
            
            ' Invoice Due Date ("AD2") Ref <Invoice is due by 03/31/23 net of any cash discount.>
            If UCase(Line) Like "*INVOICE*IS*DUE*BY*" Then
                InvoiceDueDate = Line
                If Not InvoiceDueDate Like "*/*/*" Then InvoiceDueDate = ""
                For Repeat = 1 To Len(InvoiceDueDate) - 9
                    If Mid(InvoiceDueDate, Repeat, 8) Like "[0-9][0-9]/[0-9][0-9]/[0-9][0-9]" Then
                        InvoiceDueDate = Mid(InvoiceDueDate, Repeat, 8)
                        'MsgBox InvoiceDueDate
                        Exit For
                    End If
                Next Repeat
                InvoiceDueDate = Replace(InvoiceDueDate, " ", "")
                ThisWorkbook.Sheets("Temp").Range("AD2") = InvoiceDueDate
            End If
         
            ' Discount date
           If Line Like "*you may deduct*" Then '
               InvoiceDiscountDate = Line
               If Not InvoiceDiscountDate Like "*/*/*" Then InvoiceDiscountDate = ""
               'If paid by 01/10/23 you may deduct $3.84
               InvoiceDiscountDate = Replace(InvoiceDiscountDate, "If paid by", "")
               InvoiceDiscountDate = Replace(InvoiceDiscountDate, "you may deduct", "")
               Discount = InvoiceDiscountDate
               If UCase(Discount) Like "*[A-Z]*" Then Discount = 0
               InvoiceDiscountDate = Left(InvoiceDiscountDate, 9)
               If Len(Discount) > 5 Then Discount = Mid(Discount, 10, Len(Discount) - 9)
               Discount = Replace(Discount, "$", " ")
               Discount = Replace(Discount, " ", "")
               ThisWorkbook.Sheets("Temp").Range("AE2") = InvoiceDiscountDate
               ThisWorkbook.Sheets("Temp").Range("AF2") = Discount
               'MsgBox "Stoneway InvoiceDiscountDate=:" & InvoiceDiscountDate & ":" & Chr(13) & "Discount->" & Discount
           End If
            
            
            'Delivery Method
            'If Line Like "*SHIP VIA*" Then '                                 <<<Delivery Method
            '    ThisWorkbook.Sheets("Temp").Range("AE2") = "Delivery"
            'End If
            'If Line Like "*Will Call*" Then '                                 <<<Delivery Method
            '    ThisWorkbook.Sheets("Temp").Range("AE2") = "Will Call"
            'End If
        
            'DECO-PO
            If Line Like "*CUSTOMER*PO*NUMBER*" Then     '
                line2 = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 1, y)
                DecoPO = UCase(line2)
                DecoPO = Replace(DecoPO, " ", "")
                'MsgBox "decoPO=:" & DecoPO & ":"
            End If
            
            'VENDOR INVOICE Number #
            If Line Like "INVOICE NUMBER*" Then '
                vendorInvoice = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 1, y)
                vendorInvoice = Replace(vendorInvoice, " ", "")
                'MsgBox "VendorInvoice=:" & VendorInvoice & ":"
            End If
            
            'TOTAL INVOICE
            If Line Like "*Amount Due*" Then '
                line2 = ""
                For yoffset = 1 To 20
                    line2 = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + y)
                    If line2 <> "" Then
                        If Not line2 Like "*.*" Then line2 = line2 & ".00"
                        If line2 Like "*.[0-9]" Then line2 = line2 & "0"
                        TotalInvoice = line2
                        Exit For
                    End If
                Next yoffset
                TotalInvoice = Replace(TotalInvoice, vbLf, "")
                If TotalInvoice Like "*.*.*" Then
                    For re = 0 To 20
                        If Left(TotalInvoice, 1) Like "." Then Exit For
                        TotalInvoice = Right(TotalInvoice, Len(TotalInvoice) - 1)
                    Next re
                    TotalInvoice = Right(TotalInvoice, Len(TotalInvoice) - 1)
                    For re = 0 To 20
                        If Not Left(TotalInvoice, 1) Like "0" Then Exit For
                        TotalInvoice = Right(TotalInvoice, Len(TotalInvoice) - 1)
                    Next re
                End If
                If Left(TotalInvoice, 1) = Chr(13) Then TotalInvoice = Right(TotalInvoice, Len(TotalInvoice) - 1)
                'MsgBox "TotalInvoice=:" & TotalInvoice & ":"
            End If
            
            'TAX
            If Line Like "*Tax*" Then '
                line2 = ""
                For yoffset = 1 To 20
                    line2 = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + y)
                    If line2 <> "" Then
                        If Not line2 Like "*.*" Then line2 = line2 & ".00"
                        If line2 Like "*.[0-9]" Then line2 = line2 & "0"
                        Tax = line2
                        Exit For
                    End If
                Next yoffset
            End If
        Next y
    Next xoffset
    

' Scrape Line Items Data
    For xoffset = 0 To 200 'Gather Macro Information 'Now input Line Items
        Line = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0)
            'DESCRIPTION
            If Line Like "*[0-9]ea*" Then '
                For y = 0 To 20   '
                    If Not Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, y) Like "*[0-9]ea*" _
                    And Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, y) <> "" Then
                        itemDesc = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, y)
                        Exit For
                    End If
                Next y
                itemDesc = Replace(itemDesc, vbLf, "")
                If Len(itemDesc) > 40 Then itemDesc = Left(itemDesc, 40)
                'MsgBox "ItemDesc=:" & Itemdesc & ":"
                If itemDesc = "" Then Possibleerror = Possibleerror + 1
                ThisWorkbook.Sheets("Temp").Range("P2").Offset(tempsheetoffset, 0) = itemDesc
                
                
                
                'UNIT
                For y = 0 To 20   '
                    If Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, y) Like "*[0-9]/*" Then
                        Unit = UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, y))
                        For re = 0 To 20
                            If Left(Unit, 1) <> "/" Then Unit = Right(Unit, Len(Unit) - 1)
                        Next re
                        Unit = Right(Unit, Len(Unit) - 1)
    
                    End If
                Next y
                If Unit = "FT" Then Unit = "EA"
                'MsgBox "Unit=:" & Unit & ":"
                ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
                
                'QUANTITY ORDERED
                'Quantity = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0)
                '        For Replace = 0 To 20
                '            If Not Right(Quantity, 1) Like "[0-9]" Then Quantity = Left(Quantity, Len(Quantity) - 1)
                '        Next Replace
                'MsgBox "Quantity=:" & Quantity & ":"
                'ThisWorkbook.Sheets("Temp").Range("R2").Offset(TempSheetOffset, 0) = Quantity
                
                'SHIPPED
                For y = 1 To 20
                    SHIP = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, y)
                    If SHIP Like "*[0-9]ea*" Then
                        'MsgBox ship
                        SHIP = Replace(SHIP, "ea", "")
                        Exit For
                    End If
                Next y
                'MsgBox "Shipped=:" & ship & ":"
                ThisWorkbook.Sheets("Temp").Range("R2").Offset(tempsheetoffset, 0) = SHIP
             
                'UNIT PRICE
                'unitprice = ""
                For y = 0 To 20   '
                    If UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, y)) Like "*[0-9]/[A-Z]*" Then
                        unitprice = UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, y))
                        For re = 0 To 20
                            'MsgBox unitprice
                            If Right(unitprice, 1) <> "/" Then unitprice = Left(unitprice, Len(unitprice) - 1)
                        Next re
                        unitprice = Left(unitprice, Len(unitprice) - 1)
                        unitprice = Replace(unitprice, vbLf, "")
                        unitprice = Replace(unitprice, "/EA", "")
                    End If
                Next y
    
                If Unit = "FT" Then Unit = "EA"
                'MsgBox "Unit=:" & Unit & ":"
                ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
                If Unit = "C" Then
                    If unitprice = "" Then unitprice = "0"
                    unitprice = Replace(unitprice, " ", "")
                    unitprice = unitprice / 100
                    Unit = "EA"
                    ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
                End If
                If Unit = "M" Then
                    unitprice = unitprice / 1000
                    Unit = "Ea"
                    ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
                End If
                'MsgBox "UnitPrice=:" & unitprice & ":"
                If unitprice Like "*\*\*" Or unitprice Like "*.*.*" Then unitprice = ""
                If unitprice = "" Then Possibleerror = Possibleerror + 1
                ThisWorkbook.Sheets("Temp").Range("S2").Offset(tempsheetoffset, 0) = unitprice
                lineprice = Right(Line, 10) '           <<<LinePrice
                lineprice = Replace(lineprice, " ", "")
                If unitprice = "" Then unitprice = "0"
                'If SHIP = 0 Then LinePrice = unitprice * QUANTITY
                'MsgBox "LinePrice=:" & LinePrice & ":"
                ThisWorkbook.Sheets("Temp").Range("T2").Offset(tempsheetoffset, 0) = lineprice
                ThisWorkbook.Sheets("Temp").Range("A2").Offset(tempsheetoffset, 0) = DecoPO
                ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
                ThisWorkbook.Sheets("Temp").Range("C2").Offset(tempsheetoffset, 0) = "263" 'Stoneway
                ThisWorkbook.Sheets("Temp").Range("AH2").Offset(tempsheetoffset, 0) = Tax
                ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = TotalInvoice
                ThisWorkbook.Sheets("Temp").Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
                ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
                ThisWorkbook.Sheets("Temp").Range("J2").Offset(tempsheetoffset, 0) = InvoiceDate
                tempsheetoffset = tempsheetoffset + 1
        End If
    Next xoffset
    

    If UCase(path) Like "*ATTACHMENT*" Then
        If ThisWorkbook.Sheets("Temp").Range("AD2") = "" Then
            MsgBox "known bug flag, didn't get invoice due date to put in cell AD2"
            Exit Sub
        End If
    End If

    If UCase(path) Like "*ATTACHMENT*" Then
        If ThisWorkbook.Sheets("Temp").Range("AE2") = "" Then MsgBox "known bug flag, didn't get discount date to put in cell AD2"
    End If

    If UCase(path) Like "*ATTACHMENT*" Then
        If ThisWorkbook.Sheets("Temp").Range("AF2") = "" Then MsgBox "known bug flag, didn't get Discount amount to put in cell AD2"
    End If

'Close Excel Sheet
    Workbooks(XLSname).Close SaveChanges:=False
    Kill XLSpath

If docno <> "" Then GoTo docno: 'Going to get better info from webstie so no sense in "Healing" the PDF data
 

    Call SelfHealTempPage

'Check if PO conforms before bothering to Enter
    TargetPO = ThisWorkbook.Sheets("Temp").Range("A2")
    Call CheckPONumber(TargetPO, Found)

docno:

    'MsgBox "in the stoneway loop"
    'Insert Call to Webscrape Here
    
    
    If docno <> "" Then
      ' Check if the workbook is open
        For Each wb In Workbooks
            If wb.Name = XLSname Then
                ' If workbook is found, close it without saving changes
                wb.Close SaveChanges:=False
                Kill XLSpath
            End If
        Next wb
        
        Call StonewayWebscrape(fpath, docno, Email, InvoiceDate, InvoiceDueDate, InvoiceDiscountDate, Discount, path, fname, emailmessage):
        If Not UCase(path) Like "*ATTACHMENT*" Then Exit Sub 'just here to scrape submittals
        'rename and move file
        TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
        & "ORDACK " & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
        & TotalInvoiceAmount & ").pdf"
        pdfoption1 = Replace(pdfoption1, "$", "")
        
        If Email = "Saved" Then
            If Dir(fpath) <> "" Then Name fpath As "\\server2\Faxes\STONEWAY\" & pdfoption1
            updatelog = pdfoption1
            Call logupdate(updatelog)
            Application.Wait (Now + TimeValue("00:00:06"))
        End If
        
        Exit Sub
        
    Else
        If Not UCase(path) Like "*ATTACHMENT*" Then Exit Sub 'just here to scrape submittals
        'MsgBox "not enough info to call for webscrape"
    End If
    
    If Found = 2 Then 'TargetPO matches Subcontract Numbers
        'rename and move file
        TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
        & "ORDACK " & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
        & TotalInvoiceAmount & ").pdf"
        pdfoption1 = "Subcontract" & Replace(pdfoption1, "$", "")
        
        If Dir("\\server2\Faxes\" & "\" & pdfoption1) = "" And UCase(path) Like "*ATTACH*" Then
        Name fpath As "\\server2\Faxes\" & pdfoption1
        updatelog = pdfoption1
        Call logupdate(updatelog)
        Application.Wait (Now + TimeValue("00:00:06"))
        End If
        
        SetCursorPos 1083, 11 '--------------------------'Sage Minimize
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Application.Wait (Now + TimeValue("00:00:01"))
        
        'MsgBox "Sent TargetPO that matches Subcontract number to Fax Folder"
        
        Exit Sub
    
    End If
            
                
    If Found = 1 And Possibleerror < 1 Then 'TargetPO is normal and there are no other errors
        Call ClickOnSage
        xoffset = 0
        emailmessage = "Stoneway Invoice"
        Call SageEnterINVOICEfromTEMP(xoffset, emailmessage, fpath)
        If emailmessage = "Temp Sheet Total Error" Then Exit Sub
        'rename and move file
        TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
        & "ORDACK " & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
        & TotalInvoiceAmount & ").pdf"
        pdfoption1 = Replace(pdfoption1, "$", "")
        
        If Dir("\\server2\Faxes\STONEWAY" & "\" & pdfoption1) = "" And UCase(path) Like "*ATTACH*" Then
            Name fpath As "\\server2\Faxes\STONEWAY\" & pdfoption1
            updatelog = pdfoption1
            Call logupdate(updatelog)
            Application.Wait (Now + TimeValue("00:00:06"))
        End If
        
        SetCursorPos 1083, 11 'Sage Minimize
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Application.Wait (Now + TimeValue("00:00:01"))
    Else
        MsgBox "Did not enter " & TargetPO & Chr(13) & "PossibleErrors =" & Possibleerror _
        & Chr(13) & "Found =" & Found & Chr(13) & "Moving to Fax File as;" & Chr(13) _
        & TargetPO & " Human must enter me"
        
        If Dir("\\server2\Faxes\" & TargetPO & " Human must enter me.pdf") <> "" Then TargetPO = TargetPO & ".1"

        If Dir(fpath) <> "" And Dir("\\server2\Faxes\" & TargetPO & " Human must enter me.pdf") = "" Then Name fpath As "\\server2\Faxes\" & TargetPO & " Human must enter me.pdf"
    
        If Dir(fpath) <> "" And UCase(path) Like "*ATTACH*" Or emailmessage Like "*already been entered*" And Dir(fpath) <> "" Then Kill (fpath)
End If
    
End Sub

Sub StonewayWebscrape(fpath, docno, Email, InvoiceDate, InvoiceDueDate, InvoiceDiscountDate, Discount, path, fname, emailmessage): 'LAUNCH Chrome to webscrape the order ack
    Dim Pic As Object
    Dim vendoritemno As String
    
'Clear Temp Sheet "BA" Field
    ThisWorkbook.Sheets("Temp").Range("BA2:CA300") = ""
    ThisWorkbook.Sheets("Temp").Range("BA2:CA300").UnMerge

    Call FormatTempSheet

'Load Chrome
    ShostName = Environ$("computername")
    
'ThisWorkbook.Sheets("Home").Range("A1") = ShostName
    file = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    If Dir("C:\Program Files\Google\Chrome\Application\chrome.exe") <> "" Then file = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    
    Shell (file)
    
'<<Maximize
    Application.Wait (Now + TimeValue("00:00:06"))
    Application.SendKeys "%{ }" '
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "x"
    Application.Wait (Now + TimeValue("00:00:01"))
    Tryagain = 0

start:
    'Click into navigation field // Navigate to SToneway document
    Application.Wait (Now + TimeValue("00:00:04"))
    Application.SendKeys ("%d"), True
    Sleep (250)
    
    URLDocNo = docno
    If Len(URLDocNo) > 10 Then URLDocNo = Left(URLDocNo, 10)
    
    TargetURL = "https://www.stoneway.com/index.cfm?dsp=member.MyAccount.MyOrders.all_orders#order_id=" & URLDocNo
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys (TargetURL), True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("~"), True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("~"), True
    Application.Wait (Now + TimeValue("00:00:08"))
    
    For Repeat = 1 To 5
        Application.SendKeys ("{tab}")
        Sleep 50
    Next Repeat
    
'COPY->PASTE Page Data
    Set Clipboard = New MSForms.DataObject
    Application.CutCopyMode = False
    Clipboard.Clear
    page_data = ""
    ThisWorkbook.Sheets("Temp").Range("BA2:BZ2000") = ""
    Application.CutCopyMode = False
    Application.SendKeys ("^a"), True
    Sleep (250)
    For Repeat = 1 To 3
        Application.SendKeys ("^c"), True
        Sleep 500
    Next Repeat
    Clipboard.GetFromClipboard
    page_data = Clipboard.GetText
    
    
ThisWorkbook.Sheets("Temp").Paste Destination:=ThisWorkbook.Sheets("Temp").Range("BA2")
    Sleep 500
    ThisWorkbook.Sheets("Temp").DrawingObjects.Delete
    ThisWorkbook.Sheets("Temp").UsedRange.UnMerge
'    ThisWorkbook.Sheets("Temp").Range("BA2").PasteSpecial Paste:=xlPasteValues
'Check if at login page
    'tab 16 times to hit "Login" with pre-populated data
    Found = 0
    For x = 0 To 100
        For y = 0 To 26
            If UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, y)) Like "*PASSWORD:*" Then Found = 1
        Next y
        If Found = 1 Then Exit For
    Next x
    If Found = 1 Then
       ' MsgBox "We're at the login page"
        For Repeat = 1 To 21
            Application.SendKeys "{Tab}"
            Sleep 150
        Next Repeat
        Application.SendKeys "~"
        Application.Wait (Now + TimeValue("00:00:04"))
        GoTo start:
    End If
'Check if at ORDER PAGE
    Found = 0
    For x = 0 To 500
        For y = 0 To 7
            If UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, y)) Like "*" & docno & "*" Then Found = 1
        Next y
        If Found = 1 Then Exit For
    Next x
    
    If Found = 1 Then
        'MsgBox "We're at the ORDER page"
        'if not scraping submittlas then close chrome
        If UCase(path) Like "*ATTACHMENT*" Then
            Application.SendKeys ("^w")
        End If
    Else
        Tryagain = Tryagain + 1
        If Tryagain > 2 And UCase(path) Like "*ATTACHMEN*" Then
            MsgBox "Didn't find order number " & docno & " on page"
        Else
            GoTo start:
        End If
        
        If Not UCase(path) Like "*ATTACHMENT*" Then
            Application.SendKeys ("^w")
        End If
        
        Exit Sub
    End If
    
    
    If UCase(path) Like "*ATTACHMENT*" Then
        Application.SendKeys ("^w")
    End If


    
    
'Process Data
    For starting_point = 0 To 500
        If ThisWorkbook.Sheets("Temp").Range("BA2").Offset(starting_point, 0) Like "*Invoice No.*" Then Exit For
    Next starting_point

    If starting_point > 499 Then MsgBox "erorr scraping data from Stoneway Page, did not find data on Temp Sheet"
    
                       
    'VENDOR INVOICE NUMBER (This is an Order Acknowledgement so this shouldn't be relevant)
    vendorInvoice = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(starting_point, 1)
    Possibleerror = 0
    
    For xoffset = (starting_point - 20) To 500 'Gather Macro Information
        yoffset = 0
       'MsgBox "Freeze"
       Line = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset)
       
       'INVOICE Date
        If Line Like "*Invoice Date:*" Then '
           InvoiceDate = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 1)
           If Not InvoiceDate Like "*/*/*" Then InvoiceDate = ""
        If InvoiceDate = "" Then InvoiceDate = Date
            MsgBox "InvoiceDate=:" & InvoiceDate & ":"
        End If
        
       'ORDER DATE
       If Line Like "*Order Date:*" Then '
           OrderDate = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 1)
           If Not OrderDate Like "*/*/*" Then OrderDate = ""
           'MsgBox "OrderDate=:" & OrderDate & ":"
       End If
               
       'DECO PO
       If Line Like "*PO Number:*" Then '
            DecoPO = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 1)
           DecoPO = Replace(DecoPO, "PO", "")
           DecoPO = Replace(DecoPO, "#", "")
           DecoPO = Replace(DecoPO, " ", "")
           'MsgBox "decoPO=:" & DecoPO & ":"
       End If
       
       'TOTAL PO AMOUNT
       'If UCase(Line) Like "*Subtotal*" Then '
       '    totalinvoice = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 1)
       '    totalinvoice = Replace(totalinvoice, " ", "")
       '    MsgBox "TotalInvoice=:" & totalinvoice & ":"
       'End If
       
        'Next yoffset
    Next xoffset
    
    'MsgBox "Got Macro Data"
    

    
    'ACQUIRE LINE ITEMS
    
    For starting_point = 0 To 500
        If ThisWorkbook.Sheets("Temp").Range("BB2").Offset(starting_point, 0) Like "*" & docno & "*" Then Exit For
    Next starting_point

    If starting_point > 499 Then MsgBox "error finding DocNo->" & docno & Chr(13) & "On Stoneway Page"
    'MsgBox DocNo
    For x = starting_point To starting_point + 50
        If ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, 0) Like "All" Then Exit For

    Next x
    starting_point = x + 1
    'MsgBox "found 'All' for " & DocNo & " at line->" & starting_point + 2
    
    TotalInvoice = ""
    For xoffset = (starting_point) To starting_point + 100 'Gather Macro Information 'Now input Line Items
        yoffset = 1
            
        'SHIPPING & HANDLING
        'MsgBox shipping
        If shipping = "" Or shipping = 0 Or shipping = Empty Then
                For y = 1 To 8
                    'MsgBox "Checking for shipping on line " & xoffset + 2 & Chr(13) & ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, y)
                    If ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, y) Like "*Shipping*andling*" Then
                        shipping = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, y + 1)
                        'MsgBox "Shipping cost found->" & shipping
                        Exit For
                    End If
                Next y
        End If
        
        'For Yoffset = 0 To 20
        Line = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset)
        Line = Replace(Line, vbLf, "")
        
        'ITEMs LIne
        If Line <> "" And Not Line Like "*Item*" Then
            itemDesc = Line
            itemDesc = Replace(itemDesc, vbLf, "")
            If Len(itemDesc) > 60 Then ItemDec = Left(itemDesc, 60)
            'MsgBox "ItemDesc=:" & ItemDesc
            If itemDesc = "" Then Possibleerror = Possibleerror + 1
            ThisWorkbook.Sheets("Temp").Range("P2").Offset(tempsheetoffset, 0) = itemDesc
            
            'VendorItemNo
            For xx = 0 To 5
                For yy = 0 To 5
                    If ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset + xx, yoffset + yy) Like "*Item*:*" Then
                        vendoritemno = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset + xx, yoffset + yy + 1)
                        vendoritemno = Replace(vendoritemno, " ", "")
                        vendoritemno = Replace(vendoritemno, " ", "")
                        'MsgBox ":" & vendoritemno & ":"
                        
                    End If
                Next yy
            Next xx
            
            'UNIT
            Unit = "E"
            'MsgBox unit
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
            'MsgBox "Unit=:" & unit & ": " & "possible errors (0/1)->" & Possibleerror
            ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
                       
            'QUANTITY ORDERED (NOT SHIPPED)
            Quantity = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 3)
            Quantity = Replace(Quantity, vbLf, "")
            Quantity = Replace(Quantity, " ", "")
            For re = 0 To 15
                If UCase(Right(Quantity, 1)) Like "[A-Z]" Then Quantity = Left(Quantity, Len(Quantity) - 1)
            Next re
            'MsgBox "Quantity=:" & Quantity & ":"
            ThisWorkbook.Sheets("Temp").Range("R2").Offset(tempsheetoffset, 0) = Quantity
            
            'TAX
            If UCase(Line) Like "*SALES*TAX*" Then '
                SalesTax = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 1)
                SalesTax = Replace(SalesTax, "SALES", "")
                SalesTax = Replace(SalesTax, "TAX", "")
                 SalesTax = Replace(SalesTax, " ", "")
              MsgBox "SalesTax=:" & SalesTax & ":"
            End If
       
            'UNIT PRICE

            unitprice = UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 4))

            If unitprice = "" Then unitprice = "0"
           
            unitprice = Replace(unitprice, vbLf, "")
            unitprice = Replace(unitprice, "(100 EA)", "")
            unitprice = Replace(unitprice, "(100 FT)", "")
            unitprice = Replace(unitprice, " ", "")
            unitprice = Replace(unitprice, "$", "")
            If unitprice <> "0" Then unitprice = unitprice / Quantity
            For re = 0 To 15
                If Right(UCase(unitprice), 1) Like "[A-Z]" Or Right(unitprice, 1) Like "/" Or Right(unitprice, 1) Like "." Then unitprice = Left(UCase(unitprice), Len(unitprice) - 1)
            Next re
            If unitprice = "" Then unitprice = "0"
            If unitprice Like "*[A-Z]*" Then unitprice = "0"
            ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = unitprice
            
            'LINE TOTAL
            ThisWorkbook.Sheets("Temp").Range("S2").Offset(tempsheetoffset, 0) = unitprice
            'LinePrice = Right(Line, 10) '           <<<LinePrice
            'LinePrice = Replace(LinePrice, " ", "")
            'MsgBox unitprice
            
            If Quantity = "" Then Quantity = 0
            'MsgBox Quantity
            
            If SHIP = 0 Then lineprice = unitprice * Quantity
        
            'MsgBox "LinePrice=:" & LinePrice & ":"
            ThisWorkbook.Sheets("Temp").Range("T2").Offset(tempsheetoffset, 0) = lineprice
            ThisWorkbook.Sheets("Temp").Range("A2").Offset(tempsheetoffset, 0) = DecoPO
            ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
            ThisWorkbook.Sheets("Temp").Range("C2").Offset(tempsheetoffset, 0) = "263" 'stoneway
            ThisWorkbook.Sheets("Temp").Range("AH2").Offset(tempsheetoffset, 0) = Tax
            ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = TotalInvoice
            ThisWorkbook.Sheets("Temp").Range("O2").Offset(tempsheetoffset, 0).NumberFormat = "@"
            ThisWorkbook.Sheets("Temp").Range("O2").Offset(tempsheetoffset, 0) = vendoritemno
            ThisWorkbook.Sheets("Temp").Range("H2").Offset(tempsheetoffset, 0) = docno
            ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = VendorTotalInvoice
            ThisWorkbook.Sheets("Temp").Range("J2").Offset(tempsheetoffset, 0) = InvoiceDate
            ThisWorkbook.Sheets("Temp").Range("AD2") = InvoiceDueDate
            ThisWorkbook.Sheets("Temp").Range("AE2") = InvoiceDiscountDate
            ThisWorkbook.Sheets("Temp").Range("AF2") = Discount
            tempsheetoffset = tempsheetoffset + 1

        End If
        
        If UCase(path) Like "*ATTACHMENT*" Then
            'If ThisWorkbook.Sheets("Temp").Range("AD2") = "" Then MsgBox "known bug flag, didn't get invoice due date to put in cell AD2"
        End If
        If UCase(path) Like "*ATTACHMENT*" Then
            'If ThisWorkbook.Sheets("Temp").Range("AE2") = "" Then MsgBox "known bug flag, didn't get discount date to put in cell AD2"
        End If
        If UCase(path) Like "*ATTACHMENT*" Then
            'If ThisWorkbook.Sheets("Temp").Range("AF2") = "" Then MsgBox "known bug flag, didn't get Discount amount to put in cell AD2"
        End If
        
        'MsgBox "BE2 Offset->" & ThisWorkbook.Sheets("Temp").Range("BE2").Offset(xoffset, yoffset - 1)
        If ThisWorkbook.Sheets("Temp").Range("BE2").Offset(xoffset, yoffset - 1) Like "Total*" Then '
           VendorTotalInvoice = ThisWorkbook.Sheets("Temp").Range("BE2").Offset(xoffset, yoffset)
           VendorTotalInvoice = Replace(VendorTotalInvoice, "$", "")
           'MsgBox "VendorTotalInvoice=:" & VendorTotalInvoice & ":"
           TotalInvoice = VendorTotalInvoice
           ThisWorkbook.Sheets("Temp").Range("N2") = VendorTotalInvoice
        End If
        
        'Order Total
        If TotalInvoice <> "" Then Exit For
    Next xoffset
    
    If shipping <> "" And shipping <> 0 And shipping <> "0" Then
        ThisWorkbook.Sheets("Temp").Range("S2").Offset(tempsheetoffset, 0) = shipping
        ThisWorkbook.Sheets("Temp").Range("A2").Offset(tempsheetoffset, 0) = DecoPO
        ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
        ThisWorkbook.Sheets("Temp").Range("C2").Offset(tempsheetoffset, 0) = "STONE"
        ThisWorkbook.Sheets("Temp").Range("AH2").Offset(tempsheetoffset, 0) = Tax
        ThisWorkbook.Sheets("Temp").Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
        ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = VendorTotalInvoice
        ThisWorkbook.Sheets("Temp").Range("P2").Offset(tempsheetoffset, 0) = "Shipping and Handling"
        ThisWorkbook.Sheets("Temp").Range("R2").Offset(tempsheetoffset, 0) = 1
        ThisWorkbook.Sheets("Temp").Range("J2").Offset(tempsheetoffset, 0) = InvoiceDate
    End If
    
    
    'MsgBox "Done scraping Web data"
    Call SelfHealTempPage
    'Check if PO conforms before bothering to Enter
    TargetPO = UCase(ThisWorkbook.Sheets("Temp").Range("A2"))
    'MsgBox TargetPO
    Call CheckPONumber(TargetPO, Found)
    
    
    
    If Found = 2 Or Found = 3 Then '2 = TargetPO matches subcontract number, 3 = Shop Invoice
        If Not UCase(path) Like "*ATTACHMENT*" Then Exit Sub 'just here to scrape submittals
        'rename and move file
        TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        
        pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
        & "INV " & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
        & TotalInvoiceAmount & ").pdf"
        
        If Found = 2 Then pdfoption1 = "Subcontract " & Replace(pdfoption1, "$", "")
        If Found = 3 Then pdfoption1 = "SHOP PO " & Replace(pdfoption1, "$", "")
        
        If emailmessage = "Saved" Then
            If Dir("\\server2\Faxes\" & pdfoption1) = "" And Dir(fpath) <> "" Then Name fpath As "\\server2\Faxes" & pdfoption1
            updatelog = pdfoption1
            Call logupdate(updatelog)
            Application.Wait (Now + TimeValue("00:00:06"))
        End If
        'MsgBox fpath
        
        If Dir(fpath) <> "" Or emailmessage Like "Saved" And Dir(fpath) <> "" Then Kill (fpath)
            SetCursorPos 1083, 11 '--------------------------'Sage Minimize
            Call Mouse_left_button_press
            Call Mouse_left_button_Letgo
            Application.Wait (Now + TimeValue("00:00:01"))
            
            'MsgBox "Sent TargetPO that matches Subcontract number, Or is Shop PO to Fax Folder"
            
        Exit Sub
    End If
    
    If Found = 0 And UCase(path) Like "*ATTACHMENT*" Then
        MsgBox "FOund->" & Found & Chr(13) & "Possibleerror->" & Possibleerror & Chr(13) & "Failed to confirm valid PO->" & TargetPO
    End If
    
    If Not UCase(path) Like "*ATTACHMENT*" Then Call Stoneway_get_Submittal(path, fname, page_data)  'just here to scrape submittals
    If Not UCase(path) Like "*ATTACHMENT*" Then Exit Sub 'just here to scrape submittals
    
    If Found = 1 And Possibleerror < 1 Then
        Call ClickOnSage
        xoffset = 0
        emailmessage = "Stoneway"
        Call SageEnterINVOICEfromTEMP(xoffset, emailmessage, fpath)
        If emailmessage = "Temp Sheet Total Error" Then Exit Sub
        'rename and move file
        TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        
        pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
        & "INV " & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
        & TotalInvoiceAmount & ").pdf"
        pdfoption1 = Replace(pdfoption1, "$", "")
        
        If emailmessage = "Saved" Then
            If Dir("\\server2\Faxes\STONEWAY - 263\" & pdfoption1) = "" And Dir(fpath) <> "" Then Name fpath As "\\server2\Faxes\STONEWAY - 263\" & pdfoption1
            updatelog = pdfoption1
            Call logupdate(updatelog)
            Application.Wait (Now + TimeValue("00:00:06"))
        End If
        'MsgBox fpath
        
        If Dir(fpath) <> "" Or emailmessage Like "Saved" And Dir(fpath) <> "" Then Kill (fpath)
            SetCursorPos 1083, 11 '--------------------------'Sage Minimize
            Call Mouse_left_button_press
            Call Mouse_left_button_Letgo
            Application.Wait (Now + TimeValue("00:00:01"))
        Else
        MsgBox "Did not enter ORDACK even after second try scraping web data" & TargetPO & Chr(13) & "PossibleErrors =" & Possibleerror _
        & Chr(13) & "Found =" & Found & Chr(13) & "TargetPO ->" & TargetPO
    End If


End Sub
Sub Stoneway_get_Submittal(path, fname, page_data)
'identify submittal folder
'\\server2\Dropbox\Acct\100 Jobs\2022 JOBS\2202 Marysville Stormwater Treatment (McClure)\300 Accounting\Backup
'\\server2\Dropbox\Acct\100 Jobs\2022 JOBS\2202 Marysville Stormwater Treatment (McClure)\Submittals
'\\server2\Dropbox\Acct\100 Jobs\2021 JOBS\2105 Arlington HS (Kassel)\300 Accounting\Backup
'\\server2\Dropbox\Acct\100 Jobs\2021 JOBS\2105 Arlington HS (Kassel)\Submittals
'
Submittal_Folder_Location = Replace(path, "300 Accounting\Backup", "Submittals")

'identify current order page
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("%d")
    Application.Wait (Now + TimeValue("00:00:01"))
    Set Clipboard = New MSForms.DataObject
    Application.CutCopyMode = False
    Clipboard.Clear
    Order_URL = ""
    Application.SendKeys ("^c")
    Sleep 150
    Clipboard.GetFromClipboard
    Order_URL = Clipboard.GetText
    If Order_URL = "" Then
        MsgBox "URL copy went wrong!"
    End If
    
'Run loop to download cut sheets
For x = 0 To 100
    try = 1

Get_item_description:
    Call Get_item_description(x, item_description, file_description, page_data)
    If item_description = "" Then GoTo close_webpage
    answer = 0 'whether or not this item has already been downlaoded
    Call Check_if_already_downloaded_this_submittal(Submittal_Folder_Location, file_description, answer)
    If answer = 2 Then 'already downloaded this submittal
        x = x + 1
        GoTo Get_item_description:
    End If
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("^f")
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys item_description, True
    Application.Wait (Now + TimeValue("00:00:01"))
    'advance cursor to selected item
    'Application.SendKeys ("+^~")
    'Application.Wait (Now + TimeValue("00:00:01"))
    'click item to advance to specific page
    Application.SendKeys ("^~")
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("~")
    'wait for page to load
wait_for_product_page_to_load:
    Application.Wait (Now + TimeValue("00:00:04"))
    'check if we have made it to product page
    Application.SendKeys ("%d")
    Application.Wait (Now + TimeValue("00:00:01"))
    Set Clipboard = New MSForms.DataObject
    Application.CutCopyMode = False
    Clipboard.Clear
    product_URL = ""
    Application.SendKeys ("^c")
    Sleep 150
    Clipboard.GetFromClipboard
    product_URL = Clipboard.GetText
    If Order_URL = "" Then
        MsgBox "URL copy went wrong!"
    End If
    If product_URL = Order_URL Then
        try = try + 1
        If try > 3 Then
            x = x + 1
            GoTo Get_item_description:
        End If
        GoTo wait_for_product_page_to_load:
    End If

try = 1
click_on_product_data:
If try = 1 Then findtext = "View Specifications"
'if failed to find specs, then close page and search for next
If try > 2 Then
    Call Chrome_close_secondary_tabs
    GoTo restore_primary_tab:
End If

'search and click on catalog link
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("^f")
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys (findtext)
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("{esc}")
    For Repeat = 1 To 20
        Sleep 5
    Next Repeat
    'advance cursor to selected item
    Application.SendKeys ("~")
    

wait_for_catalog_page_to_load:
    Application.Wait (Now + TimeValue("00:00:04"))
    Application.SendKeys ("%d")
    Application.Wait (Now + TimeValue("00:00:01"))
    Set Clipboard = New MSForms.DataObject
    Application.CutCopyMode = False
    Clipboard.Clear
    Check_URL = ""
    Application.SendKeys ("^c")
    Sleep 150
    Clipboard.GetFromClipboard
    Check_URL = Clipboard.GetText
    If Check_URL = product_URL Then
        try = try + 1
        If try > 3 Then
            'x = x + 1
            Call Chrome_close_secondary_tabs
            GoTo restore_primary_tab:
        End If
        GoTo click_on_product_data
    End If

Call Submittals_click_download_PDF(file_description)
    
restore_primary_tab:
    'Restore ORDER_URL in primary tab
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("%d")
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys Order_URL
    Application.SendKeys "~"
wait_for_order_page_to_Load:
    Application.Wait (Now + TimeValue("00:00:05"))
    

Next x

close_webpage:

'Update submittals page with download data
    For x = 0 To 10000
        If ThisWorkbook.Sheets("Submittals").Range("A1").Offset(x, 0) = fname Or ThisWorkbook.Sheets("Submittals").Range("A1").Offset(x, 0) = "" Then Exit For
    Next x
    If fname = "" Then MsgBox "In stoneway module, lost fname"
    ThisWorkbook.Sheets("Submittals").Range("A1").Offset(x, 0) = fname

For Repeat = 1 To 5
    Application.SendKeys ("^w")
    Application.Wait (Now + TimeValue("00:00:01"))
Next Repeat



End Sub
