Attribute VB_Name = "M43NorthCoast"
Sub IfNorthCoast(path, XLSname, XLSpath, fpath, fname, emailmessage)
    Dim InvoiceDate As String
    Dim i As Long
    Dim URL As String
    Dim IE As Object
    Dim objElement As Object
    Dim objCollection As Object
    Dim try As Integer
    Dim vendoritemno As String

'Path = source of pdf's, either attachment folder or backup folder in job fodler for submittal scrape
'Dont_go_to_sage = 1

start:

    Found = 0
    'MsgBox XLSname
    If XLSname <> "" Then GoTo Skip_Conversion: 'tossed here from Stoneway module because already converted and confirmed it's North Coast
    fname = ""
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If path = "" Then
        MsgBox "no path"
        Exit Sub
    End If


    Set objFolder = objFSO.GetFolder(path)
    For Each objFile In objFolder.Files
        fname = objFile.Name
        'MsgBox fname
        fpath = objFile.path
        ' Reference S010203436-0002_26163
        If UCase(fname) Like "*S[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]*.PDF" Or _
        UCase(fname) Like "*NORTHCOAST*" Then
        Found = 1
        Exit For
        End If '
    Next objFile
    If Found = 0 Then Exit Sub
    Call FormatTempSheet
    Call Convert_PDF_to_Excel(fname, fpath, XLSpath, XLSname, emailmessage)


'Verify it's a NorthCoast Item
    Found = 0
    For x = 0 To 100
        For y = 0 To 3
            'MsgBox UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(x, y))
            If UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(x, y)) Like "*NORTH*COAST*" Then Found = 1
        Next y
        If Found = 1 Then Exit For
    Next x
    If Found = 0 Then ' Not an Aknowledgemenbt
        Workbooks(XLSname).Close SaveChanges:=False
        Kill XLSpath
        Exit Sub
    End If

Skip_Conversion:
    
    
'Verify if it's an Acknowledgement
    Found = 0
    For x = 0 To 100
        For y = 0 To 20
            
            If Workbooks(XLSname).Sheets(1).Range("A1").Offset(x, y) Like "*Acknowledge*" Then Found = 1
            If UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(x, y)) Like "*INVOICE*" Then Found = 2
        Next y
        If Found <> 0 Then Exit For
    Next x
    
    If Found = 1 Then
        
        ' Check if the filename contains a dash
        Found = InStr(fname, "-")
        
        If Found > 0 Then
            ' Extract the substring before the dash
            docno = Left(fname, Found - 1)
            'MsgBox "Document number extracted: " & docno
            Workbooks(XLSname).Close SaveChanges:=False
            Kill XLSpath
            Call northcoastORDACKWebscrape(fpath, fname, docno, PDFtype, InvoiceDate, InvoiceDueDate, InvoiceDiscountDate, Discount, path, emailmessage):   'LAUNCH Chrome to webscrape the order ack
            Exit Sub
        Else
            MsgBox "NorthCoast Order Acknowledgement found but No dash found in the filename."
        End If
    End If


    If Found = 0 Then
        Workbooks(XLSname).Close SaveChanges:=False
        Kill XLSpath
        Exit Sub
    End If
    If Found = 2 Then
        'MsgBox "detected North Coast Invoice"
        Call IfNorthCoastInvoice(fpath, fname, XLSname, XLSpath, path, emailmessage):
        Exit Sub
    End If
    docno = ""

' TEST CODE>>>> Do not enter acknowledgements any more!! Just kill them!!
'if made it to here, then code determined the PDF is Northcoast order Acknowledgement, if here for submittlas, now exit
'If Not UCase(path) Like "*ATTACH*" Then
    Workbooks(XLSname).Close SaveChanges:=False
    Kill XLSpath
    Exit Sub
'End If



read_pdf:
    lrow = Workbooks(XLSname).Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
' Get Macro Info
    For xoffset = 0 To lrow
    'MsgBox "Freeze"
        Line = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0)
        
        'NORTH COAST ORDER NUMBER // Doc Number
        If fpath Like "*\\server2\Dropbox\Attachments\*" Then
            docno = Replace(fpath, "\\server2\Dropbox\Attachments", "")
            MsgBox docno
            MsgBox ""
            MsgBox ""
            MsgBox ""
            
        End If
        
        If Line Like "*S[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]*" Then
            docno = Line
            docno = Replace(docno, " ", "")
            docno = Right(docno, 10)
            'MsgBox "North Coast Doc No. is->" & docno
        End If
        'ORDER DATE
        If Line Like "*ORDER*DATE *" Then '
            OrderDate = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 1, 0)
            For re = 0 To 300
                If Left(OrderDate, 1) = " " Then OrderDate = Right(OrderDate, Len(OrderDate) - 1)
            Next re
            'MsgBox "OrderDate=:" & OrderDate & ":"
            rebuild = OrderDate
            OrderDate = ""
            For re = 1 To 20
                If Mid(rebuild, re, 1) = " " Then Exit For
                'MsgBox re
                OrderDate = OrderDate & Mid(rebuild, re, 1)
            Next re
            If Not OrderDate Like "*/*/*" Then OrderDate = ""
            'MsgBox "OrderDate=:" & OrderDate & ":"
        End If
        
        'DECO PO
        If Line Like "*CUSTOMER*ORDER*NUMBER*" Then '
            DecoPO = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 1, 0)
            DecoPO = Replace(DecoPO, "3156", "")
            DecoPO = Replace(DecoPO, " ", "")
            DecoPO = Replace(DecoPO, vbLf, "")
            DecoPO = Replace(DecoPO, "ORDEREDBY", "")
            'MsgBox "North Coast Order Ack" & Chr(13) & "decoPO=:" & DecoPO & ":"
        End If
                
        'TOTAL PO AMOUNT
        If UCase(Line) Like "*AMOUNT*DUE*" Then '
            TotalInvoice = UCase(Line)
            TotalInvoice = Replace(TotalInvoice, "AMOUNT", "")
            TotalInvoice = Replace(TotalInvoice, "DUE", "")
            TotalInvoice = Replace(TotalInvoice, " ", "")
            'MsgBox "TotalInvoice=:" & Totalinvoice & ":"
        End If
        
        'TAX
        If UCase(Line) Like "*SALES*TAX*" Then '
            SalesTax = UCase(Line)
            SalesTax = Replace(SalesTax, "SALES", "")
            SalesTax = Replace(SalesTax, "TAX", "")
            SalesTax = Replace(SalesTax, " ", "")
            'MsgBox "SalesTax=:" & SalesTax & ":"
        End If
    Next xoffset
    

    TargetPO = DecoPO
    Call CheckPONumber(TargetPO, Found)
    If Found = 0 And docno <> "" Then
            'MsgBox "failed to get Target PO or DocNo"
            Workbooks(XLSname).Close SaveChanges:=False
            Kill XLSpath
            
        If ThisWorkbook.Sheets("Temp").Range("E2") Like "*/*" Then
            MsgBox "When scraping North Coast document, temp sheet E2 looks like date instaed of job number, hit enter to exit and invetigate"
            Exit Sub
        End If
        
        Call northcoastORDACKWebscrape(fpath, fname, docno, PDFtype, InvoiceDate, InvoiceDueDate, InvoiceDiscountDate, Discount, path, emailmessage):
        If Not UCase(path) Like "*ATTACHMENT*" Then Exit Sub
        
        GoTo NextPDF:
     End If
     
    'ACQUIRE ITEM LINE
    For xoffset = 0 To lrow 'Gather Macro Information 'Now input Line Items
        Line = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0)
        
        'ITEM DESCRIPTION
        If Line Like "*[0-9]ea *" Then '
            itemDesc = Mid(Line, 5, 60) '
            'Remove Sapce from right side
            For re = 1 To 60
                If Right(itemDesc, 1) = " " Then itemDesc = Left(itemDesc, Len(itemDesc) - 1)
            Next re
            'ItemDesc = Replace(ItemDesc, " ", "")
            'MsgBox "ItemDesc=:" & Itemdesc & ":"
            If itemDesc = "" Then Possibleerror = Possibleerror + 1
            ThisWorkbook.Sheets("Temp").Range("P2").Offset(tempsheetoffset, 0) = itemDesc
            
            'UNIT
            Unit = Line
            Unit = Replace(Unit, vbLf, "")
            Unit = UCase(Mid(Line, 77, 20))
            Unit = Replace(Unit, " ", "")
            For re = 0 To 15
                If Left(Unit, 1) Like "[0-9]" Or Left(Unit, 1) Like "/" Or Left(Unit, 1) Like "." Then Unit = Right(Unit, Len(Unit) - 1)
            Next re
            If Unit = "FT" Then Unit = "EA"
            'MsgBox "Unit=:" & Unit & ":"
            ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
                       
            'QUANTITY
            Quantity = UCase(Mid(Line, 1, 5)) '
            Quantity = Replace(Quantity, " ", "")
            Quantity = Replace(Quantity, "EA", "")
            'MsgBox "Quantity=:" & Quantity & ":"
            ThisWorkbook.Sheets("Temp").Range("R2").Offset(tempsheetoffset, 0) = Quantity
                        
            'UNIT PRICE
            unitprice = UCase(Mid(Line, 77, 20)) '
            'MsgBox UnitPrice
            unitprice = Replace(unitprice, " ", "")
            For re = 0 To 15
                If Right(unitprice, 1) Like "[A-Z]" Or Right(unitprice, 1) Like "/" Then unitprice = Left(unitprice, Len(unitprice) - 1)
            Next re
            
            If Unit = "C" Then
                If unitprice = "" Then unitprice = 0
                unitprice = unitprice / 100
                Unit = "EA"
                ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
            End If
            If Unit = "M" Then
                unitprice = unitprice / 1000
                Unit = "Ea"
                ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
            End If
            
            'LINE TOTAL
            ThisWorkbook.Sheets("Temp").Range("S2").Offset(tempsheetoffset, 0) = unitprice
            lineprice = Right(Line, 10) '           <<<LinePrice
            lineprice = Replace(lineprice, " ", "")
            If unitprice = "" Then unitprice = 0
            If SHIP = 0 Then lineprice = unitprice * Quantity
            ThisWorkbook.Sheets("Temp").Range("T2").Offset(tempsheetoffset, 0) = lineprice
            ThisWorkbook.Sheets("Temp").Range("A2").Offset(tempsheetoffset, 0) = DecoPO
            ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
            ThisWorkbook.Sheets("Temp").Range("C2").Offset(tempsheetoffset, 0) = "218"
            ThisWorkbook.Sheets("Temp").Range("AH2").Offset(tempsheetoffset, 0) = Tax
            ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = TotalInvoice
            ThisWorkbook.Sheets("Temp").Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
            ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
            tempsheetoffset = tempsheetoffset + 1
        End If
    Next xoffset
    
    Workbooks(XLSname).Close SaveChanges:=False
    Kill XLSpath
    Call SelfHealTempPage
    'Check if PO conforms before bothering to Enter
    TargetPO = ThisWorkbook.Sheets("Temp").Range("A2")
    Call CheckPONumber(TargetPO, Found)
    If Found = 1 And Possibleerror < 1 Then
        Call ClickOnSage
        xoffset = 0
        Call SageEnterPOfromTEMP(xoffset, emailmessage)
        If emailmessage = "Job entered was not valid in sage" Then
            sourcePath = fpath
            TargetPath = "\\server2\Dropbox\Attachments\_Re Run\" & fname
            Call PDF_MoveToFolder(sourcePath, TargetPath, specialmessage)
            updatelog = "Job entered was not valid in sage " & fname
            Call logupdate(updatelog)
            Exit Sub
        End If
        'rename and move file
        TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
        & "ORDACK " & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
        & TotalInvoiceAmount & ").pdf"
        pdfoption1 = Replace(pdfoption1, "$", "")
        
        If Dir("\\server2\Faxes\NORTH COAST - 218" & "\" & pdfoption1) = "" Then
        Name fpath As "\\server2\Faxes\NORTH COAST - 218\" & pdfoption1
        updatelog = pdfoption1
        Call logupdate(updatelog)
        Application.Wait (Now + TimeValue("00:00:06"))
        End If
        
        SetCursorPos 1083, 11 '--------------------------'Sage Minimize
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Application.Wait (Now + TimeValue("00:00:01"))
        Else
        'MsgBox "Did not enter ORDACK " & TargetPO & Chr(13) & "PossibleErrors =" & Possibleerror _
        & Chr(13) & "Found =" & Found
        If docno <> "" Then
            If ThisWorkbook.Sheets("Temp").Range("E2") Like "*/*" Then
                MsgBox "Error, hit break to investigate job number looks like date"
                Exit Sub
            End If
            'MsgBox "Will Attempt to scrape the missing data from the web"
            If ThisWorkbook.Sheets("Temp").Range("E2") Like "*/*" Then
                MsgBox "Exit and invesitgate temp sheet E2 looks like date and not job number"
                Exit Sub
            End If
            
            Call northcoastORDACKWebscrape(fpath, fname, docno, PDFtype, InvoiceDate, InvoiceDueDate, InvoiceDiscountDate, Discount, path, emailmessage)
            If Not UCase(path) Like "*ATTACHMENT*" Then Exit Sub
        End If
    End If
NextPDF:
If emailmessage Like "*already been entered*" And Dir(fpath) <> "" Then Kill (fpath)
XLSname = ""
GoTo start:
End Sub
Sub IfNorthCoastInvoice(fpath, fname, XLSname, XLSpath, path, emailmessage) 'CALLED FROM ifNorthCoast()
Dim InvoiceDate As String
Dim i As Long
Dim URL As String
Dim IE As Object
Dim objElement As Object
Dim objCollection As Object
Dim try As Integer

'Path is the source PDF folder, either attachements, or backup folder from job folder for submittals download


            If ThisWorkbook.Sheets("Temp").Range("E2") Like "*/*" Then
                MsgBox "Break and investigate temp sheet E2 like date, not job#"
                Exit Sub
            End If
            
            PDFtype = "Invoice"
            docno = ""
            Found = 0
            'MsgBox XLSname
            For xoffset = 0 To 100 'Gather Macro Information
                For yoffset = 0 To 20
                   'MsgBox "Freeze"
                   'MsgBox XLSname
                   Line = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset)
                   
                'NORTH COAST ORDER NUMBER // Doc Number
                If Line Like "*S[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]*" And docno = "" Then
                    docno = Line
                    'MsgBox DocNo
                    docno = Replace(docno, " ", "")
                    docno = Replace(docno, "Invoice", "")
                    If docno Like "S[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].[0-9][0-9][0-9]" Then
                        'do nothing
                        Else
                        MsgBox "Need to edit DocNo, it's currently " & Chr(13) & docno
                        docno = Left(docno, 10)
                    End If
                    'MsgBox "North Coast Doc No. is->" & DocNo
                End If
                   
                   'INVOICE DATE
                   If Line Like "*INVOICE*DATE*" Then '
                       OrderDate = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 1, yoffset)
                       If Not OrderDate Like "*/*/*" Then OrderDate = ""
                       'MsgBox "InvoiceDate=:" & OrderDate & ":"
                   End If
                   
                   'DISOCUNT DATE / DISCOUNT AMOUNT
                   If Line Like "*you may deduct*" Then '
                       InvoiceDiscountDate = Line
                       If Not InvoiceDiscountDate Like "*/*/*" Then InvoiceDiscountDate = ""
                       'If paid by 01/10/23 you may deduct $3.84
                       InvoiceDiscountDate = Replace(InvoiceDiscountDate, "If paid by", "")
                       InvoiceDiscountDate = Replace(InvoiceDiscountDate, "you may deduct", "")
                       Discount = InvoiceDiscountDate
                       InvoiceDiscountDate = Left(InvoiceDiscountDate, 9)
                       Discount = Mid(Discount, 10, Len(Discount) - 9)
                       Discount = Replace(Discount, "$", " ")
                       Discount = Replace(Discount, " ", "")
                       
                       'MsgBox "North Coast InvoiceDiscountDate=:" & InvoiceDiscountDate & ":" & Chr(13) & "Discount->" & Discount
                       'MsgBox "Freeze"
                       'MsgBox "Freeze"
                       'MsgBox "Freeze"
                       'MsgBox "Freeze"
                       'MsgBox "Freeze"
                   End If
                   
                    'DUE DATE
                   If UCase(Line) Like "*INVOICE*DUE*BY*" Then '
                       InvoiceDueDate = Line
                       If Not InvoiceDueDate Like "*/*/*" Then InvoiceDueDate = ""
                       InvoiceDueDate = Replace(InvoiceDueDate, "Invoice is due by", "")
                       InvoiceDueDate = Replace(InvoiceDueDate, " ", "")
                       'MsgBox "InvoiceDueDate=:" & InvoiceDueDate & ":"
                   End If
                   
                   'DELIVERY METHOD
                   'ThisWorkbook.Sheets("Temp").Range("AE2") = ""
    
                   'DECO PO
                   If Line Like "*CUSTOMER*PO*NO*" Then '
                       DecoPO = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 1, yoffset)
                       DecoPO = Replace(DecoPO, "PPP - PAPER PRNT", "")
                       DecoPO = Replace(DecoPO, " ", "")
                       'MsgBox "decoPO=:" & DecoPO & ":"
                   End If
                   
                   'VENDOR INVOICE NUMBER (This is an Order Acknowledgement so this shouldn't be relevant)
                   If Line Like "PO BOX 418759*" Then '
                       vendorInvoice = Mid(Line, (Len(Line) - 16), 9)
                       vendorInvoice = Replace(vendorInvoice, " ", "")
                       MsgBox "VendorInvoice=:" & vendorInvoice & ":"
                   End If
                   
                   'TOTAL PO AMOUNT
                   If UCase(Line) Like "*AMOUNT*DUE*" Then '
                       TotalInvoice = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 3, yoffset)
                       TotalInvoice = Replace(TotalInvoice, " ", "")
                       'MsgBox "TotalInvoice=:" & Totalinvoice & ":"
                   End If
                   
                   'TAX
                   'If UCase(Line) Like "*SALES*TAX*" Then '
                   '    SalesTax = UCase(Line)
                   '    SalesTax = Replace(SalesTax, "SALES", "")
                   '    SalesTax = Replace(SalesTax, "TAX", "")
                   '    SalesTax = Replace(SalesTax, " ", "")
                   '    MsgBox "SalesTax=:" & SalesTax & ":"
                   'End If
                Next yoffset
            Next xoffset
            
            TargetPO = DecoPO
            
            'MsgBox "Freeze, break and check if temp sheet E2 is date or job no."
            'MsgBox "Freeze, break and check if temp sheet E2 is date or job no."
            'MsgBox "Freeze, break and check if temp sheet E2 is date or job no."
            'MsgBox "Freeze, break and check if temp sheet E2 is date or job no."
            
            Call CheckPONumber(TargetPO, Found)
            
            If docno <> "" Then
                Workbooks(XLSname).Close SaveChanges:=False
                Kill XLSpath
                'MsgBox "NorthCoast Invoice -> about to webscrape for Doc Number" & Chr(13) & DocNo

                
                Call northcoastORDACKWebscrape(fpath, fname, docno, PDFtype, InvoiceDate, InvoiceDueDate, InvoiceDiscountDate, Discount, path, emailmessage)
                If Not UCase(path) Like "*ATTACHMENT*" Then Exit Sub
                
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
                    If Dir(fpath) <> "" Then
                        Name fpath As "\\server2\Faxes\NORTH COAST - 218\" & pdfoption1
                        Application.Wait (Now + TimeValue("00:00:06"))
                    End If
                End If
                
                GoTo NextPDF:
            Else
                MsgBox "Couldn't get Document NO. to call for webscrape"
            End If
            
            'ACQUIRE LINE ITEMS
            For xoffset = 0 To lrow 'Gather Macro Information 'Now input Line Items
                yoffset = 0
                'For Yoffset = 0 To 20
                    Line = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset)
                    Line = Replace(Line, vbLf, "")
                    Line = Replace(Line, " ", "")
                    'ITEM DESCRIPTION
                    If Line Like "[0-9]" Or Line Like "[0-9][0-9]" Then '
                        For y = 0 To 20
                            If UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + y)) Like "*[A-Z][A-Z][A-Z]*" Then Exit For
                        Next y
                        If y > 19 Then MsgBox "located item row, but Could't find item description"
                        itemDesc = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + y)
                        itemDesc = Replace(itemDesc, vbLf, "")
                        If Len(itemDesc) > 60 Then ItemDec = Left(itemDesc, 60)
                        MsgBox "ItemDesc=:" & itemDesc
                        If itemDesc = "" Then Possibleerror = Possibleerror + 1
                        ThisWorkbook.Sheets("Temp").Range("P2").Offset(tempsheetoffset, 0) = itemDesc
                        
                        'UNIT
                        For y = 0 To 20
                            If UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + y)) Like "*[0-9].[0-9]*[A-Z]*" Then Exit For
                        Next y
                        Unit = UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + y))
                        Unit = Replace(Unit, vbLf, "")
                        Unit = Replace(Unit, " ", "")
                        If Unit Like "*E*E*" Then Possibleerror = Possibleerror + 1 'idicates that rows are combined in excel conversion
                        For re = 0 To 15
                            If Left(Unit, 1) Like "[0-9]" Or Left(Unit, 1) Like "/" Or Left(Unit, 1) Like "." Then Unit = Right(Unit, Len(Unit) - 1)
                        Next re
                        If Unit = "FT" Then Unit = "EA"
                        MsgBox "Unit=:" & Unit & ":"
                        ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
                                   
                        'QUANTITY
                        For y = 1 To 20
                            If UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + y)) Like "*[0-9]*" _
                            And Not UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + y)) Like "*[A-Z]*" Then Exit For
                        Next y
                        For nexty = y + 1 To 25
                            If UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + nexty)) Like "*[0-9]*" _
                            And Not UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + nexty)) Like "*[A-Z]*" Then Exit For
                        Next nexty
                        Quantity = UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + nexty))
                        Quantity = Replace(Quantity, vbLf, "")
                        Quantity = Replace(Quantity, " ", "")
                        MsgBox "Quantity=:" & Quantity & ":"
                        ThisWorkbook.Sheets("Temp").Range("R2").Offset(tempsheetoffset, 0) = Quantity
                        
                        'UNIT PRICE
                        For y = 0 To 20
                            If UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + y)) Like "*[0-9].[0-9]*[A-Z]*" Then Exit For
                        Next y
                        'MsgBox y
                        unitprice = UCase(Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, yoffset + y))
                        'MsgBox unitprice
                        unitprice = Replace(unitprice, vbLf, "")
                        unitprice = Replace(unitprice, " ", "")
                        For re = 0 To 15
                            If Right(UCase(unitprice), 1) Like "[A-Z]" Or Right(unitprice, 1) Like "/" Or Right(unitprice, 1) Like "." Then unitprice = Left(UCase(unitprice), Len(unitprice) - 1)
                        Next re
                        'MsgBox unitprice
                        If unitprice = "FT" Then unitprice = "EA"
                        ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = unitprice
                        
                        If Unit = "C" Then
                            If unitprice = "" Then unitprice = 0
                            unitprice = unitprice / 100
                            Unit = "EA"
                            ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
                        End If
                        If Unit = "M" Then
                            unitprice = unitprice / 1000
                            Unit = "Ea"
                            ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
                        End If
                        MsgBox "UnitPrice=:" & unitprice & ":"
                        If unitprice = "" Then unitprice = "0"
                        If unitprice Like "*[A-Z]*" Then unitprice = "0"
                        
                        'LINE TOTAL
                        
                        ThisWorkbook.Sheets("Temp").Range("S2").Offset(tempsheetoffset, 0) = unitprice
                        'LinePrice = Right(Line, 10) '           <<<LinePrice
                        'LinePrice = Replace(LinePrice, " ", "")
                        'MsgBox unitprice
                        'MsgBox Quantity
                        
                        
                        If SHIP = 0 Then lineprice = unitprice * Quantity
                        'MsgBox "LinePrice=:" & LinePrice & ":"
                        ThisWorkbook.Sheets("Temp").Range("T2").Offset(tempsheetoffset, 0) = lineprice
                        ThisWorkbook.Sheets("Temp").Range("A2").Offset(tempsheetoffset, 0) = DecoPO
                        ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
                        ThisWorkbook.Sheets("Temp").Range("C2").Offset(tempsheetoffset, 0) = "218"
                        ThisWorkbook.Sheets("Temp").Range("AH2").Offset(tempsheetoffset, 0) = Tax
                        ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = TotalInvoice
                        ThisWorkbook.Sheets("Temp").Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
                        ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
                        ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = InvoiceDate
                        tempsheetoffset = tempsheetoffset + 1
                    End If
                'Next yOffset
            Next xoffset
            'MsgBox "Done scraping PDF sheet data"
            'MsgBox "Freeze"
            'MsgBox "Freeze"
            'MsgBox "Freeze"
            'MsgBox "Freeze"
            
            
            Workbooks(XLSname).Close SaveChanges:=False
            Kill XLSpath
            Call SelfHealTempPage
            'Check if PO conforms before bothering to Enter
            TargetPO = ThisWorkbook.Sheets("Temp").Range("A2")
            
                        
            If Not UCase(path) Like "*ATTACHMENT*" Then Call Northcoast_download_product_data(path, fname, emailmessage)
            If Not UCase(path) Like "*ATTACHMENT*" Then Exit Sub
            
            If Found = 1 And Possibleerror < 1 Then
            
                Call ClickOnSage
                xoffset = 0
                
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
                
                If Dir("\\server2\Faxes\NORTH COAST - 218\" & pdfoption1) = "" And UCase(path) Like "*ATTACH*" Then
                Name fpath As "\\server2\Faxes\NORTH COAST - 218\" & pdfoption1
                Application.Wait (Now + TimeValue("00:00:06"))
                End If
                'If Dir("\\server2\Faxes\NORTH COAST" & "\" & PDFOption1) = "" And UCase(path) Like "*ATTACH*" Then _
                Name fpath As "\\server2\Faxes\NORTH COAST\" & PDFOption1
                
                SetCursorPos 1083, 11 '--------------------------'Sage Minimize
                Call Mouse_left_button_press
                Call Mouse_left_button_Letgo
                Application.Wait (Now + TimeValue("00:00:01"))
                Else
                MsgBox "Did not enter INVOICE " & fpath & Chr(13) & TargetPO & Chr(13) & "PossibleErrors =" & Possibleerror _
                & Chr(13) & "Found =" & Found
            End If

NextPDF:

End Sub

 Sub northcoastORDACKWebscrape(fpath, fname, docno, PDFtype, InvoiceDate, InvoiceDueDate, InvoiceDiscountDate, Discount, path, emailmessage):  'LAUNCH Chrome to webscrape the order ack
Dim Pic As Object

'path = source folder for PDF, either attachments or backup folder from job folder for submittlas download
    Dont_go_to_sage = 0
    
    If fpath <> "" Then GoTo SkipDirectDocumentNumberInput:
    
    docno = InputBox("Enter the North Coast Document Number")

SkipDirectDocumentNumberInput:


'Clear Temp Sheet "BA" Field
    ThisWorkbook.Sheets("Temp").Range("BA2:CA300") = ""
    ThisWorkbook.Sheets("Temp").Range("BA2:CA300").UnMerge

'Load Chrome
    ShostName = Environ$("computername")
    file = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    If Dir("C:\Program Files\Google\Chrome\Application\chrome.exe") <> "" Then file = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    Shell (file)

' Maximize Chrome
    Application.Wait (Now + TimeValue("00:00:02"))
    Application.SendKeys "%{ }" '
    Sleep 250
    Application.SendKeys "x"
    Sleep 250
start:

'Click into navigation field // Navigate to NorthCoast
    Application.SendKeys ("%d"), True
    Application.Wait (Now + TimeValue("00:00:01"))
    TargetURL = "https://www.northcoast.com/orders/detail/" & docno & "                   "
    Sleep (250)
    Application.SendKeys (TargetURL), True
    Sleep (250)
    Application.SendKeys ("~"), True
    Application.Wait (Now + TimeValue("00:00:10"))

'cycle while waiting to see if we're at the landing page

    For Repeat = 1 To 5
        'set curseor on page
        SetCursorPos 2, 85
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Set Clipboard = New MSForms.DataObject
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.CutCopyMode = False
        Clipboard.Clear
        Page_text = ""
        Application.SendKeys ("^a")
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.SendKeys ("^c")
        Clipboard.GetFromClipboard
        Page_text = Clipboard.GetText
        Clipboard.Clear
        If UCase(Page_text) Like "*NORTHCOAST*" Then
            Exit For
        End If
        Application.Wait (Now + TimeValue("00:00:10"))
    Next Repeat

    

'COPY->PASTE Page Data
copy_paste_page:
    For Repeat = 1 To 16 '(line up over login button if needed, else, just tabbing onto page for successful copy process)
        Application.SendKeys "{Tab}"
        Sleep 150
    Next Repeat

    Application.CutCopyMode = False
        For Repeat = 1 To 3
        Application.SendKeys ("^a"), True
        Sleep (250)
    Next Repeat
    
    For Repeat = 1 To 3
        Application.SendKeys ("^c"), True
        Application.Wait (Now + TimeValue("00:00:01"))
    Next Repeat
    
    Application.Wait (Now + TimeValue("00:00:02"))
    
    ThisWorkbook.Sheets("Temp").Paste Destination:=ThisWorkbook.Sheets("Temp").Range("BA2")
    Sleep 500
    ThisWorkbook.Sheets("Temp").DrawingObjects.Delete
    ThisWorkbook.Sheets("Temp").Range("BA2:CA300").UnMerge

'Check if at login page
'tab 16 times to hit "Login" with pre-populated data
    Found = 0
    For x = 0 To 100
        For y = 0 To 26
            If UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, y)) Like "*PASSWORD?*" Or _
            UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, y)) Like "*PASSWORD?*" Then Found = 1
        Next y
        If Found = 1 Then Exit For
    Next x
    If Found = 1 Then
       ' MsgBox "We're at the login page"
        Application.SendKeys "~"
        Sleep 10000
        GoTo start:
    End If
    'Check if at ORDER PAGE
    Found = 0
    For x = 0 To 100
        For y = 0 To 26
            If UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, y)) Like "*" & docno & "*" Then Found = 1
        Next y
        If Found = 1 Then Exit For
    Next x

    If Found = 1 Then
        'MsgBox "We're at the ORDER page"
        'close chrome
        If UCase(path) Like "*ATTACHMENT*" Then
            Application.SendKeys ("^w")
        End If
    End If '


GoTo SkipGettingProductNumbers:
' Get product numbers
    Dim myRange As Range
    Dim cell As Range
    Dim myString As String
    Dim partialString As String
    Set myRange = ThisWorkbook.Sheets("Temp").Range("BC2:BC100") 'change to your desired range
    partialString = "Our Part" 'change to your desired partial string
    For Each cell In myRange
        'If cell.Text <> "" Then MsgBox cell.Text
        If InStr(1, cell.Text, partialString, vbTextCompare) > 0 Then 'check if cell contains partial string
            Product_Link = cell.Text 'copy cell contents to variable
            Product_Link = Replace(Product_Link, "Our Part#:", "")
            Product_Link = Replace(Product_Link, " ", "")
            'MsgBox Product_Link
            'Do something with myString
            Vendor = "NorthCoast"
            Call Save_Product_Links(Product_Link, Vendor, workbook_open_status)
        End If
    Next cell
    If workbook_open_status = 1 Then
        Workbooks("Product_links.xlsx").Save
        Workbooks("Product_links.xlsx").Close
    End If
SkipGettingProductNumbers:


'Process Data
Possibleerror = 0
For xoffset = 0 To 200 'Gather Macro Information
    For yoffset = 0 To 20
       'MsgBox "Freeze"
       Line = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset)
       
       'INVOICE Date / Due Date
        If Line Like "*Invoice Date:*" Then '
           InvoiceDate = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 1)
           If Not InvoiceDate Like "*/*/*" Then InvoiceDate = ""
          ' MsgBox "InvoiceDate=:" & InvoiceDate & ":"
       End If
       'ORDER DATE
       If Line Like "*Order Date:*" Then '
           OrderDate = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 1)
           If Not OrderDate Like "*/*/*" Then OrderDate = ""
           'MsgBox "OrderDate=:" & OrderDate & ":"
       End If
       
       'DELIVERY METHOD
       'ThisWorkbook.Sheets("Temp").Range("AE2") = ""
       
        'Order Total
        If Line Like "*Order Total*" Then '
           VendorTotalInvoice = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 1)
           VendorTotalInvoice = Replace(VendorTotalInvoice, "$", "")
           'MsgBox "VendorTotalInvoice=:" & VendorTotalInvoice & ":"
       End If
        
       'DECO PO
       If Line Like "*PO*#*" And DecoPO = "" Then '
            DecoPO = Line
           DecoPO = Replace(DecoPO, "PO", "")
           DecoPO = Replace(DecoPO, "#", "")
           DecoPO = Replace(DecoPO, " ", "")
           'MsgBox "decoPO=:" & DecoPO & ":"
           
       End If
       
       'VENDOR INVOICE NUMBER (This is an Order Acknowledgement so this shouldn't be relevant)

        vendorInvoice = docno
       
       'Tax
       If UCase(Line) Like "*TAX*" And yoffset = 0 Then '
           SalesTax = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 1)
           SalesTax = Replace(SalesTax, " ", "")
           For Repeat = 1 To 100
                'MsgBox "NorthCoast SalesTax=:" & SalesTax & ":"
           Next Repeat
       End If
       
        'Shipping and Handling
       If UCase(Line) Like "*HANDLING*" And yoffset = 0 Then '
           handling = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + 1)
           handling = Replace(handling, " ", "")
           'MsgBox "NorthCoast SalesTax=:" & SalesTax & ":"
       End If
    Next yoffset
Next xoffset

For Repeat = 1 To 10
'    MsgBox "Freeze"
Next Repeat


'ACQUIRE LINE ITEMS
For xoffset = 0 To 500 'Gather Macro Information 'Now input Line Items
    yoffset = 0
    'For Yoffset = 0 To 20
        Line = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset)
        Line = Replace(Line, vbLf, "")
        
        'Find Line items
        If UCase(Line) Like "[0-9][A-Z][A-Z]" Or UCase(Line) Like "[0-9][0-9][A-Z][A-Z]" Or _
        UCase(Line) Like "[0-9][0-9][0-9][A-Z][A-Z]" Or UCase(Line) Like "[A-Z][0-9][0-9][0-9][A-Z][A-Z]" Or _
        UCase(Line) Like "[0-9][0-9][0-9][0-9][0-9][A-Z][A-Z]" Or _
        UCase(Line) Like "[0-9][0-9][0-9][0-9][A-Z][A-Z]" Or _
        UCase(Line) Like "-[0-9][A-Z][A-Z]" Or UCase(Line) Like "-[0-9][0-9][A-Z][A-Z]" Or _
        UCase(Line) Like "-[0-9][0-9][0-9][A-Z][A-Z]" Or UCase(Line) Like "-[A-Z][0-9][0-9][0-9][A-Z][A-Z]" Or _
        UCase(Line) Like "-[0-9][0-9][0-9][0-9][0-9][A-Z][A-Z]" _
        Then
        
        ' Error check for lines that are "Call For Price" / "AM" / "PM"

        If Not UCase(Line) Like "*AM*" And Not UCase(Line) Like "*PM*" _
            And Not LCase(ThisWorkbook.Sheets("Temp").Range("BD2").Offset(xoffset, yoffset)) Like "*call*for*price*" Then
            
            'Item Desc
            For y = 0 To 20
                If UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + y)) Like "*[A-Z][A-Z][A-Z]*" Then Exit For
            Next y
            If y > 19 Then MsgBox "located item row, but Could't find item description / Line->" & xoffset + 2
            itemDesc = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + y)
            itemDesc = Replace(itemDesc, vbLf, "")
            If Len(itemDesc) > 60 Then ItemDec = Left(itemDesc, 60)
            'MsgBox "ItemDesc=:" & ItemDesc
            If itemDesc = "" Then
                MsgBox "Northcoast Invoice Wescrape" & Chr(13) & "picked up blank description cell on temp sheet" & Chr(13) & "xoffset->" & xoffset & Chr(13) & "yoffset->" & yoffset
                Possibleerror = Possibleerror + 1
            End If
            ThisWorkbook.Sheets("Temp").Range("P2").Offset(tempsheetoffset, 0) = itemDesc
            
           
            'UNIT
            For y = 0 To 20
                If UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + y)) Like "*[0-9].[0-9]*/*[A-Z]*" And _
                Len(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + y)) < 30 Then Exit For
            Next y
            Unit = UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + y))
            'If y <> 20 Then MsgBox "Pre'unit' ->" & unit & Chr(13) & "y->" & y
            Unit = Replace(Unit, vbLf, "")
            Unit = Replace(Unit, "(100 EA)", "")
            Unit = Replace(Unit, "(100 FT)", "")
            Unit = Replace(Unit, "(1000 EA)", "")
            Unit = Replace(Unit, "(1000 FT)", "")
            Unit = Replace(Unit, " ", "")
            
            If Unit Like "*E*E*" Then
                MsgBox "Northcoast Invoice Wescrape" & Chr(13) & "UNIT -> picked up two -E-s" & Chr(13) & "xoffset->" & xoffset & Chr(13) & "yoffset->" & yoffset & Chr(13) & Unit
                Possibleerror = Possibleerror + 1
            End If
            For re = 0 To 15
                If Left(Unit, 1) Like "[0-9]" Or Left(Unit, 1) Like "/" Or Left(Unit, 1) Like "." _
                Or Left(Unit, 1) Like "$" Or Left(Unit, 1) Like "," Then Unit = Right(Unit, Len(Unit) - 1)
            Next re
            If Unit = "FT" Then Unit = "EA"
            'MsgBox "Unit->" & unit
            'MsgBox "Unit=:" & unit & ": " & "possible errors (0/1)->" & Possibleerror
            ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
                     
                     
            'Vendor Item NO
            For y = 0 To 20
                For xadd = 0 To 1
                    vendoritemno = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset + xadd, yoffset + y)
                    If vendoritemno Like "*Our*Part*:*" Then
                        vendoritemno = Replace(vendoritemno, "Our Part#:", "")
                        vendoritemno = Replace(vendoritemno, " ", "")
                        'MsgBox ":" & vendoritemno & ":"
                        Exit For
                    End If
                Next xadd
            Next y
                     
            'QUANTITY ORDERED (NOT SHIPPED)
            'Quantity = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset)
            'Quantity = Replace(Quantity, vbLf, "")
            'Quantity = Replace(Quantity, " ", "")
            'For re = 0 To 15
            '    If UCase(Right(Quantity, 1)) Like "[A-Z]" Then Quantity = Left(Quantity, Len(Quantity) - 1)
            'Next re
            
            'condition check
            'If Quantity Like "*-*" Then
            '    For Repeat = 1 To 100
                    'MsgBox "Detected North Coast Credit->" & Quantity
            '    Next Repeat
            'End If
            'MsgBox "Quantity=:" & Quantity & ":"
            'ThisWorkbook.Sheets("Temp").Range("R2").Offset(tempsheetoffset, 0) = Quantity
            
        'QUANTITY SHIPPED (NOT Qty OPrdered)
            Shipped = ThisWorkbook.Sheets("Temp").Range("BB2").Offset(xoffset, yoffset)
            Shipped = Replace(Shipped, vbLf, "")
            Shipped = Replace(Shipped, " ", "")
            For re = 0 To 15
                If UCase(Right(Shipped, 1)) Like "[A-Z]" Then Shipped = Left(Shipped, Len(Shipped) - 1)
            Next re
            For Repeat = 1 To 100
                'MsgBox "Northcoast " & itemDesc & Chr(13) & "Shipped=:" & Shipped & ":"
            Next Repeat
            Quantity = Shipped
            ThisWorkbook.Sheets("Temp").Range("R2").Offset(tempsheetoffset, 0) = Quantity
            
        'UNIT PRICE
            For y = 0 To 20
                ' Condition 1 - normal, cost is found
                If UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + y)) Like _
                    "*[0-9].[0-9][0-9][0-9]*/[A-Z]*" Then
                    unitprice = UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + y))
                    Exit For
                End If
                
                ' Condition 2 - free item "Call For Price" in pricing line
                If UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(xoffset, yoffset + y)) Like _
                    "*CALL*FOR*PRICE*" Then
                        For Repeat = 1 To 100
                            MsgBox "On Northcoast invoice, hit line with *call for price* that slipped through error checking earlier in code"
                        Next Repeat
                    Exit For
                End If
            Next y
            
            
            If y > 10 Then
                For Repeat = 1 To 100
                    MsgBox "While webscraping North Coast doc, unable to get unit price, <Break> to investigate"
                Next Repeat
            End If
            'MsgBox y
            
            unitPriceForMessage = unitprice
            unitprice = Replace(unitprice, vbLf, "")
            unitprice = Replace(unitprice, "/", "")
            unitprice = Replace(unitprice, "(100 EA)", "")
            unitprice = Replace(unitprice, "(100 FT)", "")
            unitprice = Replace(unitprice, "(1000 EA)", "")
            unitprice = Replace(unitprice, "(1000 FT)", "")
            unitprice = Replace(unitprice, "m", "")
            unitprice = Replace(unitprice, "ea", "")
            unitprice = Replace(unitprice, "c", "")
            unitprice = Replace(unitprice, " ", "")
            unitprice = Replace(unitprice, "$", "")
            For re = 0 To 15
                If Right(UCase(unitprice), 1) Like "[A-Z]" Or Right(unitprice, 1) Like "/" Or Right(unitprice, 1) Like "." Then unitprice = Left(UCase(unitprice), Len(unitprice) - 1)
            Next re
            'MsgBox unitprice
            If unitprice = "FT" Then unitprice = "EA"
            ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = unitprice
            
            If UCase(Unit) = "C" Then
                If unitprice = "" Then unitprice = 0
                If UCase(unitprice) Like "*[A-Z]*" Then unitprice = 0
                'MsgBox unitprice
                unitprice = unitprice / 100
                Unit = "EA"
                ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
            End If
            If UCase(Unit) = "M" Then
                unitprice = unitprice / 1000
                Unit = "Ea"
                ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
            End If
            'MsgBox "UnitPrice=:" & unitprice & ":"
            If CStr(unitprice) = "" Or CStr(unitprice) = "0" Or UCase(CStr(unitprice)) Like "*[A-Z]*" Then
                For Repeat = 1 To 100
                    MsgBox "Webscraping North Coast Invoice and unable to get unit price" & Chr(13) & unitPriceForMessage
                Next Repeat
                unitprice = "0"
            End If
     
            'LINE TOTAL
            ThisWorkbook.Sheets("Temp").Range("S2").Offset(tempsheetoffset, 0) = unitprice
            
            'LinePrice = Right(Line, 10) '           <<<LinePrice
            'LinePrice = Replace(LinePrice, " ", "")
            'MsgBox unitprice
            'MsgBox Quantity
            
            If SHIP = 0 Then lineprice = unitprice * Quantity
            'MsgBox "LinePrice=:" & LinePrice & ":"
            '(fpath, fname, docno, PDFtype, InvoiceDate, InvoiceDueDate, InvoiceDiscountDate, Discount, path)
            ThisWorkbook.Sheets("Temp").Range("T2").Offset(tempsheetoffset, 0) = lineprice
            ThisWorkbook.Sheets("Temp").Range("U2").Offset(tempsheetoffset, 0) = Quantity
            ThisWorkbook.Sheets("Temp").Range("A2").Offset(tempsheetoffset, 0) = DecoPO
            ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
            ThisWorkbook.Sheets("Temp").Range("C2").Offset(tempsheetoffset, 0) = "218"
            ThisWorkbook.Sheets("Temp").Range("AH2").Offset(tempsheetoffset, 0) = SalesTax
            ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = TotalInvoice
            ThisWorkbook.Sheets("Temp").Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
            ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = VendorTotalInvoice
            ThisWorkbook.Sheets("Temp").Range("O2").Offset(tempsheetoffset, 0).NumberFormat = "@"
            ThisWorkbook.Sheets("Temp").Range("O2").Offset(tempsheetoffset, 0) = vendoritemno
            ThisWorkbook.Sheets("Temp").Range("J2").Offset(tempsheetoffset, 0) = InvoiceDate
            ThisWorkbook.Sheets("Temp").Range("AD2") = InvoiceDueDate
            ThisWorkbook.Sheets("Temp").Range("AE2") = InvoiceDiscountDate
            ThisWorkbook.Sheets("Temp").Range("AF2") = Discount
            tempsheetoffset = tempsheetoffset + 1
            
        End If
    End If
    'Next yOffset
Next xoffset

' Add Sales Tax (SalesTax)
    If SalesTax <> 0 Then
        ThisWorkbook.Sheets("Temp").Range("P2").Offset(tempsheetoffset, 0) = "Sales Tax on " & vendorInvoice
        ThisWorkbook.Sheets("Temp").Range("S2").Offset(tempsheetoffset, 0) = SalesTax
        ThisWorkbook.Sheets("Temp").Range("R2").Offset(tempsheetoffset, 0) = 1
        ThisWorkbook.Sheets("Temp").Range("A2").Offset(tempsheetoffset, 0) = DecoPO
        ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
        ThisWorkbook.Sheets("Temp").Range("C2").Offset(tempsheetoffset, 0) = "218"
        ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = TotalInvoice
        ThisWorkbook.Sheets("Temp").Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
        ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = VendorTotalInvoice
        ThisWorkbook.Sheets("Temp").Range("O2").Offset(tempsheetoffset, 0).NumberFormat = "@"
        tempsheetoffset = tempsheetoffset + 1
        
        For Repeat = 1 To 100
            'MsgBox "Found sales tax, break to investigate price has been wrong in the past"
        Next Repeat
    
    End If
    

' Add shipping and handling
    If handling <> "" And handling <> "0" Then
        ThisWorkbook.Sheets("Temp").Range("P2").Offset(tempsheetoffset, 0) = "Shipping and Handling"
        ThisWorkbook.Sheets("Temp").Range("S2").Offset(tempsheetoffset, 0) = handling
        ThisWorkbook.Sheets("Temp").Range("R2").Offset(tempsheetoffset, 0) = 1
        ThisWorkbook.Sheets("Temp").Range("A2").Offset(tempsheetoffset, 0) = DecoPO
        ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
        ThisWorkbook.Sheets("Temp").Range("C2").Offset(tempsheetoffset, 0) = "218"
        ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = TotalInvoice
        ThisWorkbook.Sheets("Temp").Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
        ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = VendorTotalInvoice
        ThisWorkbook.Sheets("Temp").Range("O2").Offset(tempsheetoffset, 0).NumberFormat = "@"
        tempsheetoffset = tempsheetoffset + 1
        
        For Repeat = 1 To 100
            MsgBox "Found shipping and hanndling, break to investigae price has been wrong in the past"
        Next Repeat
    End If

    For Repeat = 1 To 10
        'MsgBox "Done scraping Web data"
    Next Repeat
    
' error check that all lines have units and pricing
    For Repeat = 0 To 100
        If ThisWorkbook.Sheets("temp").Range("P2").Offset(Repeat, 0) = "" Then Exit For
        If ThisWorkbook.Sheets("temp").Range("S2").Offset(Repeat, 0) = "0" Or _
        ThisWorkbook.Sheets("temp").Range("S2").Offset(Repeat, 0) = "" Then
            For repeatmessage = 1 To 100
                MsgBox "While webscraping North Coast, at final error check found a line item with zero or blank as the cost"
            Next repeatmessage
        End If
    Next Repeat
    
    Call SelfHealTempPage
    'Check if PO conforms before bothering to Enter
    TargetPO = ThisWorkbook.Sheets("Temp").Range("A2")
    'MsgBox TargetPO
    Call CheckPONumber(TargetPO, Found)
        
    If ThisWorkbook.Sheets("Temp").Range("E2") Like "*/*" Then
        MsgBox "Error after scraping North Coast Web Page, Job Number looks like date on temp sheet cell E2" & Chr(13) & "Hit enter to exit and investigate"
        Exit Sub
    End If
    
' Save Product Data for future use
    If PDFtype = "Invoice" Then
            'If Not UCase(path) Like "*ATTACHMENT*" Then Call Northcoast_download_product_data(path, fname, emailmessage)
            If Not UCase(path) Like "*ATTACHMENT*" Then Exit Sub
    End If
            
     If Found = 2 Then 'TargetPO matches a subcontract number
        'rename and move file
        TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        
        pdfoption1 = "Contract " & ThisWorkbook.Sheets("Temp").Range("A2") & " " _
                     & "NorthCoast " & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
                     & Replace(TotalInvoiceAmount, "$", "") & ").pdf"
        
        Dim serverPath As String
        serverPath = "\\server2\Faxes\"
        
        If Dir(serverPath & pdfoption1) = "" Then
            Name fpath As serverPath & pdfoption1
        End If
        
        SetCursorPos 1083, 11 '--------------------------'Sage Minimize
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Application.Wait (Now + TimeValue("00:00:01"))
        
        'MsgBox "Sent TargetPO that matches Subcontract number to Fax Folder" & Chr(13) & pdfoption1
        
        Exit Sub
    End If

    For Repeat = 1 To 100
        'MsgBox "Stop and check whats going on"
    Next Repeat
    
    If Found = 1 And Possibleerror < 1 And UCase(path) Like "*ATTACHMENT*" Then
        'For Repeat = 1 To 100
            'MsgBox "Entering Order Ack"
        'Next Repeat
        xoffset = 0
        emailmessage = "NorthCoast"
        If PDFtype = "Invoice" Then
            If Dont_go_to_sage <> 1 Then Call ClickOnSage
            If Dont_go_to_sage <> 1 Then Call SageEnterINVOICEfromTEMP(xoffset, emailmessage, fpath)
            If emailmessage = "Temp Sheet Total Error" Then Exit Sub
        Else
            If Not UCase(path) Like "*ATTACHMENT*" Then Exit Sub
            If Dont_go_to_sage <> 1 Then Call ClickOnSage
            If Dont_go_to_sage <> 1 Then
                Call SageEnterPOfromTEMP(xoffset, emailmessage)
                If emailmessage = "Job entered was not valid in sage" Then
                    sourcePath = fpath
                    TargetPath = "\\server2\Dropbox\Attachments\_Re Run\" & fname
                    Call PDF_MoveToFolder(sourcePath, TargetPath, specialmessage)
                    updatelog = "Job entered was not valid in sage " & fname
                    Call logupdate(updatelog)
                    Exit Sub
                End If
            End If
        End If
        
        'rename and move file
        TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        
        If UCase(PDFtype) Like "*INV*" Then
            variable = "INV"
        Else
            variable = "ORDACK"
        End If
        
        pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
                     & variable & " " & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
                     & TotalInvoiceAmount & ").pdf"
        pdfoption1 = Replace(pdfoption1, "$", "")
        
        If Dont_go_to_sage = 1 Then Kill (fpath)
        
        If emailmessage = "Saved" And Dir(fpath) <> "" Then
            ' Construct the full path of the file
            Dim fullFilePath As String
            fullFilePath = "\\server2\Faxes\NORTH COAST - 218\" & pdfoption1
            
            ' Check if the file exists
            If Dir(fullFilePath) = "" Then
                ' File doesn't exist, try renaming it
                Name fpath As fullFilePath
                
                ' Introduce a delay to allow time for the renaming operation
                Application.Wait (Now + TimeValue("00:00:06"))
            End If
        End If
        
        If Dir(fpath) <> "" And UCase(fpath) Like "*ATTACH*" Or emailmessage Like "*already been entered*" And Dir(fpath) <> "" Then
            Kill (fpath)
        End If
        
        SetCursorPos 1083, 11 '--------------------------'Sage Minimize
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Application.Wait (Now + TimeValue("00:00:01"))
    Else
        MsgBox "Did not enter ORDACK even after second try scraping web data" & Chr(13) & "PO->" & TargetPO & Chr(13) & "PossibleErrors =" & Possibleerror _
               & Chr(13) & "Found =" & Found
        
        MsgBox "Did not enter " & TargetPO & Chr(13) & "PossibleErrors =" & Possibleerror _
               & Chr(13) & "Found =" & Found & Chr(13) & "Moving to Fax File as;" & Chr(13) _
               & TargetPO & " Northcoast"
        
        If Dir("\\server2\Faxes\" & TargetPO & "Northcoast.pdf") <> "" Then TargetPO = TargetPO & ".1"
        If Dir(fpath) <> "" And Dir("\\server2\Faxes\" & TargetPO & "Northcoast.pdf") = "" Then
            Name fpath As "\\server2\Faxes\" & TargetPO & "Northcoast.pdf"
            Application.Wait (Now + TimeValue("00:00:06"))
        End If
        
        If Dir(fpath) <> "" And UCase(path) Like "*ATTACH*" Or emailmessage Like "*already been entered*" And Dir(fpath) <> "" Then
            Kill (fpath)
        End If
    End If
End Sub


Sub Northcoast_download_product_data(path, fname, emailmessage)

'MsgBox "Made it to Northcoast_download_product_data(path)"

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
    Clipboard.Clear
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
If try = 2 Then findtext = "Specification Sheet"
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
    If fname = "" Then MsgBox "In NorthCoast module, lost fname"
    ThisWorkbook.Sheets("Submittals").Range("A1").Offset(x, 0) = fname

For Repeat = 1 To 5
    Application.SendKeys ("^w")
    Application.Wait (Now + TimeValue("00:00:01"))
Next Repeat




End Sub
