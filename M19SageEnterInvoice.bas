Attribute VB_Name = "M19SageEnterInvoice"

'Declare mouse events
Public Declare PtrSafe Function SetCursorPos Lib "User32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare PtrSafe Sub mouse_event Lib "User32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10
Declare PtrSafe Function GetSystemMetrics32 Lib "User32" _
() '-----------------------------------------------------------------
Option Compare Text

Sub SageEnterINVOICEfromTEMP(xoffset As Long, emailmessage As String, fpath As String)
    Dim sourcePath As String
    Dim fname As String
    Dim fso As Object
    Dim Repeat As Long
    Dim total_rows As Long
    Dim vendorInvoice As String
    Dim Found As Long
    Dim check_for_invoice As Long
    Dim xcolumn As Long
    Dim Vendor As String
    Dim InvoiceNO As String

    ' Set the source path
    sourcePath = fpath

    ' Create an instance of the FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Extract the file name from the source path
    fname = fso.GetFileName(sourcePath)

    ' Send original invoice data to SageXferSheet in case we need to come back to it later
    ' Update line total before checking if they add up correctly
    For Repeat = 1 To 100
        If ThisWorkbook.Sheets("Temp").Range("S2").Offset(Repeat, 0).Value <> 0 Then
            ThisWorkbook.Sheets("Temp").Range("T2").Offset(Repeat, 0).Value = _
                ThisWorkbook.Sheets("Temp").Range("R2").Offset(Repeat, 0).Value * _
                ThisWorkbook.Sheets("Temp").Range("S2").Offset(Repeat, 0).Value
        End If
    Next Repeat
    
    Call Move_data_to_Sage_Xfer_Sheet

    ' Close Chrome by hot keys if it's open
    Application.SendKeys ("^w")

    GoTo skip_product_link_save

    ' Count how many lines are to be transferred
    For total_rows = 0 To 100
        If ThisWorkbook.Sheets("Temp").Range("A2").Offset(total_rows, 0) = "" Then Exit For
    Next total_rows

    ' Check if this invoice has already been added to the products sheet
    vendorInvoice = ThisWorkbook.Sheets("Temp").Range("H2").Value
    If vendorInvoice <> "" Then
        Found = 0
        For check_for_invoice = 1 To 100000
            If vendorInvoice = Workbooks("Product_links.xlsx").Sheets("DECOInvoiceHistory").Range("H2").Offset(check_for_invoice, 0).Value Then
                Found = 1
                Exit For
            End If
        Next check_for_invoice
    
        If Found = 0 Then
            ' Insert new rows on Product links sheet so that newest stuff is always on top
            For Repeat = 1 To (total_rows + 2)
                Workbooks("Product_links.xlsx").Sheets("DECOInvoiceHistory").Range("A1").EntireRow.Insert Shift:=xlDown
                Application.Wait Now + TimeValue("00:00:01")
            Next Repeat
        
            ' Write new data to Product links sheet
            For Repeat = 0 To (total_rows + 1)
                With Workbooks("Product_links.xlsx").Sheets("DECOInvoiceHistory")
                    .Range("O1").Offset(Repeat, 0).NumberFormat = "@"
                    For xcolumn = 0 To 60
                        .Range("A1").Offset(Repeat, xcolumn).Value = _
                            ThisWorkbook.Sheets("Temp").Range("A1").Offset(Repeat, xcolumn).Value
                    Next xcolumn
                End With
            Next Repeat
        End If
    End If

    If workbook_open_status = 1 Then
        Workbooks("Product_links.xlsx").Save
        Workbooks("Product_links.xlsx").Close
    End If

skip_product_link_save:

Thebeggining:
    Vendor = ThisWorkbook.Sheets("Temp").Range("C2").Value
    Sleep 250

    ' Click Through Main System Menu
    ' Sage, Click 4 Accounts Payable
    SetCursorPos 20, 234
    Sleep 250
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo

    ' Sage, Click 2 Payable invoices
    SetCursorPos 90, 274
    Sleep 250
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Application.SendKeys "~", True
    Sleep 250
    Application.SendKeys "~", True
    
    ' Wait for report window to generate, then maximize
    Application.Wait Now + TimeValue("00:00:10")
    Application.SendKeys "%{ }", True
    Sleep 250
    Application.SendKeys "x", True
    Sleep 250

    ' Sage INVOICE WINDOW click into Vendor Invoice Number
    SetCursorPos 197, 93
    Sleep 2500
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    
    If ThisWorkbook.Sheets("Temp").Range("H2").Offset(xoffset, 0).Value = "" Then
        MsgBox "Freeze // there is no Invoice number on the temp sheet to input into Sage! Hit enter to see the PDF"
        If Dir(fpath) <> "" Then ThisWorkbook.FollowHyperlink Address:=fpath
        MsgBox "freeze, break to exit and enter Doc No manually, PDF " & fpath & " will be killed if you hit enter and break after"
        Kill Dir(fpath)
        MsgBox "PDF killed, break now"
        Exit Sub
    End If
    Sleep 250

    ' Enter Invoice Number
    InvoiceNO = ThisWorkbook.Sheets("Temp").Range("H2").Offset(xoffset, 0).Value
    
    ' Last check that NorthCoast invoice number conforms
    If UCase(ThisWorkbook.Sheets("Temp").Range("C2").Value) Like "*NORTH*COAST*" And _
        Not InvoiceNO Like "S[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].[0-9][0-9][0-9]" Then
        For Repeat = 1 To 100
            MsgBox "About to Enter North Coast Invoice but getting odd Invoice Number -> " & InvoiceNO
        Next Repeat
    End If
    
    ' Last check that Stoneway invoice number conforms
    If UCase(ThisWorkbook.Sheets("Temp").Range("C2").Value) Like "*STONE*WAY*" And _
        Not InvoiceNO Like "S[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].[0-9][0-9][0-9]" Then
        For Repeat = 1 To 100
            MsgBox "About to Enter Stoneway Invoice but getting odd Invoice Number -> " & InvoiceNO
        Next Repeat
    End If
    
    ' Last check that Platt invoice number conforms
    If UCase(ThisWorkbook.Sheets("Temp").Range("C2").Value) Like "*PLATT*" And _
        Not InvoiceNO Like "[0-9][A-Z][0-9][0-9][0-9][0-9][0-9]" Then
        For Repeat = 1 To 100
            MsgBox "About to Enter Platt Invoice but getting odd Invoice Number -> " & InvoiceNO
        Next Repeat
    End If
        
    ' Adjust for Wesco Invoices that start with 0
    If UCase(ThisWorkbook.Sheets("Temp").Range("C2").Value) Like "*WESCO*" And _
        InvoiceNO Like "*[A-Z]*" Then
        For Repeat = 1 To 100
            MsgBox "About to Enter Wesco Invoice but getting odd Invoice Number -> " & InvoiceNO
        Next Repeat
    End If
    If UCase(ThisWorkbook.Sheets("Temp").Range("C2").Value) Like "*WESCO*" Then
        If Len(InvoiceNO) < 6 Then
            Application.SendKeys "0"
        End If
    End If

    ' Enter the Invoice Number
    Application.SendKeys ThisWorkbook.Sheets("Temp").Range("H2").Offset(xoffset, 0).Value
    Application.Wait Now + TimeValue("00:00:03")
    
    ' Hit F9
    Application.SendKeys "{F9}"
    Application.Wait Now + TimeValue("00:00:03")
    ThisWorkbook.Sheets("Temp").Range("A1").Value = ""

    ' Sage INVOICE WINDOW click into Grid to see if INVOICE already Exists
    SetCursorPos 275, 307
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Sleep 250
    Application.CutCopyMode = False
    Sleep 500
    For Repeat = 1 To 2
        Application.SendKeys "^c", True
        Application.Wait Now + TimeValue("00:00:01")
    Next Repeat
    Application.Wait Now + TimeValue("00:00:01")
    ThisWorkbook.Sheets("Temp").Range("A1").Value = ""
    ThisWorkbook.Sheets("Temp").Paste Destination:=ThisWorkbook.Sheets("Temp").Range("A1")
    Application.Wait Now + TimeValue("00:00:01")
    
' Check if Invoice is already entered
    If ThisWorkbook.Sheets("Temp").Range("A1") <> "" Then
        For x = 0 To 200
                If ThisWorkbook.Sheets("Temp").Range("H2").Offset(xoffset + x, 0) = ThisWorkbook.Sheets("Temp").Range("H2").Offset(xoffset, 0) _
                        Then ThisWorkbook.Sheets("Temp").Range("AC2").Offset(xoffset + x, 0) = "Yes"
        Next x
        emailmessage = "Saved"
        Application.SendKeys "%{F4}", True
        Application.Wait (Now + TimeValue("00:00:01"))
        'Application.SendKeys "{Tab}", True
        Sleep (250)
        'Application.SendKeys "~", True
            ' Check if the file exists before trying to delete
        If Dir(fpath) <> "" Then
            Kill fpath
            Application.Wait (Now + TimeValue("00:00:06"))
        End If
        Exit Sub
    End If

' ENTER Purchase Order Number / Vendor / Description
    SetCursorPos 210, 120
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    If ThisWorkbook.Sheets("Temp").Range("A2").Offset(xoffset, 0) Like "*Invoice*" Then
        'enter the other fields that - when there's a PO - are pre-populated
        Application.SendKeys "~"
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.SendKeys "~"
        Application.Wait (Now + TimeValue("00:00:01"))
        'Vendor Name
        Application.SendKeys ThisWorkbook.Sheets("Temp").Range("C2"), True
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.SendKeys "~"
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.SendKeys "~"
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.SendKeys "~"
        Application.Wait (Now + TimeValue("00:00:01"))
        'Description Field
        Application.SendKeys ThisWorkbook.Sheets("Temp").Range("D2"), True
        Application.Wait (Now + TimeValue("00:00:03"))
        Application.SendKeys "~"
        Application.Wait (Now + TimeValue("00:00:01"))
        GoTo skipPurchaseOrder:
    End If
    Application.SendKeys ThisWorkbook.Sheets("Temp").Range("A2"), True
    Sleep (250)
    Application.SendKeys "~"
    Application.Wait (Now + TimeValue("00:00:02"))

' Copy->Paste existing PURCHASE ORDER INFO field to BA2
    ThisWorkbook.Sheets("Temp").Range("BA1:BN500").Clear
    SetCursorPos 275, 307
    Application.Wait (Now + TimeValue("00:00:01"))
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    For x = 1 To 7
        Application.SendKeys "+{Right}", True
    Next x
    For x = 1 To 100
        Application.SendKeys "+{Down}", True
    Next x
    Application.CutCopyMode = False
    Application.Wait (Now + TimeValue("00:00:02"))
    'For Repeat = 1 To 2
        Application.SendKeys "^c", True
        Application.Wait (Now + TimeValue("00:00:02"))
    'Next Repeat
    ThisWorkbook.Sheets("Temp").Paste Destination:=ThisWorkbook.Sheets("Temp").Range("BA2")
    Application.Wait (Now + TimeValue("00:00:01"))
    ThisWorkbook.Sheets("Temp").Range("BA1") = "PO Desc"
    ThisWorkbook.Sheets("Temp").Range("BD1") = "PO Qty"
    ThisWorkbook.Sheets("Temp").Range("BE1") = "PO Price"
    ThisWorkbook.Sheets("Temp").Range("BF1") = "PO Total"
    ' Error Check
    ' If ThisWorkbook.Sheets("Temp").Range("BA2") = "" Then MsgBox "Error when pasting -> No data xferred to cell BA2"

' If No Purchase Order Entered
    If ThisWorkbook.Sheets("Temp").Range("BA2") = "" Then
        Application.SendKeys "%{F4}", True
        Sleep (250)
        Application.SendKeys "{Tab}", True
        Sleep (250)
        Application.SendKeys "~", True
        Sleep (250)
        SetCursorPos 20, 234 '---Sage, Click 4 Roll Up Accounts Payable
        Sleep (250)
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        
        Call SageEnterPOfromTEMP(xoffset, emailmessage)
        Application.Wait (Now + TimeValue("00:00:01"))
        If emailmessage = "Job entered was not valid in sage" Then
            sourcePath = fpath
            TargetPath = "\\server2\Dropbox\Attachments\_Re Run\" & fname
            Call PDF_MoveToFolder(sourcePath, TargetPath, specialmessage)
            updatelog = "Job entered was not valid in sage " & fname
            Call logupdate(updatelog)
            Exit Sub
        End If
        
        GoTo Thebeggining:
    End If

'PROBLEM Check // Purchase order exists but is closed
    If ThisWorkbook.Sheets("Temp").Range("BA2") Like "*---------------*" Then
        If ThisWorkbook.Sheets("Temp").Range("BA5") Like "*closed or not found*" Then emailmessage = "closedPO"
        'MsgBox "freeze"
        Application.SendKeys "~"
        Sleep (250)
        Application.SendKeys "%{F4}"
        Sleep (250)
        Application.SendKeys "{Tab}"
        Sleep (250)
        Application.SendKeys "~"
        Sleep (1000)
        For x = 0 To 200
            With ThisWorkbook.Sheets("Temp")
                If .Range("H2").Offset(xoffset + x, 0) = .Range("H2").Offset(xoffset, 0) _
                        Then .Range("AC2").Offset(xoffset + x, 0) = "Yes"
            End With
        Next x
        ' Found that Invoice number was not entered, but related PO was closed or doesn't exist could not enter
        ' Sage, Click 4 Roll Up Accounts Payable
        SetCursorPos 20, 234
        Sleep (250)
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Call SageEnterPOfromTEMP(xoffset, emailmessage)
        If emailmessage = "Job entered was not valid in sage" Then
            sourcePath = fpath
            TargetPath = "\\server2\Dropbox\Attachments\_Re Run\" & fname
            Call PDF_MoveToFolder(sourcePath, TargetPath, specialmessage)
            updatelog = "Job entered was not valid in sage " & fname
            Call logupdate(updatelog)
            Exit Sub
        End If
        ' Error Checks Discount
        GoTo Thebeggining:
        Exit Sub
    End If


GoTo skip_exit_invoice:
    If ThisWorkbook.Sheets("Temp").Range("H2") Like "S[0-9][0-9][0-9][0-9][0-9][0-9]*" Then
        emailmessage = "saved" 'So program will route PDF to Fax File
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.SendKeys "%{F4}"
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.SendKeys "{Tab}"
        Sleep 250
        Application.SendKeys "~"
        Sleep 250
    
        SetCursorPos 20, 234 '--------------------------Sage, Click 4 Accounts Payable
        Application.Wait (Now + TimeValue("00:00:01"))
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Application.Wait (Now + TimeValue("00:00:01"))
        Exit Sub
    End If
skip_exit_invoice:


' Check if Invoice total exact matches SAGE total
    SetCursorPos 1134, 677
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Application.Wait (Now + TimeValue("00:00:01"))
    For x = 1 To 10
        Application.SendKeys "+{right}"
    Next x
    For Repeat = 1 To 3
        Application.SendKeys "^c", True
        Sleep 250
    Next Repeat
    
    Application.Wait (Now + TimeValue("00:00:01"))
    ThisWorkbook.Sheets("Temp").Paste Destination:=Sheets("Temp").Range("A1")
    invoiceTotal = ThisWorkbook.Sheets("Temp").Range("N2")
    invoiceTotal = Replace(invoiceTotal, "$", "")
    
    'Message to troubleshoot system
    ifmatch = "Doesn't Match"
    If invoiceTotal Like ThisWorkbook.Sheets("Temp").Range("A1") And invoiceTotal <> "" Then ifmatch = "Matches"
    'MsgBox "Just checked Invoice Total" & Chr(13) & "PDF Invoice Total->" & Invoicetotal & Chr(13) & "Sage Total ->" & ThisWorkbook.Sheets("Temp").Range("A1") & Chr(13) & ifmatch
       
    Password = ""
    If invoiceTotal Like ThisWorkbook.Sheets("Temp").Range("A1") And invoiceTotal <> "" Then
        For x = 0 To 200
            With ThisWorkbook.Sheets("Temp")
                If .Range("H2").Offset(xoffset + x, 0) = .Range("H2").Offset(xoffset, 0) _
                        Then .Range("AC2").Offset(xoffset + x, 0) = "Yes"
            End With
        Next x
        Password = "Invoice Total Match"
        'MsgBox "Sending to invoice match"
        GoTo skipPurchaseOrder:
        'skipPurchaseOrder:
        'InvoiceTotalMatch
    End If


' Remove Quotes from item descriptions first
    For x = 0 To 100
        Line = ThisWorkbook.Sheets("Temp").Range("P2").Offset(x, 0)
        If Line = "" Then Exit For
        Line = Replace(Line, Chr(34), "") 'Quote Marks
        Line = Replace(Line, "(", "")
        Line = Replace(Line, ")", "")
        ThisWorkbook.Sheets("Temp").Range("P2").Offset(x, 0) = Line
    Next x

'MsgBox "Freeze before checking if items on PDF exceed items in Sage"

' Error check line items sum to invoice total 1
    invoiceTotal = 0
    For Repeat = 0 To 100
        invoiceTotal = invoiceTotal + ThisWorkbook.Sheets("Temp").Range("T2").Offset(Repeat, 0).Value
    Next Repeat
    Dim tolerance As Double
    tolerance = 0.03
    If Abs(invoiceTotal - ThisWorkbook.Sheets("Temp").Range("N2").Value) > tolerance Then
        For Repeat = 1 To 100
            MsgBox "WARNING1: Temp sheet sum of items is not within the acceptable tolerance of the total invoice.", vbExclamation, "Error"
        Next Repeat
    End If


' Check PO and add any missing items
Total_Found = 0
rewritePO = 0
POadder = ""
reasonstring = ""
For Invoiceitem = 0 To 100
    If ThisWorkbook.Sheets("Temp").Range("P2").Offset(Invoiceitem, 0) = "" Then Exit For

    Dim InvoiceDesc As String
    Dim InvoiceQty As Double
    Dim InvoicePrice As Variant
    
    InvoiceDesc = Trim(Left(ThisWorkbook.Sheets("Temp").Range("P2").Offset(Invoiceitem, 0), 25))
    InvoiceQty = CDbl(ThisWorkbook.Sheets("Temp").Range("R2").Offset(Invoiceitem, 0))
    InvoicePrice = ThisWorkbook.Sheets("Temp").Range("S2").Offset(Invoiceitem, 0)
    StringToClean = InvoiceDesc
    Call CleanString(StringToClean)
    If StringToClean = "" Then
        For Repeat = 1 To 50
            MsgBox "Error, sent to cleanString module but stringtoclean is empty"
        Next Repeat
        Exit Sub
    End If
    CleanedInvoiceDesc = StringToClean
    
    Found = 0
    For POitem = 0 To 100
        If ThisWorkbook.Sheets("Temp").Range("BA2").Offset(POitem, 0) = "" Then Exit For
        
        Dim PODesc As String
        Dim POQty As Double
        Dim POPrice As Variant
        
        PODesc = Trim(Left(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(POitem, 0), 25))
        POQty = CDbl(ThisWorkbook.Sheets("Temp").Range("BD2").Offset(POitem, 0))
        POPrice = ThisWorkbook.Sheets("Temp").Range("BE2").Offset(POitem, 0)
        StringToClean = PODesc
        Call CleanString(StringToClean)
        If StringToClean = "" Then
            For Repeat = 1 To 50
                MsgBox "Error, sent to cleanString module but stringtoclean is empty"
            Next Repeat
        Exit Sub
        End If
        CleanedPODesc = StringToClean
        
        Sleep 5
        
        ' Debug
        'MsgBox "Comparing:" & Chr(13) & ":" & CleanedInvoiceDesc & ":" & Chr(13) & ":" & CleanedPODesc & ":"
        
        If CleanedInvoiceDesc Like "*" & CleanedPODesc & "*" Or _
            CleanedPODesc Like "*" & CleanedInvoiceDesc & "*" Then
            
            'debug
            'MsgBox "MATCHED"
            If Round(InvoicePrice, 2) <> Round(POPrice, 2) Then
                For Repeat = 1 To 100
                    'MsgBox "Matched items->" & CleanedInvoiceDesc & Chr(13) & "But Pricing doesn't match->" & Round(InvoicePrice, 6)
                Next Repeat
            End If
                        
            If Round(InvoicePrice, 6) = Round(POPrice, 6) Then
                Found = 1
                If InvoiceQty > POQty Then
                    ThisWorkbook.Sheets("Temp").Range("BD2").Offset(POitem, 0) = InvoiceQty
                    If POPrice = "" Then ThisWorkbook.Sheets("Temp").Range("BE2").Offset(POitem, 0) = InvoicePrice
                    rewritePO = 1
                    reasonstring = "Because item was already on PO but, there was more qty invoiced than was present on the PO"
                End If
                Exit For
            End If
        End If
    Next POitem

    'If item was not found, append it to the existing invoice
    If Found = 0 Then
        'MsgBox "item on Invoice was not found on existing Sage PO" & Chr(13) & "InvoiceDesc->" & _
            InvoiceDesc & Chr(13) & "CleanedInvoiceDesc->" & CleanedInvoiceDesc & Chr(13) & _
            "Could it be invoice price mismatch?"
        ' Find end of row at column PA2
        For last_row = 0 To 100
            If ThisWorkbook.Sheets("Temp").Range("BA2").Offset(last_row, 0) = "" Then Exit For
        Next last_row
        
        With ThisWorkbook.Sheets("Temp")
            ' Description
            .Range("BA2").Offset(last_row, 0).Value = .Range("P2").Offset(Invoiceitem, 0)
            ' Unit
            .Range("BC2").Offset(last_row, 0).Value = "EA"
            ' Qty
            .Range("BD2").Offset(last_row, 0).Value = .Range("R2").Offset(Invoiceitem, 0)
            ' Price
            .Range("BE2").Offset(last_row, 0).Value = .Range("S2").Offset(Invoiceitem, 0)
        End With
        rewritePO = 1
    End If
Next Invoiceitem


' Error check line items sum to invoice total 2
invoiceTotal = 0
For Repeat = 0 To 100
    invoiceTotal = invoiceTotal + ThisWorkbook.Sheets("Temp").Range("T2").Offset(Repeat, 0).Value
Next Repeat
tolerance = 0.03
If Abs(invoiceTotal - ThisWorkbook.Sheets("Temp").Range("N2").Value) > tolerance Then
    For Repeat = 1 To 100
        MsgBox "WARNING2: Temp sheet sum of items is not within the acceptable tolerance of the total invoice.", vbExclamation, "Error"
    Next Repeat
End If

'If there were missing items or modded items in the previose loop, prepare to rewrite the PO in the following loop
If Total_Found > 0 Or rewritePO = 1 Then
    'clear the current P2:U100 field
    ThisWorkbook.Sheets("Temp").Range("P2:U100").Clear
    For Repeat = 0 To 100
        If ThisWorkbook.Sheets("Temp").Range("BA2").Offset(Repeat, 0) = "" Then Exit For
        ' write header data
        For y = 0 To 15
            ThisWorkbook.Sheets("temp").Range("A2").Offset(Repeat, y) = ThisWorkbook.Sheets("Temp").Range("A2").Offset(0, y)
        Next y
        ' Write Description
        ThisWorkbook.Sheets("Temp").Range("P2").Offset(Repeat, 0) = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(Repeat, 0)
        ' Write Unit
        ThisWorkbook.Sheets("Temp").Range("Q2").Offset(Repeat, 0) = "EA"
        ' Write Qty
        ThisWorkbook.Sheets("Temp").Range("R2").Offset(Repeat, 0) = ThisWorkbook.Sheets("Temp").Range("BD2").Offset(Repeat, 0)
        ' Write Price
        ThisWorkbook.Sheets("Temp").Range("S2").Offset(Repeat, 0) = ThisWorkbook.Sheets("Temp").Range("BE2").Offset(Repeat, 0)
        ' WRite Total
        ThisWorkbook.Sheets("Temp").Range("T2").Offset(Repeat, 0) = _
        ThisWorkbook.Sheets("Temp").Range("R2").Offset(Repeat, 0).Value * ThisWorkbook.Sheets("Temp").Range("S2").Offset(Repeat, 0).Value

    Next Repeat
    
    'Error Check 1
    If ThisWorkbook.Sheets("Temp").Range("P2") = "" Then
        MsgBox "Error copying->Pasting data from sage onto the temp worksheet. Blank cell at P2"
    End If
    
    Application.SendKeys "%{F4}", True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "{tab}", True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "~", True
    Sleep (250)
    
    SetCursorPos 20, 234 '--------------------------Sage, Click 4 Accounts Payable
    Sleep (250)
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    
    'go add items to PO
    If emailmessage = "PO Overwritten" Then MsgBox "Error-> Already overwrote PO but Invoice total STILL doesn't match"
    emailmessage = "Overwrite PO"
    Call SelfHealTempPage
    For Repeat = 1 To 100
        'MsgBox "About to rewrite PO, check that quantities are assigned correctly" & Chr(13) & reasonstring
    Next Repeat
    Call SageEnterPOfromTEMP(xoffset, emailmessage)
    If emailmessage = "Job entered was not valid in sage" Then
        sourcePath = fpath
        TargetPath = "\\server2\Dropbox\Attachments\_Re Run\" & fname
        Call PDF_MoveToFolder(sourcePath, TargetPath, specialmessage)
        updatelog = "Job entered was not valid in sage " & fname
        Call logupdate(updatelog)
        Exit Sub
    End If
    emailmessage = "NotSaved"
    
    'Restore the original invoice
    ThisWorkbook.Sheets("Temp").Range("A2:Z100") = ""
    Call Move_data_to_Sage_Temp_Sheet
    
    ' Error check line items sum to invoice total 4
    invoiceTotal = 0
    For Repeat = 0 To 100
        invoiceTotal = invoiceTotal + ThisWorkbook.Sheets("Temp").Range("T2").Offset(Repeat, 0).Value
    Next Repeat
    tolerance = 0.03
    If Abs(invoiceTotal - ThisWorkbook.Sheets("Temp").Range("N2").Value) > tolerance Then
        For Repeat = 1 To 100
            MsgBox "WARNING4: Temp sheet sum of items is not within the acceptable tolerance of the total invoice.", vbExclamation, "Error"
        Next Repeat
    End If
    GoTo Thebeggining:
End If

' Error check line items sum to invoice total 3
invoiceTotal = 0
For Repeat = 0 To 100
    invoiceTotal = invoiceTotal + ThisWorkbook.Sheets("Temp").Range("T2").Offset(Repeat, 0).Value
Next Repeat
tolerance = 0.02
If Abs(invoiceTotal - ThisWorkbook.Sheets("Temp").Range("N2").Value) > tolerance Then
    MsgBox "WARNING3: Temp sheet sum of items is not within the acceptable tolerance of the total invoice.", vbExclamation, "Error"
End If


skipPurchaseOrder:

' Sage INVOICE WINDOW Invoice Date
    SetCursorPos 612, 95
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Sleep (250)
    
    InvoiceDate = ThisWorkbook.Sheets("Temp").Range("J2")
    If InvoiceDate = "" Then InvoiceDate = Now
    InvoiceDate = Format(InvoiceDate, "MM/DD/YYYY")
    Application.SendKeys InvoiceDate, True
    Application.SendKeys "~", True
    Sleep (250)

' Set intial Due Date net 30
    DueDate = ThisWorkbook.Sheets("Temp").Range("AD2")
    If DueDate = "" Then
    ' if not available, set intial invoice date as "now"
        InvoiceDate = ThisWorkbook.Sheets("Temp").Range("J2")
        If InvoiceDate = "" Then InvoiceDate = Now
        InvoiceDate = Format(InvoiceDate, "MM/DD/YYYY")
        ' Set intial due date 1 month from "now"
        DueDate = DateAdd("M", 1, InvoiceDate)
        DueDate = Format(DueDate, "MM/DD/YYYY")
    End If
    
' Stoneway, NorthCast or Platt DueDate Modifications
If Vendor Like "*263*" Or Vendor Like "*218*" Or Vendor Like "*234*" Then 'due on the 25th
    DueDate = ThisWorkbook.Sheets("Temp").Range("AD2")
    DueDate = Replace(DueDate, ".", "")
    DueDate = Replace(DueDate, " ", "")
    If DueDate = "" Then
        ' error check
        If InvoiceDate = "" Then
            For Repeat = 1 To 100
                MsgBox "Error: When setting due date, invoicedate is empty"
            Next Repeat
        End If
        
        ' Check if the invoice date is between the 25th and the end of the month
        If Day(InvoiceDate) >= 25 Then
            ' Set the due date to the 25th of the next month
            DueDate = DateSerial(Year(InvoiceDate), Month(InvoiceDate) + 2, 25)
        Else
            ' Set the due date to the 25th of the current month
            DueDate = DateSerial(Year(InvoiceDate), Month(InvoiceDate) + 1, 25)
        End If
        
        ' Adjust the year if the due date month is January
        'If Month(DueDate) = 1 And Month(InvoiceDate) <> 12 Then
        '    DueDate = DateSerial(Year(InvoiceDate) + 1, 1, 25)
        'End If

    End If
End If
   
' Send Due Date, Hit Enter
    Application.SendKeys DueDate, True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "~", True
    Application.Wait (Now + TimeValue("00:00:01"))
    ' Set Vendor Sage Number
    Vendor = ThisWorkbook.Sheets("Temp").Range("C2")

' Set initial Discount Date
    DiscountDate = Now
    DiscountDate = Format("MM/DD/YYYY")

' Stoneway, NorthCast, Platt, WESCO DiscountDate Modifier
    If Vendor Like "*263*" Or Vendor Like "*218*" Or Vendor Like "*234*" Or Vendor Like "*430*" Then 'Discount on the 10th, due on the 25th
        ' error check
        If InvoiceDate = "" Then
            For Repeat = 1 To 100
                MsgBox "Error: When setting discount date, invoicedate is empty"
            Next Repeat
        End If
        
        ' Get the Discount date from the "Temp" sheet
        DiscountDate = ThisWorkbook.Sheets("Temp").Range("AE2")
        
        ' Check if the discount date in cell AE2 is a valid date
        If DiscountDate = "" Then
  
            ' If no valid discount date in cell AE2, extrapolate the discount date
            If Day(InvoiceDate) >= 25 Then
                
                ' Set the discount date to the 10th of the month, at least 30 days out from the invoice date
                DiscountDate = DateSerial(Year(InvoiceDate), Month(InvoiceDate) + 2, 10)
 
            Else
                DiscountDate = DateSerial(Year(InvoiceDate), Month(InvoiceDate) + 1, 10)
            End If
              
            'If Month(DiscountDate) > 12 Then
            '    DiscountDate = DateSerial(Year(Date) + 1, 1, 10)
            'End If
        
        End If
    End If

    
'Error check discount date
    If UCase(DiscountDate) Like "*[A-Z]*" Then
        For Repeat = 1 To 10
            MsgBox "ERROR discount date has letters in it!"
        Next Repeat
    End If
    
'Enter Discount Date, hit Enter
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys (DiscountDate)
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "~", True
    Application.Wait (Now + TimeValue("00:00:01"))
' Check for Invoice EXACT MATCH
If Password = "Invoice Total Match" Then GoTo InvoiceTotalMatch:
    

'Continue cross checks
skipaddingitemstoPO:

    For x = 0 To 100
        If ThisWorkbook.Sheets("Temp").Range("N2") = "" Then
            ThisWorkbook.Sheets("Temp").Range("A2:AK2").Offset(x, 0) = ""
        End If
    Next x

skipInvoicePOMatch:

' Sage INVOICE WINDOW click into Grid and start entering Data
    SetCursorPos 275, 307
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Sleep (250)
    
' Begin Loop entering Invoice Items
    TotalPOItems = 0
    
LoopInvoiceStart:

'Error check
    If TotalPOItems = 0 And ThisWorkbook.Sheets("Temp").Range("A2") Like "*Invoice*" Then
        'MsgBox "Ignoring there is nothing in the first field so that the invoice will be entered"
        GoTo skipInputCheck:
    End If
    
'copy item description from Sage
    Application.SendKeys "^c", True
    Application.Wait (Now + TimeValue("00:00:01"))
    
' Paste item Desc into cell A1
    ThisWorkbook.Sheets("Temp").Paste Destination:=Sheets("Temp").Range("A1")
    Application.Wait (Now + TimeValue("00:00:01"))

' If turns out cell A1 is blank, then assume no more line items
If ThisWorkbook.Sheets("Temp").Range("A1") = "" Then GoTo AllLineItemsInput:


skipInputCheck:


TotalPOItems = TotalPOItems + 1
    For x = 0 To 200
        ' find item in sage on PDF (on temp sheet)
        If Not ThisWorkbook.Sheets("Temp").Range("A2") Like "*Invoice*" Then
            'copy item description in sage field
            Application.SendKeys "^c", True
            Application.Wait (Now + TimeValue("00:00:01"))
            ThisWorkbook.Sheets("Temp").Paste Destination:=Sheets("Temp").Range("A1")
            Application.Wait (Now + TimeValue("00:00:01"))
            
            If ThisWorkbook.Sheets("Temp").Range("A1") = "" Then Exit For
                SageDesc = ""
                CleanedSageDesc = ""
                InvoiceDesc = ""
                CleanedInvoiceDesc = ""
                Found = 0
                For b = 0 To 100
                    SageDesc = Left(ThisWorkbook.Sheets("Temp").Range("A1").Offset(0, 0), 25)
                    InvoiceDesc = Left(ThisWorkbook.Sheets("Temp").Range("P2").Offset(b, 0), 25)
                    If InvoiceDesc = "" Then Exit For
                    
                    'Get Sage desc
                    StringToClean = SageDesc
                    Call CleanString(StringToClean)
                    If StringToClean = "" Then
                        For Repeat = 1 To 50
                            MsgBox "Error, sent to cleanString module but stringtoclean is empty"
                        Next Repeat
                        Exit Sub
                    End If
                    CleanedSageDesc = StringToClean
                    
                    'Get Invoice Desc
                    StringToClean = InvoiceDesc
                    Call CleanString(StringToClean)
                    If StringToClean = "" Then
                        For Repeat = 1 To 50
                            MsgBox "Error, sent to cleanString module but stringtoclean is empty"
                        Next Repeat
                        Exit Sub
                    End If
                    CleanedInvoiceDesc = StringToClean
                    
                    'Compare
                    If CleanedSageDesc Like "*" & CleanedInvoiceDesc & "*" Or CleanedInvoiceDesc Like "*" & CleanedSageDesc & "*" Then
                        If ThisWorkbook.Sheets("Temp").Range("O2").Offset(b, 0) <> "X" Then
                            Found = 1
                            Exit For
                        End If
                    End If
                    
                Next b
                'If Found = 0 Then MsgBox "ERROR, confirmed invoice item existed intitally but failed to find it when" & _
                "entering quantities" & Chr(13) & ThisWorkbook.Sheets("Temp").Range("A1")
        
            'Mark item as entered into sage to prevent entering an item twice
            ThisWorkbook.Sheets("Temp").Range("O2").Offset(b, 0) = "X"
            
        Else
            'Enter Description when 'Invoice Only'
            If ThisWorkbook.Sheets("Temp").Range("P2").Offset(x, 0) = "" Then Exit For
            Application.SendKeys ThisWorkbook.Sheets("Temp").Range("P2").Offset(x, 0), True
            Application.Wait (Now + TimeValue("00:00:04"))
            Z = 0
        End If
        
            
        'This Item is on the Invoice, enter Data
        
        If Found = 1 Then
            'Unit
            Application.SendKeys "~", True
            Sleep (250)
            Application.SendKeys "~", True
            Application.Wait (Now + TimeValue("00:00:01"))
            
            'Quantity (or shipped)
            Shipped = ThisWorkbook.Sheets("Temp").Range("U2").Offset(b, 0)
            If Shipped = "" Then
                Quan = ThisWorkbook.Sheets("Temp").Range("R2").Offset(b, 0)
            Else
                Quan = Shipped
            End If
            If Quan = "" Then Quan = "0"
            Application.SendKeys Quan, True
            Sleep (250)
            Application.SendKeys "~", True
            Application.Wait (Now + TimeValue("00:00:01"))
            
            'Price
            pricing = ThisWorkbook.Sheets("Temp").Range("S2").Offset(b, 0)
            If ThisWorkbook.Sheets("Temp").Range("Q2").Offset(Z, 0) = "C" Then pricing = pricing / 100
            If ThisWorkbook.Sheets("Temp").Range("Q2").Offset(Z, 0) = "M" Then pricing = pricing / 1000
            Application.SendKeys pricing, True
            Sleep (250)
            Application.SendKeys "~", True
            Sleep (250)
            
            'Account
            If ThisWorkbook.Sheets("Temp").Range("A2") Like "*Invoice*" Then
                Application.SendKeys ThisWorkbook.Sheets("Temp").Range("Z2"), True
            'do this
            Else
                If ThisWorkbook.Sheets("Temp").Range("AH2").Offset(b, 0) = "Y" Then
                    Application.SendKeys "5003 - COGS- Equipment", True
                    Else
                    Application.SendKeys "5001 - COGS-Material", True
                End If
            End If
            
            Application.SendKeys "~", True
            Application.Wait (Now + TimeValue("00:00:02"))
            
            
            Application.SendKeys "~", True '--------------------Notes
            Application.Wait (Now + TimeValue("00:00:01"))
            
            Application.SendKeys "~", True '--------------------NextLine "Part#"
            Sleep (250)
            
            
            ThisWorkbook.Sheets("Temp").Range("AC2").Offset(Z, 0) = "Yes" 'Update "Temp Sheet" with "Entered in Sage"
        Else
            'MsgBox "This item is not on the invoice but is on the PO, enter zero quantities"
            'Unit
            Application.SendKeys "~", True
            Sleep (500)
            Application.SendKeys "~", True
            Application.Wait (Now + TimeValue("00:00:01"))
            
            'Quantity
            Quan = 0
            Application.SendKeys Quan, True
            Sleep (500)
            Application.SendKeys "~", True
            Application.Wait (Now + TimeValue("00:00:01"))
            
            For Repeat = 1 To 5
                Application.SendKeys "~", True
                Application.Wait (Now + TimeValue("00:00:02"))
            Next Repeat
            'MsgBox "ERROR, Found item on PDF invoice that is not on Purchase Order" & Chr(13) & ":" & ThisWorkbook.Sheets("Temp").Range("A1") & ":"
        End If
    Next x
    
GoTo LoopInvoiceStart:



AllLineItemsInput:

' Check for Tax
vendorInvoice = ThisWorkbook.Sheets("Temp").Range("N2")
'When entering an invoice: Tax =
VendorInvoiceTax = ThisWorkbook.Sheets("Temp").Range("AH2")


' Enter Tax Info
    'If VendorInvoiceTax > 0 Then
    '    AMOUNT = VendorInvoiceTax
    '    AMOUNT = Replace(AMOUNT, "$", "")
    '    Application.SendKeys "Sales Tax", True
    '    Application.Wait (Now + TimeValue("00:00:01"))
    '    Application.SendKeys "~", True
    '    Application.Wait (Now + TimeValue("00:00:01"))
    '    Application.SendKeys "~", True
    '    Application.Wait (Now + TimeValue("00:00:01"))
    '    Application.SendKeys "1", True
     '   Application.SendKeys "~", True
      '  Application.Wait (Now + TimeValue("00:00:01"))
       ' Application.SendKeys AMOUNT, True
       ' Application.Wait (Now + TimeValue("00:00:01"))
      '  Application.SendKeys "~", True
      '  Application.Wait (Now + TimeValue("00:00:01"))
      '  Application.SendKeys "5005"
      '  Application.Wait (Now + TimeValue("00:00:01"))
      '  Application.SendKeys "~", True
    'End If

InvoiceTotalMatch:


SetCursorPos 151, 620  '--------------------------------------------------------Sage INVOICE WINDOW Discount Field
Call Mouse_left_button_press
Call Mouse_left_button_Letgo
Application.Wait (Now + TimeValue("00:00:01"))

Discount = ThisWorkbook.Sheets("Temp").Range("AF2")
If Discount = "Yes" Then
    MsgBox "Error - cell AF2 is 'Yes' when expecting a discount amount"
End If

If Discount <> "" Then
    For Repeat = 1 To 10
        Application.SendKeys ("{BS}")
    Next Repeat
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys Discount, True
    Application.Wait (Now + TimeValue("00:00:01"))
    'MsgBox "entered discount"
End If


Found = 0
If Password = "Invoice Total Match" Then GoTo skipmatchcheck:
    
    
    ' Over-ride check if all items entered by checking if sage total
    ' matches temp sheet invoiced amount
    Found = 0
    'Check if invoiced amount matches Sage total
            Sleep 500
            SetCursorPos 1134, 677 '--------------set mouse into "Net Due" field
            Sleep 250
            Call Mouse_left_button_press
            Sleep 250
            Call Mouse_left_button_Letgo
            Application.Wait (Now + TimeValue("00:00:01"))
            
            For x = 1 To 10
            Application.SendKeys "+{right}"
            Next x
            
            Application.SendKeys "^c", True
            Application.Wait (Now + TimeValue("00:00:01"))
            ThisWorkbook.Sheets("Temp").Paste Destination:=Sheets("Temp").Range("A1")
            invoiceTotal = ThisWorkbook.Sheets("temp").Range("N2").Offset(xoffset, 0)
            invoiceTotal = Replace(invoiceTotal, "$", "")
            'MsgBox "copy and pasted this=:" & ThisWorkbook.Sheets("Temp").Range("A1") & ", and invoice total on temp page is=:" & Invoicetotal
            
            Tax = ThisWorkbook.Sheets("temp").Range("AH2")
            If Tax <> 0 And Tax <> "Yes" Then invoiceTotal = invoiceTotal + Tax
            invoiceTotal = Replace(invoiceTotal, "Yes", "")
            If invoiceTotal < ThisWorkbook.Sheets("Temp").Range("N2") Then invoiceTotal = ThisWorkbook.Sheets("Temp").Range("N2")
            If invoiceTotal Like ThisWorkbook.Sheets("Temp").Range("A1") Or ItemMatch = 1 And invoiceTotal <= ThisWorkbook.Sheets("Temp").Range("A1") Then
                For x = 0 To 200
                    With ThisWorkbook.Sheets("Temp")
                        If .Range("H2").Offset(xoffset + x, 0) = .Range("H2").Offset(xoffset, 0) _
                                Then .Range("AC2").Offset(xoffset + x, 0) = "Yes"
                    End With
                Next x
                Found = 0
                'MsgBox "found the items equal"
            Else
                Found = 1
                For Repeat = 1 To 100
                    MsgBox "found the total not equal on invoice compared to in sage" & "Sage Total->" & ThisWorkbook.Sheets("Temp").Range("A1") _
                    & Chr(13) & "Temp Sheet Total->" & invoiceTotal
                    emailmessage = "Temp Sheet Total Error"
                Next Repeat
                Exit Sub
            End If
    
skipmatchcheck:
    'MsgBox "Freeze"
    If Found = 0 Then 'And TotalPOItems = totalInvoiceItems Then
            '-------------------------------------------------------------------------------Sage INVOICE WINDOW Save
            'Hit tab & enter to save if amount exceeded invoice
            
            emailmessage = "Saved"
            Application.Wait (Now + TimeValue("00:00:01"))
            Application.SendKeys "%{F4}"
            Application.Wait (Now + TimeValue("00:00:01"))
            Application.SendKeys "~", True
            Application.Wait (Now + TimeValue("00:00:02"))
            Application.SendKeys "^s"
            Application.Wait (Now + TimeValue("00:00:06"))
            
               
            'Check if hung on window "INVOICE EXCEEDS PO BALANCE". Break Code and wait for user input
            
            'Check if need to enter Job Data or not?
            
            SetCursorPos 616, 94  '---------------------Sage CHECK IF ON JOBCOST SCREEN
            Call Mouse_left_button_press
            Call Mouse_left_button_Letgo
            Application.Wait (Now + TimeValue("00:00:01"))
            ThisWorkbook.Sheets("Temp").Range("CA1:CA100") = ""
            Application.SendKeys "^c", True
            Application.Wait (Now + TimeValue("00:00:01"))
            ThisWorkbook.Sheets("Temp").Paste Destination:=Sheets("Temp").Range("CA1")
            Application.Wait (Now + TimeValue("00:00:02"))
            
            'Check if hung on window "INVOICE EXCEEDS PO BALANCE". Break Code and wait for user input
            For check = 0 To 10
                Line = ThisWorkbook.Sheets("Temp").Range("CA1").Offset(check, 0)
                If UCase(Line) Like "*BALANCE*" Then
                MsgBox "Invoice Exceeds PO Balance. Break Code here, finish user input, and resume code"
                MsgBox "Freeze"
                MsgBox "Freeze"
                MsgBox "Freeze"
                MsgBox "Freeze"
                MsgBox "Freeze"
                MsgBox "Freeze"
                MsgBox "Freeze"
                End If
            Next check
            
            
            Todaydate = Format(Date, "MM/DD/YYYY")
            'MsgBox "comparing pasted date of " & ThisWorkbook.Sheets("Temp").Range("A1") & " and todays date of " & Todaydate
            'If ThisWorkbook.Sheets("Temp").Range("BB1") Like "*" & Todaydate & "*" Then
            '    MsgBox "dates match"
                Application.SendKeys "%{F4}" 'We are in 4-2 module so exit one more time
                Application.Wait (Now + TimeValue("00:00:02"))
                Application.SendKeys "{Tab}", True
                Sleep 250
                Application.SendKeys "~", True
                Sleep 250
            'End If

            'application.SendKeys "%{F4}", True
            'application.Wait (Now + TimeValue("00:00:04"))
            'application.SendKeys "{Tab}", True
            'application.Wait (Now + TimeValue("00:00:01"))
            'application.SendKeys "~", True
            'application.Wait (Now + TimeValue("00:00:01"))
                    
            SetCursorPos 20, 234 '--------------------------Sage, Click 4 Accounts Payable
            Application.Wait (Now + TimeValue("00:00:01"))
            Call Mouse_left_button_press
            Call Mouse_left_button_Letgo
            
            emailmessage = "Saved"
            'MsgBox "freeze / did i do it!?"
    Else
            ' do not save
            emailmessage = "notsaved"
            Application.Wait (Now + TimeValue("00:00:01"))
            Application.SendKeys "%{F4}"
            Application.Wait (Now + TimeValue("00:00:01"))
            Application.SendKeys "{Tab}"
            Sleep 250
            Application.SendKeys "~"
            Sleep 250

            SetCursorPos 20, 234 '--------------------------Sage, Click 4 Accounts Payable
            Application.Wait (Now + TimeValue("00:00:01"))
            Call Mouse_left_button_press
            Call Mouse_left_button_Letgo
    
    End If
    



'MsgBox "How we doing"
Application.Wait (Now + TimeValue("00:00:02"))

End Sub
