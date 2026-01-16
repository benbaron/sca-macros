Attribute VB_Name = "Module2"
'Password to unhide sheets
Const pWord = "SCoE"
Const LpWord = "KCoE"
'import from ledger routines
'
Sub importfromledger()
Dim sPath As String
Dim reportname As String
Dim ReportVersion As String
Dim ReportSub As String
Dim ledgername As String
Dim s As String
Dim badledger As Boolean
Dim strPath() As String
Dim LedgerVers() As String
Dim CSZ() As String
Dim lngIndex As Long
Dim procarray() As String
Dim qtrarray(4) As Boolean

ReportVersion = Sheets("contents").Range("B39")
ReportSub = Sheets("contents").Range("C15")
' 0. won't work for PAYPAL form
If ReportVersion = "PAYPAL" Then
   MsgBox ("This doesn't apply to PAYPAL form.")
   Exit Sub
End If
If ReportSub = "Non-US" Then
   MsgBox ("Can't import to Non-US report form.")
   Exit Sub
End If

' 1. ask for save information (new file, same file) and create new file if necessary
msg = "You are about to import from a ledger form! This may overwrite ALL UNSAVED data already in this workbook."
msg = msg & Chr(13) & Chr(13) & "The report will be saved in a new file based on the imported ledger's branch name."
Title = "IMPORT Ledger"
style = vbOKCancel + vbExclamation + vbDefaultButton1
doitresponse = MsgBox(msg, style, Title)
If doitresponse <> vbOK Then Exit Sub

Select Case Sheets("contents").Range("C12").Value
Case 1
   qtrarray(1) = True
   msg = "First quarter report. "
Case 2
   qtrarray(2) = True
   msg = "Second quarter report. "
   If Sheets("contents").Range("C13").Value = "Cumulative" Then
      qtrarray(1) = True
   End If
Case 3
   qtrarray(3) = True
   msg = "Third quarter report. "
   If Sheets("contents").Range("C13").Value = "Cumulative" Then
      qtrarray(1) = True
      qtrarray(2) = True
   End If
Case 4
   qtrarray(4) = True
   msg = "Fourth quarter report. "
   If Sheets("contents").Range("C13").Value = "Cumulative" Then
      qtrarray(1) = True
      qtrarray(2) = True
      qtrarray(3) = True
   End If
End Select
msg = msg & " ONLY Import these Ledger Quarters: "
If qtrarray(1) Then msg = msg & "First "
If qtrarray(2) Then msg = msg & "Second "
If qtrarray(3) Then msg = msg & "Third "
If qtrarray(4) Then msg = msg & "Fourth "
doitresponse = MsgBox(msg, style, Title)
If doitresponse <> vbOK Then Exit Sub

Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.DisplayAlerts = False

' 2. get ledger file name
filePath = CurDir
reportname = ActiveWorkbook.Name
ActiveWorkbook.Sheets("Contents").Select

ledgername = Module3.mygetfile
If LCase(ledgername) = "false" Then
   ' windows says bad file
   Exit Sub
End If
Application.StatusBar = "Opening " & ledgername
Application.ScreenUpdating = False
Workbooks.Open (ledgername)
If Left(ledgername, 7) = "file://" Or Left(ledgername, 1) = "/" Then
    splitChar = "/"
Else
    splitChar = "\"
End If
strPath() = Split(ledgername, splitChar)  'Put the Parts of our path into an array
lngIndex = UBound(strPath)
ledgername = strPath(lngIndex)   'Get the File Name from our array

Workbooks(ledgername).Activate
badledger = False

' location of definition cells
LedgerVers() = Split(Sheets("Contents").Range("F46"), " ")
lngIndex = UBound(LedgerVers)
If Left(LedgerVers(lngIndex), 1) = "3" Then
   ledgerversion = 3
   lq1 = "Ledger_Q1"
   lq2 = "Ledger_Q2"
   lq3 = "Ledger_Q3"
   lq4 = "Ledger_Q4"
   eql = "Equipment_List"
Else
   Workbooks(ledgername).Close
   Workbooks(reportname).Activate
   MsgBox ("Please convert your source Ledger to version 3 to enable Import.")
   Exit Sub
End If

' unhide all sheets
Sheets("Contents").Select
ActiveWorkbook.Unprotect (LpWord)
Sheets("Summary").Visible = True
Sheets(lq1).Visible = True
Sheets(lq2).Visible = True
Sheets(lq3).Visible = True
Sheets(lq4).Visible = True
ActiveWorkbook.Protect (LpWord)

' unlock
Application.StatusBar = "Unlocking " & ledgername
For Each sh In Worksheets
  sh.Unprotect (LpWord)
Next sh

' if the report workbook is empty, don't bother checking to see if the name and timeframe matches
If Workbooks(reportname).Sheets("Contents").Range("C8") <> "" Or Workbooks(reportname).Sheets("Contents").Range("C11") > 0 Then
    ' if we have a branch name in the report workbook, make sure the ledger name matches
    If Workbooks(reportname).Sheets("Contents").Range("C8") <> Workbooks(ledgername).Sheets("Contents").Range("C4") Then
        ' branch name doesn't match
        MsgBox ("Branch name does not match. Ending...")
        badledger = True
    End If
    ' check year
    If badledger = False And Workbooks(reportname).Sheets("Contents").Range("C11") > 0 Then
        If Workbooks(reportname).Sheets("Contents").Range("C11") <> Workbooks(ledgername).Sheets("Contents").Range("C5") Then
            ' year doesn't match
            badledger = True
            MsgBox ("Year does not match. Ending...")
        End If
    End If
    ' check subsidiary
    If badledger = False And Workbooks(ledgername).Sheets("Contents").Range("C6") = ReportSub And ReportSub = "Corporate" Then
        ' both corporate
    ElseIf badledger = False And ReportSub = Workbooks(ledgername).Sheets("Contents").Range("C6") Then
        ' both state subsidiary
    Else
        ' subsidiary status doesn't match
        badledger = True
        MsgBox ("Corporate/Subsidiary status does not match. Ending...")
    End If
End If

' 3. validate that incoming ledger will fit into this version
'       validate that corp/state difference is okay
'       cancel if incoming ledger is incompatible
Application.StatusBar = "Validating " & ledgername & " will fit in this report..."
' check for acceptable report version (no PAYPAL here cuz the button to do this gets deleted)
If badledger = False And ReportVersion = "SMALL" Then
   ' check for too many secondary accounts
   If badledger = False And Sheets("Summary").Range("C15").Value = "" Then
   Else
      ' can't put too many secondary funds in small form
      badledger = True
      MsgBox ("SMALL form does not have enough room for the secondary accounts, use MEDIUM form...")
   End If
   
   ' make sure asset funds are empty
   If Workbooks(ledgername).Sheets("Summary").Range("D27") = 0 And _
      Workbooks(ledgername).Sheets("Summary").Range("E27") = 0 And _
      Workbooks(ledgername).Sheets("Summary").Range("D28") = 0 And _
      Workbooks(ledgername).Sheets("Summary").Range("E28") = 0 And _
      Workbooks(ledgername).Sheets("Summary").Range("D29") = 0 And _
      Workbooks(ledgername).Sheets("Summary").Range("E29") = 0 Then
   Else
      ' can't put assets in small form
      badledger = True
      MsgBox ("SMALL form does not have room for assets, use MEDIUM form...")
   End If
   
   If badledger = False And Workbooks(ledgername).Sheets("Summary").Range("D32") = 0 And _
      Workbooks(ledgername).Sheets("Summary").Range("E32") = 0 Then
   Else
      ' can't put newsletters in small form
      badledger = True
      MsgBox ("SMALL form does not have room for newsletters, use MEDIUM form...")
   End If
   
   If badledger = False And Workbooks(ledgername).Sheets("Summary").Range("G11") = "" And _
      Workbooks(ledgername).Sheets("Summary").Range("G12") = "" Then
   Else
      ' can't put funds in small form
      badledger = True
      MsgBox ("SMALL form does not have room for dedicated funds, use MEDIUM form...")
   End If
   
   If badledger = False And getledgervalue(ledgername, "AX19", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 24 Then
   ElseIf badledger = False And getledgervalue(ledgername, "AX20", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 16 Then
   Else
      ' can't put too many 'transfers in' in small form
      badledger = True
      MsgBox ("SMALL form does not have enough room for the transfers in, use MEDIUM form...")
   End If
   
   If badledger = False And getledgervalue(ledgername, "Ay49", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 14 Then
   ElseIf badledger = False And getledgervalue(ledgername, "Ay50", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 10 Then
   ElseIf badledger = False And getledgervalue(ledgername, "Ay51", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 10 Then
   Else
      ' can't put too many 'transfers out' in small form
      badledger = True
      MsgBox ("SMALL form does not have enough room for the transfers out, use MEDIUM form...")
   End If
ElseIf badledger = False And ReportVersion = "MEDIUM" Then
   ' check for too many secondary accounts
   If badledger = False And Sheets("Summary").Range("C15").Value = "" Then
   Else
      ' can't put too many secondary funds in medium form
      badledger = True
      MsgBox ("MEDIUM form does not have enough room for the secondary accounts, use LARGE form...")
   End If
   
   ' check for too many assets
   If badledger = False And getledgervalue(ledgername, "AV12", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 11 Then ' receivables
   ElseIf badledger = False And getledgervalue(ledgername, "Av16", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 5 Then ' prepaid expenses
   ElseIf badledger = False And getledgervalue(ledgername, "Av17", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 8 Then ' other assets
   Else
      ' can't put too many assets in medium form
      badledger = True
      MsgBox ("MEDIUM form does not have enough room for the asset items, use LARGE form...")
   End If
   
   ' check for too many liabilities
   If badledger = False And getledgervalue(ledgername, "AV19", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 15 Then ' deferred revenue
   ElseIf badledger = False And getledgervalue(ledgername, "AV20", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 7 Then ' payables
   ElseIf badledger = False And getledgervalue(ledgername, "AV21", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 7 Then ' other liabilities
   Else
      ' can't put too many liabilities in medium form
      badledger = True
      MsgBox ("MEDIUM form does not have enough room for the liability items, use LARGE form...")
   End If

   ' check for too many depreciation items
   If badledger = False And Workbooks(ledgername).Sheets(eql).Range("U9").Value + Sheets(eql).Range("U10").Value + Sheets(eql).Range("U11").Value <= 10 Then
   Else
      ' can't put too many depreciating items in medium form
      badledger = True
      MsgBox ("MEDIUM form does not have enough room for the 5-year depreciable items, use LARGE form...")
   End If
   If badledger = False And Workbooks(ledgername).Sheets(eql).Range("U12").Value <= 10 Then
   Else
      ' can't put too many depreciating items in medium form
      badledger = True
      MsgBox ("MEDIUM form does not have enough room for the 7-year depreciable items, use LARGE form...")
   End If
   
   ' check for too many regalia items
   If badledger = False And Workbooks(ledgername).Sheets(eql).Range("U8").Value <= 12 Then
   Else
      ' can't put too many regalia items in medium form
      badledger = True
      MsgBox ("MEDIUM form does not have enough room for the regalia items, use LARGE form...")
   End If
   
   ' check for too many inventory items
   If badledger = False And Workbooks(ledgername).Sheets(eql).Range("U7").Value <= 8 Then
   Else
      ' can't put too many inventory items in medium form
      badledger = True
      MsgBox ("MEDIUM form does not have enough room for the inventory items, use LARGE form...")
   End If
   
   ' check for too many transfers
   If badledger = False And getledgervalue(ledgername, "AX19", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 45 Then ' internal transfers
   ElseIf badledger = False And getledgervalue(ledgername, "AX20", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 34 Then ' external transfers
   Else
      ' can't put too many 'transfers in' in medium form
      badledger = True
      MsgBox ("MEDIUM form does not have enough room for the transfers in, use LARGE form...")
   End If
   
   If badledger = False And getledgervalue(ledgername, "Ay49", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 31 Then
   ElseIf badledger = False And getledgervalue(ledgername, "Ay50", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 20 Then
   ElseIf badledger = False And getledgervalue(ledgername, "Ay51", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 20 Then
   Else
      ' can't put too many 'transfers out' in medium form
      badledger = True
      MsgBox ("MEDIUM form does not have enough room for the transfers out, use LARGE form...")
   End If
   
   ' if subsidiary, check for too many expense donations
   If badledger = False And getledgervalue(ledgername, "Ay48", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) <= 9 Then
   Else
      ' can't put too many expense donations in medium form
      badledger = True
      MsgBox ("MEDIUM form does not have enough room for the donations, use LARGE form...")
   End If
End If

' okay to proceed
If badledger = False Then
    ' 4. fill in contents ----------------------------------------------------------------------------------------------
    Application.StatusBar = "Contents..."
    If Workbooks(reportname).Sheets("Contents").Range("C8") = "" Then
        ' fill in branch name, year and corporate/subsidiary info
        lyear = Workbooks(ledgername).Sheets("Contents").Range("C5")
        Workbooks(reportname).Sheets("Contents").Range("C8") = Workbooks(ledgername).Sheets("Contents").Range("C4")
        Workbooks(reportname).Sheets("Contents").Range("C10") = Workbooks(ledgername).Sheets("Signatories").Range("D5")
        Workbooks(reportname).Sheets("Contents").Range("C11") = Workbooks(ledgername).Sheets("Contents").Range("C5")
        If Workbooks(ledgername).Sheets("Contents").Range("C6") = "Corporate" Then
        Else
           Workbooks(reportname).Sheets("Contents").Range("C15") = Workbooks(ledgername).Sheets("Contents").Range("C6")
        End If
    End If
    lyear = Workbooks(reportname).Sheets("Contents").Range("C11")
    
    ' 4.a fill in exchequer information from signatory page ------------------------------------------------------------
    Application.StatusBar = "Contact Info..."
    Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("D12") = Workbooks(ledgername).Sheets("Signatories").Range("E5")
    If Workbooks(ledgername).Sheets("Signatories").Range("E6") = "" Then
    Else
        addrstr = Workbooks(ledgername).Sheets("Signatories").Range("E6")
        commaloc = InStr(addrstr, ",")
        If commaloc > 0 Then
            Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("D13") = Left(addrstr, commaloc) 'city
            addrstr = Right(addrstr, Len(addrstr) - commaloc)
            LTrim (addrstr) ' remove any spaces in front of state
            CSZ() = Split(addrstr, " ")
            Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("H13") = CSZ(UBound(CSZ)) 'zip
            If UBound(CSZ) = 1 Then
                Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("F13") = CSZ(0) 'state
            Else
                Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("F13") = Left(addrstr, InStr(addrstr, CSZ(UBound(CSZ)))) 'state
            End If
        Else
            Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("D13") = addrstr ' just put the whole thing in city, user can fix
        End If
    End If
    Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("D14") = Workbooks(ledgername).Sheets("Signatories").Range("E8")
    Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("F14") = Workbooks(ledgername).Sheets("Signatories").Range("F8")
    Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("D15") = Workbooks(ledgername).Sheets("Signatories").Range("D8")
    Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("D16") = Workbooks(ledgername).Sheets("Signatories").Range("D6")
    Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("H15") = Workbooks(ledgername).Sheets("Signatories").Range("F5")
    Workbooks(reportname).Sheets("CONTACT_INFO_1").Range("H16") = Workbooks(ledgername).Sheets("Signatories").Range("F6")
    
    ' 5. fill in accounts ----------------------------------------------------------------------------------------------
    Application.StatusBar = "Accounts..."
    ' checking - fill in with bottom balance from last quarter involved.
    If qtrarray(4) Then
        Workbooks(reportname).Sheets("Primary_Account_2a").Range("h37") = Workbooks(ledgername).Sheets(lq4).Range("c110").Value
        Workbooks(reportname).Sheets("Primary_Account_2a").Range("h19") = Workbooks(ledgername).Sheets("Balances").Range("N5").Value
    ElseIf qtrarray(3) Then
        Workbooks(reportname).Sheets("Primary_Account_2a").Range("h37") = Workbooks(ledgername).Sheets(lq3).Range("c110").Value
        Workbooks(reportname).Sheets("Primary_Account_2a").Range("h19") = Workbooks(ledgername).Sheets("Balances").Range("K5").Value
    ElseIf qtrarray(2) Then
        Workbooks(reportname).Sheets("Primary_Account_2a").Range("h37") = Workbooks(ledgername).Sheets(lq2).Range("c110").Value
        Workbooks(reportname).Sheets("Primary_Account_2a").Range("h19") = Workbooks(ledgername).Sheets("Balances").Range("H5").Value
    ElseIf qtrarray(1) Then
        Workbooks(reportname).Sheets("Primary_Account_2a").Range("h37") = Workbooks(ledgername).Sheets(lq1).Range("c110").Value
        Workbooks(reportname).Sheets("Primary_Account_2a").Range("h19") = Workbooks(ledgername).Sheets("Balances").Range("E5").Value
    End If
    Workbooks(reportname).Sheets("Primary_Account_2a").Range("e15") = Workbooks(ledgername).Sheets("Balances").Range("b3").Value ' bank account type
    Workbooks(reportname).Sheets("Primary_Account_2a").Range("e16") = Workbooks(ledgername).Sheets("Balances").Range("b4").Value ' bank account number
    Workbooks(reportname).Sheets("Primary_Account_2a").Range("h15") = Workbooks(ledgername).Sheets("Balances").Range("b5").Value ' signature requirement
    Workbooks(reportname).Sheets("Primary_Account_2a").Range("f38") = Workbooks(ledgername).Sheets("Balances").Range("b7").Value ' interest bearing
    Workbooks(reportname).Sheets("Primary_Account_2a").Range("e13") = Workbooks(ledgername).Sheets("Balances").Range("b8").Value ' bank name
    Workbooks(reportname).Sheets("Primary_Account_2a").Range("e14") = Workbooks(ledgername).Sheets("Balances").Range("e8").Value ' account title
    Workbooks(reportname).Sheets("Primary_Account_2a").Range("f17") = Workbooks(ledgername).Sheets("Balances").Range("c9").Value ' bank contact info
        
    ' fill in outstanding checks pt 1
    lrow = 4
    lcol = 16 ' p4
    rrow = 27
    rcol = 3
    
    ' Clear existing
    For I = 27 To 34
        Workbooks(reportname).Sheets("Primary_Account_2a").Cells(I, 3) = ""
        Workbooks(reportname).Sheets("Primary_Account_2a").Cells(I, 4) = ""
        Workbooks(reportname).Sheets("Primary_Account_2a").Cells(I, 5) = ""
        Workbooks(reportname).Sheets("Primary_Account_2a").Cells(I, 6) = ""
        Workbooks(reportname).Sheets("Primary_Account_2a").Cells(I, 7) = ""
        Workbooks(reportname).Sheets("Primary_Account_2a").Cells(I, 8) = ""
    Next I
   
    Dim monthsArray() As String
    ReDim monthsArray(11)
    monthsArray(0) = "Jan"
    monthsArray(1) = "Feb"
    monthsArray(2) = "Mar"
    monthsArray(3) = "Apr"
    monthsArray(4) = "May"
    monthsArray(5) = "Jun"
    monthsArray(6) = "Jul"
    monthsArray(7) = "Aug"
    monthsArray(8) = "Sep"
    monthsArray(9) = "Oct"
    monthsArray(10) = "Nov"
    monthsArray(11) = "Dec"
    
    If qtrarray(4) = True Then
        ' do nothing
    ElseIf qtrarray(3) = True Then
        ReDim Preserve monthsArray(8)
    ElseIf qtrarray(2) = True Then
        ReDim Preserve monthsArray(5)
    Else
        ReDim Preserve monthsArray(2)
    End If
    
    For I = 1 To 30
        ' check if not reconciled
        If Not Module3.inArray(monthsArray, Workbooks(ledgername).Sheets("Balances").Cells(lrow, lcol + 3).Value) Then
            ' check to see if a negative amount
            If Workbooks(ledgername).Sheets("Balances").Cells(lrow, lcol + 2).Value < 0 Then
                If Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rrow, rcol + 2) = "" Then
                    Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rrow, rcol) = _
                        Workbooks(ledgername).Sheets("Balances").Cells(lrow, lcol).Value ' check #
                    Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rrow, rcol + 1) = _
                        Workbooks(ledgername).Sheets("Balances").Cells(lrow, lcol + 1).Value ' date
                    Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rrow, rcol + 2) = _
                        Workbooks(ledgername).Sheets("Balances").Cells(lrow, lcol + 2).Value * -1 ' amt
                    rrow = rrow + 1
                Else
                    tst = Split(Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rrow, rcol).Value)
                    
                    Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rrow, rcol) = tst(0) _
                        & " - " & Workbooks(ledgername).Sheets("Balances").Cells(lrow, lcol).Value ' check #
                    ' date ignored
                    Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rrow, rcol + 2) = _
                        Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rrow, rcol + 2) + _
                        (Workbooks(ledgername).Sheets("Balances").Cells(lrow, lcol + 2).Value * -1) ' amt
                End If
      
                If rrow = 35 Then
                   If rcol = 6 Then ' ran out of room on report
                       rrow = 34
                   Else
                       rrow = 27
                       rcol = rcol + 3
                   End If
                End If
            End If
        End If
        lrow = lrow + 1
        If lrow = 10 Then ' move over set on ledger
           lrow = 4
           lcol = lcol + 5
        End If
    Next I
    ' fill in outstanding checks pt 2
    If rrow < 35 Then
        getoutstandingchecks ledgername, reportname, rrow, rcol, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4), monthsArray
    End If
    
    ' signatories
    rsrow = 44
    rscol = 3
    lsrow = 9
    lscol = 3
    For I = 2 To 20
        If Workbooks(ledgername).Sheets("Signatories").Cells(lsrow, lscol + 4).Value = "X" Then
            Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rsrow, rscol) = Workbooks(ledgername).Sheets("Signatories").Cells(lsrow, lscol).Value ' title
            Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rsrow, rscol + 2) = Workbooks(ledgername).Sheets("Signatories").Cells(lsrow, lscol + 1).Value ' legal name
            Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rsrow, rscol + 3) = Workbooks(ledgername).Sheets("Signatories").Cells(lsrow, lscol + 2).Value ' address 1
            Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rsrow + 1, rscol + 3) = Workbooks(ledgername).Sheets("Signatories").Cells(lsrow + 1, lscol + 2).Value ' address 2
            Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rsrow, rscol + 5) = Workbooks(ledgername).Sheets("Signatories").Cells(lsrow, lscol + 3).Value ' membership #
            Workbooks(reportname).Sheets("Primary_Account_2a").Cells(rsrow + 1, rscol + 5) = Workbooks(ledgername).Sheets("Signatories").Cells(lsrow + 1, lscol + 3).Value ' membership exp
            rsrow = rsrow + 2
            lsrow = lsrow + 4
        End If
        If rsrow = 54 Then
           Exit For
        End If
    Next I
    
    ' other accounts
    Workbooks(reportname).Sheets("SECONDARY_ACCOUNTS_2b").Range("D13:g21,D25:g25,D27:g41").ClearContents
    If thisversion = "LARGE" Or thisversion = "MASTER" Then
       Workbooks(reportname).Sheets("SECONDARY_ACCOUNTS_2c").Range("D13:g21,D25:g25,D27:g41").ClearContents
       Workbooks(reportname).Sheets("SECONDARY_ACCOUNTS_2d").Range("D13:g21,D25:g25,D27:g41").ClearContents
    End If
    lrow = 14
    lcol = 16 ' p14
    rrow = 13
    rcol = 4 ' d13
    tsheetname = "Secondary_Accounts_2b"
    For I = 11 To 22
        If Workbooks(ledgername).Sheets("Summary").Cells(I, 3) = "" Then
           Exit For
        End If
        If ReportVersion = "LARGE" Then
        Else
           If I = 15 Then
              Exit For
           End If
        End If
        Workbooks(reportname).Sheets(tsheetname).Cells(13, rcol) = Workbooks(ledgername).Sheets("Summary").Cells(I, 3) ' account/bank name
        Workbooks(reportname).Sheets(tsheetname).Cells(16, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow - 1, lcol - 14) ' account type
        Workbooks(reportname).Sheets(tsheetname).Cells(14, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow, lcol - 14)  ' account number
        Workbooks(reportname).Sheets(tsheetname).Cells(15, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow + 1, lcol - 14)  ' signature req
        Workbooks(reportname).Sheets(tsheetname).Cells(17, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow + 3, lcol - 14)  ' interest bearing?
        Workbooks(reportname).Sheets(tsheetname).Cells(21, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow, 42) ' outstanding checks
        If qtrarray(4) Then
            Workbooks(reportname).Sheets(tsheetname).Cells(19, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow + 1, lcol - 2).Value ' bank balance
            Workbooks(reportname).Sheets(tsheetname).Cells(25, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow - 1, lcol - 2).Value ' ledger balance
        ElseIf qtrarray(3) Then
            Workbooks(reportname).Sheets(tsheetname).Cells(19, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow + 1, lcol - 5).Value  ' bank balance
            Workbooks(reportname).Sheets(tsheetname).Cells(25, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow - 1, lcol - 5).Value ' ledger balance
        ElseIf qtrarray(2) Then
            Workbooks(reportname).Sheets(tsheetname).Cells(19, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow + 1, lcol - 8).Value  ' bank balance
            Workbooks(reportname).Sheets(tsheetname).Cells(25, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow - 1, lcol - 8).Value ' ledger balance
        Else
            Workbooks(reportname).Sheets(tsheetname).Cells(19, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow + 1, lcol - 11).Value  ' bank balance
            Workbooks(reportname).Sheets(tsheetname).Cells(25, rcol) = Workbooks(ledgername).Sheets("Balances").Cells(lrow - 1, lcol - 11).Value ' ledger balance
        End If
        
        ' signatories
        rsrow = 27
        lscol = 4
        lsrow = 5
        For j = 1 To 20
            If Workbooks(ledgername).Sheets("Signatories").Cells(lsrow, I - 3).Value = "X" Then
                Workbooks(reportname).Sheets(tsheetname).Cells(rsrow, rcol) = Workbooks(ledgername).Sheets("Signatories").Cells(lsrow, lscol).Value ' legal name
                Workbooks(reportname).Sheets(tsheetname).Cells(rsrow + 1, rcol) = Workbooks(ledgername).Sheets("Signatories").Cells(lsrow, lscol + 2).Value ' membership #
                Workbooks(reportname).Sheets(tsheetname).Cells(rsrow + 2, rcol) = Workbooks(ledgername).Sheets("Signatories").Cells(lsrow + 1, lscol + 2).Value ' membership exp
                rsrow = rsrow + 3
            End If
            lsrow = lsrow + 4
            If rsrow = 45 Then
               Exit For
            End If
        Next j
        
        rcol = rcol + 1
        lrow = lrow + 10
        
        If I = 14 Then
            tsheetname = "Secondary_Accounts_2c"
            rcol = 4
        ElseIf I = 18 Then
            tsheetname = "Secondary_Accounts_2d"
            rcol = 4
        End If
    Next I
    
    ' 6a. fill in balance statement previous balance, using the balances page to get whether interest/non-interest bearing,
    ' and the summary page to get the starting ledger balance ----------------------------------------------------------------------
    noninttot = 0
    inttot = 0
    lrow = 6
    ldgradd = 10
    ldgraddval = 0
    For I = 1 To 13
        ' check for non-cumulative and non-first quarter to add in net from prior quarters to get to quarter start figures
        If qtrarray(1) = False Then
            If qtrarray(2) = False Then
                If qtrarray(3) = False Then
                    ' add net from first, second and third quarters to get to beginning of 4th quarter
                    ldgraddval = Workbooks(ledgername).Sheets(lq1).Range("AO" & (I + ldgradd)) + _
                        Workbooks(ledgername).Sheets(lq2).Range("AO" & (I + ldgradd)) + _
                        Workbooks(ledgername).Sheets(lq3).Range("AO" & (I + ldgradd))
                Else
                    ' add net from first and second quarters to get beginning of 3rd quarter
                    ldgraddval = Workbooks(ledgername).Sheets(lq1).Range("AO" & (I + ldgradd)) + _
                        Workbooks(ledgername).Sheets(lq2).Range("AO" & (I + ldgradd))
                End If
            Else
                ' add net from first quarter to get to beginning of 2nd quarter
                ldgraddval = Workbooks(ledgername).Sheets(lq1).Range("AO" & (I + ldgradd))
            End If
        End If
        If Workbooks(ledgername).Sheets("Balances").Cells(lrow + 1, 2) = "YES" Then ' interest earning
            inttot = inttot + Workbooks(ledgername).Sheets("Summary").Range("D" & (I + 9)) + ldgraddval
        ElseIf Workbooks(ledgername).Sheets("Balances").Cells(lrow + 1, 2) = "NO" Then  ' non interest earning
            noninttot = noninttot + Workbooks(ledgername).Sheets("Summary").Range("D" & (I + 9)) + ldgraddval
        End If
        lrow = lrow + 10
        If I = 1 Then ldgradd = 20
    Next I
    Workbooks(reportname).Sheets("BALANCE_3").Range("G19") = noninttot
    Workbooks(reportname).Sheets("BALANCE_3").Range("G20") = inttot
    
    ' 6b. fill in income statement -------------------------------------------------------------------------------------------------
    Application.StatusBar = "Income Statement..."
    Workbooks(reportname).Activate
    Module6.ClearIncomeExpense (ReportVersion)
    Workbooks(ledgername).Activate
    ' interest
    Workbooks(reportname).Sheets("INCOME_4").Range("J18").Value = getledgervalue(ledgername, "AS21", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' donations
    Workbooks(reportname).Sheets("INCOME_DTL_11a").Range("E33").Value = getledgervalue(ledgername, "AS14", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' stale checks
    Workbooks(reportname).Sheets("INCOME_DTL_11a").Range("E34").Value = getledgervalue(ledgername, "AS15", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' recovered bad checks
    Workbooks(reportname).Sheets("INCOME_DTL_11a").Range("E35").Value = getledgervalue(ledgername, "AS16", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' newsletter income
    If ReportVersion = "SMALL" Or ReportVersion = "PAYPAL" Then
    Else
        Workbooks(reportname).Sheets("NEWSLETTER_15").Range("I11").Value = getledgervalue(ledgername, "AS24", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    
    ' bank charges
    Workbooks(reportname).Sheets("INCOME_4").Range("G29").Value = getledgervalue(ledgername, "AU15", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("H29").Value = getledgervalue(ledgername, "AU16", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("I29").Value = getledgervalue(ledgername, "AU17", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' equip rental
    Workbooks(reportname).Sheets("INCOME_4").Range("G31").Value = getledgervalue(ledgername, "AU18", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("H31").Value = getledgervalue(ledgername, "AU19", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("I31").Value = getledgervalue(ledgername, "AU20", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' food
    Workbooks(reportname).Sheets("INCOME_4").Range("G33").Value = getledgervalue(ledgername, "AU24", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("H33").Value = getledgervalue(ledgername, "AU25", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("I33").Value = getledgervalue(ledgername, "AU26", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' general supplies
    Workbooks(reportname).Sheets("INCOME_4").Range("G34").Value = getledgervalue(ledgername, "AU27", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("H34").Value = getledgervalue(ledgername, "AU28", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("I34").Value = getledgervalue(ledgername, "AU29", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' occupancy
    Workbooks(reportname).Sheets("INCOME_4").Range("G36").Value = getledgervalue(ledgername, "AU31", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("H36").Value = getledgervalue(ledgername, "AU32", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("I36").Value = getledgervalue(ledgername, "AU33", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' postage
    Workbooks(reportname).Sheets("INCOME_4").Range("G37").Value = getledgervalue(ledgername, "AU34", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("H37").Value = getledgervalue(ledgername, "AU35", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("I37").Value = getledgervalue(ledgername, "AU36", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' printing
    Workbooks(reportname).Sheets("INCOME_4").Range("G38").Value = getledgervalue(ledgername, "AU37", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("H38").Value = getledgervalue(ledgername, "AU38", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("I38").Value = getledgervalue(ledgername, "AU39", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' telephone
    Workbooks(reportname).Sheets("INCOME_4").Range("G40").Value = getledgervalue(ledgername, "AU41", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("H40").Value = getledgervalue(ledgername, "AU42", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("I40").Value = getledgervalue(ledgername, "AU43", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    ' travel
    Workbooks(reportname).Sheets("INCOME_4").Range("G41").Value = getledgervalue(ledgername, "AU44", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("H41").Value = getledgervalue(ledgername, "AU45", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    Workbooks(reportname).Sheets("INCOME_4").Range("I41").Value = getledgervalue(ledgername, "AU46", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    
    ' 7. fill in assets -----------------------------------------------------------------------------------------------
    Application.StatusBar = "Assets..."
    Workbooks(reportname).Sheets("ASSET_DTL_5a").Range("c15:g18,c24:g34,c41:g45,c52:g59").ClearContents
    If thisversion = "LARGE" Or thisversion = "PAYPAL" Or thisversion = "MASTER" Then
       Workbooks(reportname).Sheets("ASSET_DTL_5c").Range("c13:f32,c39:f43,c50:f57").ClearContents
    End If
       
    ' receivables
    If getledgervalue(ledgername, "av12", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("An12")
        reportpage = "ASSET_DTL_5a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 24, 34, 14, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' prepaid expenses
    If getledgervalue(ledgername, "av16", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("An16")
        reportpage = "ASSET_DTL_5a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 41, 45, 14, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' other assets
    If getledgervalue(ledgername, "av17", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("An17")
        reportpage = "ASSET_DTL_5a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 52, 59, 14, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    
    ' 8. fill in liabilities ------------------------------------------------------------------------------------------
    Application.StatusBar = "Liabilities..."
    Workbooks(reportname).Sheets("LIABILITY_DTL_5b").Range("c16:f30,c37:f43,c49:f55").ClearContents
    
    If thisversion = "LARGE" Or thisversion = "PAYPAL" Or thisversion = "MASTER" Then
       Workbooks(reportname).Sheets("LIABILITY_DTL_5d").Range("c11:f28,c33:f46,c51:f55").ClearContents
       
       If thisversion = "PAYPAL" Or thisversion = "MASTER" Then
          Workbooks(reportname).Sheets("LIABILITY_DTL_5e").Range("c11:f55").ClearContents
          Workbooks(reportname).Sheets("LIABILITY_DTL_5f").Range("c11:f55").ClearContents
          Workbooks(reportname).Sheets("LIABILITY_DTL_5g").Range("c11:f55").ClearContents
       End If
    End If
    ' deferred revenue
    If getledgervalue(ledgername, "av19", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("An19")
        reportpage = "LIABILITY_DTL_5b"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 16, 30, 14, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' payables
    If getledgervalue(ledgername, "av20", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("An20")
        reportpage = "LIABILITY_DTL_5b"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 37, 43, 14, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' other liabilities
    If getledgervalue(ledgername, "av21", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("An21")
        reportpage = "LIABILITY_DTL_5b"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 49, 55, 14, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    
    ' fill in non-cash assets and funds for larger reports --------------------------------------------------------------------------
    If ReportVersion = "SMALL" Then
    Else
        Application.StatusBar = "Clearing Non-cash Assets..."
        Workbooks(reportname).Sheets("INVENTORY_DTL_6").Range("E13:l14,E16:l17,E19:l20,E24:l25,E30:l30").ClearContents
        Workbooks(reportname).Sheets("REGALIA_SALES_DTL_7").Range("C20:H31,c37:I46,c49:g51,i49:I51").ClearContents
        Workbooks(reportname).Sheets("DEPR_DTL_8").Range("d14:g23,j14:j23,e32:g41,j32:j41").ClearContents
         
        If thisversion = "LARGE" Or thisversion = "MASTER" Then
           Workbooks(reportname).Sheets("INVENTORY_DTL_6b").Range("E13:l14,E16:l17,E19:l20,E24:l25,E30:l30").ClearContents
           Workbooks(reportname).Sheets("REGALIA_SALES_DTL_7b").Range("C20:H31,c37:I46,c49:g51,i49:I51").ClearContents
           Workbooks(reportname).Sheets("DEPR_DTL_8b").Range("d14:g53,j14:j53").ClearContents
           Workbooks(reportname).Sheets("DEPR_DTL_8c").Range("e14:g53,j14:j53").ClearContents
        End If ' LARGE
        
        ' 9. fill in depreciation
        Application.StatusBar = "Depreciation..."
        fiveyrtotal = Workbooks(ledgername).Sheets(eql).Range("U9").Value + Workbooks(ledgername).Sheets(eql).Range("U10").Value + Workbooks(ledgername).Sheets(eql).Range("U11").Value
        If fiveyrtotal > 0 Then
           MsgBox ("There are " & fiveyrtotal & " 5-year Depreciation Assets marked on the Equipment List.")
           If Workbooks(reportname).Sheets("DEPR_DTL_8").Range("E14") <> "" Then
              msg = "Depreciation information is already in this Report Form. Overwrite?"
              style = vbYesNo + vbExclamation + vbDefaultButton1
              doitresponse = MsgBox(msg, style, Title)
           Else
              doitresponse = vbYes
           End If
           If doitresponse = vbYes Then
                reportline = 14
                targetpage = "DEPR_DTL_8"
                For ledgerline = 11 To 260
                   If InStr(Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 8), "5yr") > 0 Then
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 4) = Left(Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 8), 2) ' type
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 5) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 4) ' desc
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 6) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 5) ' qty
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 7) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 6) ' year acquired
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 10) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 7) ' current value
                      reportline = reportline + 1
                      If reportline = 24 And targetpage = "DEPR_DTL_8" Then
                         If ReportVersion = "LARGE" Then
                            targetpage = "DEPR_DTL_8b" ' go to next page
                            reportline = 14
                         Else
                           Exit For
                         End If
                      End If
                   End If
                Next ledgerline
           End If
        End If
        
        If Workbooks(ledgername).Sheets(eql).Range("U12").Value > 0 Then
           MsgBox ("There are " & Workbooks(ledgername).Sheets(eql).Range("U12").Value & " 7-year Depreciation Assets marked on the Equipment List.")
           If Workbooks(reportname).Sheets("DEPR_DTL_8").Range("E32") <> "" Then
              msg = "Depreciation information is already in this Report Form. Overwrite?"
              style = vbYesNo + vbExclamation + vbDefaultButton1
              doitresponse = MsgBox(msg, style, Title)
           Else
              doitresponse = vbYes
           End If
           If doitresponse = vbYes Then
                reportline = 32
                targetpage = "DEPR_DTL_8"
                For ledgerline = 11 To 260
                   If InStr(Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 8), "7yr") > 0 Then
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 5) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 4) ' desc
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 6) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 5) ' qty
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 7) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 6) ' year acquired
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 10) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 7) ' old value
                      reportline = reportline + 1
                      If reportline = 41 And targetpage = "DEPR_DTL_8" Then
                         If ReportVersion = "LARGE" Then
                            targetpage = "DEPR_DTL_8c" ' go to next page
                            reportline = 14
                         Else
                           Exit For
                         End If
                      End If
                   End If
                Next ledgerline
           End If
        End If
       
        ' 10. fill in regalia
        Application.StatusBar = "Regalia..." & Workbooks(ledgername).Sheets(eql).Range("U8")
        If Workbooks(ledgername).Sheets(eql).Range("U8").Value > 0 Then
           MsgBox ("There are " & Workbooks(ledgername).Sheets(eql).Range("U8") & " Regalia Assets marked on the Equipment List. Only copying those with value >= $500.")
           If Workbooks(reportname).Sheets("REGALIA_SALES_DTL_7").Range("C20") <> "" Then
              msg = "Regalia information is already in this Report Form. Overwrite?"
              style = vbYesNo + vbExclamation + vbDefaultButton1
              doitresponse = MsgBox(msg, style, Title)
           Else
              doitresponse = vbYes
           End If
           If doitresponse = vbYes Then
                reportline = 20
                targetpage = "REGALIA_SALES_DTL_7"
                For ledgerline = 11 To 260
                   If Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 8) = "Regalia" And _
                      Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 7) >= 500 Then ' only copy those above the limit
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 3) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 4) ' desc
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 4) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 5) ' qty
                      Workbooks(reportname).Sheets(targetpage).Cells(reportline, 5) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 6) ' year acquired
                      If Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 6) = lyear Then
                         Workbooks(reportname).Sheets(targetpage).Cells(reportline, 7) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 7) ' new value
                      Else
                         Workbooks(reportname).Sheets(targetpage).Cells(reportline, 6) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 7) ' old value
                      End If
                      reportline = reportline + 1
                      If reportline = 32 Then
                         If ReportVersion = "LARGE" Then
                            targetpage = "REGALIA_SALES_DTL_7b" ' go to next page
                            reportline = 20
                         Else
                           Exit For
                         End If
                      End If
                   End If
                Next ledgerline
           End If
        End If
        
        ' 11. fill in inventory
        Application.StatusBar = "Inventory..." & Workbooks(ledgername).Sheets(eql).Range("U7")
        If Workbooks(ledgername).Sheets(eql).Range("U7") > 0 Then
           MsgBox ("There are " & Workbooks(ledgername).Sheets(eql).Range("U7") & " Inventory Assets marked on the Equipment List.")
           If Workbooks(reportname).Sheets("INVENTORY_DTL_6").Range("E13") <> "" Then
              msg = "Inventory information is already in this Report Form. Overwrite?"
              style = vbYesNo + vbExclamation + vbDefaultButton1
              doitresponse = MsgBox(msg, style, Title)
           Else
              doitresponse = vbYes
           End If
           If doitresponse = vbYes Then
                reportcol = 5
                targetpage = "INVENTORY_DTL_6"
                For ledgerline = 11 To 260
                   If Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 8) = "Inventory" Then
                      Workbooks(reportname).Sheets(targetpage).Cells(13, reportcol) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 4) ' desc
                      If WorksheetFunction.IsNumber(Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 5)) Then
                          Workbooks(reportname).Sheets(targetpage).Cells(16, reportcol) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 5) ' qty
                      End If
                      If Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 6) < lyear Then
                         Workbooks(reportname).Sheets(targetpage).Cells(17, reportcol) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 7) ' prior value
                      Else
                         Workbooks(reportname).Sheets(targetpage).Cells(20, reportcol) = Workbooks(ledgername).Sheets(eql).Cells(ledgerline, 7) ' new value
                      End If
                      reportcol = reportcol + 1
                      If reportcol = 13 Then
                        If ReportVersion = "LARGE" Then
                           If targetpage = "INVENTORY_DTL_6" Then
                              targetpage = "INVENTORY_DTL_6b" ' go to next page
                           reportcol = 5
                        Else
                           Exit For
                        End If
                      End If
                   End If
                Next ledgerline
           End If
        End If
    
        ' 12. fill in funds
        Application.StatusBar = "Funds..."
        Workbooks(reportname).Sheets("FUNDS_14").Range("F14:F55,D15:E55").ClearContents
        Workbooks(reportname).Sheets("FUNDS_14").Range("f14") = Workbooks(ledgername).Sheets("Summary").Range("i10")
        If Workbooks(reportname).Sheets("FUNDS_14").Range("D15") <> "" Then
           msg = "Fund information is already in this Report Form. Overwrite?"
           style = vbYesNo + vbExclamation + vbDefaultButton1
           doitresponse = MsgBox(msg, style, Title)
        Else
           doitresponse = vbYes
        End If
        If doitresponse = vbYes Then
            For I = 10 To 51
                If Workbooks(ledgername).Sheets("Summary").Cells(I, 7).Value = "" Then
                   Exit For
                End If
                If I > 10 Then
                    Workbooks(reportname).Sheets("FUNDS_14").Cells(I + 4, 4) = Workbooks(ledgername).Sheets("Summary").Cells(I, 7).Value
                End If
                fundtot = Workbooks(ledgername).Sheets("Summary").Cells(I, 8).Value ' start value
                ' add the net up to the max quarter to get the balance at the end of the quarter
                If qtrarray(4) Then
                    fundtot = fundtot + Workbooks(ledgername).Sheets("Ledger_Q1").Range("AQ" & I + 1).Value + _
                              Workbooks(ledgername).Sheets("Ledger_Q2").Range("AQ" & I + 1).Value + _
                              Workbooks(ledgername).Sheets("Ledger_Q3").Range("AQ" & I + 1).Value + _
                              Workbooks(ledgername).Sheets("Ledger_Q4").Range("AQ" & I + 1).Value
                ElseIf qtrarray(3) Then
                    fundtot = fundtot + Workbooks(ledgername).Sheets("Ledger_Q1").Range("AQ" & I + 1).Value + _
                              Workbooks(ledgername).Sheets("Ledger_Q2").Range("AQ" & I + 1).Value + _
                              Workbooks(ledgername).Sheets("Ledger_Q3").Range("AQ" & I + 1).Value
                ElseIf qtrarray(2) Then
                    fundtot = fundtot + Workbooks(ledgername).Sheets("Ledger_Q1").Range("AQ" & I + 1).Value + _
                              Workbooks(ledgername).Sheets("Ledger_Q2").Range("AQ" & I + 1).Value
                ElseIf qtrarray(1) Then
                    fundtot = fundtot + Workbooks(ledgername).Sheets("Ledger_Q1").Range("AQ" & I + 1).Value
                End If
                Workbooks(reportname).Sheets("FUNDS_14").Cells(I + 4, 6) = fundtot
            Next I
        End If
    End If
    
    ' 13. fill in transfers in -----------------------------------------------------------------------------------
    Application.StatusBar = "Transfers In..."
    If getledgervalue(ledgername, "ax19", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
       ' inkingdom transfers
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AR19").Value
        reportpage = "TRANSFER_IN_9"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 14, 39, 15, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    If getledgervalue(ledgername, "ax20", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
       ' outkingdom transfers
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AR20").Value
        reportpage = "TRANSFER_IN_9"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 42, 57, 15, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    
    ' 14. fill in transfers out ----------------------------------------------------------------------------------
    Application.StatusBar = "Transfers Out..."
    If getledgervalue(ledgername, "ay49", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
       ' inkingdom transfers
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT49").Value
        reportpage = "TRANSFER_OUT_10"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 11, 24, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    If getledgervalue(ledgername, "ay50", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
       ' outkingdom transfers
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT50").Value
        reportpage = "TRANSFER_OUT_10"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 41, 50, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    If getledgervalue(ledgername, "ay51", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
       ' corporate transfers
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT51").Value
        reportpage = "TRANSFER_OUT_10"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 29, 38, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
   
    ' 16. fill in income detail -----------------------------------------------------------------------------------
    Application.StatusBar = "Income Detail..."
    ' internal fundraising
    If getledgervalue(ledgername, "ax11", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AR11")
        reportpage = "INCOME_DTL_11a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 11, 20, 15, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
       
    ' external fundraising
    If getledgervalue(ledgername, "ax12", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AR12")
        reportpage = "INCOME_DTL_11a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 23, 29, 15, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' demos/heraldic fees
    If getledgervalue(ledgername, "ax13", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AR13")
        reportpage = "INCOME_DTL_11a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 40, 51, 15, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' event income
    If getledgervalue(ledgername, "ax17", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AR17")
        reportpage = "INCOME_DTL_11b"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 12, 26, 15, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' refunds
    If getledgervalue(ledgername, "ay52", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("At52")
        reportpage = "INCOME_DTL_11b"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 12, 26, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' event income - PAYPAL
    If getledgervalue(ledgername, "ax18", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AR18")
        reportpage = "INCOME_DTL_11b"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 29, 35, 15, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' non-inventory sales
    If getledgervalue(ledgername, "ax23", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AR23")
        reportpage = "REGALIA_SALES_DTL_7"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 37, 46, 15, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' advertising income
    If getledgervalue(ledgername, "ax25", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AR25")
        reportpage = "INCOME_DTL_11b"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 40, 46, 15, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' other income
    If getledgervalue(ledgername, "ax26", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AR26")
        reportpage = "INCOME_DTL_11b"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 50, 56, 15, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    
    ' 17. fill in expense detail --------------------------------------------------------------------------------
    Application.StatusBar = "Expense Detail..."
    ' non-sca advertising
    If getledgervalue(ledgername, "ay11", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT11")
        reportpage = "EXPENSE_DTL_12a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 12, 22, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' bad debts
    If getledgervalue(ledgername, "ay12", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT12")
        reportpage = "EXPENSE_DTL_12a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 27, 38, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    If getledgervalue(ledgername, "ay13", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT13")
        reportpage = "EXPENSE_DTL_12a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 27, 38, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    If getledgervalue(ledgername, "ay14", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT14")
        reportpage = "EXPENSE_DTL_12a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 27, 38, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' fees
    If getledgervalue(ledgername, "ay21", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT21")
        reportpage = "EXPENSE_DTL_12a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 43, 54, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    If getledgervalue(ledgername, "ay22", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT22")
        reportpage = "EXPENSE_DTL_12a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 43, 54, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    If getledgervalue(ledgername, "ay23", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT23")
        reportpage = "EXPENSE_DTL_12a"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 43, 54, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' non-sca insurance
    If getledgervalue(ledgername, "ay30", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT30")
        reportpage = "EXPENSE_DTL_12b"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 12, 21, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' other expenses
    If getledgervalue(ledgername, "ay47", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT47")
        reportpage = "EXPENSE_DTL_12b"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 27, 41, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If
    ' donations to other non-profits
    If getledgervalue(ledgername, "ay48", qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4)) > 0 Then
        matchname = Workbooks(ledgername).Sheets("Ledger_Q1").Range("AT48")
        reportpage = "EXPENSE_DTL_12b"
        reportline = writereportline(ledgername, reportname, matchname, reportpage, 47, 55, 16, qtrarray(1), qtrarray(2), qtrarray(3), qtrarray(4))
    End If

    ' 16. save updated report form --------------------------------------------------------------------------
    Application.StatusBar = "Closing " & ledgername
    Workbooks(ledgername).Saved = True
    Workbooks(ledgername).Close
    Workbooks(reportname).Activate
    Application.StatusBar = "Saving " & reportname
    
    bname = Module3.sanitize(Workbooks(reportname).Sheets("Contents").Range("C8"))
    newName = "IMP_LDGR_" & (bname) & "_" & Workbooks(reportname).Sheets("Contents").Range("C11") _
                     & "_Q" & Workbooks(reportname).Sheets("Contents").Range("C12")
    Application.StatusBar = "Saving " & newName
    Module4.mysavefile (newName)
Else
    Application.StatusBar = "Closing " & ledgername
    Workbooks(ledgername).Saved = True
    Workbooks(ledgername).Close
    Workbooks(reportname).Activate
End If
Application.DisplayStatusBar = False
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox ("Done!")
 
End Sub
Function writereportline(lname, rname, matchtxt, reportpage, rptlinestart, rptlineend, ledgercolstart, q1f, q2f, q3f, q4f)
rptline = rptlinestart
Do While rptline < rptlineend + 1
    If matchtxt = "Advert Non-SCA" Or matchtxt = "Insurance - NON SCA" Then
        Exit Do ' not needed
    ElseIf (Workbooks(rname).Sheets(reportpage).Cells(rptline, 3) <> "") Then
        rptline = rptline + 1
    Else
        Exit Do ' found a blank line to fill in
    End If
Loop

If q1f Then
   rptline = writereportsub(lname, rname, rptline, "Ledger_Q1", matchtxt, reportpage, rptlineend, ledgercolstart)
End If
If rptline < rptlineend And q2f Then
   rptline = writereportsub(lname, rname, rptline, "Ledger_Q2", matchtxt, reportpage, rptlineend, ledgercolstart)
End If
If rptline < rptlineend And q3f Then
   rptline = writereportsub(lname, rname, rptline, "Ledger_Q3", matchtxt, reportpage, rptlineend, ledgercolstart)
End If
If rptline < rptlineend And q4f Then
   rptline = writereportsub(lname, rname, rptline, "Ledger_Q4", matchtxt, reportpage, rptlineend, ledgercolstart)
End If
writereportline = rptline
End Function
Function writereportsub(ledgername, reportname, rline, ledgerpage, matchtxt, reportpage, rptlineend, ledgercolstart)
Dim xarray() As String
For ledgerline = 11 To Workbooks(ledgername).Sheets(ledgerpage).Range("BR10") + 10 ' loop through all transaction lines
    If (Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart) = matchtxt) Or _
       (Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 5) = matchtxt) Or _
       (Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 11) = matchtxt) Or _
       (Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 16) = matchtxt) Then
        
        ledgerfld1 = Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, 8) ' to/from
        ledgerfld2 = Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, 9) ' memo
        ledgerfld3 = 0
        If Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart) = matchtxt Then
           ledgerfld3 = Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, 13) ' amt 1
           incflag = (Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 1) <> "") ' assets only: true if income category nonblank
        End If
        If Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 5) = matchtxt Then
           ledgerfld3 = ledgerfld3 + Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, 18) ' amt 2
           incflag = (Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 6) <> "") ' assets only: true if income category nonblank
        End If
        If Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 11) = matchtxt Then
           ledgerfld3 = ledgerfld3 + Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, 24) ' amt 3
           incflag = (Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 12) <> "") ' assets only: true if income category nonblank
        End If
        If Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 16) = matchtxt Then
           ledgerfld3 = ledgerfld3 + Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, 29) ' amt 4
           incflag = (Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 17) <> "") ' assets only: true if income category nonblank
        End If
        ledgerfld4 = Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, 4) ' date
        ledgerfld5 = Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, 5) ' check #
        ledgerfld6 = Workbooks(ledgername).Sheets(ledgerpage).Cells(ledgerline, 10) ' budget category or event name
        
        Select Case reportpage
        Case "ASSET_DTL_5a"
            If incflag Then
            Else: ledgerfld3 = ledgerfld3 * -1
            End If
            If matchtxt = "Receivables" Then
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 ' to/from
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld2 ' memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 7).Value = ledgerfld3 ' amt
            Else
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 7).Value = ledgerfld3 ' amt
            End If
            If rline = rptlineend Then
                If Workbooks(reportname).Sheets("CONTENTS").Range("B39") = "LARGE" Then
                   reportpage = "ASSET_DTL_5c"
                   Select Case matchtxt
                   Case "Receivables"
                      rline = 13
                      rptlineend = 32
                   Case "Prepaid Expenses"
                      rline = 39
                      rptlineend = 43
                   Case "Other Assets"
                      rline = 50
                      rptlineend = 57
                   End Select
                Else
                   Exit For
                End If
            End If
       Case "ASSET_DTL_5c"
            If incflag Then
            Else: ledgerfld3 = ledgerfld3 * -1
            End If
            If matchtxt = "Receivables" Then
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 ' to/from
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld2 ' memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            Else
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            End If
        Case "LIABILITY_DTL_5b"
            If incflag Then
            Else: ledgerfld3 = ledgerfld3 * -1
            End If
            If matchtxt = "Deferred Revenue" Then
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            Else
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 ' to/from
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld2 ' memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            End If
            If rline = rptlineend Then
                If Workbooks(reportname).Sheets("CONTENTS").Range("B39") = "LARGE" Then
                   reportpage = "LIABILITY_DTL_5d"
                   Select Case matchtxt
                   Case "Deferred Revenue"
                      rline = 11
                      rptlineend = 28
                   Case "Payables"
                      rline = 33
                      rptlineend = 46
                   Case "Other Liabilities"
                      rline = 51
                      rptlineend = 55
                   End Select
                Else
                   Exit For
                End If
            End If
        Case "LIABILITY_DTL_5d"
            If incflag Then
            Else: ledgerfld3 = ledgerfld3 * -1
            End If
            If matchtxt = "Deferred Revenue" Then
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            Else
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 ' to/from
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld2 ' memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            End If
        Case "REGALIA_SALES_DTL_7"
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 9).Value = ledgerfld3 ' amt
        Case "INCOME_DTL_11a"
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 ' to/from
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld2 ' memo
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 5).Value = ledgerfld3 ' amt
        Case "INCOME_DTL_11b"
            If matchtxt = "Other Income" Then
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            ElseIf matchtxt = "Refund" Then
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 5).Value = ledgerfld3 ' amt
            ElseIf matchtxt = "Advertising Income" Then
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld6 ' event name
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld3 ' amt
            Else
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld6 ' event name
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld3 ' amt
            End If
        Case "EXPENSE_DTL_12a"
            If matchtxt = "Advert Non-SCA" Then
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            Else
                xarray() = Split(matchtxt)
                lngIndex = UBound(xarray)
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = xarray(lngIndex) ' OA/AR/FR
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld1 ' to/from
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 5).Value = ledgerfld2 ' memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            End If
        Case "EXPENSE_DTL_12b"
            If matchtxt = "Insurance - NON SCA" Then
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            ElseIf matchtxt = "Other Expense" Then
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 ' to/from
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 5).Value = ledgerfld2 ' memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            Else
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
                Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
                If rline = rptlineend Then
                    If Sheets("contents").Range("C15") <> "Corporate" Then
                        reportpage = "EXPENSE_DTL_12c"
                        rline = 11
                        rptlineend = 54
                    Else
                        Exit For
                    End If
                End If
            End If
        Case "EXPENSE_DTL_12c"
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 5).Value = ledgerfld3 ' amt
        Case "TRANSFER_IN_9"
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            If rline = rptlineend Then
                If Workbooks(reportname).Sheets("CONTENTS").Range("B39") <> "SMALL" Then
                   reportpage = "TRANSFER_IN_9b"
                   If matchtxt = "Transfer In - In Kingdom" Then
                      rline = 11
                      rptlineend = 31
                   Else
                      rline = 36
                      rptlineend = 53
                   End If
                Else
                   Exit For
                End If
            End If
        Case "TRANSFER_IN_9b"
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            If rline = rptlineend Then
                If Workbooks(reportname).Sheets("CONTENTS").Range("B39") <> "MEDIUM" Then
                   reportpage = "TRANSFER_IN_9c"
                   If matchtxt = "Transfer In - In Kingdom" Then
                      rline = 11
                      rptlineend = 31
                   Else
                      rline = 36
                      rptlineend = 53
                   End If
                Else
                    Exit For
                End If
            End If
        Case "TRANSFER_IN_9c"
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            If rline = rptlineend Then
                If Workbooks(reportname).Sheets("CONTENTS").Range("B39") <> "MEDIUM" And matchtxt = "Transfer In - In Kingdom" Then
                   reportpage = "TRANSFER_IN_9d"
                   rline = 11
                   rptlineend = 54
                Else
                    Exit For
                End If
            End If
        Case "TRANSFER_OUT_10"
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 5).Value = ledgerfld4 ' check #
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld5 ' date
            If rline = rptlineend Then
                If Workbooks(reportname).Sheets("CONTENTS").Range("B39") <> "SMALL" Then
                   reportpage = "TRANSFER_OUT_10b"
                   Select Case matchtxt
                   Case "Transfer Out - In Kingdom"
                      rline = 11
                      rptlineend = 27
                   Case "Transfer Out - SCA Corp"
                      rline = 32
                      rptlineend = 41
                   Case "Transfer Out - Out Kingdom"
                      rline = 44
                      rptlineend = 53
                   End Select
                Else
                    Exit For
                End If
            End If
        Case "TRANSFER_OUT_10b"
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 5).Value = ledgerfld4 ' check #
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld5 ' date
            If rline = rptlineend Then
                If Workbooks(reportname).Sheets("CONTENTS").Range("B39") <> "MEDIUM" Then
                   reportpage = "TRANSFER_OUT_10c"
                   Select Case matchtxt
                   Case "Transfer Out - In Kingdom"
                      rline = 11
                      rptlineend = 27
                   Case "Transfer Out - SCA Corp"
                      rline = 32
                      rptlineend = 41
                   Case "Transfer Out - Out Kingdom"
                      rline = 44
                      rptlineend = 53
                   End Select
                Else
                    Exit For
                End If
            End If
        Case "TRANSFER_OUT_10c"
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1 & ", " & ledgerfld2 ' to/from, memo
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3 ' amt
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 5).Value = ledgerfld4 ' check #
            Workbooks(reportname).Sheets(reportpage).Cells(rline, 4).Value = ledgerfld5 ' date
            If rline = rptlineend Then
                If Workbooks(reportname).Sheets("CONTENTS").Range("B39") <> "MEDIUM" Then
                   reportpage = "TRANSFER_OUT_10d"
                   Select Case matchtxt
                   Case "Transfer Out - In Kingdom"
                      rline = 11
                      rptlineend = 27
                   Case "Transfer Out - SCA Corp"
                      rline = 32
                      rptlineend = 41
                   Case "Transfer Out - Out Kingdom"
                      rline = 44
                      rptlineend = 53
                   End Select
                Else
                    Exit For
                End If
            End If
        End Select
        rline = rline + 1
        If rline > rptlineend Then ' no more lines for this
           MsgBox ("Too many " & matchtxt & " items - please fill in manually. Continuing...")
           Exit For
        End If
    End If
Next ledgerline
writereportsub = rline
End Function
Function getledgervalue(ledgername, rngname, q1f, q2f, q3f, q4f)
Dim q1, q2, q3, q4

q1 = 0
q2 = 0
q3 = 0
q4 = 0
If q1f Then
   q1 = Workbooks(ledgername).Sheets("Ledger_Q1").Range(rngname).Value
End If
If q2f Then
   q2 = Workbooks(ledgername).Sheets("Ledger_Q2").Range(rngname).Value
End If
If q3f Then
   q3 = Workbooks(ledgername).Sheets("Ledger_Q3").Range(rngname).Value
End If
If q4f Then
   q4 = Workbooks(ledgername).Sheets("Ledger_Q4").Range(rngname).Value
End If
getledgervalue = q1 + q2 + q3 + q4
End Function
Sub getoutstandingchecks(ledgername, reportname, reprow, repcol, q1f, q2f, q3f, q4f, monthsArray)
Dim doIt As Boolean

doIt = False

y = UBound(monthsArray) + 2
ReDim Preserve monthsArray(y)
monthsArray(UBound(monthsArray) - 1) = "N/A"
monthsArray(UBound(monthsArray)) = "STALE"

' go through each ledger page looking for unreconciled transactions
For j = 1 To 4
    If j = 1 And q1f Then
        doIt = True
    ElseIf j = 2 And q2f Then
        doIt = True
    ElseIf j = 3 And q3f Then
        doIt = True
    ElseIf j = 4 And q4f Then
        doIt = True
    End If
    ledgerpage = "Ledger_Q" & j
If doIt = True Then
For I = 11 To 110
    ' if not reconciled in curent or previous period
    If Not Module3.inArray(monthsArray, Workbooks(ledgername).Sheets(ledgerpage).Cells(I, 7).Value) Then
    
        ' if has a negative value
        If Workbooks(ledgername).Sheets(ledgerpage).Cells(I, 6).Value < 0 Then
            If Workbooks(reportname).Sheets("Primary_Account_2a").Cells(reprow, repcol + 2) = "" Then
                Workbooks(reportname).Sheets("Primary_Account_2a").Cells(reprow, repcol) = Workbooks(ledgername).Sheets(ledgerpage).Cells(I, 5).Value ' check #
                Workbooks(reportname).Sheets("Primary_Account_2a").Cells(reprow, repcol + 1) = Workbooks(ledgername).Sheets(ledgerpage).Cells(I, 4).Value ' date
                Workbooks(reportname).Sheets("Primary_Account_2a").Cells(reprow, repcol + 2) = Workbooks(ledgername).Sheets(ledgerpage).Cells(I, 6).Value * -1 ' amt
                reprow = reprow + 1
            Else
                If Workbooks(ledgername).Sheets(ledgerpage).Cells(I, 5).Value = "" Then
                    Workbooks(reportname).Sheets("Primary_Account_2a").Cells(reprow, repcol) = _
                        Workbooks(reportname).Sheets("Primary_Account_2a").Cells(reprow, repcol) & ", --" ' concatenate something to show for that check if no check number
                Else
                    Workbooks(reportname).Sheets("Primary_Account_2a").Cells(reprow, repcol) = _
                        Workbooks(reportname).Sheets("Primary_Account_2a").Cells(reprow, repcol) & ", " _
                        & Workbooks(ledgername).Sheets(ledgerpage).Cells(I, 5).Value ' concatenate check #
                End If
                Workbooks(reportname).Sheets("Primary_Account_2a").Cells(reprow, repcol + 2) = _
                    Workbooks(reportname).Sheets("Primary_Account_2a").Cells(reprow, repcol + 2) + _
                    (Workbooks(ledgername).Sheets(ledgerpage).Cells(I, 6).Value * -1) ' add amt
            End If
            If reprow = 35 Then
                If repcol = 6 Then ' ran out of room on report
                   reprow = 34
                Else
                   reprow = 27
                   repcol = repcol + 3
                End If
            End If
        End If
    End If
Next I
End If
    doIt = False
Next j
End Sub


