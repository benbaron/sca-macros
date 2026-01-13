Attribute VB_Name = "Module4"
'Password to unhide sheets
Const pWord = "KCoE"
Sub CreateReportIncExp()
Dim newsheetname, newlinename As String
'  JL - array used when deleting pages
Dim curSheetNames() As String

Msg = "You are about to overwrite any existing Income and Expense reports."
Title = "Income and Expense Reports"
Style = vbOKCancel + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)
If response = vbOK Then
    ' since they are read-only, just delete them and start over
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect (pWord)
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect 'for OO
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ' set up status bar to show progress
    Application.DisplayStatusBar = True
    Application.StatusBar = "Removing old Income and Expense Reports..."
    
    ' remove extra subfund ledgers
    ' JL - use an independent array because sheet numbers change as being deleted.
    x = Sheets.Count
    ReDim curSheetNames(x - 1)
    For i = 0 To x - 1
        curSheetNames(i) = Sheets(i + 1).Name
    Next i
        
    For Each sh In curSheetNames
      If Left(UCase(sh), 6) = UCase("SubInc") Or Left(UCase(sh), 6) = UCase("SubExp") Then
         Sheets(sh).Delete
      End If
    Next sh
    
    Application.DisplayAlerts = True
    Sheets("Ledger Report Template").Visible = xlSheetVisible
    
    Application.ScreenUpdating = False
    Sheets("Summary").Select
    Application.StatusBar = "Creating new Income and Expense Reports..."


    For incomeline = 11 To 27
        If Sheets("Ledger_Q1").Cells(incomeline, 50).Value + Sheets("Ledger_Q2").Cells(incomeline, 50).Value + _
            Sheets("Ledger_Q3").Cells(incomeline, 50).Value + Sheets("Ledger_Q4").Cells(incomeline, 50).Value > 0 Then
            ' do the lines and delete if empty afterwards
            newlinename = Sheets("Ledger_Q1").Cells(incomeline, 44).Value
            newsheetname = Left("SubInc " & newlinename, 30)
            Application.StatusBar = "Creating " & newsheetname & "..."
            
            ' copy the report template with new account name
            If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect (pWord)
            If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect 'for OO
            Sheets("Summary").Select

            Sheets("Ledger Report Template").Copy After:=Sheets(Sheets.Count)
            If ActiveSheet.Name = "Summary" Then
               Sheets("Ledger Report Template_2").Select
            End If
            
            ActiveSheet.Unprotect (pWord)
            'for OO
            If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect
            ActiveSheet.Name = newsheetname
            ' reset the header to match the account name
            Sheets(newsheetname).Range("B5") = "LEDGER AND JOURNAL FOR INCOME CATEGORY: " & _
                                               newlinename
            Sheets(newsheetname).Range("C9:E9").ClearContents
            Sheets(newsheetname).Range("F9").ClearContents
            
            ActiveWorkbook.Protect (pWord)
            
            Sheets(newsheetname).Select
            Range("A1").Select
            
            newincomeline = 11
            ' copy transactions from 1st quarter that match this fund
            If Sheets("Ledger_Q1").Cells(incomeline, 45) <> 0 Then
              For ledgerline = 11 To 110
                 ' does this transaction affect this fund?
                 If (Sheets("Ledger_Q1").Cells(ledgerline, 15) = newlinename And _
                    Sheets("Ledger_Q1").Cells(ledgerline, 13) <> 0) Or _
                    (Sheets("Ledger_Q1").Cells(ledgerline, 20) = newlinename And _
                    Sheets("Ledger_Q1").Cells(ledgerline, 18) <> 0) Or _
                    (Sheets("Ledger_Q1").Cells(ledgerline, 26) = newlinename And _
                    Sheets("Ledger_Q1").Cells(ledgerline, 24) <> 0) Or _
                    (Sheets("Ledger_Q1").Cells(ledgerline, 31) = newlinename And _
                    Sheets("Ledger_Q1").Cells(ledgerline, 29) <> 0) Then
                    ' copy transaction to report
                    Module5.CopyLedgerEntryNarrow "Ledger_Q1", newsheetname, ledgerline, newincomeline, newlinename, 1
                    
                    newincomeline = newincomeline + 5
                 End If
              Next ledgerline
            End If
            If Sheets("Ledger_Q2").Cells(incomeline, 45) <> 0 Then
              ' copy transactions from 2nd quarter that match this fund
              For ledgerline = 11 To 110
                 ' does this transaction affect this account?
                 If (Sheets("Ledger_Q2").Cells(ledgerline, 15) = newlinename And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 13) <> 0) Or _
                    (Sheets("Ledger_Q2").Cells(ledgerline, 20) = newlinename And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 18) <> 0) Or _
                    (Sheets("Ledger_Q2").Cells(ledgerline, 26) = newlinename And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 24) <> 0) Or _
                    (Sheets("Ledger_Q2").Cells(ledgerline, 31) = newlinename And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 29) <> 0) Then
                    ' copy transaction to report
                    Module5.CopyLedgerEntryNarrow "Ledger_Q2", newsheetname, ledgerline, newincomeline, newlinename, 1
                    
                    newincomeline = newincomeline + 5
                 End If
              Next ledgerline
            End If
            If Sheets("Ledger_Q3").Cells(incomeline, 45) <> 0 Then
              ' copy transactions from 3rd quarter that match this fund
              For ledgerline = 11 To 110
                 ' does this transaction affect this fund?
                 If (Sheets("Ledger_Q3").Cells(ledgerline, 15) = newlinename And _
                    Sheets("Ledger_Q3").Cells(ledgerline, 13) <> 0) Or _
                    (Sheets("Ledger_Q3").Cells(ledgerline, 20) = newlinename And _
                    Sheets("Ledger_Q3").Cells(ledgerline, 18) <> 0) Or _
                    (Sheets("Ledger_Q3").Cells(ledgerline, 26) = newlinename And _
                    Sheets("Ledger_Q3").Cells(ledgerline, 24) <> 0) Or _
                    (Sheets("Ledger_Q3").Cells(ledgerline, 31) = newlinename And _
                    Sheets("Ledger_Q3").Cells(ledgerline, 29) <> 0) Then
                    ' copy transaction to report
                    Module5.CopyLedgerEntryNarrow "Ledger_Q3", newsheetname, ledgerline, newincomeline, newlinename, 1
                    
                    newincomeline = newincomeline + 5
                 End If
              Next ledgerline
            End If
            ' copy transactions from 4th quarter that match this fund
            If Sheets("Ledger_Q4").Cells(incomeline, 45) <> 0 Then
              For ledgerline = 11 To 110
                 ' does this transaction affect this fund?
                 If (Sheets("Ledger_Q4").Cells(ledgerline, 15) = newlinename And _
                    Sheets("Ledger_Q4").Cells(ledgerline, 13) <> 0) Or _
                    (Sheets("Ledger_Q4").Cells(ledgerline, 20) = newlinename And _
                    Sheets("Ledger_Q4").Cells(ledgerline, 18) <> 0) Or _
                    (Sheets("Ledger_Q4").Cells(ledgerline, 26) = newlinename And _
                    Sheets("Ledger_Q4").Cells(ledgerline, 24) <> 0) Or _
                    (Sheets("Ledger_Q4").Cells(ledgerline, 31) = newlinename And _
                    Sheets("Ledger_Q4").Cells(ledgerline, 29) <> 0) Then
                    ' copy transaction to report
                    Module5.CopyLedgerEntryNarrow "Ledger_Q4", newsheetname, ledgerline, newincomeline, newlinename, 1
                    
                    newincomeline = newincomeline + 5
                 End If
              Next ledgerline
            End If ' Q4
            Sheets(newsheetname).Select
            If newincomeline = 11 Then
                ' ditch this page
                Sheets("Ledger_Q1").Select
                ActiveWorkbook.Unprotect (pWord)
                If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect 'for OO
                Application.DisplayAlerts = False
                Sheets(newsheetname).Delete
            Else
                ActiveSheet.PageSetup.PrintArea = "$b$3:$k$" & (newincomeline)
                hiderange = (newincomeline) & ":510"
                ActiveSheet.Rows(hiderange).Hidden = True
                   
                ActiveSheet.Protect (pWord)
                On Error Resume Next ' for OO
                ActiveSheet.EnableSelection = 0 'xlNoRestrictions
                On Error GoTo 0
                Range("A1").Select
            End If
        End If
    Next incomeline
    
    For expenseline = 11 To 53
        If Sheets("Ledger_Q1").Cells(expenseline, 51).Value + Sheets("Ledger_Q2").Cells(expenseline, 51).Value + _
            Sheets("Ledger_Q3").Cells(expenseline, 51).Value + Sheets("Ledger_Q4").Cells(expenseline, 51).Value > 0 Then
            ' do the lines and delete if empty afterwards
            newlinename = Sheets("Ledger_Q1").Cells(expenseline, 46).Value
            newsheetname = Left("SubExp " & newlinename, 30)
            Application.StatusBar = "Creating " & newsheetname & "..."
            
            ' copy the report template with new account name
            If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect (pWord)
            If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect 'for OO
            Sheets("Summary").Select
            Sheets("Ledger Report Template").Copy After:=Sheets(Sheets.Count)
            If ActiveSheet.Name = "Summary" Then
               Sheets("Ledger Report Template_2").Select
            End If
            
            ActiveSheet.Unprotect (pWord)
            'for OO
            If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect
            ActiveSheet.Name = newsheetname
            ' reset the header to match the account name
            Sheets(newsheetname).Range("B5") = "LEDGER AND JOURNAL FOR EXPENSE CATEGORY: " & _
                                               newlinename
            Sheets(newsheetname).Range("C9:E9").ClearContents
            Sheets(newsheetname).Range("F9").ClearContents
            
            ActiveWorkbook.Protect (pWord)
            
            Sheets(newsheetname).Select
            Range("A1").Select
            
            newexpenseline = 11
            ' copy transactions from 1st quarter that match this fund
            If Sheets("Ledger_Q1").Cells(expenseline, 47) <> 0 Then
              For ledgerline = 11 To 110
                 ' does this transaction affect this fund?
                 If (Sheets("Ledger_Q1").Cells(ledgerline, 16) = newlinename And _
                    Sheets("Ledger_Q1").Cells(ledgerline, 13) <> 0) Or _
                    (Sheets("Ledger_Q1").Cells(ledgerline, 21) = newlinename And _
                    Sheets("Ledger_Q1").Cells(ledgerline, 18) <> 0) Or _
                    (Sheets("Ledger_Q1").Cells(ledgerline, 27) = newlinename And _
                    Sheets("Ledger_Q1").Cells(ledgerline, 24) <> 0) Or _
                    (Sheets("Ledger_Q1").Cells(ledgerline, 32) = newlinename And _
                    Sheets("Ledger_Q1").Cells(ledgerline, 29) <> 0) Then
                    ' copy transaction to report
                    Module5.CopyLedgerEntryNarrow "Ledger_Q1", newsheetname, ledgerline, newexpenseline, newlinename, 2
                    
                    newexpenseline = newexpenseline + 5
                 End If
              Next ledgerline
            End If
            If Sheets("Ledger_Q2").Cells(expenseline, 47) <> 0 Then
              ' copy transactions from 2nd quarter that match this fund
              For ledgerline = 11 To 110
                 ' does this transaction affect this account?
                 If (Sheets("Ledger_Q2").Cells(ledgerline, 16) = newlinename And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 13) <> 0) Or _
                    (Sheets("Ledger_Q2").Cells(ledgerline, 21) = newlinename And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 18) <> 0) Or _
                    (Sheets("Ledger_Q2").Cells(ledgerline, 27) = newlinename And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 24) <> 0) Or _
                    (Sheets("Ledger_Q2").Cells(ledgerline, 32) = newlinename And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 29) <> 0) Then
                    ' copy transaction to report
                    Module5.CopyLedgerEntryNarrow "Ledger_Q2", newsheetname, ledgerline, newexpenseline, newlinename, 2
                    
                    newexpenseline = newexpenseline + 5
                 End If
              Next ledgerline
            End If
            If Sheets("Ledger_Q3").Cells(expenseline, 47) <> 0 Then
              ' copy transactions from 3rd quarter that match this fund
              For ledgerline = 11 To 110
                 ' does this transaction affect this fund?
                 If (Sheets("Ledger_Q3").Cells(ledgerline, 16) = newlinename And _
                    Sheets("Ledger_Q3").Cells(ledgerline, 13) <> 0) Or _
                    (Sheets("Ledger_Q3").Cells(ledgerline, 21) = newlinename And _
                    Sheets("Ledger_Q3").Cells(ledgerline, 18) <> 0) Or _
                    (Sheets("Ledger_Q3").Cells(ledgerline, 27) = newlinename And _
                    Sheets("Ledger_Q3").Cells(ledgerline, 24) <> 0) Or _
                    (Sheets("Ledger_Q3").Cells(ledgerline, 32) = newlinename And _
                    Sheets("Ledger_Q3").Cells(ledgerline, 29) <> 0) Then
                    ' copy transaction to report
                    Module5.CopyLedgerEntryNarrow "Ledger_Q3", newsheetname, ledgerline, newexpenseline, newlinename, 2
                    
                    newexpenseline = newexpenseline + 5
                 End If
              Next ledgerline
            End If
            ' copy transactions from 4th quarter that match this fund
            If Sheets("Ledger_Q4").Cells(expenseline, 47) <> 0 Then
              For ledgerline = 11 To 110
                 ' does this transaction affect this fund?
                 If (Sheets("Ledger_Q4").Cells(ledgerline, 16) = newlinename And _
                    Sheets("Ledger_Q4").Cells(ledgerline, 13) <> 0) Or _
                    (Sheets("Ledger_Q4").Cells(ledgerline, 21) = newlinename And _
                    Sheets("Ledger_Q4").Cells(ledgerline, 18) <> 0) Or _
                    (Sheets("Ledger_Q4").Cells(ledgerline, 27) = newlinename And _
                    Sheets("Ledger_Q4").Cells(ledgerline, 24) <> 0) Or _
                    (Sheets("Ledger_Q4").Cells(ledgerline, 32) = newlinename And _
                    Sheets("Ledger_Q4").Cells(ledgerline, 29) <> 0) Then
                    ' copy transaction to report
                    Module5.CopyLedgerEntryNarrow "Ledger_Q4", newsheetname, ledgerline, newexpenseline, newlinename, 2
                    
                    newexpenseline = newexpenseline + 5
                 End If
              Next ledgerline
            End If ' Q4
            Sheets(newsheetname).Select
            If newexpenseline = 11 Then
                ' ditch this page
                Sheets("Ledger_Q1").Select
                ActiveWorkbook.Unprotect (pWord)
                If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect 'for OO
                Application.DisplayAlerts = False
                Sheets(newsheetname).Delete
            Else
                ActiveSheet.PageSetup.PrintArea = "$b$3:$k$" & (newexpenseline)
                hiderange = (newexpenseline) & ":510"
                ActiveSheet.Rows(hiderange).Hidden = True
                   
                ActiveSheet.Protect (pWord)
                On Error Resume Next ' for OO
                ActiveSheet.EnableSelection = 0 'xlNoRestrictions
                On Error GoTo 0
                
                Range("A1").Select
            End If
        End If
    Next expenseline
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect (pWord)
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect 'for OO
    Sheets("Ledger Report Template").Visible = False
    ActiveWorkbook.Protect (pWord)
     
    ' reset protection, save workbook, and notify user that we're done!
    Module5.cleanupsub
    
End If ' response

End Sub



