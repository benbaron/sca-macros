Attribute VB_Name = "Module5"
'Password to unhide sheets
Const pWord = "KCoE"

Sub CreateReportFunds()
Dim newsheetname, newfundname As String
'  JL - array used when deleting pages
Dim curSheetNames() As String

Msg = "You are about to overwrite any existing Fund reports."
Title = "Fund Reports"
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
   Application.StatusBar = "Removing old SubFund Reports..."

   ' remove extra subfund ledgers
   ' JL - use an independent array because sheet numbers change as being deleted.
   x = Sheets.Count
   ReDim curSheetNames(x - 1)
   For i = 0 To x - 1
       curSheetNames(i) = Sheets(i + 1).Name
   Next i
       
   For Each sh In curSheetNames
     If Left(UCase(sh), 8) = Left(UCase("SubFund "), 8) Then
        Sheets(sh).Delete
     End If
   Next sh
   
   Application.DisplayAlerts = True
   Sheets("Ledger Report Template").Visible = xlSheetVisible

   Application.ScreenUpdating = False
   Sheets("Summary").Select
   Application.StatusBar = "Creating new Fund Reports..."

   For fundline = 10 To 51
       ' if there's a fund with activity
       If Sheets("Summary").Cells(fundline, 7) = "" Then
          Exit For
       Else
          newfundname = Sheets("Summary").Cells(fundline, 7).Value
          newsheetname = Left("SubFund " & newfundname, 30)
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
          Sheets(newsheetname).Range("B5") = "LEDGER AND JOURNAL FOR FUND: " & _
                                             newfundname
          Sheets(newsheetname).Range("C9") = newfundname & " Balance Forward: "
          Sheets(newsheetname).Range("F9") = Sheets("Summary").Cells(fundline, 8).Value
          
          ActiveWorkbook.Protect (pWord)
          
          Sheets(newsheetname).Select
          Range("A1").Select
        
          newfundline = 11
          ' copy transactions from 1st quarter that match this fund
          If Sheets("Summary").Cells(fundline, 16) Then
            For ledgerline = 11 To 110
               ' does this transaction affect this fund?
               If (Sheets("Ledger_Q1").Cells(ledgerline, 17) = newfundname And _
                  Sheets("Ledger_Q1").Cells(ledgerline, 13) <> 0) Or _
                  (Sheets("Ledger_Q1").Cells(ledgerline, 22) = newfundname And _
                  Sheets("Ledger_Q1").Cells(ledgerline, 18) <> 0) Or _
                  (Sheets("Ledger_Q1").Cells(ledgerline, 28) = newfundname And _
                  Sheets("Ledger_Q1").Cells(ledgerline, 24) <> 0) Or _
                  (Sheets("Ledger_Q1").Cells(ledgerline, 33) = newfundname And _
                  Sheets("Ledger_Q1").Cells(ledgerline, 29) <> 0) Then
                  ' copy transaction to subacct
                  CopyLedgerEntryNarrow "Ledger_Q1", newsheetname, ledgerline, newfundline, newfundname, 3
                  
                  newfundline = newfundline + 5
               End If
            Next ledgerline
          End If
          If Sheets("Summary").Cells(fundline, 17) Then
            ' copy transactions from 2nd quarter that match this fund
            For ledgerline = 11 To 110
               ' does this transaction affect this account?
               If (Sheets("Ledger_Q2").Cells(ledgerline, 17) = newfundname And _
                  Sheets("Ledger_Q2").Cells(ledgerline, 13) <> 0) Or _
                  (Sheets("Ledger_Q2").Cells(ledgerline, 22) = newfundname And _
                  Sheets("Ledger_Q2").Cells(ledgerline, 18) <> 0) Or _
                  (Sheets("Ledger_Q2").Cells(ledgerline, 28) = newfundname And _
                  Sheets("Ledger_Q2").Cells(ledgerline, 24) <> 0) Or _
                  (Sheets("Ledger_Q2").Cells(ledgerline, 33) = newfundname And _
                  Sheets("Ledger_Q2").Cells(ledgerline, 29) <> 0) Then
                  ' copy transaction to subacct
                  CopyLedgerEntryNarrow "Ledger_Q2", newsheetname, ledgerline, newfundline, newfundname, 3
                  
                  newfundline = newfundline + 5
               End If
            Next ledgerline
         End If
          If Sheets("Summary").Cells(fundline, 18) Then
            ' copy transactions from 3rd quarter that match this fund
            For ledgerline = 11 To 110
               ' does this transaction affect this fund?
               If (Sheets("Ledger_Q3").Cells(ledgerline, 17) = newfundname And _
                  Sheets("Ledger_Q3").Cells(ledgerline, 13) <> 0) Or _
                  (Sheets("Ledger_Q3").Cells(ledgerline, 22) = newfundname And _
                  Sheets("Ledger_Q3").Cells(ledgerline, 18) <> 0) Or _
                  (Sheets("Ledger_Q3").Cells(ledgerline, 28) = newfundname And _
                  Sheets("Ledger_Q3").Cells(ledgerline, 24) <> 0) Or _
                  (Sheets("Ledger_Q3").Cells(ledgerline, 33) = newfundname And _
                  Sheets("Ledger_Q3").Cells(ledgerline, 29) <> 0) Then
                  ' copy transaction to subacct
                  CopyLedgerEntryNarrow "Ledger_Q3", newsheetname, ledgerline, newfundline, newfundname, 3
                  
                  newfundline = newfundline + 5
               End If
            Next ledgerline
         End If
         ' copy transactions from 4th quarter that match this fund
          If Sheets("Summary").Cells(fundline, 19) Then
            For ledgerline = 11 To 110
               ' does this transaction affect this fund?
               If (Sheets("Ledger_Q4").Cells(ledgerline, 17) = newfundname And _
                  Sheets("Ledger_Q4").Cells(ledgerline, 13) <> 0) Or _
                  (Sheets("Ledger_Q4").Cells(ledgerline, 22) = newfundname And _
                  Sheets("Ledger_Q4").Cells(ledgerline, 18) <> 0) Or _
                  (Sheets("Ledger_Q4").Cells(ledgerline, 28) = newfundname And _
                  Sheets("Ledger_Q4").Cells(ledgerline, 24) <> 0) Or _
                  (Sheets("Ledger_Q4").Cells(ledgerline, 33) = newfundname And _
                  Sheets("Ledger_Q4").Cells(ledgerline, 29) <> 0) Then
                  ' copy transaction to subacct
                  CopyLedgerEntryNarrow "Ledger_Q4", newsheetname, ledgerline, newfundline, newfundname, 3
                  
                  newfundline = newfundline + 5
               End If
            Next ledgerline
          End If ' Q4
       End If  'check for subfund exists
       Sheets(newsheetname).Select
       ActiveSheet.PageSetup.PrintArea = "$b$3:$k$" & (newfundline + 5)
       hiderange = (newfundline + 5) & ":510"
       ActiveSheet.Rows(hiderange).Hidden = True
          
       ActiveSheet.Protect (pWord)
       On Error Resume Next ' for OO
       ActiveSheet.EnableSelection = 0 'xlNoRestrictions
       On Error GoTo 0
       
       Range("A1").Select
   Next fundline
   If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect (pWord)
   If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect ' for OO
   Sheets("Ledger Report Template").Visible = False
   ActiveWorkbook.Protect (pWord)
    
   ' reset protection, save workbook, and notify user that we're done!
   cleanupsub

End If ' response

End Sub
Sub CreateReportAccounts()
Dim curSheetNames() As String
Dim newsheetname, newacctname As String

Msg = "You are about to overwrite any existing Account Reports."
Title = "SubAccount Reports"
Style = vbOKCancel + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)
If response = vbOK Then
   Application.ScreenUpdating = False

   ' since they are read-only, just delete them and start over
   If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect (pWord)
   If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect 'for OO
   Application.DisplayAlerts = False
   ' set up status bar to show progress
   Application.DisplayStatusBar = True
   Application.StatusBar = "Removing old SubAcct Reports..."
   
   ' remove extra subaccount ledgers
   ' JL - use an independent array because sheet numbers change as being deleted.
   x = Sheets.Count
   ReDim curSheetNames(x - 1)
   For i = 0 To x - 1
       curSheetNames(i) = Sheets(i + 1).Name
   Next i
         
   For Each sh In curSheetNames
     If Left(UCase(sh), 8) = Left(UCase("SubAcct "), 8) Then
        Sheets(sh).Delete
     End If
   Next sh
   ' JL
   Application.DisplayAlerts = True
   Sheets("Ledger Report Template").Visible = xlSheetVisible
   
   Application.ScreenUpdating = False
   Sheets("Summary").Select
   Application.StatusBar = "Creating new SubAcct Reports..."

   For acctline = 10 To 35
       acct = Sheets("Summary").Cells(acctline, 3).Value
       If acct = "" Or acctline = 23 Or acctline = 24 Or acctline = 25 Then
       ' skip as they aren't account lines
       
       ' if there's activity for this account
       ElseIf Application.WorksheetFunction.CountBlank(Sheets("Summary").Cells(acctline, 3)) = 0 And _
          (Sheets("Summary").Cells(acctline, 12) Or _
           Sheets("Summary").Cells(acctline, 13) Or _
           Sheets("Summary").Cells(acctline, 14) Or _
           Sheets("Summary").Cells(acctline, 15)) Then
          
          newacctname = Sheets("Summary").Cells(acctline, 3).Value
          newsheetname = Left("SubAcct " & newacctname, 30)
          Application.StatusBar = "Creating " & newsheetname & "..."
          
          ' copy the ledger report template sheet with new account name
          If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect (pWord)
          If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect ' for OO
          Sheets("Summary").Select
          Sheets("Ledger Report Template").Copy After:=Sheets(Sheets.Count)
          If ActiveSheet.Name = "Summary" Then
             Sheets("Ledger Report Template_2").Select
          End If
    
          ActiveSheet.Name = newsheetname
          ActiveSheet.Unprotect (pWord)
          'for OO
          If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect
          ' reset the header to match the account name
          Sheets(newsheetname).Range("B5") = "LEDGER AND JOURNAL FOR ACCT: " & _
                                             newacctname
          Sheets(newsheetname).Range("C9") = newacctname & " Balance Forward: "
          Sheets(newsheetname).Range("F9") = Sheets("Summary").Cells(acctline, 4).Value
         
          ActiveWorkbook.Protect (pWord)
         
          Sheets(newsheetname).Select
          Range("A1").Select
       
          newacctline = 11
          ' copy transactions from 1st quarter that match this account
          ' check to see if there's any transactions in this quarter
          If Sheets("Summary").Cells(acctline, 12) Then
            For ledgerline = 10 To 109
               ' does this transaction affect this account?
               If (Sheets("Ledger_Q1").Cells(ledgerline, 14).Value = newacctname And _
                  Sheets("Ledger_Q1").Cells(ledgerline, 13).Value <> 0) Or _
                  (Sheets("Ledger_Q1").Cells(ledgerline, 19).Value = newacctname And _
                  Sheets("Ledger_Q1").Cells(ledgerline, 18).Value <> 0) Or _
                  (Sheets("Ledger_Q1").Cells(ledgerline, 25).Value = newacctname And _
                  Sheets("Ledger_Q1").Cells(ledgerline, 24).Value <> 0) Or _
                  (Sheets("Ledger_Q1").Cells(ledgerline, 30).Value = newacctname And _
                  Sheets("Ledger_Q1").Cells(ledgerline, 29).Value <> 0) Then
                  ' copy transaction to subacct
                  CopyLedgerEntryNarrow "Ledger_Q1", newsheetname, ledgerline, newacctline, newacctname, 0
                  
                  newacctline = newacctline + 5
               End If
            Next ledgerline
          End If
          ' check to see if there's any transactions in this quarter
          If Sheets("Summary").Cells(acctline, 13) Then
             ' copy transactions from 2nd quarter that match this account
             For ledgerline = 10 To 109
                ' does this transaction affect this account?
                If (Sheets("Ledger_Q2").Cells(ledgerline, 14).Value = newacctname And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 13).Value <> 0) Or _
                   (Sheets("Ledger_Q2").Cells(ledgerline, 19).Value = newacctname And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 18).Value <> 0) Or _
                   (Sheets("Ledger_Q2").Cells(ledgerline, 25).Value = newacctname And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 24).Value <> 0) Or _
                   (Sheets("Ledger_Q2").Cells(ledgerline, 30).Value = newacctname And _
                    Sheets("Ledger_Q2").Cells(ledgerline, 29).Value <> 0) Then
                   ' copy transaction to subacct
                   CopyLedgerEntryNarrow "Ledger_Q2", newsheetname, ledgerline, newacctline, newacctname, 0
                
                   newacctline = newacctline + 5
                End If
             Next ledgerline
          End If
          If Sheets("Summary").Cells(acctline, 14) Then
            ' copy transactions from 3rd quarter that match this account
            For ledgerline = 10 To 109
               ' does this transaction affect this account?
               If (Sheets("Ledger_Q3").Cells(ledgerline, 14).Value = newacctname And _
                  Sheets("Ledger_Q3").Cells(ledgerline, 13).Value <> 0) Or _
                  (Sheets("Ledger_Q3").Cells(ledgerline, 19).Value = newacctname And _
                  Sheets("Ledger_Q3").Cells(ledgerline, 18).Value <> 0) Or _
                  (Sheets("Ledger_Q3").Cells(ledgerline, 25).Value = newacctname And _
                  Sheets("Ledger_Q3").Cells(ledgerline, 24).Value <> 0) Or _
                  (Sheets("Ledger_Q3").Cells(ledgerline, 30).Value = newacctname And _
                  Sheets("Ledger_Q3").Cells(ledgerline, 29).Value <> 0) Then
                  ' copy transaction to subacct
                  CopyLedgerEntryNarrow "Ledger_Q3", newsheetname, ledgerline, newacctline, newacctname, 0
                  
                  newacctline = newacctline + 5
               End If
            Next ledgerline
          End If
          If Sheets("Summary").Cells(acctline, 15) Then
            ' copy transactions from 4th quarter that match this account
            For ledgerline = 10 To 109
               ' does this transaction affect this account?
               If (Sheets("Ledger_Q4").Cells(ledgerline, 14).Value = newacctname And _
                  Sheets("Ledger_Q4").Cells(ledgerline, 13).Value <> 0) Or _
                  (Sheets("Ledger_Q4").Cells(ledgerline, 19).Value = newacctname And _
                  Sheets("Ledger_Q4").Cells(ledgerline, 18).Value <> 0) Or _
                  (Sheets("Ledger_Q4").Cells(ledgerline, 25).Value = newacctname And _
                  Sheets("Ledger_Q4").Cells(ledgerline, 24).Value <> 0) Or _
                  (Sheets("Ledger_Q4").Cells(ledgerline, 30).Value = newacctname And _
                  Sheets("Ledger_Q4").Cells(ledgerline, 29).Value <> 0) Then
                  ' copy transaction to subacct
                  CopyLedgerEntryNarrow "Ledger_Q4", newsheetname, ledgerline, newacctline, newacctname, 0
                  
                  newacctline = newacctline + 5
               End If
            Next ledgerline
          End If
          Sheets(newsheetname).Select
          ' hide the unused rows
          ActiveSheet.PageSetup.PrintArea = "$b$3:$k$" & (newacctline + 5)
          hiderange = (newacctline + 5) & ":510"
          ActiveSheet.Rows(hiderange).Hidden = True
          ' turn the sheet into a report - non-editable and all white
          ActiveSheet.Protect (pWord)
          On Error Resume Next ' for OO
          ActiveSheet.EnableSelection = 0 'xlNoRestrictions
          On Error GoTo 0
          
          Range("A1").Select
          
       End If  'check for skipping non-account lines
    Next acctline
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect (pWord)
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect 'for OO
    Sheets("Ledger Report Template").Visible = False
    ActiveWorkbook.Protect (pWord)
    
    ' reset protection, save workbook, and notify user that we're done!
    cleanupsub
End If ' response

End Sub
Sub getRowCol(Col_0, Col_1, Col_2, Col_3, row, tranNo, AcctNum)
If tranNo < 7 Then
        Col_0 = "P"
        Col_1 = "Q"
        Col_2 = "R"
        Col_3 = "S"
        row = tranNo + 3 + ((AcctNum - 1) * 10)
    ElseIf tranNo < 13 Then
        Col_0 = "U"
        Col_1 = "V"
        Col_2 = "W"
        Col_3 = "X"
        row = tranNo - 6 + 3 + ((AcctNum - 1) * 10)
    ElseIf tranNo < 19 Then
        Col_0 = "Z"
        Col_1 = "AA"
        Col_2 = "AB"
        Col_3 = "AC"
        row = tranNo - 12 + 3 + ((AcctNum - 1) * 10)
    ElseIf tranNo < 25 Then
        Col_0 = "AE"
        Col_1 = "AF"
        Col_2 = "AG"
        Col_3 = "AH"
        row = tranNo - 18 + 3 + ((AcctNum - 1) * 10)
    Else
        Col_0 = "AJ"
        Col_1 = "AK"
        Col_2 = "AL"
        Col_3 = "AM"
        row = tranNo - 24 + 3 + ((AcctNum - 1) * 10)
    End If

End Sub

Sub getOldTrans(thisTran, tranNo, AcctNum)
Dim Col_0 As String
Dim Col_1 As String
Dim Col_2 As String
Dim Col_3 As String
Dim row As String

    Call getRowCol(Col_0, Col_1, Col_2, Col_3, row, tranNo, AcctNum)
    thisTran(0) = Sheets("Balances").Range(Col_0 & row).Value
    thisTran(1) = Sheets("Balances").Range(Col_1 & row).Value
    thisTran(2) = Sheets("Balances").Range(Col_2 & row).Value
    thisTran(3) = Sheets("Balances").Range(Col_3 & row).Value
End Sub


Sub putOldTrans(thisTran, tranNo, AcctNum)
Dim Col_0 As String
Dim Col_1 As String
Dim Col_2 As String
Dim Col_3 As String
Dim row As String

    Call getRowCol(Col_0, Col_1, Col_2, Col_3, row, tranNo, AcctNum)
    Sheets("Balances").Range(Col_0 & row).Value = thisTran(0)
    Sheets("Balances").Range(Col_1 & row).Value = thisTran(1)
    Sheets("Balances").Range(Col_2 & row).Value = thisTran(2)
    Sheets("Balances").Range(Col_3 & row).Value = thisTran(3)
End Sub
Sub ResetYear()
Dim curSheetNames() As String
Dim trans(3) As Variant
   
Msg = "You are about to reset the entire ledger!"
Title = "RESET Ledger"
Style = vbOKCancel + vbExclamation + vbDefaultButton1
doitresponse = MsgBox(Msg, Style, Title)
If doitresponse = vbOK Then
   Application.ScreenUpdating = False
   Application.DisplayStatusBar = True
   
   Msg = "Do you want to save this reset ledger to a new file?"
   Style = vbYesNo + vbExclamation + vbDefaultButton1
   newfileresponse = False
   newfileresponse = MsgBox(Msg, Style, Title)
   If newfileresponse = vbYes Then
       ' if there's a branch name, use it
       Sheets("Contents").Select
       If Application.WorksheetFunction.CountBlank(Sheets("Contents").Range("C4")) = 0 Then
          saveasname = "Ledger_" & Sheets("Contents").Range("C4").Value & "_" & (Sheets("Contents").Range("C5").Value + 1) & ".xls"
       Else
          saveasname = "New_" & ActiveWorkbook.Name
       End If
       Module1.mySaveFile (saveasname)
   End If
    
    ' Do un-reconciled Transactions
    numAccts = 0
    For i = 1 To 13
        If Sheets("Summary").Range("C" & 9 + i).Value <> "" Then
            numAccts = numAccts + 1
        Else
            Exit For
        End If
    Next i
    
    For i = 5 To 125 Step 10 ' clean old bank balances
        If Sheets("Balances").Range("N" & i).Value <> "" Then
            Sheets("Balances").Range("B" & i + 1).Value = Sheets("Balances").Range("N" & i).Value
        Else
            Sheets("Balances").Range("B" & i + 1).Value = Sheets("Balances").Range("N" & i - 2).Value
        End If
        Sheets("Balances").Range("C" & i & ":" & "N" & i).ClearContents
    Next i
    
    ' keep or clear old un-reconciled transactions
    For i = 1 To numAccts
        Tgt = 1
        For j = 1 To 30
            Call getOldTrans(trans, j, i)
            If trans(2) <> 0 And trans(3) = "-" Then
                If (j <> Tgt) Then ' if not the same cells
                    Call putOldTrans(trans, Tgt, i)
                    trans(0) = ""
                    trans(1) = ""
                    trans(2) = ""
                    trans(3) = "-"
                    Call putOldTrans(trans, j, i)
                End If
                Tgt = Tgt + 1
            Else ' clear data
                trans(0) = ""
                trans(1) = ""
                trans(2) = ""
                trans(3) = "-"
                Call putOldTrans(trans, j, i)
            End If
        Next j
    Next i
    
    ' get new un-reconciled Transactions
    trans(3) = "-"
    For i = 1 To numAccts
        row = 4 + ((i - 1) * 10)
        Account = Sheets("Summary").Range("C" & 9 + i).Value
        For j = 1 To 4
            numTrans = Sheets("Ledger_Q" & j).Range("BQ10").Value
            TransFound = 0
            For k = 11 To 110
                If Sheets("Ledger_Q" & j).Range("BQ" & k).Value > 0 Then ' if a transaction
                    TransFound = TransFound + 1
                    If Sheets("Ledger_Q" & j).Range("G" & k).Value = "-" Or Sheets("Ledger_Q" & j).Range("G" & k).Value = "" Then ' if un reconciled
                        If Sheets("Ledger_Q" & j).Range("N" & k).Value = Account Or _
                        Sheets("Ledger_Q" & j).Range("S" & k).Value = Account Or _
                        Sheets("Ledger_Q" & j).Range("Y" & k).Value = Account Or _
                        Sheets("Ledger_Q" & j).Range("AD" & k).Value = Account Then  ' got one for this account
                            
                            x = 1
                            trans(0) = Sheets("Ledger_Q" & j).Range("E" & k).Value
                            trans(1) = Sheets("Ledger_Q" & j).Range("D" & k).Value
                            trans(2) = 0
                            If Sheets("Ledger_Q" & j).Range("N" & k).Value = Account Then
                                If Sheets("Ledger_Q" & j).Range("O" & k).Value = "" Then x = -1
                                trans(2) = trans(2) + Sheets("Ledger_Q" & j).Range("M" & k).Value * x
                            End If
                            If Sheets("Ledger_Q" & j).Range("S" & k).Value = Account Then
                                If Sheets("Ledger_Q" & j).Range("T" & k).Value = "" Then x = -1
                                trans(2) = trans(2) + Sheets("Ledger_Q" & j).Range("R" & k).Value * x
                            End If
                            If Sheets("Ledger_Q" & j).Range("Y" & k).Value = Account Then
                                If Sheets("Ledger_Q" & j).Range("Z" & k).Value = "" Then x = -1
                                trans(2) = trans(2) + Sheets("Ledger_Q" & j).Range("X" & k).Value * x
                            End If
                            If Sheets("Ledger_Q" & j).Range("AD" & k).Value = Account Then
                                If Sheets("Ledger_Q" & j).Range("AE" & k).Value = "" Then x = -1
                                trans(2) = trans(2) + Sheets("Ledger_Q" & j).Range("AC" & k).Value * x
                            End If
                            Tgt = Sheets("Balances").Range("AO" & row).Value + 1
                            If trans(2) <> 0 Then Call putOldTrans(trans, Tgt, i)
                        End If ' end unrec matches account
                    End If ' end unrec
                End If ' end found transaction
                If TransFound = numTrans Then Exit For  ' we've seen em all
            Next k
        Next j
    Next i
        
    ' START RESETTING!!  increment year
    Application.StatusBar = "Resetting..."
    Sheets("Contents").Range("C5") = Sheets("Contents").Range("C5").Value + 1
    
    ' reset starting balances - do this before messing with any transactions
    Sheets("Summary").Unprotect (pWord)
    'for OO
    If Sheets("Summary").ProtectContents = True Then Sheets("Summary").Unprotect
    Sheets("Summary").Range("D10:D22") = Sheets("Summary").Range("e10:e22").Value
    Sheets("Summary").Range("D26:D35") = Sheets("Summary").Range("e26:e35").Value
    Sheets("Summary").Range("h10:h51") = Sheets("Summary").Range("i10:i51").Value
    Sheets("Summary").Protect (pWord)
    
    ' clear ledger pages
    ResetYearSub "Ledger_Q1"
    ResetYearSub "Ledger_Q2"
    ResetYearSub "Ledger_Q3"
    ResetYearSub "Ledger_Q4"
      
    Application.DisplayAlerts = True
    ActiveWorkbook.Protect (pWord)
       
    ' reset Equipment List"
    Sheets("Equipment_List").Range("l11:o260").ClearContents
    Sheets("Equipment_List").Range("l11:l260") = "No"
    
    ' remove extra calculated sub-ledgers
    Application.StatusBar = "Removing extra Ledger Report pages..."
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect (pWord)
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect 'for OO
    Application.DisplayAlerts = False
    ' remove extra sub-ledgers
    ' JL - use an independent array because sheet numbers change as being deleted.
    x = Sheets.Count
    ReDim curSheetNames(x - 1)
    For i = 0 To x - 1
        curSheetNames(i) = Sheets(i + 1).Name
    Next i
        
    For Each sh In curSheetNames
      If Left(UCase(sh), 8) = Left(UCase("SubFund "), 8) Then
            Sheets(sh).Delete
        ElseIf Left(UCase(sh), 8) = Left(UCase("SubAcct "), 8) Then
            Sheets(sh).Delete
        ElseIf Left(UCase(sh), 7) = Left(UCase("SubInc "), 7) Then
            Sheets(sh).Delete
        ElseIf Left(UCase(sh), 7) = Left(UCase("SubExp "), 7) Then
            Sheets(sh).Delete
        ElseIf Left(UCase(sh), 13) = Left(UCase("Unreconciled "), 13) And _
            UCase(sh) <> UCase(unreconciledsheetname) Then
            Sheets(sh).Delete
        ElseIf UCase(sh) = UCase("Ledger Report Template") Then
            ' ignore this page
        Else
        On Error Resume Next
            Sheets(sh).Select
            Range("A1").Select
        On Error GoTo 0
        End If
    Next sh
    ' JL
    
    ' reset protection, save workbook, and notify user that we're done!
    cleanupsub '(isOO)
End If 'doitresponse

End Sub
Sub DeleteReports()
'Dim isOO As Boolean
Dim curSheetNames() As String

Msg = "You are about to delete all Ledger Reports!"
Title = "DELETE Ledger Reports"
Style = vbOKCancel + vbExclamation + vbDefaultButton1
doitresponse = MsgBox(Msg, Style, Title)
If doitresponse = vbOK Then
   Application.ScreenUpdating = False
   Application.DisplayStatusBar = True
   Application.StatusBar = "Removing extra report pages..."

   ' remove extra calculated sub-ledgers
   If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect (pWord)
   If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect 'for OO
   Application.DisplayAlerts = False
   ' remove extra sub-ledgers
   ' JL - use an independent array because sheet numbers change as being deleted.
   x = Sheets.Count
   ReDim curSheetNames(x - 1)
   For i = 0 To x - 1
       curSheetNames(i) = Sheets(i + 1).Name
   Next i
        
   For Each sh In curSheetNames
      If Left(UCase(sh), 3) = UCase("Sub") Then
            Sheets(sh).Delete
        ElseIf Left(UCase(sh), 13) = Left(UCase("Unreconciled "), 13) And _
            UCase(sh) <> UCase(unreconciledsheetname) Then
            Sheets(sh).Delete
        ElseIf UCase(sh) = UCase("Ledger Report Template") Then
            ' ignore this page
        Else
        On Error Resume Next
            Sheets(sh).Select
            Range("A1").Select
        On Error GoTo 0
        End If
   Next sh
   ' reset protection, save workbook, and notify user that we're done!
   cleanupsub
    
End If 'response
End Sub
Sub cleanupsub()
Application.DisplayAlerts = False

Application.StatusBar = "Resetting locks... "
' reset protection
For sheetnum = 1 To 6
    Sheets(sheetnum).Protect (pWord)
Next sheetnum
Sheets("Free_Form").Unprotect (pWord) ' just to make sure
'for OO
If Sheets("Free_Form").ProtectContents = True Then Sheets("Free_Form").Unprotect
Sheets("contents").Select
Application.ScreenUpdating = True

'save workbook
   Module1.mySaveFile (ActiveWorkbook.Name)


If Not ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Protect (pWord)

' notify user that we're done and maybe saved!
Msg = "Done!"
'If Not isOOlocal Then
Msg = Msg & " File Saved."
Style = vbOKOnly + vbExclamation + vbDefaultButton1
newfileresponse = MsgBox(Msg, Style, Title)
Sheets("Contents").Select
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = False

End Sub
Sub ResetYearSub(loopsheet As String)

Application.StatusBar = "Resetting " & loopsheet & "... "
Sheets(loopsheet).Range("d11:e110").ClearContents ' date & check number

Sheets(loopsheet).Range("g11:j110").ClearContents ' cleared through budget

Sheets(loopsheet).Range("m11:v110").ClearContents ' page 2 transactions
Sheets(loopsheet).Range("x11:ag110").ClearContents ' page 3 transactions

' fill in defaults for account and fund
Sheets(loopsheet).Range("n11:n110") = Sheets("Summary").Range("C10").Value 'account 1
Sheets(loopsheet).Range("s11:s110") = Sheets("Summary").Range("C10").Value 'account 2
Sheets(loopsheet).Range("y11:y110") = Sheets("Summary").Range("C10").Value 'account 3
Sheets(loopsheet).Range("ad11:ad110") = Sheets("Summary").Range("C10").Value 'account 4

Sheets(loopsheet).Range("q11:q110") = Sheets("Summary").Range("G10").Value 'fund 1
Sheets(loopsheet).Range("v11:v110") = Sheets("Summary").Range("G10").Value 'fund 2
Sheets(loopsheet).Range("ab11:ab110") = Sheets("Summary").Range("G10").Value 'fund 3
Sheets(loopsheet).Range("ag11:ag110") = Sheets("Summary").Range("G10").Value 'fund 4
    
End Sub
Sub CopyLedgerEntry(sourcesheet, targetsheet, sourceline, targetline, targetName, checkoffset)

Sheets(targetsheet).Cells(targetline, 4) = Sheets(sourcesheet).Cells(sourceline, 4).Value 'date
Sheets(targetsheet).Cells(targetline, 5) = Sheets(sourcesheet).Cells(sourceline, 5).Value 'check #
Sheets(targetsheet).Cells(targetline, 7) = Sheets(sourcesheet).Cells(sourceline, 7).Value 'clear bank
Sheets(targetsheet).Cells(targetline, 8) = Sheets(sourcesheet).Cells(sourceline, 8).Value 'to/from
Sheets(targetsheet).Cells(targetline, 9) = Sheets(sourcesheet).Cells(sourceline, 9).Value 'memo/notes
Sheets(targetsheet).Cells(targetline, 10) = Sheets(sourcesheet).Cells(sourceline, 10).Value 'budget/tracking
                
If Sheets(sourcesheet).Cells(sourceline, 14 + checkoffset) = targetName And _
   Sheets(sourcesheet).Cells(sourceline, 13).Value <> 0 Then
   Sheets(targetsheet).Cells(targetline, 13) = Sheets(sourcesheet).Cells(sourceline, 13).Value 'amount 1
   Sheets(targetsheet).Cells(targetline, 14) = Sheets(sourcesheet).Cells(sourceline, 14).Value 'account 1
   Sheets(targetsheet).Cells(targetline, 15) = Sheets(sourcesheet).Cells(sourceline, 15).Value 'income 1
   Sheets(targetsheet).Cells(targetline, 16) = Sheets(sourcesheet).Cells(sourceline, 16).Value 'expense 1
   Sheets(targetsheet).Cells(targetline, 17) = Sheets(sourcesheet).Cells(sourceline, 17).Value 'fund 1
End If
If Sheets(sourcesheet).Cells(sourceline, 19 + checkoffset) = targetName And _
   Sheets(sourcesheet).Cells(sourceline, 18).Value <> 0 Then
   Sheets(targetsheet).Cells(targetline, 18) = Sheets(sourcesheet).Cells(sourceline, 18).Value 'amount 2
   Sheets(targetsheet).Cells(targetline, 19) = Sheets(sourcesheet).Cells(sourceline, 19).Value 'account 2
   Sheets(targetsheet).Cells(targetline, 20) = Sheets(sourcesheet).Cells(sourceline, 20).Value 'income 2
   Sheets(targetsheet).Cells(targetline, 21) = Sheets(sourcesheet).Cells(sourceline, 21).Value 'expense 2
   Sheets(targetsheet).Cells(targetline, 22) = Sheets(sourcesheet).Cells(sourceline, 22).Value 'fund 2
End If
If Sheets(sourcesheet).Cells(sourceline, 25 + checkoffset) = targetName And _
   Sheets(sourcesheet).Cells(sourceline, 24).Value <> 0 Then
   Sheets(targetsheet).Cells(targetline, 24) = Sheets(sourcesheet).Cells(sourceline, 24).Value 'amount 3
   Sheets(targetsheet).Cells(targetline, 25) = Sheets(sourcesheet).Cells(sourceline, 25).Value 'account 3
   Sheets(targetsheet).Cells(targetline, 26) = Sheets(sourcesheet).Cells(sourceline, 26).Value 'income 3
   Sheets(targetsheet).Cells(targetline, 27) = Sheets(sourcesheet).Cells(sourceline, 27).Value 'expense 3
   Sheets(targetsheet).Cells(targetline, 28) = Sheets(sourcesheet).Cells(sourceline, 28).Value 'fund 3
End If
If Sheets(sourcesheet).Cells(sourceline, 30 + checkoffset) = targetName And _
   Sheets(sourcesheet).Cells(sourceline, 29).Value <> 0 Then
   Sheets(targetsheet).Cells(targetline, 29) = Sheets(sourcesheet).Cells(sourceline, 29).Value 'amount 4
   Sheets(targetsheet).Cells(targetline, 30) = Sheets(sourcesheet).Cells(sourceline, 30).Value 'account 4
   Sheets(targetsheet).Cells(targetline, 31) = Sheets(sourcesheet).Cells(sourceline, 31).Value 'income 4
   Sheets(targetsheet).Cells(targetline, 32) = Sheets(sourcesheet).Cells(sourceline, 32).Value 'expense 4
   Sheets(targetsheet).Cells(targetline, 33) = Sheets(sourcesheet).Cells(sourceline, 33).Value 'fund 4
End If

End Sub
Sub CopyLedgerEntryNarrow(sourcesheet, targetsheet, sourceline, targetline, targetName, checkoffset)

Sheets(targetsheet).Cells(targetline, 4) = Sheets(sourcesheet).Cells(sourceline, 4).Value 'date
Sheets(targetsheet).Cells(targetline, 5) = Sheets(sourcesheet).Cells(sourceline, 5).Value 'check #
Sheets(targetsheet).Cells(targetline, 7) = Sheets(sourcesheet).Cells(sourceline, 7).Value 'clear bank
Sheets(targetsheet).Cells(targetline, 8) = Sheets(sourcesheet).Cells(sourceline, 8).Value 'to/from
Sheets(targetsheet).Cells(targetline, 9) = Sheets(sourcesheet).Cells(sourceline, 9).Value 'memo/notes
Sheets(targetsheet).Cells(targetline, 10) = Sheets(sourcesheet).Cells(sourceline, 10).Value 'budget/tracking
               
lineoffset = 1
If Sheets(sourcesheet).Cells(sourceline, 14 + checkoffset) = targetName And _
   Sheets(sourcesheet).Cells(sourceline, 13).Value <> 0 Then
   Sheets(targetsheet).Cells(targetline + lineoffset, 7) = Sheets(sourcesheet).Cells(sourceline, 13).Value 'amount 1
   Sheets(targetsheet).Cells(targetline + lineoffset, 8) = Sheets(sourcesheet).Cells(sourceline, 14).Value 'account 1
   If Application.WorksheetFunction.CountBlank(Sheets(sourcesheet).Cells(sourceline, 15)) = 0 Then
      Sheets(targetsheet).Cells(targetline + lineoffset, 9) = Sheets(sourcesheet).Cells(sourceline, 15).Value 'income 1
   Else
      Sheets(targetsheet).Cells(targetline + lineoffset, 9) = Sheets(sourcesheet).Cells(sourceline, 16).Value 'expense 1
      Sheets(targetsheet).Cells(targetline + lineoffset, 7) = Sheets(targetsheet).Cells(targetline + 1, 7) * -1
   End If
   Sheets(targetsheet).Cells(targetline + lineoffset, 10) = Sheets(sourcesheet).Cells(sourceline, 17).Value 'fund 1
   lineoffset = lineoffset + 1
End If

If Sheets(sourcesheet).Cells(sourceline, 19 + checkoffset) = targetName And _
   Sheets(sourcesheet).Cells(sourceline, 18).Value <> 0 Then
   Sheets(targetsheet).Cells(targetline + lineoffset, 7) = Sheets(sourcesheet).Cells(sourceline, 18).Value 'amount 2
   Sheets(targetsheet).Cells(targetline + lineoffset, 8) = Sheets(sourcesheet).Cells(sourceline, 19).Value 'account 2
   If Application.WorksheetFunction.CountBlank(Sheets(sourcesheet).Cells(sourceline, 20)) = 0 Then
      Sheets(targetsheet).Cells(targetline + lineoffset, 9) = Sheets(sourcesheet).Cells(sourceline, 20).Value 'income 2
   Else
      Sheets(targetsheet).Cells(targetline + lineoffset, 9) = Sheets(sourcesheet).Cells(sourceline, 21).Value 'expense 2
      Sheets(targetsheet).Cells(targetline + lineoffset, 7) = Sheets(targetsheet).Cells(targetline + 2, 7) * -1
   End If
   Sheets(targetsheet).Cells(targetline + lineoffset, 10) = Sheets(sourcesheet).Cells(sourceline, 22).Value 'fund 2
   lineoffset = lineoffset + 1
End If
If Sheets(sourcesheet).Cells(sourceline, 25 + checkoffset) = targetName And _
   Sheets(sourcesheet).Cells(sourceline, 24).Value <> 0 Then
   Sheets(targetsheet).Cells(targetline + lineoffset, 7) = Sheets(sourcesheet).Cells(sourceline, 24).Value 'amount 3
   Sheets(targetsheet).Cells(targetline + lineoffset, 8) = Sheets(sourcesheet).Cells(sourceline, 25).Value 'account 3
   If Application.WorksheetFunction.CountBlank(Sheets(sourcesheet).Cells(sourceline, 26)) = 0 Then
      Sheets(targetsheet).Cells(targetline + lineoffset, 9) = Sheets(sourcesheet).Cells(sourceline, 26).Value 'income 3
   Else
      Sheets(targetsheet).Cells(targetline + lineoffset, 9) = Sheets(sourcesheet).Cells(sourceline, 27).Value 'expense 3
      Sheets(targetsheet).Cells(targetline + lineoffset, 7) = Sheets(targetsheet).Cells(targetline + 3, 7) * -1
   End If
   Sheets(targetsheet).Cells(targetline + lineoffset, 10) = Sheets(sourcesheet).Cells(sourceline, 28).Value 'fund 3
   lineoffset = lineoffset + 1
End If
If Sheets(sourcesheet).Cells(sourceline, 30 + checkoffset) = targetName And _
   Sheets(sourcesheet).Cells(sourceline, 29).Value <> 0 Then
   Sheets(targetsheet).Cells(targetline + lineoffset, 7) = Sheets(sourcesheet).Cells(sourceline, 29).Value 'amount 4
   Sheets(targetsheet).Cells(targetline + lineoffset, 8) = Sheets(sourcesheet).Cells(sourceline, 30).Value 'account 4
   If Application.WorksheetFunction.CountBlank(Sheets(sourcesheet).Cells(sourceline, 31)) = 0 Then
      Sheets(targetsheet).Cells(targetline + lineoffset, 9) = Sheets(sourcesheet).Cells(sourceline, 31).Value 'income 4
   Else
      Sheets(targetsheet).Cells(targetline + lineoffset, 9) = Sheets(sourcesheet).Cells(sourceline, 32).Value 'expense 4
      Sheets(targetsheet).Cells(targetline + lineoffset, 7) = Sheets(targetsheet).Cells(targetline + lineoffset, 7) * -1
   End If
   Sheets(targetsheet).Cells(targetline + lineoffset, 10) = Sheets(sourcesheet).Cells(sourceline, 33).Value 'fund 4
End If

End Sub
Sub copyunreconciled(sourcesheet, targetsheet, sourceline, targetline)
   
Sheets(targetsheet).Cells(targetline, 4) = Sheets(sourcesheet).Cells(sourceline, 4).Value 'date
Sheets(targetsheet).Cells(targetline, 5) = Sheets(sourcesheet).Cells(sourceline, 5).Value 'check #
Sheets(targetsheet).Cells(targetline, 7) = Sheets(sourcesheet).Cells(sourceline, 7).Value 'clear bank
Sheets(targetsheet).Cells(targetline, 8) = Sheets(sourcesheet).Cells(sourceline, 8).Value 'to/from
Sheets(targetsheet).Cells(targetline, 9) = Sheets(sourcesheet).Cells(sourceline, 9).Value 'memo/notes
Sheets(targetsheet).Cells(targetline, 10) = Sheets(sourcesheet).Cells(sourceline, 10).Value 'budget/tracking
           
Sheets(targetsheet).Cells(targetline, 13) = Sheets(sourcesheet).Cells(sourceline, 13).Value 'amount 1
Sheets(targetsheet).Cells(targetline, 14) = Sheets(sourcesheet).Cells(sourceline, 14).Value 'account 1
Sheets(targetsheet).Cells(targetline, 15) = Sheets(sourcesheet).Cells(sourceline, 15).Value 'income 1
Sheets(targetsheet).Cells(targetline, 16) = Sheets(sourcesheet).Cells(sourceline, 16).Value 'expense 1
Sheets(targetsheet).Cells(targetline, 17) = Sheets(sourcesheet).Cells(sourceline, 17).Value 'fund 1
Sheets(targetsheet).Cells(targetline, 18) = Sheets(sourcesheet).Cells(sourceline, 18).Value 'amount 2
Sheets(targetsheet).Cells(targetline, 19) = Sheets(sourcesheet).Cells(sourceline, 19).Value 'account 2
Sheets(targetsheet).Cells(targetline, 20) = Sheets(sourcesheet).Cells(sourceline, 20).Value 'income 2
Sheets(targetsheet).Cells(targetline, 21) = Sheets(sourcesheet).Cells(sourceline, 21).Value 'expense 2
Sheets(targetsheet).Cells(targetline, 22) = Sheets(sourcesheet).Cells(sourceline, 22).Value 'fund 2
Sheets(targetsheet).Cells(targetline, 24) = Sheets(sourcesheet).Cells(sourceline, 24).Value 'amount 3
Sheets(targetsheet).Cells(targetline, 25) = Sheets(sourcesheet).Cells(sourceline, 25).Value 'account 3
Sheets(targetsheet).Cells(targetline, 26) = Sheets(sourcesheet).Cells(sourceline, 26).Value 'income 3
Sheets(targetsheet).Cells(targetline, 27) = Sheets(sourcesheet).Cells(sourceline, 27).Value 'expense 3
Sheets(targetsheet).Cells(targetline, 28) = Sheets(sourcesheet).Cells(sourceline, 28).Value 'fund 3
Sheets(targetsheet).Cells(targetline, 29) = Sheets(sourcesheet).Cells(sourceline, 29).Value 'amount 4
Sheets(targetsheet).Cells(targetline, 30) = Sheets(sourcesheet).Cells(sourceline, 30).Value 'account 4
Sheets(targetsheet).Cells(targetline, 31) = Sheets(sourcesheet).Cells(sourceline, 31).Value 'income 4
Sheets(targetsheet).Cells(targetline, 32) = Sheets(sourcesheet).Cells(sourceline, 32).Value 'expense 4
Sheets(targetsheet).Cells(targetline, 33) = Sheets(sourcesheet).Cells(sourceline, 33).Value 'fund 4

End Sub


