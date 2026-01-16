Attribute VB_Name = "Module7"
'=========================================================================
' Module: Module7 (Account maintenance)
' Purpose: Add, edit, delete, and reflow summary account rows.
' Key routines: addAccount, deleteAccount, editAccount
'=========================================================================
'Password to unhide sheets
Const pWord = "KCoE"

Sub addAccount()
Dim myValue As String
Dim i As Integer
    myValue = InputBox("Name of Account to Add", "Add an Account")
    myValue = Trim(myValue)
    If myValue = "" Then
        MsgBox ("No Account name given")
        Exit Sub
    End If
    For i = 10 To 22
        If myValue = Sheets("Summary").Range("C" & i).Value Then
            MsgBox ("Account already exist")
            Exit Sub
        End If
    Next i
    Sheets("Summary").Unprotect (pWord)
    If Sheets("Summary").ProtectContents = True Then Sheets("Summary").Unprotect ' for OO
    For i = 10 To 22
        setOff ("C" & i)
        If Sheets("Summary").Range("C" & i).Value = "" Then
            Sheets("Summary").Range("C" & i).Value = myValue
            setOn ("D" & i)
            Exit For
        ElseIf i = 22 Then
            MsgBox ("Unable to add since there are no blank accounts")
        End If
    Next i
    Sheets("Summary").Protect (pWord)
    hideShowAccounts
    
    
    
End Sub

Sub deleteAccount()
Dim reponse As Integer
Dim Title As String
Dim Style As Integer
Dim i As Integer
    If ActiveCell.Column < 3 Or ActiveCell.Column > 5 Or ActiveCell.row < 10 Or ActiveCell.row > 22 Or _
                Sheets("Summary").Range("C" & ActiveCell.row).Value = "" Then
        MsgBox ("Select the Account you wish to delete")
        Exit Sub
    End If
    
    Msg = "Do you reall wish to delete " & Sheets("Summary").Range("C" & ActiveCell.row).Value
    Title = "Delete Account"
    Style = vbOKCancel + vbExclamation + vbDefaultButton1
    response = MsgBox(Msg, Style, Title)
    If response = vbOK Then
        If Sheets("Summary").Range("L" & ActiveCell.row).Value = True Or _
            Sheets("Summary").Range("M" & ActiveCell.row).Value = True Or _
            Sheets("Summary").Range("N" & ActiveCell.row).Value = True Or _
            Sheets("Summary").Range("O" & ActiveCell.row).Value = True Then
                MsgBox ("Account is in use.  Unable to Delete")
                Exit Sub
        End If
        Sheets("Summary").Unprotect (pWord)
        If Sheets("Summary").ProtectContents = True Then Sheets("Summary").Unprotect ' for OO
        Sheets("Summary").Range("C" & ActiveCell.row).Value = ""
        Sheets("Summary").Range("D" & ActiveCell.row).Value = 0
        setOff ("C" & ActiveCell.row & ":D" & ActiveCell.row)
        For i = ActiveCell.row + 1 To 22
            If Sheets("Summary").Range("C" & i).Value <> "" Then
                Sheets("Summary").Range("C" & i - 1).Value = Sheets("Summary").Range("C" & i).Value
                Sheets("Summary").Range("D" & i - 1).Value = Sheets("Summary").Range("D" & i).Value
                Sheets("Summary").Range("D" & i).Value = 0
                Sheets("Summary").Range("C" & i).Value = ""
                setOn ("D" & i - 1)
                setOff ("C" & i & ":D" & i)
            End If
        Next i
        Sheets("Summary").Protect (pWord)
        scrunchAccounts (ActiveCell.row)
        hideShowAccounts
    End If

End Sub

Sub editAccount()
Dim newValue As String
Dim oldValue As String
Dim i As Integer
    ' is an account chosen
    If ActiveCell.Column < 3 Or ActiveCell.Column > 5 Or ActiveCell.row < 10 Or ActiveCell.row > 22 Or _
                Sheets("Summary").Range("C" & ActiveCell.row).Value = "" Then
        MsgBox ("Select an Existing Account to Edit")
        Exit Sub
    End If
        
    oldValue = Sheets("Summary").Range("C" & ActiveCell.row).Value
    newValue = InputBox("Edit Account " & oldValue, "Edit an Account", oldValue)
    newValue = Trim(newValue)
    
    ' is acceptable content
    If newValue = "" Or newValue = oldValue Then
        MsgBox ("No change or blank value")
        Exit Sub
    End If
    
    'check dupes
    For i = 10 To 22
        If i <> ActiveCell.row Then
            If Sheets("Summary").Range("C" & i).Value = newValue Then
                MsgBox (newValue & " is already in use")
                Exit Sub
            ElseIf Sheets("Summary").Range("C" & i).Value = "" Then
                Exit For
            End If
        End If
    Next i
    
    ' do it
    Call updateLedgerAccountNames(oldValue, newValue)
    Sheets("Summary").Unprotect (pWord)
    If Sheets("Summary").ProtectContents = True Then Sheets("Summary").Unprotect ' for OO
    Sheets("Summary").Range("C" & ActiveCell.row).Value = newValue
    Sheets("Summary").Protect (pWord)
End Sub
Sub addFund()
Dim myValue As String
Dim i As Integer
    myValue = InputBox("Name of Fund to Add", "Add a Fund")
    myValue = Trim(myValue)
    If myValue = "" Then
        MsgBox ("No Fund name given")
        Exit Sub
    End If
    For i = 10 To 51
        If myValue = Sheets("Summary").Range("G" & i).Value Then
            MsgBox ("Fund already exists")
            Exit Sub
        End If
    Next i
    Sheets("Summary").Unprotect (pWord)
    If Sheets("Summary").ProtectContents = True Then Sheets("Summary").Unprotect ' for OO
    For i = 11 To 51
        setOff ("G" & i)
        If Sheets("Summary").Range("G" & i).Value = "" Then
            Sheets("Summary").Range("G" & i).Value = myValue
            setOn ("H" & i)
            Exit For
        ElseIf i = 51 Then
            MsgBox ("Unable to add since there are no blank Funds")
        End If
    Next i
    Sheets("Summary").Protect (pWord)
    
End Sub

Sub deleteFund()
Dim reponse As Integer
Dim Title As String
Dim Style As Integer
Dim i As Integer
    If (ActiveCell.Column = 7 Or ActiveCell.Column = 8) And ActiveCell.row = 10 Then
        MsgBox ("Can't delete general Fund")
        Exit Sub
    ElseIf ActiveCell.Column < 7 Or ActiveCell.Column > 8 Or ActiveCell.row < 11 Or ActiveCell.row > 51 Then
        MsgBox ("Select the Fund you wish to delete")
        Exit Sub
    End If
    
    Msg = "Do you reall wish to delete " & Sheets("Summary").Range("G" & ActiveCell.row).Value
    Title = "Delete Fundt"
    Style = vbOKCancel + vbExclamation + vbDefaultButton1
    response = MsgBox(Msg, Style, Title)
    If response = vbOK Then
        If Sheets("Summary").Range("P" & ActiveCell.row).Value = True Or _
            Sheets("Summary").Range("Q" & ActiveCell.row).Value = True Or _
            Sheets("Summary").Range("R" & ActiveCell.row).Value = True Or _
            Sheets("Summary").Range("S" & ActiveCell.row).Value = True Then
                MsgBox ("Fund is in use.  Unable to Delete")
                Exit Sub
        End If
        Sheets("Summary").Unprotect (pWord)
        If Sheets("Summary").ProtectContents = True Then Sheets("Summary").Unprotect ' for OO
        Sheets("Summary").Range("G" & ActiveCell.row).Value = ""
        Sheets("Summary").Range("H" & ActiveCell.row).Value = 0
        setOff ("G" & ActiveCell.row & ":H" & ActiveCell.row)
        For i = ActiveCell.row + 1 To 51
            If Sheets("Summary").Range("G" & i).Value <> "" Then
                Sheets("Summary").Range("G" & i - 1).Value = Sheets("Summary").Range("G" & i).Value
                Sheets("Summary").Range("H" & i - 1).Value = Sheets("Summary").Range("H" & i).Value
                Sheets("Summary").Range("H" & i).Value = 0
                Sheets("Summary").Range("G" & i).Value = ""
                setOn ("H" & i - 1)
                setOff ("G" & i & ":H" & i)
            End If
        Next i
    End If
    Sheets("Summary").Protect (pWord)
End Sub

Sub editFund()
Dim newValue As String
Dim oldValue As String
Dim i As Integer
    ' is an Fund chosen
    If (ActiveCell.Column = 7 Or ActiveCell.Column = 8) And ActiveCell.row = 10 Then
        MsgBox ("Can't edit General Fund")
        Exit Sub
    ElseIf ActiveCell.Column < 7 Or ActiveCell.Column > 8 Or ActiveCell.row < 11 Or ActiveCell.row > 51 Then
        MsgBox ("Select the Fund you wish to delete")
        Exit Sub
    End If
        
    oldValue = Sheets("Summary").Range("G" & ActiveCell.row).Value
    newValue = InputBox("Edit Fund " & oldValue, "Edit a Fund", oldValue)
    newValue = Trim(newValue)
    
    ' is acceptable content
    If newValue = "" Or newValue = oldValue Then
        MsgBox ("No change or blank value")
        Exit Sub
    End If
    
    'check dupes
    For i = 10 To 51
        If i <> ActiveCell.row Then
            If Sheets("Summary").Range("G" & i).Value = newValue Then
                MsgBox (newValue & " is already in use")
                Exit Sub
            ElseIf Sheets("Summary").Range("G" & i).Value = "" Then
                Exit For
            End If
        End If
    Next i
    
    ' do it
    Call updateLedgerFundNames(oldValue, newValue)
    Sheets("Summary").Unprotect (pWord)
    If Sheets("Summary").ProtectContents = True Then Sheets("Summary").Unprotect ' for OO
    Sheets("Summary").Range("G" & ActiveCell.row).Value = newValue
    Sheets("Summary").Protect (pWord)
End Sub

Sub setOff(rng As String)
    Sheets("Summary").Range(rng).Interior.ColorIndex = xlNone
    Sheets("Summary").Range(rng).Locked = True
    Sheets("Summary").Range(rng).FormulaHidden = False
End Sub

Sub setOn(rng As String)
    Sheets("Summary").Range(rng).Interior.ColorIndex = 34
    Sheets("Summary").Range(rng).Locked = False
    Sheets("Summary").Range(rng).FormulaHidden = False
End Sub

Sub updateLedgerAccountNames(old, cur)
Dim i, j, k As Integer
    For i = 1 To 4  'ledger quarters
        For j = 1 To 4 ' account columns
            If j = 1 Then
                col = "N"
            ElseIf j = 2 Then
                col = "S"
            ElseIf j = 3 Then
                col = "Y"
            Else: col = "AD"
            End If
            For k = 10 To 110 'ledger rows
                If Sheets("Ledger_Q" & i).Range(col & k).Value = old Then
                    Sheets("Ledger_Q" & i).Range(col & k).Value = cur
                End If
            Next k
        Next j
    Next i
End Sub

Sub updateLedgerFundNames(old, cur)
Dim i, j, k As Integer
    For i = 1 To 4  'ledger quarters
        For j = 1 To 4 ' account columns
            If j = 1 Then
                col = "Q"
            ElseIf j = 2 Then
                col = "V"
            ElseIf j = 3 Then
                col = "AB"
            Else: col = "AG"
            End If
            For k = 10 To 110 'ledger rows
                If Sheets("Ledger_Q" & i).Range(col & k).Value = old Then
                    Sheets("Ledger_Q" & i).Range(col & k).Value = cur
                End If
            Next k
        Next j
    Next i
End Sub

Sub hideShowAccounts()
    Application.ScreenUpdating = False
    Sheets("Balances").Unprotect (pWord)
    If Sheets("Balances").ProtectContents = True Then Sheets("Balances").Unprotect ' for OO
    Sheets("Balances").Select
    For i = 1 To 12
        If Sheets("Summary").Range("C" & i + 10) = "" Then
            Rows(i & "1" & ":130").Select
            Selection.EntireRow.Hidden = True
            Exit For
        Else
            Rows(i & "1" & ":" & i & "9").Select
            Selection.EntireRow.Hidden = False
        End If
    Next i
    Rows(10).Select
    Selection.EntireRow.Hidden = True
    Rows(130).Select
    Selection.EntireRow.Hidden = True
    Sheets("Balances").Protect (pWord)
    Sheets("Balances").Range("A1").Select
    
    Sheets("Signatories").Select
    ActiveSheet.Unprotect (pWord)
    If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect ' for OO
    For i = 1 To 12
        If Sheets("Summary").Range("C" & i + 10) = "" Then
            Columns(Chr(Asc("G") + i) & ":S").Select
            Selection.EntireColumn.Hidden = True
            Exit For
        Else
            Columns(Chr(Asc("G") + i) & ":" & Chr(Asc("G") + i)).Select
            Selection.EntireColumn.Hidden = False
        End If
    Next i
    ActiveSheet.Protect (pWord)
    Sheets("Signatories").Range("A1").Select
 
    Sheets("Summary").Select
    Application.ScreenUpdating = True
End Sub

Sub scrunchAccounts(row)
    Application.ScreenUpdating = False
    baseRow = (row - 10) * 10
    Sheets("Balances").Unprotect (pWord)
    If Sheets("Balances").ProtectContents = True Then Sheets("Balances").Unprotect ' for OO
    For i = row + 1 To 22
        If Sheets("Summary").Range("C" & i - 1).Value <> "" Then
            srcRow = (i - 10) * 10
            For j = 3 To 8
                Sheets("Balances").Range("B" & baseRow + j).Value = Sheets("Balances").Range("B" & srcRow + j).Value
                Sheets("Balances").Range("B" & srcRow + j).Value = ""
            Next j
            Sheets("Balances").Range("E" & baseRow + 8).Value = Sheets("Balances").Range("E" & srcRow + 8).Value
            Sheets("Balances").Range("E" & srcRow + 8).Value = ""
            Sheets("Balances").Range("C" & baseRow + 98).Value = Sheets("Balances").Range("C" & srcRow + 9).Value
            Sheets("Balances").Range("C" & srcRow + 9).Value = ""
            For j = 3 To 14
                Sheets("Balances").Cells(baseRow + 5, j).Value = Sheets("Balances").Cells(srcRow + 5, j).Value
                Sheets("Balances").Cells(srcRow + 5, j).Value = ""
            Next j
            For j = 4 To 9
                For k = 16 To 19
                    Sheets("Balances").Cells(baseRow + j, k).Value = Sheets("Balances").Cells(srcRow + j, k).Value
                    Sheets("Balances").Cells(srcRow + j, k).Value = ""
                Next k
            Next j
            
            For j = 4 To 9
                For k = 21 To 24
                    Sheets("Balances").Cells(baseRow + j, k).Value = Sheets("Balances").Cells(srcRow + j, k).Value
                    Sheets("Balances").Cells(srcRow + j, k).Value = ""
                Next k
            Next j
            
            For j = 4 To 9
                For k = 26 To 29
                    Sheets("Balances").Cells(baseRow + j, k).Value = Sheets("Balances").Cells(srcRow + j, k).Value
                    Sheets("Balances").Cells(srcRow + j, k).Value = ""
                Next k
            Next j
            
            For j = 4 To 9
                For k = 31 To 34
                    Sheets("Balances").Cells(baseRow + j, k).Value = Sheets("Balances").Cells(srcRow + j, k).Value
                    Sheets("Balances").Cells(srcRow + j, k).Value = ""
                Next k
            Next j
            
            For j = 4 To 9
                For k = 36 To 39
                    Sheets("Balances").Cells(baseRow + j, k).Value = Sheets("Balances").Cells(srcRow + j, k).Value
                    Sheets("Balances").Cells(srcRow + j, k).Value = ""
                Next k
            Next j
        Else
            Exit For
        End If
        baseRow = (i - 10) * 10
    Next i

    Sheets("Balances").Protect (pWord)
    Sheets("Signatories").Unprotect (pWord)
    If Sheets("Signatories").ProtectContents = True Then Sheets("Signatories").Unprotect ' for OO
    Sheets("Signatories").Select
    baseCol = Chr(Asc("G") + (row - 10))
    Application.DisplayAlerts = False
    For i = row + 1 To 22
    
        If Sheets("Summary").Range("C" & i - 1).Value <> "" Then
            
            srcCol = Chr(Asc("G") + (i - 10))
            For j = 5 To 81 Step 4
                Sheets("Signatories").Range(baseCol & j).Value = Sheets("Signatories").Range(srcCol & j).Value
                Sheets("Signatories").Range(srcCol & j).Value = "-"
            Next j
            baseCol = srcCol
        Else
            Exit For
        End If
    Next i
    Application.DisplayAlerts = True
    Sheets("Signatories").Protect (pWord)
    Application.ScreenUpdating = True
End Sub


