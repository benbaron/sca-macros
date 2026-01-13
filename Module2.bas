Attribute VB_Name = "Module2"
'Password to unhide sheets
Const pWord = "KCoE"
Sub PrintAll()
Msg = "You are about to print the entire ledger."
Title = "Print Ledger"
Style = vbOKCancel + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)
If response = vbOK Then
   Application.ScreenUpdating = False
   Sheets("Contents").Select
   ActiveSheet.PrintOut
   Sheets(2).Select
   ActiveSheet.PrintOut

   PrintLedger "Ledger_Q1"
   PrintLedger "Ledger_Q2"
   PrintLedger "Ledger_Q3"
   PrintLedger "Ledger_Q4"
   PrintEquipment
   PrintSubFunds
   PrintSubAccts
   PrintSubIncExp
   PrintBalances
   PrintSignatories
End If

Sheets("Contents").Select
Application.ScreenUpdating = True

Msg = "Print Complete!"
Style = vbOK + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)

End Sub

Sub PrintFirstQuarter()
    Msg = "Print First Quarter Ledger."
    Title = "Print Ledger"
    Style = vbOKCancel + vbExclamation + vbDefaultButton1
    response = MsgBox(Msg, Style, Title)
    If response = vbOK Then
       Application.ScreenUpdating = False

       PrintLedger "Ledger_Q1"
       Sheets("Contents").Select
    
       Application.ScreenUpdating = True
   End If
End Sub

Sub PrintSecondQuarter()
    Msg = "Print Second Quarter Ledger."
    Title = "Print Ledger"
    Style = vbOKCancel + vbExclamation + vbDefaultButton1
    response = MsgBox(Msg, Style, Title)
    If response = vbOK Then
       Application.ScreenUpdating = False

       PrintLedger "Ledger_Q2"
       Sheets("Contents").Select
    
       Application.ScreenUpdating = True
   End If
End Sub

Sub PrintThirdQuarter()
    Msg = "Print Third Quarter Ledger."
    Title = "Print Ledger"
    Style = vbOKCancel + vbExclamation + vbDefaultButton1
    response = MsgBox(Msg, Style, Title)
    If response = vbOK Then
       Application.ScreenUpdating = False

       PrintLedger "Ledger_Q3"
       Sheets("Contents").Select
    
       Application.ScreenUpdating = True
   End If
End Sub

Sub PrintFourthQuarter()
    Msg = "Print Fourth Quarter Ledger."
    Title = "Print Ledger"
    Style = vbOKCancel + vbExclamation + vbDefaultButton1
    response = MsgBox(Msg, Style, Title)
    If response = vbOK Then
       Application.ScreenUpdating = False

       PrintLedger "Ledger_Q4"
       Sheets("Contents").Select
    
       Application.ScreenUpdating = True
   End If
End Sub

Sub PrintLedger(selsheet As String)
    ' check to see how far down to print
    
    Sheets(selsheet).Select
    For counter = 110 To 11 Step -1
       ' find lowest non-blank cell
       If Application.WorksheetFunction.CountBlank(Cells(counter, 8)) = 0 Or _
          ActiveSheet.Cells(counter, 13).Value <> 0 Then
          ' print all 3 pages across or only 2
          If Application.WorksheetFunction.Sum(Range("x11:x" & counter)) + _
             Application.WorksheetFunction.Sum(Range("ac11:ac" & counter)) = 0 Then
             ' 2 pages
             ActiveSheet.PageSetup.PrintArea = "$b$3:$v$" & counter + 1
          Else
             ' 3 pages
             ActiveSheet.PageSetup.PrintArea = "$b$3:$ag$" & counter + 1
          End If
          Exit For
        End If
    Next counter
    
    ' check for blank ledger to save paper
    If counter = 11 And _
          (Application.WorksheetFunction.CountBlank(Cells(counter + 2, 8)) = 1 And _
          ActiveSheet.Cells(counter + 2, 13).Value = 0) Then
       Msg = "This ledger is blank! Print blank ledger?"
       Title = "Print Blank Ledger"
       Style = vbYesNo + vbExclamation + vbDefaultButton1
       response = MsgBox(Msg, Style, Title)
       If response = vbYes Then
          ActiveSheet.PageSetup.PrintArea = "$b$3:$ag$110"
          ActiveSheet.PrintOut
       End If
    Else
       ActiveSheet.PrintOut
    End If
    
    ActiveSheet.PageSetup.PrintArea = "$b$3:$ag$110"
End Sub

Sub PrintEquip()
    Msg = "Print Equipment List"
    Title = "Print Equipment List"
    Style = vbOKCancel + vbExclamation + vbDefaultButton1
    response = MsgBox(Msg, Style, Title)
    If response = vbOK Then
       PrintEquipment
    End If
End Sub

Sub PrintEquipment()
    Application.ScreenUpdating = False

    Sheets("Equipment_List").Select
   
    ' check to see how far down to print
    For counter = 112 To 11 Step -1
       If Application.WorksheetFunction.CountBlank(Cells(counter, 4)) = 0 Then
          ActiveSheet.PageSetup.PrintArea = "$b$2:$p$" & counter
          Exit For
        End If
    Next counter
   
    If counter = 9 And Application.WorksheetFunction.CountBlank(Cells(counter + 1, 4)) <> 0 Then
       Msg = "This equipment list is blank! Print blank equipment list?"
       Title = "Print Equipment List"
       Style = vbYesNo + vbExclamation + vbDefaultButton1
       response = MsgBox(Msg, Style, Title)
       If response = vbYes Then
          ActiveSheet.PageSetup.PrintArea = "$b$2:$p$110"
          ActiveSheet.PrintOut
       End If
    Else
       ActiveSheet.PrintOut
    End If
    
    ' reset print area
    ActiveSheet.PageSetup.PrintArea = "$b$2:$p$110"
    Sheets("Contents").Select
    
    Application.ScreenUpdating = True
End Sub
Sub PrintSubAcc()

Msg = "You are about to print the SubAcct reports."
Title = "Print SubAccounts"
Style = vbOKCancel + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)
If response = vbOK Then
   PrintSubAccts
End If
Sheets("Contents").Select
Application.ScreenUpdating = True

Msg = "Print Complete!"
Style = vbOK + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)

End Sub
Sub PrintSubFun()

Msg = "You are about to print the SubFund reports."
Title = "Print SubFund"
Style = vbOKCancel + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)
If response = vbOK Then
   PrintSubFunds
End If
Sheets("Contents").Select
Application.ScreenUpdating = True

Msg = "Print Complete!"
Style = vbOK + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)

End Sub
Sub PrintSubFunds()
Application.ScreenUpdating = False

For Each sh In Worksheets
    If Left(UCase(sh.Name), 8) = Left(UCase("SubFund "), 8) Then
       Sheets(sh.Name).Select
       ActiveSheet.PrintOut
    End If
Next sh

End Sub
Sub PrintSubAccts()
Application.ScreenUpdating = False

For Each sh In Worksheets
    If Left(UCase(sh.Name), 8) = Left(UCase("SubAcct "), 8) Then
       Sheets(sh.Name).Select
       ActiveSheet.PrintOut
    End If
Next sh

End Sub
Sub printsubIE()
Msg = "You are about to print the SubCategory reports."
Title = "Print SubCategories"
Style = vbOKCancel + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)
If response = vbOK Then
   PrintSubIncExp
End If
Sheets("Contents").Select
Application.ScreenUpdating = True

Msg = "Print Complete!"
Style = vbOK + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)

End Sub
Sub PrintSubIncExp()
Application.ScreenUpdating = False

For Each sh In Worksheets
    If Left(UCase(sh.Name), 7) = Left(UCase("SubInc "), 7) Or Left(UCase(sh.Name), 7) = Left(UCase("SubExp "), 7) Then
       Sheets(sh.Name).Select
       ActiveSheet.PrintOut
    End If
Next sh

End Sub
Sub PrintBal()

Msg = "You are about to print the Balances page."
Title = "Print Balances"
Style = vbOKCancel + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)
If response = vbOK Then
   PrintBalances
End If
Sheets("Contents").Select
Application.ScreenUpdating = True

Msg = "Print Complete!"
Style = vbOK + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)

End Sub
Sub PrintBalances()

    ' check to see how far down to print
    Sheets("Balances").Select
    ActiveSheet.PageSetup.CenterHeader = Format(Sheets("contents").Range("E3").Value & vbCr & vbCr & Sheets("contents").Range("E5").Value)
    lowrow = 129
    rightcol = "am"
    For counter = 122 To 12 Step -10
       ' find lowest non-blank section
       If Range("A" & counter).Value <> "No Account" Then
          lowrow = counter + 7
          Exit For
        End If
    Next counter
    
    ' check to see how far across to print
    If Range("al10").Value + Range("al20").Value + Range("al30").Value + Range("al40").Value + _
        Range("al50").Value + Range("al60").Value + Range("al70").Value + Range("al80").Value + _
        Range("al90").Value + Range("al100").Value + Range("al110").Value + Range("al120").Value + Range("al130").Value > 0 Then
        rightcol = "am"
    ElseIf Range("ag10").Value + Range("ag20").Value + Range("ag30").Value + Range("ag40").Value + _
        Range("ag50").Value + Range("ag60").Value + Range("ag70").Value + Range("ag80").Value + _
        Range("ag90").Value + Range("ag100").Value + Range("ag110").Value + Range("ag120").Value + Range("ag130").Value > 0 Then
        rightcol = "ah"
    ElseIf Range("ab10").Value + Range("ab20").Value + Range("ab30").Value + Range("ab40").Value + _
        Range("ab50").Value + Range("ab60").Value + Range("ab70").Value + Range("ab80").Value + _
        Range("ab90").Value + Range("ab100").Value + Range("ab110").Value + Range("ab120").Value + Range("ab130").Value > 0 Then
        rightcol = "ac"
    ElseIf Range("w10").Value + Range("w20").Value + Range("w30").Value + Range("w40").Value + _
        Range("w50").Value + Range("w60").Value + Range("w70").Value + Range("w80").Value + _
        Range("w90").Value + Range("w100").Value + Range("w110").Value + Range("w120").Value + Range("w130").Value > 0 Then
        rightcol = "x"
    ElseIf Range("r10").Value + Range("r20").Value + Range("r30").Value + Range("r40").Value + _
        Range("r50").Value + Range("r60").Value + Range("r70").Value + Range("r80").Value + _
        Range("r90").Value + Range("r100").Value + Range("r110").Value + Range("r120").Value + Range("r130").Value > 0 Then
        rightcol = "s"
    Else
        rightcol = "N"
    End If
    
    ActiveSheet.PageSetup.PrintArea = "$a$2:$" + rightcol + "$" & lowrow
    ActiveSheet.PrintOut
    ActiveSheet.PageSetup.PrintArea = "$a$2:$am$129"
End Sub
Sub PrintSig()

Msg = "You are about to print the Signatories page."
Title = "Print Signatories"
Style = vbOKCancel + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)
If response = vbOK Then
   PrintSignatories
End If
Sheets("Contents").Select
Application.ScreenUpdating = True

Msg = "Print Complete!"
Style = vbOK + vbExclamation + vbDefaultButton1
response = MsgBox(Msg, Style, Title)

End Sub
Sub PrintSignatories()
    ' check to see how far down to print
    Sheets("Signatories").Select
    ActiveSheet.PageSetup.CenterHeader = Format(Sheets("contents").Range("E3").Value & vbCr & vbCr & Sheets("contents").Range("E5").Value)
    lowrow = 84
    rightcol = "s"
    For counter = 81 To 5 Step -4
       ' find lowest non-blank section
       If Range("D" & counter).Value <> "" Then
          lowrow = counter + 3
          Exit For
        End If
    Next counter
    
    ' check to see how far across to print
    For counter = Asc("S") To Asc("G") Step -1
        If Range(Chr(counter) & "85").Value > 0 Then
            rightcol = Chr(counter)
            Exit For
        End If
    Next counter
    
    ActiveSheet.PageSetup.PrintArea = "$b$2:$" + rightcol + "$" & lowrow
    ActiveSheet.PrintOut
    ActiveSheet.PageSetup.PrintArea = "$b$2:$s$84"

End Sub



