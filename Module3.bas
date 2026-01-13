Attribute VB_Name = "Module3"
'Password to unhide sheets
Const pWord = "KCoE"
Dim keepOff As Boolean


Sub unhidetemplate()
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect (pWord)
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Unprotect 'for OO
    Sheets("Ledger Report Template").Visible = xlSheetVisible
    ActiveWorkbook.Protect (pWord)
End Sub

Sub unprotectit()
    For Each sh In Worksheets
        If sh.Visible Then sh.Select
        sh.Unprotect pWord
        If sh.ProtectContents = True Then sh.Unprotect
        ActiveWindow.DisplayHeadings = True
    Next sh
    keepOff = True
End Sub

Sub protectit()
    For Each sh In Worksheets
        If sh.Name <> "Free_Form" Then
            sh.Protect (pWord)
            If sh.Visible Then
                sh.Select
                ActiveWindow.DisplayHeadings = False
                sh.Range("A1").Select
            End If
        End If
    Next sh
    keepOff = False
End Sub
Sub showstuff()
    Msg = "Do you wish to also leave the Sheets unprotected?"
    Title = "Hi There!"
    Style = vbYesNo + vbDefaultButton1
    doitresponse = MsgBox(Msg, Style, Title)
    
    last = ActiveSheet.Name
    unprotectit
    Sheets("Contents").Rows("32:115").Hidden = False
    Sheets("Contents").Columns("H").Hidden = False
    Sheets("Summary").Columns("L:T").Hidden = False
    Sheets("Summary").Rows("100:180").Hidden = False
    Sheets("Ledger_Q1").Columns("AI:BQ").Hidden = False
    Sheets("Ledger_Q2").Columns("AI:BQ").Hidden = False
    Sheets("Ledger_Q3").Columns("AI:BQ").Hidden = False
    Sheets("Ledger_Q4").Columns("AI:BQ").Hidden = False
    Sheets("Equipment_List").Columns("T:U").Hidden = False
    Sheets("Balances").Columns("AO:BS").Hidden = False
    Sheets("Balances").Rows("10:130").Hidden = False
    Sheets("Signatories").Columns("H:S").Hidden = False
    Sheets("Signatories").Columns("X:AC").Hidden = False
    If doitresponse = vbNo Then protectit
    ' reset cursor to upper left
    For Each sh In Worksheets
        If sh.Visible Then
            sh.Select
            Range("A1").Select
            ActiveWindow.SmallScroll Up:=100, ToLeft:=100
        End If
    Next sh
    Sheets(last).Select
End Sub
Sub hidestuff()
    last = ActiveSheet.Name
    unprotectit
    Sheets("Contents").Rows("32:115").Hidden = True
    Sheets("Contents").Columns("H").Hidden = True
    Sheets("Summary").Columns("L:T").Hidden = True
    Sheets("Summary").Rows("100:180").Hidden = True
    Sheets("Ledger_Q1").Columns("AI:BQ").Hidden = True
    Sheets("Ledger_Q2").Columns("AI:BQ").Hidden = True
    Sheets("Ledger_Q3").Columns("AI:BQ").Hidden = True
    Sheets("Ledger_Q4").Columns("AI:BQ").Hidden = True
    Sheets("Equipment_List").Columns("T:U").Hidden = True
    Sheets("Balances").Columns("AO:BS").Hidden = True
    Sheets("Signatories").Columns("X:AC").Hidden = True
    Module7.hideShowAccounts
    protectit
    ' reset cursor to upper left
    For Each sh In Worksheets
        If sh.Visible Then
            sh.Select
            Range("A1").Select
            ActiveWindow.SmallScroll Up:=100, ToLeft:=100
        End If
    Next sh
    Sheets(last).Select

End Sub
Sub DeleteUnused()
  
Dim myLastRow As Long
Dim myLastCol As Long
Dim wks As Worksheet
Dim dummyRng As Range

For Each wks In ActiveWorkbook.Worksheets
  With wks
    myLastRow = 0
    myLastCol = 0
    Set dummyRng = .UsedRange
    On Error Resume Next
    myLastRow = _
      .Cells.Find("*", After:=.Cells(1), _
        LookIn:=xlFormulas, LookAt:=xlWhole, _
        searchdirection:=xlPrevious, _
        SearchOrder:=xlByRows).row
    myLastCol = _
      .Cells.Find("*", After:=.Cells(1), _
        LookIn:=xlFormulas, LookAt:=xlWhole, _
        searchdirection:=xlPrevious, _
        SearchOrder:=xlByColumns).Column
    On Error GoTo 0

    If myLastRow * myLastCol = 0 Then
        .Columns.Delete
    Else
        .Range(.Cells(myLastRow + 1, 1), _
          .Cells(.Rows.Count, 1)).EntireRow.Delete
        .Range(.Cells(1, myLastCol + 1), _
          .Cells(1, .Columns.Count)).EntireColumn.Delete
    End If
  End With
Next wks
End Sub

Function doProtect()
    doProtect = keepOff
End Function




