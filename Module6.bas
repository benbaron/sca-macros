Attribute VB_Name = "Module6"
'Password to unhide sheets
Const pWord = "KCoE"
Public MyChng As String
Dim tgtWB As Object
Dim srcWB As Object
'import from Source routines
'
Sub importFromLedger()
Dim sPath As String
Dim targetName As String
Dim TargetVersion As String
Dim TargetSub As String
Dim sourceName As String
Dim badSource As Boolean
Dim strPath() As String
Dim SourceVers() As String
Dim lngIndex As Long
Dim openWBs() As String
Dim secAutomation As MsoAutomationSecurity

MyChng = "Off"

TargetSub = Sheets("contents").Range("C6")

' 1. ask for save information (new file, same file) and create new file if necessary
Msg = "You are about to import from a different Ledger form! This may overwrite ALL UNSAVED data already in this workbook."
Msg = Msg & Chr(13) & Chr(13) & "The Ledger will be saved in a new file based on the imported Ledger's branch name."
Title = "IMPORT Ledger"
Style = vbOKCancel + vbExclamation + vbDefaultButton1
doitresponse = MsgBox(Msg, Style, Title)
If doitresponse <> vbOK Then Exit Sub

Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.DisplayAlerts = False

' 2. get Source file name
filePath = CurDir
targetName = ActiveWorkbook.Name
ActiveWorkbook.Sheets("Contents").Select

sourceName = Module1.myGetFile
If LCase(sourceName) = "false" Then
   ' windows says bad file
   Exit Sub
End If

On Error Resume Next ' don't work in OO
secAutomation = Application.AutomationSecurity
Application.AutomationSecurity = msoAutomationSecurityForceDisable
On Error GoTo 0

'  get list of workbooks so we can check if the newly opened  wb is in the list
'  OO linux is barfing on wb names with spaces
'  so we will open the new wb by index rather than by name
'  unfortunately OO doesn't make the last opened the same as workbooks.count
'  so we get to  jump through a bunch of hoops
ReDim openWBs(Workbooks.Count)
i = 1
For Each wb In Workbooks
   openWBs(i) = wb.Name
   i = i + 1
Next wb

Application.StatusBar = "Opening " & sourceName
Application.ScreenUpdating = False
Workbooks.Open (sourceName)
If Left(sourceName, 7) = "file://" Then
    splitChar = "/"
Else
    splitChar = "\"
End If
strPath() = Split(sourceName, splitChar)  'Put the Parts of our path into an array
lngIndex = UBound(strPath)
sourceName = strPath(lngIndex)   'Get the File Name from our array

For Each wb In Workbooks
   tst = wb.Name
   If inArray(openWBs, tst) = False Then
       Set srcWB = wb
       sourceName = srcWB.Name
   End If
   If tst = targetName Then
       Set tgtWB = wb
   End If
Next wb

srcWB.Activate
badSource = False

' unlock
Application.StatusBar = "Unlocking " & sourceName
For Each sh In Worksheets
  sh.Unprotect pWord
  If sh.ProtectContents = True Then sh.Unprotect ' for OO
Next sh
Application.StatusBar = "Unlocking " & targetName
tgtWB.Activate
For Each sh In Worksheets
  sh.Unprotect pWord
  If sh.ProtectContents = True Then sh.Unprotect ' for OO
Next sh
srcWB.Activate

' location of definition cells
SourceVers() = Split(Sheets("Contents").Range("F46"), " ")
lngIndex = UBound(SourceVers)
If SourceVers(lngIndex) = "3" Then
   sourceversion = 3
   lq1 = "Ledger_Q1"
   lq2 = "Ledger_Q2"
   lq3 = "Ledger_Q3"
   lq4 = "Ledger_Q4"
   eql = "Equipment_List"
ElseIf SourceVers(lngIndex) = "2" Then
   sourceversion = 2
   lq1 = "Ledger Q1"
   lq2 = "Ledger Q2"
   lq3 = "Ledger Q3"
   lq4 = "Ledger Q4"
   eql = "Equipment List"
Else
   sourceversion = 1
   lq1 = "Ledger Q1"
   lq2 = "Ledger Q2"
   lq3 = "Ledger Q3"
   lq4 = "Ledger Q4"
   eql = "Equipment List"
End If

' unhide all sheets
Sheets("Contents").Select
If ActiveWorkbook.ProtectStructure = True Then ActiveWorkbook.Unprotect pWord
If ActiveWorkbook.ProtectStructure = True Then ActiveWorkbook.Unprotect 'for OO
Sheets("Summary").Visible = True
Sheets(lq1).Visible = True
Sheets(lq2).Visible = True
Sheets(lq3).Visible = True
Sheets(lq4).Visible = True
Sheets(eql).Visible = True
ActiveWorkbook.Protect pWord

' if the Target workbook is empty, don't bother checking to see if the name and timeframe matches
If tgtWB.Sheets("Contents").Range("C4") <> "" Or tgtWB.Sheets("Contents").Range("C5") > 0 Then
    ' if we have a branch name in the Target workbook, make sure the Source name matches
    If tgtWB.Sheets("Contents").Range("C4") <> srcWB.Sheets("Contents").Range("C4") Then
        ' branch name doesn't match
        MsgBox ("Branch name does not match. Ending...")
        badSource = True
    End If
    ' check year
    If badSource = False And tgtWB.Sheets("Contents").Range("C5") > 0 Then
        If tgtWB.Sheets("Contents").Range("C5") <> srcWB.Sheets("Contents").Range("C5") Then
            ' year doesn't match
            badSource = True
            MsgBox ("Year does not match. Ending...")
        End If
    End If
    ' check subsidiary
    If badSource = False And srcWB.Sheets("Contents").Range("C6") = tgtWB.Sheets("Contents").Range("C6") Then
    Else
        ' subsidiary status doesn't match
        badSource = True
        MsgBox ("Corporate/Subsidiary status does not match. Ending...")
    End If
End If

' okay to proceed
If badSource = False Then
    ' 4. fill in contents
    Application.StatusBar = "Contents..."
    If tgtWB.Sheets("Contents").Range("C4") = "" Then
        ' fill in branch name, year and corporate/subsidiary info
        tgtWB.Sheets("Contents").Range("C4") = srcWB.Sheets("Contents").Range("C4")
        tgtWB.Sheets("Contents").Range("C5") = srcWB.Sheets("Contents").Range("C5")
        If sourceversion = 3 Then
            tgtWB.Sheets("Contents").Range("C6") = srcWB.Sheets("Contents").Range("C6")
        End If
    End If

    ' 5. fill in Summary
    Application.StatusBar = "Summary..."
    ' checking - fill in with bottom balance from last quarter involved.
    tgtWB.Sheets("Summary").Range("C10:D22") = srcWB.Sheets("Summary").Range("C10:D22").Value
    tgtWB.Sheets("Summary").Range("D26:D35") = srcWB.Sheets("Summary").Range("D26:D35").Value
    tgtWB.Sheets("Summary").Range("G10:H51") = srcWB.Sheets("Summary").Range("G10:H51").Value
    For i = 10 To 22
        rng = "D" & i
        If tgtWB.Sheets("Summary").Range("C" & i).Value <> "" Then
            tgtWB.Sheets("Summary").Range(rng).Interior.ColorIndex = 34
            tgtWB.Sheets("Summary").Range(rng).Locked = False
            tgtWB.Sheets("Summary").Range(rng).FormulaHidden = False
        Else
            tgtWB.Sheets("Summary").Range(rng).Interior.ColorIndex = xlNone
            tgtWB.Sheets("Summary").Range(rng).Locked = True
            tgtWB.Sheets("Summary").Range(rng).FormulaHidden = False
        End If
    Next i
    For i = 11 To 51
        rng = "H" & i
        If tgtWB.Sheets("Summary").Range("G" & i).Value <> "" Then
            tgtWB.Sheets("Summary").Range(rng).Interior.ColorIndex = 34
            tgtWB.Sheets("Summary").Range(rng).Locked = False
            tgtWB.Sheets("Summary").Range(rng).FormulaHidden = False
        Else
            tgtWB.Sheets("Summary").Range(rng).Interior.ColorIndex = xlNone
            tgtWB.Sheets("Summary").Range(rng).Locked = True
            tgtWB.Sheets("Summary").Range(rng).FormulaHidden = False
        End If
    Next i
    

    ' 6. fill in Ledgers
    Application.StatusBar = "Ledger Q1 ..."
    Call copyLedgerRow("Ledger_Q1", lq1, sourceversion)

    Application.StatusBar = "Ledger Q2 ..."
    Call copyLedgerRow("Ledger_Q2", lq2, sourceversion)

    Application.StatusBar = "Ledger Q3 ..."
    Call copyLedgerRow("Ledger_Q3", lq3, sourceversion)

    Application.StatusBar = "Ledger Q4 ..."
    Call copyLedgerRow("Ledger_Q4", lq4, sourceversion)

    ' 7. fill in Equipment
    Application.StatusBar = "Assets..."
    tgtWB.Sheets("Equipment_List").Range("d11:q250") = srcWB.Sheets(eql).Range("d11:q250").Value

    ' fill in version 3 sheets
    If sourceversion = 3 Then
        For i = 1 To 20
            x = 1 + (i * 4)
            r1 = "C" & x
            If srcWB.Sheets("Signatories").Range(r1) = "" Then
              Exit For
            End If
            r2 = "D" & x & ":f" & (x + 1)
            r3 = "D" & (x + 3) & ":F" & (x + 3)
            r4 = "G" & x & ":S" & x
            With tgtWB.Sheets("Signatories")
                If r1 = "C5" Then
                Else
                   .Range(r1) = srcWB.Sheets("Signatories").Range(r1).Value
                End If
                .Range(r2) = srcWB.Sheets("Signatories").Range(r2).Value
                .Range(r3) = srcWB.Sheets("Signatories").Range(r3).Value
                .Range(r4) = srcWB.Sheets("Signatories").Range(r4).Value
            End With
        Next i
        


        For i = 1 To 12
            x = (i * 10) - 7
            If srcWB.Sheets("Balances").Range("A" & (x - 1)).Value = "No Account" Then
               Exit For
            End If
            r1a = "B" & x & ":B" & (x + 5)
          '  r1b = "B" & (x + 4) & ":B" & (x + 5)
            r2 = "c" & (x + 2) & ":n" & (x + 2)
            r3 = "p" & (x + 1) & ":s" & (x + 6)
            r4 = "u" & (x + 1) & ":x" & (x + 6)
            r5 = "z" & (x + 1) & ":ac" & (x + 6)
            r6 = "ae" & (x + 1) & ":ah" & (x + 6)
            r7 = "aj" & (x + 1) & ":am" & (x + 6)
            With tgtWB.Sheets("Balances")
                .Range(r1a) = srcWB.Sheets("Balances").Range(r1a).Value
             '   .Range(r1b) = srcWB.Sheets("Balances").Range(r1b).Value
                .Range(r2) = srcWB.Sheets("Balances").Range(r2).Value
                .Range("E" & (x + 5)) = srcWB.Sheets("Balances").Range("E" & (x + 5))
                .Range("C" & (x + 6)) = srcWB.Sheets("Balances").Range("C" & (x + 6))
                .Range(r3) = srcWB.Sheets("Balances").Range(r3).Value
                .Range(r4) = srcWB.Sheets("Balances").Range(r4).Value
                .Range(r5) = srcWB.Sheets("Balances").Range(r5).Value
                .Range(r6) = srcWB.Sheets("Balances").Range(r6).Value
                .Range(r7) = srcWB.Sheets("Balances").Range(r7).Value
            End With
            ' sho hide stuff on Balances and Signatories
            If tgtWB.Sheets("Summary").Range("C" & i + 10) = "" Then  ' hide the rest of the rows & columns
                tgtWB.Sheets("Signatories").Columns(Chr(Asc("G") + i) & ":S").EntireColumn.Hidden = True
                tgtWB.Sheets("Balances").Rows(i & "1" & ":130").EntireRow.Hidden = True
            Else
                tgtWB.Sheets("Signatories").Columns(Chr(Asc("G") + i) & ":" & Chr(Asc("G") + i)).EntireColumn.Hidden = False 'show the rows & columns
                tgtWB.Sheets("Balances").Rows(i & "1" & ":" & i & "9").EntireRow.Hidden = False
            End If
        Next i
    End If

End If

' 16. save updated Target Ledger
Application.StatusBar = "Closing " & sourceName
srcWB.Saved = True
srcWB.Close
tgtWB.Activate
Application.StatusBar = "Saving " & targetName

On Error Resume Next
Application.AutomationSecurity = secAutomation
On Error GoTo 0

If badSource = False Then
    newName = "IMP_LDGR_" & targetName
    Application.StatusBar = "Saving " & newName
    mySaveFile (newName)
End If

Application.DisplayStatusBar = False
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MyChng = "On"
MsgBox ("Done!")

End Sub

Sub copyLedgerRow(tSheetName, sSheetName, sVersion)
Dim blankDateCount As Integer

    blankDateCount = 0
    With tgtWB.Sheets(tSheetName)
        .Range("D11:E110") = srcWB.Sheets(sSheetName).Range("D11:E110").Value
        If sVersion = 3 Then
            .Range("G11:G110") = srcWB.Sheets(sSheetName).Range("G11:G110").Value
        Else
            For i = 11 To 110
                If .Range("D" & i) = "" Then
                    blankDateCount = blankDateCount + 1
                    ' allow for more than one blank row in the active part of the ledger
                    If blankDateCount > 5 Then Exit For
                End If
    
                If InStr(UCase(Left(srcWB.Sheets(sSheetName).Range("g" & i).Value, 3)), "JAN;FEB;MAR;APR;MAY;JUN;JUL;AUG;SEP;OCT;NOV;DEC") Then
                    .Range("g" & i) = StrConv(srcWB.Sheets(sSheetName).Range("g" & i).Value, vbProperCase)
                Else
                   Select Case srcWB.Sheets(sSheetName).Range("g" & i).Value
                   Case 1
                        .Range("G" & i) = "Jan"
                   Case 2
                        .Range("G" & i) = "Feb"
                   Case 3
                        .Range("G" & i) = "Mar"
                   Case 4
                        .Range("G" & i) = "Apr"
                   Case 5
                        .Range("G" & i) = "May"
                   Case 6
                        .Range("G" & i) = "Jun"
                   Case 7
                        .Range("G" & i) = "Jul"
                   Case 8
                        .Range("G" & i) = "Aug"
                   Case 9
                        .Range("G" & i) = "Sep"
                   Case 10
                        .Range("G" & i) = "Oct"
                   Case 11
                        .Range("G" & i) = "Nov"
                   Case 12
                        .Range("G" & i) = "Dec"
                  End Select
                End If
            Next i
        End If
        .Range("h11:j110") = srcWB.Sheets(sSheetName).Range("h11:j110").Value
        .Range("m11:v110") = srcWB.Sheets(sSheetName).Range("m11:v110").Value
        .Range("x11:ag110") = srcWB.Sheets(sSheetName).Range("x11:ag110").Value
    End With
End Sub


