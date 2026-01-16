Attribute VB_Name = "Module3"
'=========================================================================
' Module: Module3 (Report import/migration)
' Purpose: Import data from another report workbook (v2/v3) into current.
' Key routines: importFromReport, importPage, importOldtoNew
'=========================================================================
'Password to unhide sheets
Const pWord = "SCoE"
Const LpWord = "KCoE"

Dim tgtWB As Object
Dim srcWB As Object
Dim isOO As Boolean

' import from other versions/old reports routines
'
Sub importFromReport()

Dim TgtName As String
Dim srcName As String
Dim fileName As String
Dim newName As String
Dim SrcSize As String
Dim filePath As String
Dim splitChar As String
Dim msg As String
Dim Title As String

Dim TgtSheets() As String
Dim SrcSheets() As String
Dim openWBs() As String

Dim lngIndex As Integer
Dim SrcVersion As Integer
Dim I As Integer
Dim style As Integer
Dim doitresponse As Integer

Dim badReport As Boolean

    '  get list of workbooks so we can check if the newly opened  wb is in the list
    '  OO linux is barfing on wb names with spaces
    '  so we will open the new wb by index rather than by name
    '  unfortunately OO doesn't make the last opened the same as workbooks.count
    '  so we get to  jump through a bunch of hoops
    ReDim openWBs(Workbooks.Count)
    I = 1
    For Each wb In Workbooks
       openWBs(I) = wb.Name
       I = I + 1
    Next wb


    isOO = False

    ' 0. won't work for PAYPAL form
    If Sheets("contents").Range("B39") = "PAYPAL" Then
       MsgBox ("This doesn't apply to PAYPAL form")
       Exit Sub
    End If

    ' 1. ask for save information (new file, same file) and create new file if necessary
    msg = "You are about to import from a different report form! This will overwrite ALL UNSAVED data already in this workbook."
    msg = msg & Chr(13) & Chr(13) & "The report will be saved in a new file based on the imported report's branch name."
    Title = "IMPORT Report"
    style = vbOKCancel + vbExclamation + vbDefaultButton1
    doitresponse = MsgBox(msg, style, Title)
    If doitresponse <> vbOK Then Exit Sub

    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True

    ' 2. get report file name
    filePath = CurDir

    TgtName = ActiveWorkbook.Name
    ActiveWorkbook.Sheets("Contents").Select

    srcName = mygetfile
    If LCase(srcName) = "false" Then
       ' windows says bad file
       Exit Sub
    End If
    Application.StatusBar = "Opening " & srcName

    Application.ScreenUpdating = False
    Workbooks.Open srcName, 0, True

     For Each wb In Workbooks
        tst = wb.Name
        If inArray(openWBs, tst) = False Then
            Set srcWB = wb
            fileName = srcWB.Name
        End If
        If tst = TgtName Then
            Set tgtWB = wb
        End If
     Next wb

    srcWB.Activate
    badReport = False

    ' version 2 & 3 location of definition cells
    SrcSize = UCase(srcWB.Sheets("Contents").Range("B39"))
    TgtSize = UCase(tgtWB.Sheets("Contents").Range("B39"))
    If Left(srcWB.Sheets("Contents").Range("B50"), 7) = "Version" Then
       SrcVersion = 2
    Else
       SrcVersion = 3
    End If

    ' Warn that there may be data loss if sizes are mismatched
    If (TgtSize = "SMALL" And TgtSize <> SrcSize) Or (TgtSize = "MEDIUM" And SrcSize = "LARGE") Then
        msg = "You are about to import from a larger sized Report Form. " & _
                "Not all data may be imported and the Report may no longer be in balance.  Do you wish to continue?"
        Title = "Continue"
        style = vbOKCancel + vbExclamation + vbDefaultButton1
        doitresponse = MsgBox(msg, style, Title)
        If doitresponse <> vbOK Then
            badReport = True
        End If
    End If

   ' 3. validate that incoming report will fit into this version
   '       validate that corp/state difference is okay
   '       cancel if incoming report is incompatible
   Application.StatusBar = "Validating " & fileName & " will fit in this report..."
    If SrcVersion = 3 Then
       If badReport = False And srcWB.Sheets("Contents").Range("C15") <> tgtWB.Sheets("Contents").Range("C15") Then
          badReport = True
          MsgBox ("Corporate/Subsidiary status does not match. Ending...")
       End If
    End If

   ' okay to proceed
   If badReport = False Then
        Application.Calculation = xlManual

        If SrcVersion = 3 Then
            ' get lists of sheets
            ReDim TgtSheets(tgtWB.Sheets.Count - 1)
            For I = 0 To tgtWB.Sheets.Count - 1
                TgtSheets(I) = tgtWB.Sheets(I + 1).Name
            Next I
            ReDim SrcSheets(srcWB.Sheets.Count - 1)
            For I = 0 To srcWB.Sheets.Count - 1
                SrcSheets(I) = srcWB.Sheets(I + 1).Name
            Next I
            ' if sheets match then import the sheet
            For Each rptPage In TgtSheets
                If inArray(SrcSheets, rptPage) Then
                    Call importPage(rptPage, False)
                Else
                    Call importPage(rptPage, True)
                End If
            Next rptPage
        Else
            ' version 2 -- do small sheets first
            Call importOldtoNew("Contents", "Contents", 8, 13, "C", 1, -3, False)
            tgtWB.Sheets("Contents").Range("C14") = "USD $"
            Call importOldtoNew("CONTACT_INFO_1", "CONTACT INFO 4", 10, 35, "D", 5, -1, False)
            Call doPrimary
            If SrcSize <> "SMALL" Then Call importOldtoNew("SECONDARY_ACCOUNTS_2b", "SECONDARY ACCOUNTS 3b", 13, 41, "D", 4, -1, False)
            Application.StatusBar = "BALANCE_3..."
            tgtWB.Sheets("BALANCE_3").Range("G19") = 0
            tgtWB.Sheets("BALANCE_3").Range("G20") = 0
            If tgtWB.Sheets("BALANCE_3").Range("G31").Locked = False Then
                tgtWB.Sheets("BALANCE_3").Range("G31") = 0
            End If
            tgtWB.Sheets("BALANCE_3").Range("G19") = srcWB.Sheets("BALANCE 1").Range("G17")
            tgtWB.Sheets("BALANCE_3").Range("G20") = srcWB.Sheets("BALANCE 1").Range("G18")
            If tgtWB.Sheets("BALANCE_3").Range("G31").Locked = False Then
                tgtWB.Sheets("BALANCE_3").Range("G31") = srcWB.Sheets("BALANCE 1").Range("G28")
            End If
            Call importOldtoNew("INCOME_4", "INCOME 2", 29, 41, "G", 3, -1, False)
            tgtWB.Sheets("INCOME_4").Range("J18") = 0
            tgtWB.Sheets("INCOME_4").Range("J18") = srcWB.Sheets("INCOME 2").Range("J17") 'interest
            Call doAssets
            Call doLiabs
            Call importOldtoNew("TRANSFER_IN_9", "TRANSFER IN 9", 13, 57, "C", 4, -4, False)
            Call importOldtoNew("TRANSFER_OUT_10", "TRANSFER OUT 10", 11, 50, "C", 4, -1, False)
            Call importOldtoNew("INCOME_DTL_11a", "INCOME DTL 11a", 11, 51, "C", 4, -1, False)
            Call doIncomeB
            Call importOldtoNew("EXPENSE_DTL_12a", "EXPENSE DTL 12a", 12, 54, "C", 4, -1, False)
            Call doExpenseB
            Call importOldtoNew("FINANCE_COMM_13", "FINANCE COMM 13", 11, 54, "C", 4, -1, False)
            Call importOldtoNew("COMMENTS", "COMMENTS", 8, 32, "C", 1, -1, False)

            ' Done with import into small!
            If TgtSize <> "SMALL" Then ' keep going
                If SrcSize = "SMALL" Then 'sheets don't exist for import
                    BlankIt = True
                Else
                    BlankIt = False
                End If

                Call importOldtoNew("INVENTORY_DTL_6", "INVENTORY DTL 6", 13, 30, "E", 8, -1, BlankIt)
                Call importOldtoNew("REGALIA_SALES_DTL_7", "REGALIA SALES DTL 7", 20, 51, "C", 7, -3, BlankIt)
                Call importOldtoNew("DEPR_DTL_8", "DEPR DTL 8", 14, 41, "A", 10, -1, BlankIt)
                Call importOldtoNew("FUNDS_14", "FUNDS 14", 14, 55, "D", 3, -1, BlankIt)
                Call importOldtoNew("NEWSLETTER_15", "NEWSLETTER 15", 11, 58, "D", 6, -1, BlankIt)
                Call importOldtoNew("TRANSFER_IN_9b", "TRANSFER IN 9b", 11, 53, "C", 4, -1, BlankIt)
                Call importOldtoNew("TRANSFER_OUT_10b", "TRANSFER OUT 10b", 11, 53, "C", 4, -1, BlankIt)

                If TgtSize <> "MEDIUM" Then ' keep going
                    If SrcSize = "MEDIUM" Then 'sheets don't exist for import
                        BlankIt = True
                    Else
                        BlankIt = False
                    End If
                    Call importOldtoNew("SECONDARY_ACCOUNTS_2c", "SECONDARY ACCOUNTS 3c", 13, 41, "D", 4, -1, BlankIt)
                    Call doAssetsOver(BlankIt)
                    Call doLiabsOver(BlankIt)
                    Call importOldtoNew("INVENTORY_DTL_6b", "INVENTORY DTL 6b", 13, 30, "E", 8, -1, BlankIt)
                    Call importOldtoNew("INVENTORY_DTL_6b", "INVENTORY DTL 6b", 13, 30, "E", 8, -1, BlankIt)
                    Call importOldtoNew("INVENTORY_DTL_6c", "INVENTORY_DTL 6c", 13, 30, "E", 8, -1, BlankIt)
                    Call importOldtoNew("REGALIA_SALES_DTL_7b", "REGALIA SALES DTL 7b", 20, 31, "C", 6, -3, BlankIt)
                    Call importOldtoNew("REGALIA_SALES_DTL_7b", "REGALIA SALES DTL 7b", 37, 46, "C", 7, -3, BlankIt)
                    Call importOldtoNew("REGALIA_SALES_DTL_7b", "REGALIA SALES DTL 7b", 49, 51, "C", 7, -3, True)
                    ' old sheet had more data so concatenate and sum on last line of current rpt
                    If BlankIt = False Then
                        For I = 44 To 50
                           Application.StatusBar = "REGALIA_SALES_DTL_7b... " & I
                           If srcWB.Sheets("REGALIA SALES_DTL 7b").Range("C" & I) <> "" Then _
                              tgtWB.Sheets("REGALIA_SALES_DTL_7b").Range("C46") = tgtWB.Sheets("REGALIA_SALES_DTL_7b").Range("C46") & _
                                   ", " & srcWB.Sheets("REGALIA SALES DTL 7b").Range("C" & I)

                           If srcWB.Sheets("REGALIA SALES DTL 7b").Range("H" & I) <> "" Then _
                              tgtWB.Sheets("REGALIA_SALES_DTL_7b").Range("H46") = tgtWB.Sheets("REGALIA_SALES_DTL_7b").Range("H46") + _
                                   srcWB.Sheets("REGALIA SALES DTL 7b").Range("H" & I)

                           If srcWB.Sheets("REGALIA SALES DTL 7b").Range("I" & I) <> "" Then _
                              tgtWB.Sheets("REGALIA_SALES_DTL_7b").Range("I46") = tgtWB.Sheets("INCOME_DTL_11b").Range("I46") + _
                                   srcWB.Sheets("REGALIA SALES DTL 7b").Range("ID" & I)
                        Next I
                    End If
                    Call importOldtoNew("DEPR_DTL_8b", "DEPR DTL 8b", 14, 28, "A", 10, -1, BlankIt)
                    Call importOldtoNew("DEPR_DTL_8b", "DEPR DTL 8b", 29, 53, "A", 10, -1, True)
                    Call importOldtoNew("DEPR_DTL_8c", "DEPR DTL 8b", 14, 28, "A", 10, 19, BlankIt)
                    Call importOldtoNew("DEPR_DTL_8c", "DEPR DTL 8b", 29, 53, "A", 10, -1, True)
                    Call importOldtoNew("TRANSFER_IN_9b", "TRANSFER IN 9b", 11, 53, "C", 4, -1, BlankIt)
                    Call importOldtoNew("TRANSFER_IN_9c", "TRANSFER IN 9c", 11, 53, "C", 4, -1, BlankIt)
                    Call importOldtoNew("TRANSFER_IN_9d", "TRANSFER IN 9d", 11, 54, "C", 4, -1, BlankIt)
                    Call importOldtoNew("TRANSFER_OUT_10b", "TRANSFER OUT 10b", 11, 53, "C", 4, -1, BlankIt)
                    Call importOldtoNew("TRANSFER_OUT_10c", "TRANSFER OUT 10c", 11, 53, "C", 4, -1, BlankIt)
                    Call importOldtoNew("TRANSFER_OUT_10d", "TRANSFER OUT 10d", 11, 53, "C", 4, -1, BlankIt)
                End If  ' end if large v2 form
           End If  ' end if medium v2 form

        End If  ' end imports
        If isOO = False Then
            Call doFreeForm
        Else
            MsgBox ("It appears you are runnning Open Office. You will have to transfer any data for the Free Form page manually!")
        End If

    End If  ' End OK to do import

    ' 16. save updated report form

    tgtWB.Activate
    tgtWB.Sheets("Contents").Select
    Application.StatusBar = "Closing " & fileName
    srcWB.Saved = True
    srcWB.Close

    If badReport = False Then ' save with new file name
        newName = tgtWB.Sheets("Contents").Range("C8")
        If newName = "" Then newName = "Unnamed Branch"
        newName = "IMP_RPT_" & sanitize(newName) & "_" & tgtWB.Sheets("Contents").Range("C11") _
                         & "_Q" & tgtWB.Sheets("Contents").Range("C12")
        Application.StatusBar = "Saving " & newName
        mysavefile (newName)
    End If
    Application.Calculation = xlAutomatic
    Calculate
    Application.DisplayStatusBar = False
    Application.ScreenUpdating = True
    MsgBox ("Done!")
End Sub
Sub importPage(rptPage, BlankIt)
    If rptPage = "Contents" Then
        StartRow = 8
        EndRow = 14
        col = "C"
        numCols = 1
    ElseIf rptPage = "CONTACT_INFO_1" Then
        StartRow = 10
        EndRow = 35
        col = "D"
        numCols = 5
    ElseIf rptPage = "PRIMARY_ACCOUNT_2a" Then
        StartRow = 13
        EndRow = 51
        col = "C"
        numCols = 7
    ElseIf rptPage = "SECONDARY_ACCOUNTS_2b" Then
        StartRow = 13
        EndRow = 41
        col = "D"
        numCols = 4
    ElseIf rptPage = "BALANCE_3" Then
        StartRow = 19
        EndRow = 31
        col = "G"
        numCols = 1
    ElseIf rptPage = "INCOME_4" Then
        StartRow = 18
        EndRow = 41
        col = "G"
        numCols = 4
    ElseIf rptPage = "ASSET_DTL_5a" Then
        StartRow = 15
        EndRow = 59
        col = "C"
        numCols = 5
    ElseIf rptPage = "LIABILITY_DTL_5b" Then
        StartRow = 16
        EndRow = 55
        col = "C"
        numCols = 5
    ElseIf rptPage = "INVENTORY_DTL_6" Or rptPage = "INVENTORY_DTL_6b" Or rptPage = "INVENTORY_DTL_6c" Then
        StartRow = 13
        EndRow = 30
        col = "E"
        numCols = 8
    ElseIf rptPage = "REGALIA_SALES_DTL_7" Then
        StartRow = 20
        EndRow = 51
        col = "C"
        numCols = 7
    ElseIf rptPage = "DEPR_DTL_8" Then
        StartRow = 14
        EndRow = 41
        col = "A"
        numCols = 10
    ElseIf rptPage = "TRANSFER_IN_9" Then
        StartRow = 14
        EndRow = 57
        col = "C"
        numCols = 4
    ElseIf rptPage = "TRANSFER_OUT_10" Then
        StartRow = 11
        EndRow = 50
        col = "C"
        numCols = 4
    ElseIf rptPage = "INCOME_DTL_11a" Then
        StartRow = 11
        EndRow = 51
        col = "C"
        numCols = 3
    ElseIf rptPage = "INCOME_DTL_11b" Then
        StartRow = 12
        EndRow = 56
        col = "C"
        numCols = 4
    ElseIf rptPage = "EXPENSE_DTL_12a" Then
        StartRow = 12
        EndRow = 54
        col = "C"
        numCols = 4
    ElseIf rptPage = "EXPENSE_DTL_12b" Then
        StartRow = 12
        EndRow = 55
        col = "C"
        numCols = 4
    ElseIf rptPage = "FINANCE_COMM_13" Then
        StartRow = 11
        EndRow = 53
        col = "C"
        numCols = 4
    ElseIf rptPage = "FUNDS_14" Then
        StartRow = 14
        EndRow = 55
        col = "D"
        numCols = 3
    ElseIf rptPage = "NEWSLETTER_15" Then
        StartRow = 11
        EndRow = 58
        col = "D"
        numCols = 6
    ElseIf rptPage = "COMMENTS" Then
        StartRow = 81
        EndRow = 32
        col = "C"
        numCols = 1
    ElseIf rptPage = "SECONDARY_ACCOUNTS_2b" Or rptPage = "SECONDARY ACCOUNTS 2c" Then
        StartRow = 13
        EndRow = 41
        col = "D"
        numCols = 4
    ElseIf rptPage = "ASSET_DTL_5c" Then
        StartRow = 13
        EndRow = 57
        col = "C"
        numCols = 5
    ElseIf rptPage = "LIABILITY_DTL_5d" Then
        StartRow = 11
        EndRow = 55
        col = "C"
        numCols = 5
    ElseIf rptPage = "REGALIA_SALES_DTL_7b" Then
        StartRow = 20
        EndRow = 51
        col = "C"
        numCols = 7
     ElseIf rptPage = "DEPR_DTL_8b" Or rptPage = "DEPR_DTL_8c" Then
        StartRow = 14
        EndRow = 53
        col = "A"
        numCols = 10
     ElseIf rptPage = "TRANSFER_IN_9b" Or rptPage = "TRANSFER_IN_9c" Or rptPage = "TRANSFER_OUT_10b" _
                Or rptPage = "TRANSFER_OUT_10c" Or rptPage = "TRANSFER_OUT_10d" Then
        StartRow = 11
        EndRow = 53
        col = "C"
        numCols = 4
     ElseIf rptPage = "TRANSFER_IN_9d" Then
        StartRow = 11
        EndRow = 54
        col = "C"
        numCols = 4
    End If
    If rptPage <> "FreeForm" Then
        Call importOldtoNew(rptPage, rptPage, StartRow, EndRow, col, numCols, 0, BlankIt)
    End If
End Sub


Sub doFreeForm()
    'Freeform
    Application.StatusBar = "FreeForm... "
    tgtWB.Activate
    tgtWB.Sheets("FreeForm").Select
    Cells.Select
    Selection.ClearContents
    Selection.Delete Shift:=xlUp
On Error Resume Next
    srcWB.Activate
    srcWB.Sheets("FreeForm").Select
    Cells.Select
    ActiveSheet.Shapes.SelectAll
    Selection.Copy

    tgtWB.Activate
    tgtWB.Sheets("FreeForm").Select
    Range("A1").Select

    ActiveSheet.Paste
    Application.CutCopyMode = False
On Error GoTo 0
    tgtWB.Sheets("Contents").Select
    srcWB.Activate
End Sub
Sub doExpenseB()
    Call importOldtoNew("EXPENSE_DTL_12b", "EXPENSE DTL 12b", 12, 21, "D", 3, -1, False)
    Call importOldtoNew("EXPENSE_DTL_12b", "EXPENSE DTL 12b", 27, 41, "C", 4, -2, False)
    Call importOldtoNew("EXPENSE_DTL_12b", "EXPENSE DTL 12b", 47, 55, "C", 4, -3, False)

    If srcWB.Sheets("EXPENSE DTL 12b").Range("C53") <> "" Then
        oldTxt = tgtWB.Sheets("EXPENSE_DTL_12b").Range("C55")
        tgtWB.Sheets("EXPENSE_DTL_12b").Range("C55") = ""
        tgtWB.Sheets("EXPENSE_DTL_12b").Range("C55") = oldTxt & ", " & srcWB.Sheets("EXPENSE DTL 12b").Range("D53")
    End If

    If srcWB.Sheets("EXPENSE DTL 12b").Range("E53") <> "" Then
        oldTxt = tgtWB.Sheets("EXPENSE_DTL_12b").Range("E55")
        tgtWB.Sheets("EXPENSE_DTL_12b").Range("E55") = ""
        tgtWB.Sheets("EXPENSE_DTL_12b").Range("E55") = oldTxt & ", " & srcWB.Sheets("EXPENSE DTL 12b").Range("E53")
    End If

    If srcWB.Sheets("EXPENSE DTL 12b").Range("F53") <> "" Then
        oldTxt = tgtWB.Sheets("EXPENSE_DTL_12b").Range("F55")
        tgtWB.Sheets("EXPENSE_DTL_12b").Range("F55") = 0
        tgtWB.Sheets("EXPENSE_DTL_12b").Range("F55") = oldTxt + srcWB.Sheets("EXPENSE DTL 12b").Range("F53")
    End If
End Sub

Sub doIncomeB()
    Call importOldtoNew("INCOME_DTL_11b", "INCOME DTL 11b", 12, 26, "C", 4, -2, False)

    For I = 25 To 31
        Application.StatusBar = "INCOME_DTL_11b... ... " & I
        If srcWB.Sheets("INCOME DTL 11b").Range("C" & I) <> "" Then
            oldTxt = tgtWB.Sheets("INCOME_DTL_11b").Range("C26")
            tgtWB.Sheets("INCOME_DTL_11b").Range("C26") = ""
            tgtWB.Sheets("INCOME_DTL_11b").Range("C26") = oldTxt & ", " & srcWB.Sheets("INCOME DTL 11b").Range("C" & I)
        End If

        If srcWB.Sheets("INCOME DTL 11b").Range("D" & I) <> "" Then
            oldTxt = tgtWB.Sheets("INCOME_DTL_11b").Range("D26")
            tgtWB.Sheets("INCOME_DTL_11b").Range("D26") = 0
            tgtWB.Sheets("INCOME_DTL_11b").Range("D26") = oldTxt + srcWB.Sheets("INCOME DTL 11b").Range("D" & I)
        End If

        If srcWB.Sheets("INCOME DTL 11b").Range("E" & I) <> "" Then
            oldTxt = tgtWB.Sheets("INCOME_DTL_11b").Range("E6")
            tgtWB.Sheets("INCOME_DTL_11b").Range("E6") = 0
            tgtWB.Sheets("INCOME_DTL_11b").Range("E26") = oldTxt + srcWB.Sheets("INCOME DTL 11b").Range("E" & I)
        End If
    Next I


    For I = 29 To 35
        For j = 0 To 2
            col = Chr(Asc("C") + j)
            Application.StatusBar = "INCOME_DTL_11b... " & col & I
            tgtWB.Sheets("INCOME_DTL_11b").Range(col & I) = ""
            If tgtWB.Sheets("INCOME_DTL_11b").Range(col & I) <> "" Then tgtWB.Sheets("INCOME_DTL_11b").Range(col & I) = 0
        Next j
    Next I

    Call importOldtoNew("INCOME_DTL_11b", "INCOME DTL 11b", 40, 46, "C", 3, -5, False)
    Call importOldtoNew("INCOME_DTL_11b", "INCOME DTL 11b", 50, 56, "C", 4, -5, False)

End Sub
Sub doLiabs()
    Application.StatusBar = "LIABILITY_DTL_5b... "
    For I = 16 To 30
        For j = 0 To 3
            col = Chr(Asc("C") + j)
            Application.StatusBar = "LIABILITY_DTL_5b... " & col & I
            tgtWB.Sheets("LIABILITY_DTL_5b").Range(col & I) = ""
            If tgtWB.Sheets("LIABILITY_DTL_5b").Range(col & I) <> "" Then tgtWB.Sheets("LIABILITY_DTL_5b").Range(col & I) = 0
        Next j
    Next I

    For I = 37 To 43
        For j = 0 To 3
            col = Chr(Asc("C") + j)
            Application.StatusBar = "LIABILITY_DTL_5b... " & col & I
            If tgtWB.Sheets("LIABILITY_DTL_5b").Range(col & I).Locked = False And I < 41 Then
                Call syncCells("LIABILITY_DTL_5b", "COMP BAL DTL 5", col & I, col & I + 7)
            Else
                tgtWB.Sheets("LIABILITY_DTL_5b").Range(col & I) = ""
                If tgtWB.Sheets("LIABILITY_DTL_5b").Range(col & I) <> "" Then tgtWB.Sheets("LIABILITY_DTL_5b").Range(col & I) = 0
            End If
        Next j
    Next I

    For I = 49 To 55
        For j = 0 To 3
            col = Chr(Asc("C") + j)
            Application.StatusBar = "LIABILITY_DTL_5b... " & col & I
            If tgtWB.Sheets("LIABILITY_DTL_5b").Range(col & I).Locked = False And I < 52 Then
                Call syncCells("LIABILITY_DTL_5b", "COMP BAL DTL 5", col & I, col & I + 2)
            Else
                tgtWB.Sheets("LIABILITY_DTL_5b").Range(col & I) = ""
                If tgtWB.Sheets("LIABILITY_DTL_5b").Range(col & I) <> "" Then tgtWB.Sheets("LIABILITY_DTL_5b").Range(col & I) = 0
            End If
        Next j
    Next I
End Sub

Sub doLiabsOver(BlankIt)
    Call importOldtoNew("LIABILITY_DTL_5d", "COMP BAL DTL 5b", 11, 28, "C", 5, -3, True)
    Call importOldtoNew("LIABILITY_DTL_5d", "COMP BAL DTL 5b", 33, 37, "C", 5, 9, BlankIt)
    Call importOldtoNew("LIABILITY_DTL_5d", "COMP BAL DTL 5b", 38, 46, "C", 5, -3, True)
    Call importOldtoNew("LIABILITY_DTL_5d", "COMP BAL DTL 5b", 51, 55, "C", 5, -1, BlankIt)
End Sub

Sub doAssets()
    Call importOldtoNew("ASSET_DTL_5a", "COMP BAL DTL 5", 15, 18, "C", 5, -2, False)
    Call importOldtoNew("ASSET_DTL_5a", "COMP BAL DTL 5", 24, 34, "C", 5, -4, False)


    If srcWB.Sheets("COMP BAL DTL 5").Range("C31") <> "" Then
        oldTxt = tgtWB.Sheets("ASSET_DTL_5a").Range("C34")
        tgtWB.Sheets("ASSET_DTL_5a").Range("C34") = ""
        tgtWB.Sheets("ASSET_DTL_5a").Range("C34") = oldTxt & ", " & srcWB.Sheets("COMP BAL DTL 5").Range("C31")

        oldTxt = tgtWB.Sheets("ASSET_DTL_5a").Range("D34")
        tgtWB.Sheets("ASSET_DTL_5a").Range("D34") = ""
        tgtWB.Sheets("ASSET_DTL_5a").Range("D34") = oldTxt & ", " & srcWB.Sheets("COMP BAL DTL 5").Range("D31")

        oldTxt = tgtWB.Sheets("ASSET_DTL_5a").Range("F34")
        tgtWB.Sheets("ASSET_DTL_5a").Range("F34") = 0
        tgtWB.Sheets("ASSET_DTL_5a").Range("F34") = oldTxt + srcWB.Sheets("COMP BAL DTL 5").Range("F31")

        oldTxt = tgtWB.Sheets("ASSET_DTL_5a").Range("G34")
        tgtWB.Sheets("ASSET_DTL_5a").Range("GD34") = 0
        tgtWB.Sheets("ASSET_DTL_5a").Range("G34") = oldTxt + srcWB.Sheets("COMP BAL DTL 5").Range("G31")
    End If

    For I = 41 To 45
        For j = 0 To 4
            col = Chr(Asc("C") + j)
            Application.StatusBar = "ASSET_DTL_5a... " & col & I
            tgtWB.Sheets("ASSET_DTL_5a").Range(col & I) = ""
            If tgtWB.Sheets("ASSET_DTL_5a").Range(col & I) <> "" Then tgtWB.Sheets("ASSET_DTL_5a").Range(col & I) = 0
        Next j
    Next I

    For I = 52 To 59
        For j = 0 To 4
            col = Chr(Asc("C") + j)
            Application.StatusBar = "ASSET_DTL_5a... " & col & I
            If tgtWB.Sheets("ASSET_DTL_5a").Range(col & I).Locked = False And I < 56 Then
                Call syncCells("ASSET_DTL_5a", "COMP BAL DTL 5", col & I, col & I - 16)
            ElseIf tgtWB.Sheets("ASSET_DTL_5a").Range(col & I).Locked = False Then
                tgtWB.Sheets("ASSET_DTL_5a").Range(col & I) = ""
                If tgtWB.Sheets("ASSET_DTL_5a").Range(col & I) <> "" Then tgtWB.Sheets("ASSET_DTL_5a").Range(col & I) = 0
            End If
        Next j
    Next I
End Sub

Sub doAssetsOver(BlankIt)
    Call importOldtoNew("ASSET_DTL_5c", "COMP BAL DTL 5b", 13, 32, "C", 5, -3, BlankIt)
    Call importOldtoNew("ASSET_DTL_5c", "COMP BAL DTL 5b", 39, 43, "C", 5, -3, True)
    Call importOldtoNew("ASSET_DTL_5c", "COMP BAL DTL 5b", 50, 55, "C", 5, -17, BlankIt)
    Call importOldtoNew("ASSET_DTL_5c", "COMP BAL DTL 5b", 56, 57, "C", 5, -3, True)
End Sub

Sub doPrimary()

    Application.StatusBar = "PRIMARY_ACCOUNT_2a..."
    ' Primary Account  C13 to G51
    ' Bank Info
    For I = 13 To 17
        col = "E"
        sCol = "D"
        If I = 17 Then
            col = "F"
            sCol = "E"
        End If
        Application.StatusBar = "PRIMARY_ACCOUNT_2a... " & col & I
        If tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(col & I).Locked = False Then
            Call syncCells("PRIMARY_ACCOUNT_2a", "PRIMARY ACCOUNT 3a", col & I, sCol & I - 1)
        End If
        If I = 16 Or I = 15 Then
            col = "H"
            sCol = "G"
            If tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(col & I).Locked = False Then
                Call syncCells("PRIMARY_ACCOUNT_2a", "PRIMARY ACCOUNT 3a", col & I, sCol & I - 1)
            End If
        End If
    Next I
    If tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h15") = 1 Then
        tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h15") = ""
        tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h15") = "Single Signature"
    ElseIf tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h15") = 2 Then
        tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h15") = ""
        tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h15") = "Dual Signature"
    End If

    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h19") = 0
    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h19") = srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("G17")

    ' outstanding deposits
    For I = 21 To 23
        For j = 0 To 6
            col = Chr(Asc("C") + j)
            If col = "F" Then
                sCol = "E"
            ElseIf col = "E" Then
                sCol = "D"
            ElseIf col = "H" Then
                sCol = "G"
            Else
                sCol = col
            End If
            Application.StatusBar = "PRIMARY_ACCOUNT_2a... " & col & I
            If tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(col & I).Locked = False Then
                Call syncCells("PRIMARY_ACCOUNT_2a", "PRIMARY ACCOUNT 3a", col & I, sCol & I - 2)
            End If
        Next j
    Next I

    ' outstanding checks
    For I = 27 To 34
        For j = 0 To 6
            col = Chr(Asc("C") + j)
            
            If col = "E" Then
                sCol = "D"
            ElseIf col = "F" Then
                sCol = "E"
            ElseIf col = "H" Then
                sCol = "G"
            Else
                sCol = col
            End If

            Application.StatusBar = "PRIMARY_ACCOUNT_2a... " & col & I
            If tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(col & I).Locked = False And col <> "D" And col <> "G" Then
                Call syncCells("PRIMARY_ACCOUNT_2a", "PRIMARY ACCOUNT 3a", col & I, sCol & I - 3)
            ElseIf tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(col & I).Locked = False Then
               tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(col & I) = ""
               If tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(col & I) <> "" Then tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(col & I) = 0
            End If
        Next j
    Next I

    ' old form had more lines here so concatenate check nums and add amts
    For I = 32 To 33
        Application.StatusBar = "PRIMARY_ACCOUNT_2a... " & I
        If srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("C" & I) <> "" Then
            oldTxt = tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("C34")
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("C34") = ""
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("C34") = oldTxt & ", " & srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("C" & I)
        End If

        If srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("E" & I) <> "" Then
            oldTxt = tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F34")
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F34") = ""
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F34") = oldTxt & ", " & srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("E" & I)
        End If

        If srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("D" & I) <> "" Then
            oldTxt = tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("E34")
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("E34") = 0
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("E34") = oldTxt + srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("D" & I)
        End If

        If srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("G" & I) <> "" Then
            oldTxt = tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("I34")
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("I34") = 0
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("I34") = oldTxt + srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("G" & I)
        End If
    Next I
    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h37") = 0
    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F38") = ""
    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h40") = ""
    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h37") = srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("G36")
    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F38") = srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("E37")
    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h40") = srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("G39")
    If tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F38") <> "Yes" Then tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F38") = "No"

    ' Signatories
    For I = 44 To 51
        For j = 0 To 6
            col = Chr(Asc("C") + j)
            If col = "F" Then
                sCol = "E"
            ElseIf col = "E" Then
                sCol = "D"
            ElseIf col = "H" Then
                sCol = "G"
            Else
                sCol = col
            End If
            Application.StatusBar = "PRIMARY_ACCOUNT_2a... " & col & I
            If tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(col & I).Locked = False And col <> "D" And col <> "H" Then
                Call syncCells("PRIMARY_ACCOUNT_2a", "PRIMARY ACCOUNT 3a", col & I, sCol & I - 1)
            End If
        Next j
    Next I
End Sub


Sub importOldtoNew(tPage, sPage, firstRow, lastRow, startCol, numCols, offset, BlankIt)
    For I = firstRow To lastRow
        For j = 0 To numCols - 1
            col = Chr(Asc(startCol) + j)
            Application.StatusBar = tPage & "... " & col & I
            If tgtWB.Sheets(tPage).Range(col & I).Locked = False Then
                If BlankIt = True Then
                    tgtWB.Sheets(tPage).Range(col & I) = ""
                Else
                    Call syncCells(tPage, sPage, col & I, col & I + offset)
                End If
            End If
        Next j
    Next I
End Sub

Sub syncCells(tPage, sPage, tCell, sCell)
    tgtWB.Sheets(tPage).Range(tCell) = ""
    If tgtWB.Sheets(tPage).Range(tCell) <> "" Then tgtWB.Sheets(tPage).Range(tCell) = 0
    tgtWB.Sheets(tPage).Range(tCell) = srcWB.Sheets(sPage).Range(sCell)
End Sub

'*******************************************************
'  Function mygetfile returns either a path & file name
'  or "false"  branches to OO on error
'*******************************************************
Option Explicit
Public Function mygetfile() As String
    Dim fd As FileDialog
    Dim selectedPath As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select Ledger Workbook to Import"
        .AllowMultiSelect = False

        ' Allow .xls, .xlsx, .xlsm
        .Filters.Clear
        .Filters.Add "Excel Workbooks (*.xls; *.xlsx; *.xlsm)", "*.xls;*.xlsx;*.xlsm"
        .Filters.Add "All Files (*.*)", "*.*"

        If .Show <> -1 Then
            mygetfile = "False"
            Exit Function
        End If

        selectedPath = .SelectedItems(1)
    End With

    mygetfile = selectedPath
End Function



'*******************************************************
'  Function GetFile returns either a path & file name
'  or "false"  for OO
'*******************************************************
Function GetFile() As String
    Dim Dlg As Object
    doc = ThisComponent

    If doc.hasLocation Then
        filePath = doc.getURL()
        FileType = Right(filePath, 4)
        Do While Right(filePath, 1) <> "/"
            filePath = Left(filePath, Len(filePath) - 1)
        Loop
    End If
    GetFile = "false" ' to trigger input box if picker is not there
    On Error GoTo oops
    Dlg = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    Dlg.setDisplayDirectory (filePath)
    If FileType = ".xls" Then
            Dlg.appendFilter("Excel Spreadsheet", "*.xls")
            Dlg.appendFilter("OpenOffice Spreadsheet", "*.ods")
        Else
            Dlg.appendFilter("OpenOffice Spreadsheet", "*.ods")
            Dlg.appendFilter("Excel Spreadsheet", "*.xls")
        End If
    Dlg.appendFilter("All Files", "*.*")
    Dlg.execute()
    On Error Resume Next
    GetFile = Dlg.Files(0)
    GoTo Done
oops:
    If GetFile = "false" Then
        GetFile = InputBox("Please enter the path and the filename you wish to import", "File to Import", filePath)
    End If
    If GetFile = "" Then GetFile = "false"
    If Left(GetFile, 7) <> "file://" And GetFile <> "false" Then
        GetFile = "file:///" & Right(GetFile, Len(GetFile) - 6)
    End If
Done:
End Function

'*******************************************************
'  Function sanitize - removes spaces and non text or numbers from file names
'*******************************************************
Function sanitize(fName As String) As String
    sanitize = ""
    lastChar = ""
    x = Len(fName)
    For I = 1 To x
        mychar = Mid(fName, I, 1)

        tst = Asc(mychar)

        If tst > 96 And tst < 123 Then
            sanitize = sanitize & mychar
            lastChar = mychar
        ElseIf tst > 64 And tst < 91 Then
            sanitize = sanitize & mychar
            lastChar = mychar
        ElseIf tst > 47 And tst < 58 Then
            sanitize = sanitize & mychar
            lastChar = mychar
        ElseIf mychar = "." Or mychar = "-" Or mychar = "&" Or mychar = "(" Or mychar = ")" Or mychar = "[" Or mychar = "]" Then
            sanitize = sanitize & mychar
            lastChar = mychar
        ElseIf lastChar <> "_" Then
            sanitize = sanitize & "_"
            lastChar = "_"
        End If
    Next I
End Function

Function inArray(myRay, str)
    inArray = False
    For Each tst In myRay
        If tst = str Then
           inArray = True
           Exit Function
        End If
    Next tst
End Function





