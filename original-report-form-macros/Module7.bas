Attribute VB_Name = "Module7"
' routines to save local versions from master version
'
'Password to unhide sheets
Const pWord = "SCoE"
Const LpWord = "KCoE"
Sub createlocalversions()

' 1. clear out master of all data
Module6.ClearReport True, True

' 2. create copies of master with proper template names
msg = "You are about to create all new local versions!"
Title = "CREATE Local Reports"
style = vbOKCancel + vbExclamation + vbDefaultButton1
doitresponse = MsgBox(msg, style, Title)
If doitresponse = vbOK Then
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = False
    
    ' unlock/unhide master workbook to make unlocked copies
    Module4.showstuff
    
    Sheets("Contents").Select
    Range("B40") = "LOCAL"
    Range("B38") = "unlocked"
    
    '   LARGE CORP UNLOCKED
    Range("C15") = "Corporate"
    Range("B39") = "LARGE"
    largecorp = "SCAFinancialReportv6_" & Sheets("Contents").Range("B39").Value & "_" & _
                 Sheets("Contents").Range("B38").Value & ".xlsm"
    mycopyfile (largecorp)
    
    '   MEDIUM CORP UNLOCKED
    Range("B39") = "MEDIUM"
    mediumcorp = "SCAFinancialReportv6_" & Sheets("Contents").Range("B39").Value & "_" & _
                 Sheets("Contents").Range("B38").Value & ".xlsm"
    mycopyfile (mediumcorp)
    
    '   SMALL CORP UNLOCKED
    Range("B39") = "SMALL"
    smallcorp = "SCAFinancialReportv6_" & Sheets("Contents").Range("B39").Value & "_" & _
                 Sheets("Contents").Range("B38").Value & ".xlsm"
    mycopyfile (smallcorp)
    
    '   PayPal CORP UNLOCKED
    Range("B39") = "PayPal"
    paypalcorp = "SCAFinancialReportv6_" & Sheets("Contents").Range("B39").Value & "_" & _
                 Sheets("Contents").Range("B38").Value & ".xlsm"
    mycopyfile (paypalcorp)
    
    '   LARGE STATE UNLOCKED
    Range("C15") = "Illinois"
    Range("B39") = "LARGE"
    largesub = "SCASubFinancialReportv6_" & Sheets("Contents").Range("B39").Value & "_" & _
                 Sheets("Contents").Range("B38").Value & ".xlsm"
    mycopyfile (largesub)
    
    '   MEDIUM STATE UNLOCKED
    Range("B39") = "MEDIUM"
    mediumsub = "SCASubFinancialReportv6_" & Sheets("Contents").Range("B39").Value & "_" & _
                 Sheets("Contents").Range("B38").Value & ".xlsm"
    mycopyfile (mediumsub)
    
    '   SMALL STATE UNLOCKED
    Range("B39") = "SMALL"
    smallsub = "SCASubFinancialReportv6_" & Sheets("Contents").Range("B39").Value & "_" & _
                 Sheets("Contents").Range("B38").Value & ".xlsm"
    mycopyfile (smallsub)
        
    '   SMALL Non-US UNLOCKED
    Range("B39") = "SMALL"
    Range("C15") = "Non-US"
    smallxsub = "SCAXUSFinancialReportv6_" & Sheets("Contents").Range("B39").Value & "_" & _
                 Sheets("Contents").Range("B38").Value & ".xlsm"
    mycopyfile (smallxsub)
        
    Range("B39") = "LARGE"
    Range("B40") = "MASTER"
    Application.DisplayAlerts = True
    ActiveWorkbook.Save
    masterworksheet = Activeworksheet
    
    ' 3. open, automatically customize each workbook according to type, and close
    finishsetup (largecorp)
    finishsetup (mediumcorp)
    finishsetup (smallcorp)
    finishsetup (paypalcorp)
    finishsetup (largesub)
    finishsetup (mediumsub)
    finishsetup (smallsub)
    finishsetup (smallxsub)
End If
    
Module4.cleanupsub False

End Sub
Sub finishsetup(workbookname)
Dim obj As OLEObject
'
Workbooks.Open (ActiveWorkbook.Path & "\" & workbookname)
Workbooks(workbookname).Activate
' unlock
For Each sh In Worksheets
    sh.Unprotect (pWord)
Next sh

' fix contents page
Application.StatusBar = "Fix Table of Contents.." & ActiveWorkbook.Name
With Sheets("contents")
    .Range("F7:H27").Locked = False
    .Range("F30:H50").Locked = False
    If .Range("B39") = "SMALL" Then
       ' fix list of links
       .Range("E15:H17").ClearContents
       .Range("E27:H27").ClearContents
       .Range("E30:H48").ClearContents
       If .Range("C15") = "Corporate" Then
          .Range("E49:H49").ClearContents
       End If
    ElseIf .Range("B39") = "MEDIUM" Then
       ' fix list of links
       .Range("E30:H43").ClearContents
       .Range("E45:H48").ClearContents
       If .Range("C15") = "Corporate" Then
          .Range("E49:H49").ClearContents
       End If
    ElseIf .Range("B39") = "LARGE" Then
       ' fix list of links
       .Range("E33:H38").ClearContents
       If .Range("C15") = "Corporate" Then
          .Range("E49:H49").ClearContents
       End If
    Else 'PayPal
       ' fix list of links
       .Range("E15:H17").ClearContents
       .Range("E26:H27").ClearContents
       .Range("E30:H32").ClearContents
       .Range("E39:H49").ClearContents
    ' delete import buttons
       .Shapes.Range("B_ImportLedger").Delete
       .Shapes.Range("B_ImportReport").Delete
    End If
    .Range("F7:H27").Locked = True
    .Range("F30:H50").Locked = True
    
    ' fix dates for non-us
    If .Range("C15") = "Non-US" Then
       .Range("C61") = "=IF(C59="""","""",TEXT(DATE(C63,C59,1),""*dd/mm/yyyy""))"
       .Range("C62") = "=IF(C60="""","""",TEXT(DATE(C63,C60,C64),""*dd/mm/yyyy""))"
       ' delete import buttons
       .Shapes.Range("B_ImportLedger").Delete
    End If
End With

' fix bug email hyperlink.
Sheets("Contents").Select
oldtext = "LARGE"
newtext = Range("B39")
For Each h In ActiveSheet.Hyperlinks
    x = InStr(1, h.Address, oldtext)
    If x > 0 Then
        h.Address = Application.WorksheetFunction.Substitute(h.Address, oldtext, newtext)
    End If
Next

' fix balance statement calculations to remove references to extra sheets not in that version
Application.StatusBar = "Fix Balance Statement.." & ActiveWorkbook.Name
If Sheets("Contents").Range("B39") = "SMALL" Then
   ' ending balance - remove reference to missing pages
   With Sheets("BALANCE_3")
        ' cash
        .Range("H19").Formula = "=IF('PRIMARY_ACCOUNT_2a'!$F$38<>""YES"",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$I$22+'ASSET_DTL_5a'!$G$19"
        .Range("H20").Formula = "=IF('PRIMARY_ACCOUNT_2a'!$F$38=""YES"",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$I$23"
        .Range("G21").Formula = "='ASSET_DTL_5a'!F35"
        .Range("H21").Formula = "='ASSET_DTL_5a'!G35"
        ' assets
        .Range("G22:H25").ClearContents
        ' cash
        .Range("G26").Formula = "='ASSET_DTL_5a'!f46"
        .Range("H26").Formula = "='ASSET_DTL_5a'!g46"
        .Range("G27").Formula = "='ASSET_DTL_5a'!f60"
        .Range("H27").Formula = "='ASSET_DTL_5a'!g60"
        ' liabilities
        .Range("G32").Formula = "='LIABILITY_DTL_5b'!e31"
        .Range("H32").Formula = "='LIABILITY_DTL_5b'!f31"
        .Range("G33").Formula = "='LIABILITY_DTL_5b'!e44"
        .Range("h33").Formula = "='LIABILITY_DTL_5b'!f44"
        .Range("G34").Formula = "='LIABILITY_DTL_5b'!e56"
        .Range("H34").Formula = "='LIABILITY_DTL_5b'!f56"
        
        ' newsletter subs - clear out end cell, reset start cell to no editing
        .Range("H31").ClearContents
        .Range("g31").Interior.Color = Sheets("BALANCE_3").Range("g32").Interior.Color
        .Range("g31").Locked = True
   End With
ElseIf Sheets("Contents").Range("B39") = "MEDIUM" Then
   With Sheets("BALANCE_3")
        ' cash
        .Range("H19").Formula = "=IF('PRIMARY_ACCOUNT_2a'!$F$38<>""YES"",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$I$22+'ASSET_DTL_5a'!$G$19"
        .Range("H20").Formula = "=IF('PRIMARY_ACCOUNT_2a'!F38=""YES"",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$I$23"
        .Range("G21").Formula = "='ASSET_DTL_5a'!F35"
        .Range("H21").Formula = "='ASSET_DTL_5a'!G35"
        ' assets
        .Range("G22").Formula = "='INVENTORY_DTL_6'!M17"
        .Range("H22").Formula = "='INVENTORY_DTL_6'!M27"
        .Range("G23").Formula = "='REGALIA_SALES_DTL_7'!F32"
        .Range("H23").Formula = "='REGALIA_SALES_DTL_7'!I32"
        .Range("G24").Formula = "='DEPR_DTL_8'!I47 + 'REGALIA_SALES_DTL_7'!F49 + 'REGALIA_SALES_DTL_7'!F50 + 'REGALIA_SALES_DTL_7'!F51"
        .Range("H24").Formula = "='DEPR_DTL_8'!J47"
        .Range("G25").Formula = "=-1*('DEPR_DTL_8'!K47+'REGALIA_SALES_DTL_7'!G49+'REGALIA_SALES_DTL_7'!G50+'REGALIA_SALES_DTL_7'!G51)"
        .Range("H25").Formula = "=('DEPR_DTL_8'!M47*-1)"
        ' cash
        .Range("G26").Formula = "='ASSET_DTL_5a'!F46"
        .Range("H26").Formula = "='ASSET_DTL_5a'!G46"
        .Range("G27").Formula = "='ASSET_DTL_5a'!F60"
        .Range("H27").Formula = "='ASSET_DTL_5a'!G60"
        ' liabilities
        .Range("G32").Formula = "='LIABILITY_DTL_5b'!e31"
        .Range("H32").Formula = "='LIABILITY_DTL_5b'!f31"
        .Range("G33").Formula = "='LIABILITY_DTL_5b'!e44"
        .Range("h33").Formula = "='LIABILITY_DTL_5b'!f44"
        .Range("G34").Formula = "='LIABILITY_DTL_5b'!e56"
        .Range("H34").Formula = "='LIABILITY_DTL_5b'!f56"
   End With
ElseIf Sheets("Contents").Range("B39") = "LARGE" Then
   With Sheets("BALANCE_3")
        .Range("G33").Formula = "='LIABILITY_DTL_5b'!e44+'LIABILITY_DTL_5d'!e47"
        .Range("H33").Formula = "='LIABILITY_DTL_5b'!f44+'LIABILITY_DTL_5d'!f47"
   End With
Else 'PayPal
   With Sheets("BALANCE_3")
        ' cash
        .Range("H19").Formula = "=IF('PRIMARY_ACCOUNT_2a'!$F$38<>""YES"",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$i$22+'ASSET_DTL_5a'!$G$19"
        .Range("H20").Formula = "=IF('PRIMARY_ACCOUNT_2a'!F38=""YES"",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$i$23"
        .Range("G21").Formula = "='ASSET_DTL_5a'!F35"
        .Range("H21").Formula = "='ASSET_DTL_5a'!G35"
        ' assets
        .Range("G22:H25").ClearContents
        ' cash
        .Range("G26").Formula = "='ASSET_DTL_5a'!F46"
        .Range("H26").Formula = "='ASSET_DTL_5a'!G46"
        .Range("G27").Formula = "='ASSET_DTL_5a'!F60"
        .Range("H27").Formula = "='ASSET_DTL_5a'!G60"
        ' liabilities
        ' newsletter subs - clear out end cell, reset start cell to no editing
        .Range("H31").ClearContents
        .Range("g31").Interior.Color = Sheets("BALANCE_3").Range("g32").Interior.Color
        .Range("g31").Locked = True
   End With
End If

' fix income statement calculations to remove references to extra sheets not in that version
Application.StatusBar = "Fix Income Statement.." & ActiveWorkbook.Name
If Sheets("Contents").Range("B39") = "SMALL" Then
   With Sheets("INCOME_4")
        ' income
        .Range("j16").Formula = "='TRANSFER_IN_9'!F38"
        .Range("j17").Formula = "='TRANSFER_IN_9'!F58"
        .Range("H19:I19").ClearContents
        .Range("j20").ClearContents
        .Range("j21").ClearContents
        ' expense
        .Range("g30:I30").ClearContents
        .Range("h39").ClearContents
        If Sheets("Contents").Range("C15") = "Corporate" Then
            .Range("j44").Formula = "='EXPENSE_DTL_12b'!I56"
        End If
        .Range("j45").Formula = "='TRANSFER_OUT_10'!F25"
        .Range("j46").Formula = "='TRANSFER_OUT_10'!F52"
   End With
ElseIf Sheets("Contents").Range("B39") = "MEDIUM" Then
   With Sheets("INCOME_4")
        ' income
        .Range("j16").Formula = "='TRANSFER_IN_9'!F38+'TRANSFER_IN_9b'!F32"
        .Range("j17").Formula = "='TRANSFER_IN_9'!F58+'TRANSFER_IN_9b'!F54"
        .Range("h19").Formula = "='INVENTORY_DTL_6'!M30"
        .Range("i19").Formula = "='INVENTORY_DTL_6'!M29"
        .Range("j20").Formula = "='REGALIA_SALES_DTL_7'!I53"
        ' expense
        .Range("g30").Formula = "=SUMIF('DEPR_DTL_8'!$D14:$D23,""OA"",'DEPR_DTL_8'!$L14:$L23)+SUMIF('DEPR_DTL_8'!$D32:$D41,""OA"",'DEPR_DTL_8'!$L32:$L41)"
        .Range("h30").Formula = "=SUMIF('DEPR_DTL_8'!$D14:$D23,""AR"",'DEPR_DTL_8'!$L14:$L23)+SUMIF('DEPR_DTL_8'!$D32:$D41,""AR"",'DEPR_DTL_8'!$L32:$L41)"
        .Range("i30").Formula = "=SUMIF('DEPR_DTL_8'!$D14:$D23,""FR"",'DEPR_DTL_8'!$L14:$L23)+SUMIF('DEPR_DTL_8'!$D32:$D41,""FR"",'DEPR_DTL_8'!$L32:$L41)"
        .Range("h39").Formula = "='REGALIA_SALES_DTL_7'!H52"
        If Sheets("Contents").Range("C15") = "Corporate" Then
            .Range("j44").Formula = "='EXPENSE_DTL_12b'!I56"
        End If
        .Range("j45").Formula = "='TRANSFER_OUT_10'!F25+'TRANSFER_OUT_10b'!F28"
        .Range("j46").Formula = "='TRANSFER_OUT_10'!F52+'TRANSFER_OUT_10b'!F42+'TRANSFER_OUT_10b'!F54"
   End With
ElseIf Sheets("Contents").Range("B39") = "LARGE" Then
   With Sheets("INCOME_4")
        ' income
        ' expense
        If Sheets("Contents").Range("C15") = "Corporate" Then
             .Range("j44").Formula = "='EXPENSE_DTL_12b'!I56"
        End If
   End With
Else 'PayPal
   With Sheets("INCOME_4")
        ' income
        .Range("H19:I19").ClearContents
        .Range("j20").ClearContents
        .Range("j21").ClearContents
        ' expense
        .Range("g30:I30").ClearContents
        .Range("h39").ClearContents
        .Range("j44").Formula = "='EXPENSE_DTL_12b'!I56"
        .Range("j45").Formula = "='TRANSFER_OUT_10'!F25+'TRANSFER_OUT_10b'!F28"
        .Range("j46").Formula = "='TRANSFER_OUT_10'!F52+'TRANSFER_OUT_10b'!F42+'TRANSFER_OUT_10b'!F54"
   End With
End If

' fix income detail calculations to remove references to extra sheets not in that version
If Sheets("Contents").Range("B39") = "SMALL" Or Sheets("Contents").Range("B39") = "PayPal" Then
   Sheets("INCOME_DTL_11a").Range("C35") = ""
   Sheets("INCOME_DTL_11a").Range("E35").ClearContents
ElseIf Sheets("Contents").Range("B39") = "MEDIUM" Then
   Sheets("INCOME_DTL_11a").Range("E35") = "='REGALIA_SALES_DTL_7'!H32"
End If

If Sheets("Contents").Range("C15") = "Corporate" Then
    ' for non-state versions, make state dropdown only have corporate and not be editable
    Sheets("Contents").Range("C15").Locked = True
    Sheets("Contents").Range("C15").Interior.Color = Sheets("contents").Range("B15").Interior.Color
    Sheets("Contents").Range("C15").Validation.Delete
ElseIf Sheets("Contents").Range("C15") = "Non-US" Then
    ' for non-us versions, make state dropdown only have non-us and not be editable
    Sheets("Contents").Range("C15").Locked = True
    Sheets("Contents").Range("C15").Interior.Color = Sheets("contents").Range("B15").Interior.Color
    Sheets("Contents").Range("C15").Validation.Delete
    ' fix donations for proper orgs
    Sheets("EXPENSE_DTL_12b").Range("C46") = Sheets("Corporations").Range("c1") ' corp
    Sheets("EXPENSE_DTL_12b").Range("E46") = Sheets("Corporations").Range("b1") ' corp
    Sheets("EXPENSE_DTL_12b").Range("C46").Interior.Color = Sheets("contents").Range("B15").Interior.Color ' white
    Sheets("EXPENSE_DTL_12b").Range("E46").Interior.Color = Sheets("contents").Range("B15").Interior.Color
    ' add extra instructions for state subsidiary transactions
    Sheets("INCOME_4").Range("D25") = "SCA, Inc. Stock Clerk expenses are General Supplies!"
    Sheets("INCOME_4").Range("D25").Interior.ColorIndex = 40
    Sheets("INCOME_DTL_11a").Range("C31") = "Transfers in from foreign branches (except PayPal) go under a) below!"
    Sheets("INCOME_DTL_11a").Range("C31").Interior.ColorIndex = 40
    Sheets("EXPENSE_DTL_12a").Range("D40") = "Transfers to SCA, Inc. for Insurance go here!"
    Sheets("EXPENSE_DTL_12a").Range("D40").Interior.ColorIndex = 40
    Sheets("EXPENSE_DTL_12b").Range("D43") = "Transfers to foreign branches and kingdom accounts go here!"
    Sheets("EXPENSE_DTL_12b").Range("D43").Interior.ColorIndex = 40
Else
    ' special changes for corporate vs. state subsidiary
    Application.StatusBar = "Fix State forms.." & ActiveWorkbook.Name
    ' copy valid state corps for reference to bottom of EXPENSE_DTL_12b and EXPENSE_DTL_12c
    j = 60
    K = 100
    For I = 3 To 59
        If Sheets("Corporations").Range("B" & I) = "" Then
        Else
            Sheets("Contents").Range("B" & K) = Sheets("Corporations").Range("A" & I)
            K = K + 1
            Sheets("EXPENSE_DTL_12b").Range("C" & j) = Sheets("corporations").Range("c" & I)
            Sheets("EXPENSE_DTL_12b").Range("e" & j) = Sheets("corporations").Range("b" & I)
            Sheets("EXPENSE_DTL_12c").Range("C" & j) = Sheets("corporations").Range("c" & I)
            Sheets("EXPENSE_DTL_12c").Range("e" & j) = Sheets("corporations").Range("b" & I)
            j = j + 1
        End If
    Next I
    
    ' add blank for pdf version
    Sheets("Contents").Range("B" & K) = Sheets("Corporations").Range("A59")
    
    ' for state versions, reset state dropdown to not allow corporate and Non-US as values
    Sheets("Contents").Range("C15").Validation.Modify xlValidateList, xlValidAlertStop, xlEqual, "=$B$100:$B$" & K - 1
    
    ' add extra instructions for state subsidiary transactions
    Sheets("INCOME_4").Range("D25") = "SCA, Inc. Stock Clerk expenses are General Supplies!"
    Sheets("INCOME_4").Range("D25").Interior.ColorIndex = 40
    Sheets("INCOME_DTL_11a").Range("C31") = "Transfers in from out-of-state branches (except PayPal) go under a) below!"
    Sheets("INCOME_DTL_11a").Range("C31").Interior.ColorIndex = 40
    Sheets("EXPENSE_DTL_12a").Range("D40") = "Transfers to SCA, Inc. for Insurance go here!"
    Sheets("EXPENSE_DTL_12a").Range("D40").Interior.ColorIndex = 40
    Sheets("EXPENSE_DTL_12b").Range("D43") = "Transfers to out-of-state branches and kingdom accounts go here!"
    Sheets("EXPENSE_DTL_12b").Range("D43").Interior.ColorIndex = 40
    
    ' fix donations for proper org
    Sheets("EXPENSE_DTL_12b").Range("C46") = Sheets("Corporations").Range("c1") ' corp
    Sheets("EXPENSE_DTL_12b").Range("E46") = Sheets("Corporations").Range("b1") ' corp
    Sheets("EXPENSE_DTL_12b").Range("C46").Interior.Color = Sheets("contents").Range("B15").Interior.Color ' white
    Sheets("EXPENSE_DTL_12b").Range("E46").Interior.Color = Sheets("contents").Range("B15").Interior.Color
    Sheets("EXPENSE_DTL_12b").Range("C46:D46").Locked = True
    Sheets("EXPENSE_DTL_12b").Range("E46").Locked = True
    
    ' change text on transfer forms to refer to states
    Sheets("TRANSFER_IN_9").Range("C12") = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_IN_9").Range("C12"), "country", "state")
    Sheets("TRANSFER_IN_9").Range("C40") = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_IN_9").Range("C40"), "country", "state")
    Sheets("TRANSFER_OUT_10").Select
    Range("C9").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_OUT_10").Range("C9"), "country", "state")
    Range("C27").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_OUT_10").Range("C27"), "country", "state")
    Rows(45).Select      'This copies the selected row
    Selection.Copy
    ActiveSheet.Paste Destination:=Worksheets("TRANSFER_OUT_10").Rows(39)
    ActiveSheet.Paste Destination:=Worksheets("TRANSFER_OUT_10").Rows(40)
    ActiveSheet.Paste Destination:=Worksheets("TRANSFER_OUT_10").Rows(51)
    Range("C52").Value = "TOTAL"
    Range("f52").Value = "=sum(f29:f51)"
    
    ' fix transfer version b
    If Sheets("Contents").Range("B39") <> "SMALL" Then
        Sheets("TRANSFER_IN_9b").Range("C9").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_IN_9b").Range("C9"), "country", "state")
        Sheets("TRANSFER_IN_9b").Range("C34").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_IN_9b").Range("C34"), "country", "state")
        Sheets("TRANSFER_OUT_10b").Select
        Range("C9").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_OUT_10b").Range("C9"), "country", "state")
        Range("C30").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_OUT_10b").Range("C30"), "country", "state")
        Rows(45).Select      'This copies the selected row
        Selection.Copy
        ActiveSheet.Paste Destination:=Worksheets("TRANSFER_OUT_10b").Rows(42)
        ActiveSheet.Paste Destination:=Worksheets("TRANSFER_OUT_10b").Rows(43)
        Range("e54").Value = "TOTAL"
        Range("f54").Value = "=sum(f32:f53)"
    End If
    ' fix transfer versions c and d
    If Sheets("Contents").Range("B39") = "LARGE" Or Sheets("Contents").Range("B39") = "PAYPAL" Then
        Sheets("TRANSFER_IN_9c").Range("C9").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_IN_9c").Range("C9"), "country", "state")
        Sheets("TRANSFER_IN_9c").Range("C34").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_IN_9c").Range("C340"), "country", "state")
        Sheets("TRANSFER_IN_9d").Range("C92").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_IN_9d").Range("C9"), "country", "state")
    End If
    If Sheets("Contents").Range("B39") = "LARGE" Then
        Sheets("TRANSFER_OUT_10c").Select
        Range("C9").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_OUT_10c").Range("C9"), "country", "state")
        Range("C30").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_OUT_10c").Range("C30"), "country", "state")
        Rows(45).Select      'This copies the selected row
        Selection.Copy
        ActiveSheet.Paste Destination:=Worksheets("TRANSFER_OUT_10c").Rows(42)
        ActiveSheet.Paste Destination:=Worksheets("TRANSFER_OUT_10c").Rows(43)
        Range("e54").Value = "TOTAL"
        Range("f54").Value = "=sum(f32:f53)"
        
        Sheets("TRANSFER_OUT_10d").Select
        Range("C9").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_OUT_10d").Range("C9"), "country", "state")
        Range("C30").Formula = Application.WorksheetFunction.Substitute(Sheets("TRANSFER_OUT_10d").Range("C30"), "country", "state")
        Rows(45).Select      'This copies the selected row
        Selection.Copy
        ActiveSheet.Paste Destination:=Worksheets("TRANSFER_OUT_10d").Rows(42)
        ActiveSheet.Paste Destination:=Worksheets("TRANSFER_OUT_10d").Rows(43)
        Range("e54").Value = "TOTAL"
        Range("f54").Value = "=sum(f32:f53)"
    End If
End If

' remove unnecessary workbook pages
ActiveWorkbook.Unprotect (pWord)
Application.StatusBar = "Remove extra pages..."
Application.DisplayAlerts = False
ActiveWorkbook.Sheets("Corporations").Delete
If Sheets("Contents").Range("B39") = "LARGE" Then
    ActiveWorkbook.Sheets("LIABILITY_DTL_5e").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5f").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5g").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5h").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5i").Delete
    If Sheets("Contents").Range("C15") = "Corporate" Then
        ActiveWorkbook.Sheets("EXPENSE_DTL_12c").Delete
    End If
ElseIf Sheets("Contents").Range("B39") = "MEDIUM" Then
    ActiveWorkbook.Sheets("SECONDARY_ACCOUNTS_2c").Delete
    ActiveWorkbook.Sheets("SECONDARY_ACCOUNTS_2d").Delete
    ActiveWorkbook.Sheets("ASSET_DTL_5c").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5d").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5e").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5f").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5g").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5h").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5i").Delete
    ActiveWorkbook.Sheets("INVENTORY_DTL_6b").Delete
    ActiveWorkbook.Sheets("REGALIA_SALES_DTL_7b").Delete
    ActiveWorkbook.Sheets("DEPR_DTL_8b").Delete
    ActiveWorkbook.Sheets("DEPR_DTL_8c").Delete
    ActiveWorkbook.Sheets("TRANSFER_IN_9c").Delete
    ActiveWorkbook.Sheets("TRANSFER_IN_9d").Delete
    ActiveWorkbook.Sheets("TRANSFER_OUT_10c").Delete
    ActiveWorkbook.Sheets("TRANSFER_OUT_10d").Delete
    If Sheets("Contents").Range("C15") = "Corporate" Then
        ActiveWorkbook.Sheets("EXPENSE_DTL_12c").Delete
    End If
ElseIf Sheets("Contents").Range("B39") = "PayPal" Then
    ActiveWorkbook.Sheets("INVENTORY_DTL_6").Delete
    ActiveWorkbook.Sheets("REGALIA_SALES_DTL_7").Delete
    ActiveWorkbook.Sheets("DEPR_DTL_8").Delete
    ActiveWorkbook.Sheets("NEWSLETTER_15").Delete
    ActiveWorkbook.Sheets("SECONDARY_ACCOUNTS_2c").Delete
    ActiveWorkbook.Sheets("SECONDARY_ACCOUNTS_2d").Delete
    ActiveWorkbook.Sheets("INVENTORY_DTL_6b").Delete
    ActiveWorkbook.Sheets("REGALIA_SALES_DTL_7b").Delete
    ActiveWorkbook.Sheets("DEPR_DTL_8b").Delete
    ActiveWorkbook.Sheets("DEPR_DTL_8c").Delete
    ActiveWorkbook.Sheets("TRANSFER_OUT_10c").Delete
    ActiveWorkbook.Sheets("TRANSFER_OUT_10d").Delete
    ActiveWorkbook.Sheets("EXPENSE_DTL_12c").Delete
Else 'small
    ActiveWorkbook.Sheets("INVENTORY_DTL_6").Delete
    ActiveWorkbook.Sheets("REGALIA_SALES_DTL_7").Delete
    ActiveWorkbook.Sheets("DEPR_DTL_8").Delete
    ActiveWorkbook.Sheets("NEWSLETTER_15").Delete
    ActiveWorkbook.Sheets("SECONDARY_ACCOUNTS_2c").Delete
    ActiveWorkbook.Sheets("SECONDARY_ACCOUNTS_2d").Delete
    ActiveWorkbook.Sheets("ASSET_DTL_5c").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5d").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5e").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5f").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5g").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5h").Delete
    ActiveWorkbook.Sheets("LIABILITY_DTL_5i").Delete
    ActiveWorkbook.Sheets("INVENTORY_DTL_6b").Delete
    ActiveWorkbook.Sheets("REGALIA_SALES_DTL_7b").Delete
    ActiveWorkbook.Sheets("DEPR_DTL_8b").Delete
    ActiveWorkbook.Sheets("DEPR_DTL_8c").Delete
    ActiveWorkbook.Sheets("TRANSFER_IN_9b").Delete
    ActiveWorkbook.Sheets("TRANSFER_IN_9c").Delete
    ActiveWorkbook.Sheets("TRANSFER_IN_9d").Delete
    ActiveWorkbook.Sheets("TRANSFER_OUT_10b").Delete
    ActiveWorkbook.Sheets("TRANSFER_OUT_10c").Delete
    ActiveWorkbook.Sheets("TRANSFER_OUT_10d").Delete
    If Sheets("Contents").Range("C15") = "Corporate" Then
        ActiveWorkbook.Sheets("EXPENSE_DTL_12c").Delete
    End If
End If
ActiveWorkbook.Protect (pWord)

' hide all the extra stuff that hasn't been deleted
Module4.hidestuff

' unprotect everything and hide row/column headers
For Each sh In Worksheets
    sh.Unprotect (pWord)
    sh.Select
    ActiveWindow.DisplayHeadings = False
Next sh
Sheets("FreeForm").Select
ActiveWindow.DisplayHeadings = True
Sheets("Contents").Select

' save unlocked version
Application.DisplayAlerts = False
Application.StatusBar = "Saving to file" & ActiveWorkbook.Name
ActiveWorkbook.Save

nworkbookname = Application.WorksheetFunction.Substitute(ActiveWorkbook.Name, "un", "")
nworkbookname = Application.WorksheetFunction.Substitute(nworkbookname, ".xlsm", "")
Sheets("contents").Range("B38").Locked = False
Sheets("contents").Range("B38") = "locked"
Sheets("contents").Range("B38").Locked = True
' protect everything
For Each sh In Worksheets
    sh.Protect (pWord)
Next sh
Sheets("FreeForm").Unprotect (pWord) ' just to make sure

' save locked version
mysavefile (nworkbookname)

' create pdf version
Sheets("contents").Unprotect (pWord)
Sheets("contents").Range("B38").Locked = False
Sheets("contents").Range("B38") = "pdf"
Sheets("contents").Range("B38").Locked = True
Sheets("contents").Range("C12:C13").ClearContents
Sheets("Primary_account_2a").Range("H15").ClearContents
Sheets("Primary_account_2a").Range("F38").ClearContents
Sheets("secondary_accounts_2b").Range("D15:G15").ClearContents
If Sheets("Contents").Range("B39") = "LARGE" Then
    Sheets("secondary_accounts_2c").Range("D15:G15").ClearContents
    Sheets("secondary_accounts_2d").Range("D15:G15").ClearContents
End If
nworkbookname = Application.WorksheetFunction.Substitute(ActiveWorkbook.Name, "locked", "pdf")
For Each sh In Worksheets
    sh.Unprotect (pWord)
Next sh

' save pdf version
mysavefile (nworkbookname)

ActiveWorkbook.Close
End Sub


