Attribute VB_Name = "Module1"
'=========================================================================
' Module: Module1 (Report printing)
' Purpose: Print required report pages in forward/reverse order or short sets.
' Key routines: PrintBackwards, PrintForwards, PrintFour, fillPage
'=========================================================================
'Password to unhide sheets
Const pWord = "SCoE"
Const LpWord = "KCoE"
' print routines
'
Sub PrintBackwards()
Dim I, num, newnum As Long
Dim msg, style, Title, lnk, response As String

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Printing Backwards..."

Pages = fillPage()
ct = 0
For I = 0 To UBound(Pages)
    If Sheets("Contents").Cells(I + 7, 7).Value = "REQUIRED" Then
       ct = ct + 1
    Else
      ' don't print that page
       Pages(I) = "noprint"
    End If
Next I

msg = "You are about to print " & ct & " pages."
Title = "Print Backwards"
style = vbOKCancel + vbExclamation + vbDefaultButton1
response = MsgBox(msg, style, Title)
If response = vbOK Then
   For I = UBound(Pages) To 0 Step -1
      If Pages(I) <> "noprint" Then
         Sheets(Pages(I)).Select  ' for OO
         Sheets(Pages(I)).PrintOut
      End If
   Next I
End If
Sheets("Contents").Select
Application.ScreenUpdating = True
Application.DisplayStatusBar = False
' notify user that we're done
msg = "Done!"
style = vbOKOnly + vbExclamation + vbDefaultButton1
response = MsgBox(msg, style, Title)

End Sub
Sub PrintForwards()
Dim I, num, newnum As Long
Dim msg, style, Title, lnk, response As String

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Printing Forwards..."

Pages = fillPage()
ct = 0
For I = 0 To UBound(Pages)
    If Sheets("Contents").Cells(I + 7, 7).Value = "REQUIRED" Then
       ct = ct + 1
    Else
      ' don't print that page
       Pages(I) = "noprint"
    End If
Next I

msg = "You are about to print " & ct & " pages."
Title = "Print Forwards"
style = vbOKCancel + vbExclamation + vbDefaultButton1
response = MsgBox(msg, style, Title)
If response = vbOK Then
   For I = 0 To UBound(Pages)
      If Pages(I) <> "noprint" Then
         Sheets(Pages(I)).Select  ' for OO
         Sheets(Pages(I)).PrintOut
      End If
   Next I
End If
Sheets("Contents").Select
Application.ScreenUpdating = True
Application.DisplayStatusBar = False
' notify user that we're done
msg = "Done!"
style = vbOKOnly + vbExclamation + vbDefaultButton1
response = MsgBox(msg, style, Title)

End Sub
Sub PrintFour()
Dim I, num, newnum As Long
Dim msg, style, Title, lnk, response As String
   
Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Printing Forwards..."


Pages = Array("Contents", "CONTACT_INFO_1", "PRIMARY_ACCOUNT_2a", _
        "SECONDARY_ACCOUNTS_2b", "SECONDARY_ACCOUNTS_2c", "SECONDARY_ACCOUNTS_2d", "BALANCE_3", "INCOME_4")
ct = 5
For I = 3 To 5
   If Pages(I) = "SECONDARY_ACCOUNTS_2c" And Sheets("Contents").Range("B39") = "LARGE" Then
      If Sheets("Contents").Range("G29").Value = "REQUIRED" Then
         ct = ct + 1
      End If
   ElseIf Pages(I) = "SECONDARY_ACCOUNTS_2c" Then
       ' don't print that page
       Pages(I) = "noprint"
   End If
   If Pages(I) = "SECONDARY_ACCOUNTS_2d" And Sheets("Contents").Range("B39") = "LARGE" Then
      If Sheets("Contents").Range("G30").Value = "REQUIRED" Then
         ct = ct + 1
      End If
   ElseIf Pages(I) = "SECONDARY_ACCOUNTS_2d" Then
       ' don't print that page
       Pages(I) = "noprint"
   End If
   If Pages(I) = "SECONDARY_ACCOUNTS_2b" And Sheets("Contents").Range("G29").Value = "REQUIRED" Then
       ct = ct + 1
   ElseIf Pages(I) = "SECONDARY_ACCOUNTS_2b" Then
       ' don't print that page
       Pages(I) = "noprint"
   End If
Next I

msg = "You are about to print " & ct & " pages."
Title = "Print Report."
style = vbOKCancel + vbExclamation + vbDefaultButton1
response = MsgBox(msg, style, Title)
If response = vbOK Then
   For I = 0 To UBound(Pages)
       If Pages(I) <> "noprint" Then
          Sheets(Pages(I)).Select  ' for OO
          Sheets(Pages(I)).PrintOut
       End If
   Next I
End If
Sheets("Contents").Select
Application.ScreenUpdating = True
Application.DisplayStatusBar = False
' notify user that we're done
msg = "Done!"
style = vbOKOnly + vbExclamation + vbDefaultButton1
response = MsgBox(msg, style, Title)

End Sub

Function fillPage()
' array needs to be in the same order as the Required / No data checks
' on the Constants page and need to have the same name as the actual pages

' There is an array for all versions.

fillPage = Array("Contents", "CONTACT_INFO_1", _
    "PRIMARY_ACCOUNT_2a", "SECONDARY_ACCOUNTS_2b", _
    "BALANCE_3", "INCOME_4", _
    "ASSET_DTL_5a", "LIABILITY_DTL_5b", _
    "INVENTORY_DTL_6", "REGALIA_SALES_DTL_7", "DEPR_DTL_8", _
    "TRANSFER_IN_9", "TRANSFER_OUT_10", _
    "INCOME_DTL_11a", "INCOME_DTL_11b", "INCOME_DTL_11c", "EXPENSE_DTL_12a", "EXPENSE_DTL_12b", _
    "FINANCE_COMM_13", "FUNDS_14", "NEWSLETTER_15", _
    "COMMENTS", "UNUSED_LINE28", "SECONDARY_ACCOUNTS_2c", "SECONDARY_ACCOUNTS_2d", _
    "ASSET_DTL_5c", "LIABILITY_DTL_5d", "LIABILITY_DTL_5e", "LIABILITY_DTL_5f", "LIABILITY_DTL_5g", _
    "LIABILITY_DTL_5h", "LIABILITY_DTL_5i", _
    "INVENTORY_DTL_6b", "REGALIA_SALES_DTL_7b", _
    "DEPR_DTL_8b", "DEPR_DTL_8c", "TRANSFER_IN_9b", "TRANSFER_IN_9c", "TRANSFER_IN_9d", _
    "TRANSFER_OUT_10b", "TRANSFER_OUT_10c", "TRANSFER_OUT_10d", "EXPENSE_DTL_12c", "FreeForm")

End Function





