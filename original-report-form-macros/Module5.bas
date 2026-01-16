Attribute VB_Name = "Module5"
' routines to reset report for next reporting year
'
'Password to unhide sheets
Const pWord = "SCoE"
Const LpWord = "KCoE"
Sub ResetReport()
Dim curSheetNames() As String
thisversion = Sheets("Contents").Range("B39")

msg = "You are about to reset the entire report workbook for a new quarter!"
Title = "RESET Report"
style = vbOKCancel + vbExclamation + vbDefaultButton1
doitresponse = MsgBox(msg, style, Title)
If doitresponse = vbOK Then
   Application.ScreenUpdating = False
   Application.DisplayStatusBar = True
   
   msg = "Do you want to save this reset report workbook to a new file?"
   style = vbYesNo + vbExclamation + vbDefaultButton1
   newfileresponse = False
   newfileresponse = MsgBox(msg, style, Title)
   If newfileresponse = vbYes Then
       ' if there's a branch name, use it
       Sheets("Contents").Select
       If Application.WorksheetFunction.CountBlank(Sheets("Contents").Range("C8")) = 0 Then
          saveasname = "Report_" & Sheets("Contents").Range("C8").Value & "_"
          If Sheets("Contents").Range("C12").Value = 4 Then
             ' new year
             saveasname = saveasname & (Sheets("Contents").Range("C11").Value + 1) & "_Q1"
          Else
             ' new quarter same year
             saveasname = saveasname & (Sheets("Contents").Range("C11").Value) & "_Q" & (Sheets("Contents").Range("C12").Value + 1)
          End If
       Else
          saveasname = "New_" & ActiveWorkbook.Name
       End If
       mysavefile (saveasname)
    End If
    
    ' START RESETTING!!  increment year or quarter
    Application.StatusBar = "Resetting..."
    If Sheets("Contents").Range("C12").Value = 4 Then
       Sheets("Contents").Range("C11") = Sheets("Contents").Range("C11").Value + 1 ' increment year
       Sheets("Contents").Range("C12") = 1                                         ' reset quarter
    Else
       Sheets("Contents").Range("C12") = Sheets("Contents").Range("C12") + 1       ' increment quarter
    End If
    
    ' CONTACT_INFO_1 - no changes for next timeframe
    ' depreciation - no changes for next timeframe

    ' BALANCE_3 - reset starting balances - do this before messing with any other information
    ' reset starting balances if new year OR sequential quarters
    If Sheets("Contents").Range("C12").Value = 1 Or Sheets("Contents").Range("C12").Value = "Sequential" Then
       Sheets("BALANCE_3").Range("g19:g20") = Sheets("BALANCE_3").Range("h19:h20").Value
       If thisversion = "MEDIUM" Or thisversion = "LARGE" Then
          ' newsletter liability reset
          Sheets("BALANCE_3").Range("g31") = Sheets("BALANCE_3").Range("h31").Value
       End If
    End If
    
    ' PRIMARY_ACCOUNT_2a
    Application.StatusBar = "Accounts..."
    With Sheets("PRIMARY_ACCOUNT_2a")
       .Range("h16").ClearContents ' statement date
       .Range("h19").ClearContents ' statement balance
       .Range("C21:g23").ClearContents ' deposits
       '.range("C27:h34").ClearContents ' uncleared checks
       .Range("h37").ClearContents ' ledger balance
    End With
    
    ' SECONDARY_ACCOUNTS_2b
    With Sheets("SECONDARY_ACCOUNTS_2b")
       .Range("D18:g21").ClearContents ' statement date through withdrawals
       .Range("D25:g25").ClearContents ' ledger balance
    End With
    If thisversion = "LARGE" Then
        With Sheets("SECONDARY_ACCOUNTS_2c")
           .Range("D18:g21").ClearContents ' statement date through withdrawals
           .Range("D25:g25").ClearContents ' ledger balance
        End With
        With Sheets("SECONDARY_ACCOUNTS_2d")
           .Range("D18:g21").ClearContents ' statement date through withdrawals
           .Range("D25:g25").ClearContents ' ledger balance
        End With
    End If
    
    If Sheets("Contents").Range("C12").Value = 1 Or Sheets("Contents").Range("C13").Value = "Sequential" Then
       ' reset for year or sequential quarters means overwriting starting numbers with ending numbers
       Application.StatusBar = "Cash Assets..."
       With Sheets("ASSET_DTL_5a")
           .Range("c15:g18").ClearContents ' undeposited funds
           .Range("f24:f34") = .Range("g24:g34").Value ' receivables
           .Range("f41:f45") = .Range("g41:g45").Value ' prepaid expenses
           .Range("f52:f59") = .Range("g52:g59").Value ' other assets
           .Range("g24:g34,g41:g45,g52:g59").ClearContents
       End With
       If thisversion = "LARGE" Or thisversion = "PAYPAL" Then
            With Sheets("ASSET_DTL_5c")
               .Range("e13:e32") = .Range("f13:f32").Value ' receivables
               .Range("f39:e43") = .Range("f39:f43").Value ' prepaid expenses
               .Range("e50:e57") = .Range("f50:f57").Value ' other assets
               .Range("f13:f32,f39:f43,f50:f57").ClearContents
            End With
       End If
       
       If thisversion = "MEDIUM" Or thisversion = "LARGE" Then
          Application.StatusBar = "Non-cash Assets..."
          With Sheets("INVENTORY_DTL_6")
             .Range("E16:L17") = .Range("E26:L27") ' starting/ending quantity/value
             .Range("E24:L25,E30:L30").ClearContents ' sold/discarded, income
          End With
          With Sheets("REGALIA_SALES_DTL_7")
             .Range("f20:f31") = .Range("I20:I31").Value ' regalia set starting value
             .Range("g20:h31").ClearContents
             .Range("c37:I46,c49:g51,i49:I51").ClearContents ' clear sales of non-inventory
          End With
        
          If thisversion = "LARGE" Then
             With Sheets("INVENTORY_DTL_6b")
                .Range("E16:L17") = .Range("E26:L27") ' starting/ending quantity/value
                .Range("E24:L25,E30:L30").ClearContents ' sold/discarded, income
             End With
             With Sheets("REGALIA_SALES_DTL_7b")
                .Range("f20:f31") = .Range("I20:I31").Value ' regalia set starting value
                .Range("g20:h31").ClearContents
                .Range("c37:I46,c49:g51,i49:I51").ClearContents ' clear sales of non-inventory
             End With
          End If ' LARGE
       End If ' MEDIUM/LARGE
       
       Application.StatusBar = "Liabilities..."
       With Sheets("LIABILITY_DTL_5b")
          .Range("e16:e30") = .Range("f16:f30").Value ' deferred revenue
          .Range("e37:e43") = .Range("f37:f43").Value ' payables
          .Range("e49:e55") = .Range("f49:f55").Value ' other liabilities
          .Range("f16:f30,f37:f42,f49:f55").ClearContents
       End With
       If thisversion = "LARGE" Or thisversion = "PAYPAL" Then
          With Sheets("LIABILITY_DTL_5d")
            .Range("e11:e28") = .Range("f11:f28").Value ' deferred revenue
            .Range("e33:e46") = .Range("f33:f46").Value ' payables
            .Range("e51:e55") = .Range("f51:f55").Value ' other liabilities
            .Range("f11:f28,f33:f46,f51:f55").ClearContents
          End With
          If thisversion = "PAYPAL" Then
             With Sheets("LIABILITY_DTL_5e")
                .Range("e11:e55") = .Range("f11:f55").Value ' payables
                .Range("f11:f55").ClearContents
             End With
             With Sheets("LIABILITY_DTL_5f")
                .Range("e11:e55") = .Range("f11:f55").Value ' payables
                .Range("f11:f55").ClearContents
             End With
             With Sheets("LIABILITY_DTL_5g")
                .Range("e11:e55") = .Range("f11:f55").Value ' payables
                .Range("f11:f55").ClearContents
             End With
          End If
       End If
       
       Application.StatusBar = "Newsletter Subscriptions..."
       If thisversion = "MEDIUM" Or thisversion = "LARGE" Then
          Sheets("NEWSLETTER_15").Range("I11,D22:E57,g22:h57,F58,I58").ClearContents
       End If
    
       ClearIncomeExpense (thisversion)
        
    Else
       ' just clear out ending asset/liability numbers, no overwriting of starting numbers, no clearing out existing transfers/income/expense details
       Application.StatusBar = "Cash Assets..."
       Sheets("ASSET_DTL_5a").Range("c15:g18,g24:g34,g41:g45,g53:g59").ClearContents
       If thisversion = "LARGE" Or thisversion = "PAYPAL" Then
            Sheets("ASSET_DTL_5c").Range("c13:f32,c39:f43,c50:f57").ClearContents
       End If
       If thisversion = "MEDIUM" Or thisversion = "LARGE" Then
          Application.StatusBar = "Clearing Ending Non-cash Asset Values..."
          Sheets("INVENTORY_DTL_6").Range("E24:L25,E30:L30").ClearContents
          Sheets("REGALIA_SALES_DTL_7").Range("f20:f31") = Sheets("REGALIA_SALES_DTL_7").Range("I20:I31").Value ' regalia set starting value
          Sheets("REGALIA_SALES_DTL_7").Range("g20:h31").ClearContents
          Sheets("REGALIA_SALES_DTL_7").Range("H37:I46").ClearContents
        
          If thisversion = "LARGE" Then
             Sheets("INVENTORY_DTL_6b").Range("E24:L25,E30:L30").ClearContents
             Sheets("REGALIA_SALES_DTL_7b").Range("f20:f31") = Sheets("REGALIA_SALES_DTL_7b").Range("I20:I31").Value ' regalia set starting value
             Sheets("REGALIA_SALES_DTL_7b").Range("g20:h31").ClearContents
             Sheets("REGALIA_SALES_DTL_7b").Range("H37:I46").ClearContents
          End If ' LARGE
       End If ' MEDIUM/LARGE
       
       Application.StatusBar = "Liabilities..."
       Sheets("LIABILITY_DTL_5b").Range("f16:f30,f37:f43,f49:f55").ClearContents
       If thisversion = "LARGE" Or thisversion = "PAYPAL" Then
          Sheets("LIABILITY_DTL_5d").Range("f11:f28,f33:f46,f51:f55").ClearContents
          If thisversion = "PAYPAL" Then
             Sheets("LIABILITY_DTL_5e").Range("f11:f55").ClearContents
             Sheets("LIABILITY_DTL_5f").Range("f11:f55").ClearContents
             Sheets("LIABILITY_DTL_5g").Range("f11:f55").ClearContents
          End If ' PAYPAL
       End If ' LARGE/PAYPAL
       
    End If
    
    Application.StatusBar = "Fund Balances..."
    If thisversion <> "SMALL" Then
        Sheets("FUNDS_14").Range("F14:F55").ClearContents
    End If
    
    Application.StatusBar = "Comments..."
    Sheets("COMMENTS").Range("C8:C32").ClearContents
        
    Sheets("FreeForm").Columns.Delete
          
    ' reset protection, save workbook, and notify user that we're done!
    Module4.cleanupsub False
End If 'doitresponse

End Sub
