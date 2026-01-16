Attribute VB_Name = "Module6"
'=========================================================================
' Module: Module6 (Test data)
' Purpose: Clear report data or populate sheets with sample/test values.
' Key routines: ClearReport, MessIncomeExpense, ClearIncomeExpense
'=========================================================================
' routines to clear test data and create test data
'
'Password to unhide sheets
Const pWord = "SCoE"
Const LpWord = "KCoE"
Sub ClearReport(clearme As Boolean, nomsg As Boolean)
Dim curSheetNames() As String
thisversion = Sheets("Contents").Range("B39")
If Sheets("Contents").Range("B40") = "MASTER" Then thisversion = "MASTER"
  
Application.ScreenUpdating = False
Application.DisplayStatusBar = True

If clearme Then
  msg = "Do you want to save this cleared report workbook to a new file?"
Else
  msg = "Do you want to save this messed up report workbook to a new file?"
End If
style = vbYesNo + vbExclamation + vbDefaultButton1
newfileresponse = False
newfileresponse = MsgBox(msg, style, Title)
If newfileresponse = vbYes Then
   Sheets("Contents").Select
   saveasname = "Report_" & Sheets("Contents").Range("b39").Value & "_" & _
                            Sheets("Contents").Range("b38").Value
   mysavefile (saveasname)
End If

' lock everything up to make sure
Module4.hidestuff

' START Clearing!
If clearme Then
    Sheets("Contents").Range("C8:C11").ClearContents
    Sheets("Contents").Range("C12") = 1
    Sheets("contents").Range("C15") = "Corporate"
Else
    With Sheets("Contents")
        .Range("c8") = "MESSED UP REPORT"
        .Range("C9:C10") = "Someone Important"
        .Range("c11") = 2009
        .Range("c12") = 4
        .Range("C15") = "Corporate"
    End With
End If

' CONTACT_INFO_1
Application.StatusBar = "Contact Info..."
If clearme Then
    With Sheets("CONTACT_INFO_1")
       .Range("D10:h10,D12:h12,D13:D14,F13,F14:h14,H13").ClearContents
       .Range("D15:f16,H15:H16,D18:h18,D19,f19,h19").ClearContents
    
       .Range("e21:h21,D22:h23,D24:D25,F24,F25:h25,H24").ClearContents
       .Range("D26:f27,H26:H27").ClearContents
    
       .Range("e29:H29,D30:h31,D32:d33,F32,F33:h33,H32").ClearContents
       .Range("D34:f35,H34:H35").ClearContents
    End With
Else
    With Sheets("CONTACT_INFO_1")
       .Range("D10,D12,D13:D14,F13,F14,H13,D15:D16,H15:H16") = 1
       .Range("D18,D19,f19,h19") = 1

       .Range("e21,D22,D23,D24:D25") = 1
       .Range("F24,F25,H24,D26:D27,H26:H27") = 1

       .Range("e29,D30,D31,D32:D33") = 1
       .Range("F32,F33,H32,D34:D35,H34:H35") = 1
    End With
End If

' PRIMARY_ACCOUNT_2a
Application.StatusBar = "Primary Account..."
If clearme Then
    With Sheets("PRIMARY_ACCOUNT_2a")
       .Range("E13:h14,E15:E16,h16,F17:h17").ClearContents
       .Range("H15").Value = .Range("C59").Value
       .Range("h16,h19,C21:h23,C27:h34,h37,F38").ClearContents
       .Range("F38") = "No"
       .Range("h40,C44:h53").ClearContents
    End With
Else
    With Sheets("PRIMARY_ACCOUNT_2a")
       .Range("E13:E14,E15:E16,h16,F17") = 1
       .Range("H15").Value = .Range("C59").Value
       .Range("h16,h19,C21:h23,C27:h34,h37") = 1
       .Range("F38") = "Yes"
       .Range("h40,C44:h53") = 1
    End With
End If

' SECONDARY_ACCOUNTS_2b
Application.StatusBar = "Secondary Accounts..."
If clearme Then
    Sheets("SECONDARY_ACCOUNTS_2b").Range("D13:g21,D25:g25,D27:g44").ClearContents
    Sheets("SECONDARY_ACCOUNTS_2b").Range("D15:g15") = Sheets("SECONDARY_ACCOUNTS_2b").Range("C47")
    If thisversion = "LARGE" Or thisversion = "MASTER" Then
       Sheets("SECONDARY_ACCOUNTS_2c").Range("D13:g21,D25:g25,D27:g44").ClearContents
       Sheets("SECONDARY_ACCOUNTS_2d").Range("D13:g21,D25:g25,D27:g44").ClearContents
       Sheets("SECONDARY_ACCOUNTS_2c").Range("D15:g15") = Sheets("SECONDARY_ACCOUNTS_2c").Range("C47")
       Sheets("SECONDARY_ACCOUNTS_2d").Range("D15:g15") = Sheets("SECONDARY_ACCOUNTS_2d").Range("C47")
    End If
Else
    Sheets("SECONDARY_ACCOUNTS_2b").Range("D13:g21,D25:g25,D27:g44") = 1
    Sheets("SECONDARY_ACCOUNTS_2b").Range("D16:g16") = "CD"
    Sheets("SECONDARY_ACCOUNTS_2b").Range("D17:g17") = "No"
    Sheets("SECONDARY_ACCOUNTS_2b").Range("D15:g15") = Sheets("SECONDARY_ACCOUNTS_2b").Range("C47")
    If thisversion = "LARGE" Or thisversion = "MASTER" Then
        Sheets("SECONDARY_ACCOUNTS_2c").Range("D13:g21,D25:g25,D27:g44") = 1
        Sheets("SECONDARY_ACCOUNTS_2c").Range("D16:g16") = "CD"
        Sheets("SECONDARY_ACCOUNTS_2c").Range("D17:g17") = "No"
        Sheets("SECONDARY_ACCOUNTS_2d").Range("D13:g21,D25:g25,D27:g44") = 1
        Sheets("SECONDARY_ACCOUNTS_2d").Range("D16:g16") = "CD"
        Sheets("SECONDARY_ACCOUNTS_2d").Range("D17:g17") = "No"
        Sheets("SECONDARY_ACCOUNTS_2c").Range("D15:g15") = Sheets("SECONDARY_ACCOUNTS_2c").Range("C47")
        Sheets("SECONDARY_ACCOUNTS_2d").Range("D15:g15") = Sheets("SECONDARY_ACCOUNTS_2d").Range("C47")
    End If
End If

' BALANCE_3
Sheets("BALANCE_3").Range("g19:g20").ClearContents
If clearme Then
    If thisversion = "MEDIUM" Or thisversion = "LARGE" Or thisversion = "MASTER" Then
       ' newsletter liability reset
       Sheets("BALANCE_3").Range("g31").ClearContents
    End If
Else
    Sheets("BALANCE_3").Range("g19:g20") = 1
    If thisversion = "MEDIUM" Or thisversion = "LARGE" Or thisversion = "MASTER" Then
       ' newsletter liability reset
       Sheets("BALANCE_3").Range("g31") = 1
    End If
End If

Application.StatusBar = "Cash Assets..."
If clearme Then
    Sheets("ASSET_DTL_5a").Range("c15:g18,c24:g34,c41:g45,c52:g59").ClearContents
    If thisversion = "LARGE" Or thisversion = "PAYPAL" Or thisversion = "MASTER" Then
       Sheets("ASSET_DTL_5c").Range("c13:f32,c39:f43,c50:f57").ClearContents
    End If
   
    If thisversion = "MEDIUM" Or thisversion = "LARGE" Or thisversion = "MASTER" Then
       Application.StatusBar = "Non-cash Assets..."
       Sheets("INVENTORY_DTL_6").Range("E13:l14,E16:l17,E19:l20,E24:l25,E30:l30").ClearContents
       Sheets("REGALIA_SALES_DTL_7").Range("C20:H31,c37:I46,c49:g51,i49:I51").ClearContents
       Sheets("DEPR_DTL_8").Range("d14:g23,j14:j23,e32:g41,j32:j41").ClearContents
        
       If thisversion = "LARGE" Or thisversion = "MASTER" Then
          Sheets("INVENTORY_DTL_6b").Range("E13:l14,E16:l17,E19:l20,E24:l25,E30:l30").ClearContents
          Sheets("REGALIA_SALES_DTL_7b").Range("C20:H31,c37:I46,c49:g51,i49:I51").ClearContents
          Sheets("DEPR_DTL_8b").Range("d14:g53,j14:j53").ClearContents
          Sheets("DEPR_DTL_8c").Range("e14:g53,j14:j53").ClearContents
       End If ' LARGE
    End If ' MEDIUM/LARGE
Else
    Sheets("ASSET_DTL_5a").Range("c15:g18,c24:g34,c41:g45,c52:g59") = 1
    If thisversion = "LARGE" Or thisversion = "PAYPAL" Or thisversion = "MASTER" Then
       Sheets("ASSET_DTL_5c").Range("c13:f32,c39:f43,c50:f57") = 1
    End If
    If thisversion = "MEDIUM" Or thisversion = "LARGE" Or thisversion = "MASTER" Then
       Application.StatusBar = "Non-cash Assets..."
       With Sheets("INVENTORY_DTL_6")
          .Range("E14:l14") = 1
          .Range("E16:l17") = 1
          .Range("E19:l20") = 1
          .Range("E24:l25") = 1
          .Range("E30:l30") = 1
       End With
       With Sheets("REGALIA_SALES_DTL_7")
          .Range("C20:H31") = 1
          .Range("c37:I46") = 1
          .Range("c49:g51") = 1
          .Range("i49:I51") = 1
       End With
       With Sheets("DEPR_DTL_8")
          .Range("d14:g23") = 1
          .Range("j14:j23") = 1
          .Range("e32:g41") = 1
          .Range("j32:j41") = 1
       End With
        
       If thisversion = "LARGE" Or thisversion = "MASTER" Then
          With Sheets("INVENTORY_DTL_6b")
            .Range("E14:l14") = 1
            .Range("E16:l17") = 1
            .Range("E19:l20") = 1
            .Range("E24:l25") = 1
            .Range("E30:l30") = 1
          End With
          With Sheets("REGALIA_SALES_DTL_7b")
            .Range("C20:H31") = 1
            .Range("c37:I46") = 1
            .Range("c49:g51") = 1
            .Range("i49:I51") = 1
          End With
          With Sheets("DEPR_DTL_8b")
            .Range("d14:g53") = 1
            .Range("j14:j53") = 1
          End With
          With Sheets("DEPR_DTL_8c")
            .Range("e14:g53") = 1
            .Range("j14:j53") = 1
          End With
       End If ' LARGE
    End If ' MEDIUM/LARGE
End If
  
Application.StatusBar = "Liabilities..."
If clearme Then
    Sheets("LIABILITY_DTL_5b").Range("c16:f30,c37:f43,c49:f55").ClearContents
    
    If thisversion = "LARGE" Or thisversion = "PAYPAL" Or thisversion = "MASTER" Then
       Sheets("LIABILITY_DTL_5d").Range("c11:f28,c33:f46,c51:f55").ClearContents
       
       If thisversion = "PAYPAL" Or thisversion = "MASTER" Then
          Sheets("LIABILITY_DTL_5e").Range("c11:f55").ClearContents
          Sheets("LIABILITY_DTL_5f").Range("c11:f55").ClearContents
          Sheets("LIABILITY_DTL_5g").Range("c11:f55").ClearContents
          Sheets("LIABILITY_DTL_5h").Range("c11:f55").ClearContents
          Sheets("LIABILITY_DTL_5i").Range("c11:f55").ClearContents
       End If
    End If
Else
    With Sheets("LIABILITY_DTL_5b")
       .Range("c16:f30") = 1
       .Range("c37:f43") = 1
       .Range("c49:f55") = 1
    End With
    If thisversion = "LARGE" Or thisversion = "PAYPAL" Or thisversion = "MASTER" Then
       With Sheets("LIABILITY_DTL_5d")
         .Range("c11:f28") = 1
         .Range("c33:f46") = 1
         .Range("c51:f55") = 1
       End With
       If thisversion = "PAYPAL" Or thisversion = "MASTER" Then
          Sheets("LIABILITY_DTL_5e").Range("c11:f55") = 1
          Sheets("LIABILITY_DTL_5f").Range("c11:f55") = 1
          Sheets("LIABILITY_DTL_5g").Range("c11:f55") = 1
          Sheets("LIABILITY_DTL_5h").Range("c11:f55") = 1
          Sheets("LIABILITY_DTL_5i").Range("c11:f55") = 1
       End If
    End If
End If
   
Application.StatusBar = "Newsletter Subscriptions..."
If clearme Then
    If thisversion = "MEDIUM" Or thisversion = "LARGE" Or thisversion = "MASTER" Then
       Sheets("NEWSLETTER_15").Range("E11:g11,I11,H15:I16,D22:E57,g22:h57,F58,I58").ClearContents
    End If
Else
    If thisversion = "MEDIUM" Or thisversion = "LARGE" Or thisversion = "MASTER" Then
       With Sheets("NEWSLETTER_15")
          .Range("E11") = 1
          .Range("I11") = 1
          .Range("H15:I16") = 1
          .Range("D22:E57") = 1
          .Range("g22:h57") = 1
          .Range("F58") = 1
          .Range("I58") = 1
       End With
    End If
End If

If clearme Then
    ClearIncomeExpense (thisversion)
Else
    MessIncomeExpense (thisversion)
End If
    
Application.StatusBar = "Financial Committee Information..."
If clearme Then
    Sheets("FINANCE_COMM_13").Range("C11:C13,D18,E17:F18,c21:f54").ClearContents
Else
    With Sheets("FINANCE_COMM_13")
        .Range("C11:C13") = "X"
        .Range("D18") = 1
        .Range("E17:F18") = 1
        .Range("c21:f54") = 1
    End With
End If
  
Application.StatusBar = "Fund Balances..."
If clearme Then
    If thisversion <> "SMALL" Then
        Sheets("FUNDS_14").Range("F14:F55,D15:E55").ClearContents
    End If
Else
    If thisversion <> "SMALL" Then
        With Sheets("FUNDS_14")
           .Range("F14:F55") = 1
           .Range("D15:E55") = 1
        End With
    End If
End If

Application.StatusBar = "Comments..."
If clearme Then
    Sheets("COMMENTS").Range("C8:C32").ClearContents
    If Sheets("Contents").Range("C15") <> "Corporate" Then
        Sheets("EXPENSE_DTL_12c").Range("c12:f54").ClearContents
    End If
Else
    Sheets("COMMENTS").Range("C8:C32") = "This is a Comment."
    If Sheets("Contents").Range("C15") <> "Corporate" Then
        Sheets("EXPENSE_DTL_12c").Range("c12:f54") = 1
    End If
End If
Sheets("FreeForm").Columns.Delete

Module4.hidestuff
' reset protection, save workbook, and notify user that we're done!
Module4.cleanupsub (nomsg)

End Sub
Sub MessIncomeExpense(thisvers As String)
    
    Application.StatusBar = "Income..."
    With Sheets("INCOME_4")
       .Range("j18") = 1
       .Range("G29:i29") = 1
       .Range("G31:i31") = 1
       .Range("G33:i34") = 1
       .Range("G36:i38") = 1
       .Range("G40:i41") = 1
    End With
        
    With Sheets("INCOME_DTL_11a")
       .Range("C11:E19") = 1
       .Range("C23:E29") = 1
       .Range("e33:E35") = 1
       .Range("C40:E51") = 1
    End With
    With Sheets("INCOME_DTL_11b")
       .Range("C12:e26") = 1
       .Range("C29:e35") = 1
       .Range("c40:e46") = 1
       .Range("C50:f56") = 1
    End With
    
    Application.StatusBar = "Expenses..."
    With Sheets("EXPENSE_DTL_12a")
       .Range("d12:f22") = 1
       .Range("C27:f38") = 1
       .Range("c43:f54") = 1
    End With
    With Sheets("EXPENSE_DTL_12b")
       .Range("d12:f21") = 1
       .Range("C27:f41") = 1
       .Range("C46:f55") = 1
    End With
        
    Application.StatusBar = "Transfers..."
    With Sheets("TRANSFER_IN_9")
       .Range("C14:f37") = 1
       .Range("C42:f57") = 1
    End With
    With Sheets("TRANSFER_OUT_10")
       .Range("C11:f24") = 1
       .Range("C29:f38") = 1
       .Range("C41:f50") = 1
    End With
        
    If thisvers <> "SMALL" Then
       With Sheets("TRANSFER_IN_9b")
         .Range("C11:f31") = 1
         .Range("C36:f53") = 1
       End With
       With Sheets("TRANSFER_OUT_10b")
         .Range("C11:f27") = 1
         .Range("C32:f41") = 1
         .Range("C44:f53") = 1
       End With
    End If
    If thisvers = "LARGE" Or thisvers = "PAYPAL" Or thisvers = "MASTER" Then
       With Sheets("TRANSFER_IN_9c")
         .Range("C11:f31") = 1
         .Range("C36:f53") = 1
       End With
       With Sheets("TRANSFER_IN_9d")
         .Range("C11:f54") = 1
       End With
    End If
    If thisvers = "LARGE" Or thisvers = "MASTER" Then
       With Sheets("TRANSFER_OUT_10c")
         .Range("C11:f27") = 1
         .Range("C32:f41") = 1
         .Range("C44:f53") = 1
       End With
       With Sheets("TRANSFER_OUT_10d")
         .Range("C11:f27") = 1
         .Range("C32:f41") = 1
         .Range("C44:f53") = 1
       End With
    End If

End Sub

Sub ClearIncomeExpense(thisvers As String)
    Application.StatusBar = "Income..."
    Sheets("INCOME_4").Range("j18,G29:i29,G31:i31,G33:i34,G36:i38,G40:i41").ClearContents
        
'    Sheets("INCOME_DTL_11a").Range("C11:E19,C23:E29,E34:E35,C40:E51").ClearContents
    Sheets("INCOME_DTL_11a").Range("C11:E19,C23:E29,C40:E51").ClearContents
'    Sheets("INCOME_DTL_11b").Range("C12:e26,C29:e35,c40:e42,C49:f56").ClearContents
    Sheets("INCOME_DTL_11b").Range("C12:e26,C29:e35,c40:e42,C49:f55").ClearContents
    Sheets("INCOME_DTL_11C").Range("C13:G23,C35:G65").ClearContents
    
    Application.StatusBar = "Expenses..."
    Sheets("EXPENSE_DTL_12a").Range("d12:f22,C27:f38,c43:f54").ClearContents
    Sheets("EXPENSE_DTL_12b").Range("d12:I15,C21:I26,c31:I55").ClearContents
    If Sheets("Contents").Range("C15") = "Corporate" Then
       Sheets("EXPENSE_DTL_12b").Range("c31:I31").ClearContents
    
       Sheets("EXPENSE_DTL_12b").Range("H31").ClearContents
    End If
    If Sheets("Contents").Range("C15") <> "Corporate" Then
        Sheets("EXPENSE_DTL_12c").Range("c11:I53").ClearContents
    End If
        
    Application.StatusBar = "Transfers..."
    Sheets("TRANSFER_IN_9").Range("C14:f37,C42:f57").ClearContents
    Sheets("TRANSFER_OUT_10").Range("C11:f24,C29:f38,C41:f50").ClearContents
    If Sheets("Contents").Range("C15") <> "Corporate" Then
          Sheets("TRANSFER_OUT_10").Unprotect pWord
          Sheets("TRANSFER_OUT_10").Range("C39:f40").ClearContents
          Sheets("TRANSFER_OUT_10").Range("C37:f37").Copy
          Set targetRange = Sheets("TRANSFER_OUT_10").Range("C39:f40")
          targetRange.PasteSpecial xlPasteFormats
          Sheets("TRANSFER_OUT_10").Protect pWord
    End If
        
    If thisvers <> "SMALL" Then
       Sheets("TRANSFER_IN_9b").Range("C11:f31,C36:f53").ClearContents
       Sheets("TRANSFER_OUT_10b").Range("C11:f27,C32:f41,C44:f53").ClearContents
       If Sheets("Contents").Range("C15") <> "Corporate" Then
          Sheets("TRANSFER_OUT_10b").Unprotect pWord
          Sheets("TRANSFER_OUT_10b").Range("C42:f43").ClearContents
          Sheets("TRANSFER_OUT_10b").Range("C37:f37").Copy
          Set targetRange = Sheets("TRANSFER_OUT_10b").Range("C42:f43")
          targetRange.PasteSpecial xlPasteFormats
          Sheets("TRANSFER_OUT_10b").Protect pWord
       End If
    End If
    
    If thisvers = "LARGE" Or thisvers = "PAYPAL" Or thisvers = "MASTER" Then
       Sheets("TRANSFER_IN_9c").Range("C11:f31,C36:f53").ClearContents
       Sheets("TRANSFER_IN_9d").Range("C11:f54").ClearContents
    End If
    If thisvers = "LARGE" Or thisvers = "MASTER" Then
       Sheets("TRANSFER_OUT_10c").Range("C11:f27,C32:f41,C44:f53").ClearContents
       Sheets("TRANSFER_OUT_10d").Range("C11:f27,C32:f41,C44:f53").ClearContents
       If Sheets("Contents").Range("C15") <> "Corporate" Then
          Sheets("TRANSFER_OUT_10c").Unprotect pWord
          Sheets("TRANSFER_OUT_10c").Range("C42:f43").ClearContents
          Sheets("TRANSFER_OUT_10c").Range("C37:f37").Copy
          Set targetRange = Sheets("TRANSFER_OUT_10c").Range("C42:f43")
          targetRange.PasteSpecial xlPasteFormats
          Sheets("TRANSFER_OUT_10c").Protect pWord
          Sheets("TRANSFER_OUT_10d").Unprotect pWord
          Sheets("TRANSFER_OUT_10d").Range("C42:f43").ClearContents
          Sheets("TRANSFER_OUT_10d").Range("C37:f37").Copy
          Set targetRange = Sheets("TRANSFER_OUT_10d").Range("C42:f43")
          targetRange.PasteSpecial xlPasteFormats
          Sheets("TRANSFER_OUT_10d").Protect pWord
       End If
    End If
End Sub
Sub messreport()
ClearReport False, False
End Sub
Sub unmessreport()
ClearReport True, False
End Sub
