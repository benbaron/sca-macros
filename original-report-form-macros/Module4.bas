Attribute VB_Name = "Module4"
' common cleanup routines and hide/show routines
'
'Password to unhide sheets
Const pWord = "SCoE"
Const LpWord = "KCoE"
' cleanup routines
Sub cleanupsub(nomsg)
Application.DisplayAlerts = False

Application.StatusBar = "Resetting locks... "
' reset protection
For Each sh In Worksheets
    sh.Protect (pWord)
Next sh
Sheets("FreeForm").Unprotect (pWord) ' just to make sure
Sheets("Contents").Select
Application.ScreenUpdating = True
If Not (ActiveWorkbook.ProtectStructure) Then ActiveWorkbook.Protect (pWord)

'save workbook
mysavefile (ActiveWorkbook.Name)

If Not (nomsg) Then
    ' notify user that we're done and saved!
    msg = "Done!"
    'If Not isOOlocal Then
    msg = msg & " File Saved."
    style = vbOKOnly + vbExclamation + vbDefaultButton1
    newfileresponse = MsgBox(msg, style, Title)
End If
Sheets("Contents").Select
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = False

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
      .Cells.Find("*", after:=.Cells(1), _
        LookIn:=xlFormulas, lookat:=xlWhole, _
        searchdirection:=xlPrevious, _
        searchorder:=xlByRows).Row
    myLastCol = _
      .Cells.Find("*", after:=.Cells(1), _
        LookIn:=xlFormulas, lookat:=xlWhole, _
        searchdirection:=xlPrevious, _
        searchorder:=xlByColumns).Column
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
Sub showstuff()
    For Each sh In Worksheets
        sh.Unprotect (pWord)
        sh.Columns("P:T").Hidden = False
    Next sh
    Sheets("Contents").Range("B38:B40").Locked = False
    Sheets("Contents").Rows("55:200").Hidden = False
    Sheets("contents").Columns("H:z").Hidden = False
    Sheets("Contents").Range("B37:B40").Font.Color = Black
End Sub
Sub hidestuff()
    For Each sh In Worksheets
        sh.Unprotect (pWord)
        If Not (sh.Name = "Contents" And sh.Name = "Free Form") Then
            If Sheets("Contents").Range("b39") = "LARGE" Then
               If Sheets("contents").Range("C15") = "Corporate" Then
                  sh.Columns("P").Hidden = False
                  sh.Columns("q:t").Hidden = True
               Else
                  sh.Columns("T").Hidden = False
                  sh.Columns("p:s").Hidden = True
               End If
            ElseIf Sheets("Contents").Range("b39") = "MEDIUM" Then
               sh.Columns("p").Hidden = True
               sh.Columns("q").Hidden = False
               sh.Columns("r:t").Hidden = True
            ElseIf Sheets("Contents").Range("b39") = "SMALL" Then
               sh.Columns("p:q").Hidden = True
               sh.Columns("r").Hidden = False
               sh.Columns("s:t").Hidden = True
            Else 'PAYPAL
               sh.Columns("p:r").Hidden = True
               sh.Columns("s").Hidden = False
               sh.Columns("t").Hidden = True
            End If
       End If
       sh.Select
       Range("A1").Select
       ActiveWindow.SmallScroll Up:=100, ToLeft:=100
    Next sh
    
    ' hide contents constants/work areas
    Sheets("Contents").Range("B38:B40").Locked = True
    Sheets("Contents").Rows("55:200").Hidden = True
    Sheets("contents").Columns("H:Z").Hidden = True
    Sheets("Contents").Range("B37:B40").Font.Color = RGB(204, 255, 204)
    
    ' protect everything
    For Each sh In Worksheets
        sh.Protect (pWord)
    Next sh
    Sheets("FreeForm").Unprotect (pWord) ' just to make sure
End Sub
Sub mycopyfile(saveasname)
Application.DisplayAlerts = False
Application.StatusBar = "Saving to new file " & saveasname
ActiveWorkbook.SaveCopyAs (ActiveWorkbook.Path & "\" & saveasname) ' , FileFormat:=52   'excel xlsm format
Application.DisplayAlerts = True
End Sub
Sub mysavefile(newsavename As String)
Application.DisplayAlerts = False
saveasname = Module3.sanitize(newsavename)
Application.StatusBar = "Saving to new file " & saveasname
On Error GoTo mustbeOO
ActiveWorkbook.SaveAs (ActiveWorkbook.Path & "\" & saveasname) ' , FileFormat:=52   'excel xlsm format
GoTo filesaved
mustbeOO:

saveOO (saveasname)
filesaved:
Application.DisplayAlerts = True
End Sub
Sub saveOO(saveasname)
    Dim oProp(1) As New com.sun.star.beans.PropertyValue
    Dim oDoc As Variant
    Dim s As String
    Dim fType As String

    oDoc = ThisComponent ' Get the active document

    If oDoc.hasLocation Then
      s = oDoc.getURL()
      fType = Right(s, 4)
      Do While Right(s, 1) <> "/"
         s = Left(s, Len(s) - 1)
       Loop
    End If

    If Right(saveasname, 4) = ".xls" Or Right(saveasname, 4) = ".ods" Then
        saveasname = Left(saveasname, Len(saveasname) - 4)
    End If
        
    saveasname = saveasname & fType
    
    oProp(0).Name = "Overwrite"
    oProp(0).Value = True
    
    oProp(1).Name = "FilterName"
    
    If fType = ".xls" Then
        oProp(1).Value = "MS Excel 97"
    Else
        oProp(1).Value = "StarOffice XML (Calc)"
    End If

    oDoc.storeAsURL(s & saveasname, oProp())
End Sub




