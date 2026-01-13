Attribute VB_Name = "Module1"

'*******************************************************
'  Function myGetFile returns either a path & file name
'  or "false"  branches to OO on error
'*******************************************************
Function myGetFile()
    myGetFile = "false"
    On Error GoTo mustbeOO
    myGetFile = Application.GetOpenFilename(Filefilter:="Excel Files (*.xls),*.xls", Title:="Open File to Import")
    Exit Function
mustbeOO:
    myGetFile = GetFile()
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
        GetFile = InputBox("Please enter the path and the filename you wish to
import", "File to Import", filePath)
    End If
    If GetFile = "" Then GetFile = "false"
    If Left(GetFile, 7) <> "file://" And GetFile <> "false" Then
        GetFile = "file:///" & Right(GetFile, Len(GetFile) - 6)
    End If
Done:
End Function


'*******************************************************
'  Save a copy not currently used needs OO version
'*******************************************************
Sub myCopyFile(saveasname)
Application.DisplayAlerts = False
Application.StatusBar = "Saving to new file " & saveasname
ActiveWorkbook.SaveCopyAs (ActiveWorkbook.Path & "\" & saveasname) ' , FileFormat:=56   'excel 97-2003 format
Application.DisplayAlerts = True
End Sub

'*******************************************************
'  Save as newname - On error tries to save as OO
'*******************************************************
Sub mySaveFile(saveasname As String)
    Application.DisplayAlerts = False
    If ActiveWorkbook.Name = saveasname Then
        Application.StatusBar = "Saving file "
    Else
        saveasname = sanitize(saveasname)
        Application.StatusBar = "Saving to new file " & saveasname
    End If
    On Error GoTo mustbeOO
    ActiveWorkbook.SaveAs (ActiveWorkbook.Path & "\" & saveasname) ' , FileFormat:=56   'excel 97-2003 format
    GoTo filesaved
mustbeOO:
    saveOO (saveasname)
filesaved:
Application.DisplayAlerts = True
End Sub

'*******************************************************
'  OO Save as newname
'*******************************************************
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

'*******************************************************
'  Function sanitize - removes spaces and non text or numbers from file names
'*******************************************************
Function sanitize(fName As String) As String
    sanitize = ""
    lastChar = ""
    x = Len(fName)
    For i = 1 To x
        myChar = Mid(fName, i, 1)

        tst = Asc(myChar)

        If tst > 96 And tst < 123 Then
            sanitize = sanitize & myChar
            lastChar = myChar
        ElseIf tst > 64 And tst < 91 Then
            sanitize = sanitize & myChar
            lastChar = myChar
        ElseIf tst > 47 And tst < 58 Then
            sanitize = sanitize & myChar
            lastChar = myChar
        ElseIf myChar = "." Or myChar = "-" Or myChar = "&" Or myChar = "(" Or myChar = ")" Or myChar = "[" Or myChar = "]" Then
            sanitize = sanitize & myChar
            lastChar = myChar
        ElseIf lastChar <> "_" Then
            sanitize = sanitize & "_"
            lastChar = "_"
        End If
    Next i
End Function
'*******************************************************
'  Function inArray tests whether a value is in an array
'*******************************************************
Function inArray(myRay, str)
    inArray = False
    For Each tst In myRay
        If tst = str Then
           inArray = True
           Exit Function
        End If
    Next tst
End Function
