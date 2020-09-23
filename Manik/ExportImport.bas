Attribute VB_Name = "Module1"
Public Function Export_html(vs As vaSpread, expopath_html As String, expopath_txt As String)
    Dim expo As Boolean
    expo = vs.ExportToHTML(expopath_html, False, "C:\Program Files\Spread30\Samples\LOGFILE.TXT")
    If expo = True Then
        MsgBox "Export complete.", , "Result"
    Else
        MsgBox "Export did not succeed.", , "Result"
    End If
End Function

Public Function Export_excel(vs As vaSpread, expopath_xls As String, expopath_txt As String)
    Dim expo As Boolean
    expo = vs.ExportToExcel(expopath_xls, "Test Sheet 1", expopath_txt)
    If expo = True Then
        MsgBox "Export complete.", , "Result"
    Else
        MsgBox "Export did not succeed.", , "Result"
    End If
    Frm_export_import.OLE1.CreateLink (expopath_xls)
    Frm_export_import.OLE1.Action = 7
End Function

Public Function Import_excel(vs As vaSpread, impopath_xls As String, impopath_txt As String)
    Dim y As Boolean, z As Boolean
    Dim Var As Variant
    Dim impo As Integer, listcount As Integer, handle As Integer, shee As Integer
    Dim List(10) As String

    impo = vs.IsExcelFile(impopath_xls)
    If impo = 1 Then
        shee = InputBox("Enter which sheet u want", "Excel Sheet")
        y = vs.GetExcelSheetList(impopath_xls, List, listcount, impopath_txt, handle, True)
        If y = True Then
            If listcount = 0 Then
                z = vs.ImportExcelSheet(handle, 0)
            Else
                z = vs.ImportExcelSheet(handle, shee)
            End If
            If z = True Then
                MsgBox "Import complete.", , "Result"
            Else
                MsgBox "Import did not succeed.", , "Result"
            End If
        Else
            MsgBox "Cannot return information for Excel file.", , "Result"
        End If
    Else
        MsgBox "File is not an Excel file or is locked and cannot be imported.", , "Invalid File Type or Locked"
    End If
End Function

Public Sub clear_texts(spr As vaSpread)
   With spr
     .Row = 1
     .Col = 1
     .Row2 = .MaxRows
     .Col2 = .MaxCols
     .BlockMode = True
     .Action = ActionClear
     .BlockMode = False
  End With
End Sub

