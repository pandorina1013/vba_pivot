Attribute VB_Name = "Module1"
Sub analysis()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim SourceBook As Workbook
    Dim SettingSheet, SourceSheet As Worksheet
    Dim Data As Variant
    Dim Path As String
    Dim f As Variant
    
    Set SettingSheet = ThisWorkbook.Sheets("Setting")
    
    Path = SettingSheet.Cells(3, 3)
    
    Set FSO = CreateObject("Scripting.FileSystemObject")

    For Each f In FSO.GetFolder(Path).Files
        Set SourceBook = Workbooks.Open(f.Path)
        Set SourceSheet = SourceBook.Sheets(1)
        Call add_pivots(SourceBook)
        SourceBook.Sheets("pivot").SaveAs SettingSheet.Cells(4, 3) & "\Analysis_" & SourceBook.Name
        SourceBook.Close False
    Next
    Application.ScreenUpdating = True
    MsgBox ("done:" & vbCrLf & SettingSheet.Cells(4, 3))
End Sub
Sub add_pivots(SourceBook)
    Dim DataSheet, PivotSheet As Worksheet
    Dim PCache As PivotCache
    
    Set DataSheet = SourceBook.Sheets(1)
    DataSheet.Cells(1, 1) = "A"
    
    Set PCache = SourceBook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=DataSheet.UsedRange)
    
    Worksheets.Add
    ActiveSheet.Name = "pivot"
    
    Set PivotSheet = SourceBook.Sheets("pivot")
    
    PCache.CreatePivotTable TableDestination:=PivotSheet.Range("A1"), TableName:="pivot1"
    With PivotSheet.PivotTables("pivot1")
        .PivotFields("A").Orientation = xlRowField
        .PivotFields("A").Position = 1
        .PivotFields("distance").Orientation = xlRowField
        .PivotFields("distance").Position = 2
        .PivotFields("key_resp_2.keys").Orientation = xlColumnField
        .PivotFields("key_resp_2.keys").Position = 1
        .AddDataField ActiveSheet.PivotTables("pivot1").PivotFields("key_resp_2.rt"), "Count / key_resp_2.rt", xlCount
    End With

    PCache.CreatePivotTable TableDestination:=PivotSheet.Range("G1"), TableName:="pivot2"
    With PivotSheet.PivotTables("pivot2")
        .PivotFields("A").Orientation = xlRowField
        .PivotFields("A").Position = 1
        .PivotFields("distance").Orientation = xlRowField
        .PivotFields("distance").Position = 2
        .PivotFields("key_resp_2.keys").Orientation = xlColumnField
        .PivotFields("key_resp_2.keys").Position = 1
        .AddDataField ActiveSheet.PivotTables("pivot1").PivotFields("key_resp_2.rt"), "Average / key_resp_2.rt", xlAverage
    End With
End Sub


Sub get_input_file()
    Dim SettingSheet As Worksheet
    Set SettingSheet = ThisWorkbook.Sheets("Setting")
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            SettingSheet.Cells(3, 3) = .SelectedItems(1)
        End If
    End With
End Sub
Sub get_output_file()
    Dim SettingSheet As Worksheet
    Set SettingSheet = ThisWorkbook.Sheets("Setting")
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            SettingSheet.Cells(4, 3) = .SelectedItems(1)
        End If
    End With
End Sub
