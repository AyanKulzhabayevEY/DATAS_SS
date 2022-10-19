Attribute VB_Name = "DATAS_PAS_SS"
    Public detail_page_count As Integer
    Public dict_codes As Scripting.Dictionary
    Public path_save As String
Sub Each_book(wb_monthly, wb_annual)
    Dim I, row_input, row, LR, WS_Count, page_count As Integer
    
    Application.DisplayAlerts = False
    WS_Count = wb_monthly.Worksheets.Count - detail_page_count
    For I = 2 To WS_Count
        LR = wb_monthly.Worksheets(I).Cells(Rows.Count, 1).End(xlUp).row
        For Each k In dict_codes.Keys
            On Error Resume Next
            With wb_monthly.Worksheets(I).Cells(Application.Match(k, wb_monthly.Worksheets(I).Range("A1:A" & LR), 0), 1)
                .Hyperlinks.Add Anchor:=wb_monthly.Worksheets(I).Cells(Application.Match(k, wb_monthly.Worksheets(I).Range("A1:A" & LR), 0), 1), Address:="", SubAddress:="'" & dict_codes(k) & "'!A1"
                .Value = dict_codes(k)
                .Font.Size = 6.5
                .Font.Name = "Arial Narrow"
            End With
        Next k
        wb_monthly.Worksheets(I).Range("$A$1:$J$" & CStr(LR)).PageBreak = xlPageBreakManual
        With wb_monthly.Worksheets(I).PageSetup
            .PrintTitleRows = "$A$1:$J$5"
            .PrintArea = "$A$1:$J$" & CStr(LR)
            .Zoom = 85
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With
    Next I
   wb_monthly.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=path_save & "\EY_Monthly", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False, _
        From:=2
        
    WS_Count = wb_annual.Worksheets.Count
    For I = 2 To WS_Count
        LR = wb_annual.Worksheets(I).Cells(Rows.Count, 1).End(xlUp).row
        For Each k In dict_codes.Keys
            On Error Resume Next
            temp_row = Application.Match(k, wb_annual.Worksheets(I).Range("A1:A" & LR), 0)
            wb_annual.Worksheets(I).Range("A" & CStr(temp_row) & ":A" & CStr(temp_row + 1)) = dict_codes(k)
        Next k

        wb_annual.Worksheets(I).Range("$A$1:$J$" & CStr(LR)).PageBreak = xlPageBreakManual
        With wb_annual.Worksheets(I).PageSetup
            .PrintTitleRows = "$A$1:$K$6"
            .PrintArea = "$A$1:$J$" & CStr(LR)
            .Zoom = 85
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With

    Next I

    wb_annual.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=path_save & "\EY_All_Annual", _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=False, _
    From:=2
    
    wb_annual.Close savechanges:=True, Filename:=path_save & "\EY_All_Annual_finished.xlsx"
    wb_monthly.Close savechanges:=True, Filename:=path_save & "\EY_Monthly_finished.xlsx"
End Sub

Sub SalarySurvey()
    Dim fp, fldr As Office.FileDialog
    Dim path_monthly, path_annual, path_detail, path_codes As String
    Dim file As Variant

    Application.ScreenUpdating = False
           
    Set fp = Application.FileDialog(msoFileDialogFilePicker)
    With fp
        .Filters.Add "Excel", "*.xlsx", 1
        .Title = "Please select Monthly, Annual, Detailed reports  & Codes"
        .AllowMultiSelect = True
        .InitialFileName = "C:\Users\Ayan.Kulzhabayev\Desktop\RPA\PAS\Salary Survey\"
        If .Show <> -1 Then GoTo err_files
        If .SelectedItems.Count <> 4 Then GoTo err_files
        For Each vrtSelectedItem In .SelectedItems
            If InStr(LCase(vrtSelectedItem), "monthly") Then
                Set wb_monthly = Workbooks.Open(vrtSelectedItem, False)
            ElseIf InStr(LCase(vrtSelectedItem), "annual") Then
                Set wb_annual = Workbooks.Open(vrtSelectedItem, False)
            ElseIf InStr(LCase(vrtSelectedItem), "detailed") Then
                Set wb_details = Workbooks.Open(vrtSelectedItem, False)
            ElseIf InStr(LCase(vrtSelectedItem), "codes") Then
                Set wb_codes = Workbooks.Open(vrtSelectedItem, False)
            Else
err_files:
                MsgBox "Please choose correct files(4), must contains - 'monthly, annual, detailed & codes' in their names", vbCritical
                Exit Sub
            End If
        
        Next
    End With
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder to save"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "Process was interrupted", vbCritical
            wb_annual.Close savechanges:=False
            wb_monthly.Close savechanges:=False
            wb_details.Close savechanges:=False
            wb_codes.Close savechanges:=False
        End If
    path_save = .SelectedItems(1)
    End With
    
    Set fp = Nothing
    Set fldr = Nothing
    
    Call change_details(wb_details, wb_codes, wb_monthly)
    Call Each_book(wb_monthly, wb_annual)
    
End Sub

Sub change_details(wb_details, wb_codes, wb_monthly):
    Dim monthly_page_count As Integer
    Dim code_replace, LR, code, code_prev As String
    Dim ws As Worksheet
    
    monthly_page_count = wb_monthly.Sheets.Count
    LR = wb_codes.Worksheets("JobCodes").Cells(Rows.Count, 1).End(xlUp).row
    Set dict_codes = New Scripting.Dictionary
    detail_page_count = wb_details.Worksheets.Count
    
    For I = 1 To detail_page_count
        On Error Resume Next
        code = wb_details.Worksheets(I).Name
        code_replace = wb_codes.Worksheets("JobCodes").Cells(Application.Match(code, wb_codes.Worksheets("JobCodes").Range("B1:B" & LR), 0), 1).Value
        dict_codes.Add code, code_replace
        If code_prev = code_replace Then
            wb_details.Worksheets(I).Range("B3") = code
        Else
            wb_details.Worksheets(I).Range("B3") = code_replace
            code_prev = code_replace
        End If
        wb_details.Worksheets(I).Name = code_replace
        Set ws = wb_monthly.Worksheets(code_replace)
        If ws Is Nothing Then
            With wb_details.Worksheets(I)
                .PageSetup.PrintArea = "$A$1:$L$" & CStr(wb_details.Worksheets(I).Cells(Rows.Count, 1).End(xlUp).row + 1)
                .Range("B3").Hyperlinks.Add Anchor:=wb_details.Worksheets(I).Range("B3"), Address:="", SubAddress:="'" & wb_monthly.Worksheets(2).Name & "'!A1"
                Application.DisplayAlerts = False
                .Copy after:=wb_monthly.Sheets(monthly_page_count)
            End With
        End If
        monthly_page_count = monthly_page_count + 1
    Next I

    wb_details.Close savechanges:=True, Filename:=path_save & "\EY_detailed report_finished.xlsx"
    wb_codes.Close savechanges:=False
End Sub
