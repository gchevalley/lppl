Attribute VB_Name = "bas_process_cockpit_datas"
Sub reverse_output()

Dim oBBG As New cls_Bloomberg_Sync


Dim i As Long, j As Long, k As Long, m As Integer

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False


'transfere header
Dim c_last As Integer, l_report As Integer, l_visu As Integer, c_vert As Integer

    c_last = 0
    l_report = 2
    l_visu = 10
    c_vert = 5

Dim tmp_wrksht As Worksheet, base_dataset_wrksht As Worksheet
    Set base_dataset_wrksht = Worksheets(1)

Dim vec_needed_wrksht() As Variant
    vec_needed_wrksht = Array("reverse", "reverse_clean", "report", "visu", "vert", "vert_clean")

Dim found_sheet As Boolean
For i = 0 To UBound(vec_needed_wrksht, 1)
    
    found_sheet = False

    For Each tmp_wrksht In Worksheets
        If tmp_wrksht.name = vec_needed_wrksht(i) Then
            found_sheet = True
        End If
    Next
    
    If found_sheet = False Then
        Set tmp_wrksht = Worksheets.Add(, base_dataset_wrksht)
            tmp_wrksht.name = vec_needed_wrksht(i)
    End If
    
Next i


Worksheets("reverse").Cells.Clear
Worksheets("reverse_clean").Cells.Clear
Worksheets("report").Cells.Clear
Worksheets("visu").Cells.Clear
Worksheets("vert").Cells.Clear
Worksheets("vert_clean").Cells.Clear
    
    
    Worksheets("vert").Cells(6, 4) = "price"
    Worksheets("vert").Cells(7, 4) = "m"
    Worksheets("vert").Cells(8, 4) = "omega"
    Worksheets("vert").Cells(9, 4) = "days until crash"
    
    Worksheets("vert").Cells(11, 4) = "m"
    Worksheets("vert").Cells(12, 4) = "omega"
    Worksheets("vert").Cells(13, 4) = "days until crash"
    
    
    Worksheets("vert_clean").Cells(6, 4) = "price"
    Worksheets("vert_clean").Cells(7, 4) = "m"
    Worksheets("vert_clean").Cells(8, 4) = "omega"
    Worksheets("vert_clean").Cells(9, 4) = "days until crash"
    
    Worksheets("vert_clean").Cells(11, 4) = "m"
    Worksheets("vert_clean").Cells(12, 4) = "omega"
    Worksheets("vert_clean").Cells(13, 4) = "days until crash"
    

k = 2
For i = 1 To 100
    If Worksheets(1).Cells(1, i) = "" Then
        c_last = i - 1
        Exit For
    Else
        Worksheets("reverse").Cells(1, i).value = Worksheets(1).Cells(1, i)
        Worksheets("reverse_clean").Cells(1, i).value = Worksheets(1).Cells(1, i)
    End If
Next i


Dim tmp_ticker As String, tmp_start_point As Long
tmp_ticker = Worksheets(1).Cells(2, 1)
tmp_start_point = 2

Worksheets("report").Select

Dim dic_ticker As New Scripting.Dictionary

For i = 3 To 600000
    
    If Worksheets(1).Cells(i, 1) = "" Then
        Exit For
    Else
    
        If Worksheets(1).Cells(i, 1) <> tmp_ticker Or Worksheets(1).Cells(i + 1, 1) = "" Then
        
            'reverse
            Dim tmp_first_line_ts As Long
            tmp_first_line_ts = k
            For j = i - 1 To tmp_start_point Step -1
                
                For m = 1 To c_last
                    Worksheets("reverse").Cells(k, m) = Worksheets(1).Cells(j, m)
                    Worksheets("reverse_clean").Cells(k, m) = Worksheets(1).Cells(j, m)
                    
                    
                    
                    
                    
                Next m
                
                Dim vec_col_to_empty_clean() As Variant
                vec_col_to_empty_clean = Array(9, 10, 11)
                If Worksheets(1).Cells(j, 9) > 650 Then
                    For m = 0 To UBound(vec_col_to_empty_clean, 1)
                        Worksheets("reverse_clean").Cells(k, vec_col_to_empty_clean(m)) = ""
                    Next m
                End If
                
                k = k + 1
                
            Next j
            
            
            'report
            Dim tmp_range As Range
            Worksheets("report").Cells(l_report, 1) = "name"
            Worksheets("report").Cells(l_report + 1, 1) = tmp_ticker
            Worksheets("report").Cells(l_report + 1, 2) = Worksheets("reverse").Cells(i - 1, 5)
                
                Set tmp_range = Range(Worksheets("report").Cells(l_report, 3).Address, Worksheets("report").Cells(l_report + 1, 3).Address)
                    tmp_range.Merge
                    
                    Worksheets("report").Range("c" & l_report).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!E" & tmp_first_line_ts & ":E" & i - 1
                    
            Worksheets("report").Cells(l_report + 2, 1) = "m"
                Worksheets("report").Cells(l_report + 2, 2) = Worksheets("reverse").Cells(i - 1, 10)
                    Worksheets("report").Range("c" & l_report + 2).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!J" & tmp_first_line_ts & ":J" & i - 1
            
            Worksheets("report").Cells(l_report + 3, 1) = "omega"
                Worksheets("report").Cells(l_report + 3, 2) = Worksheets("reverse").Cells(i - 1, 11)
                    Worksheets("report").Range("c" & l_report + 3).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!K" & tmp_first_line_ts & ":K" & i - 1
            
            Worksheets("report").Cells(l_report + 4, 1) = "tc"
                Worksheets("report").Cells(l_report + 4, 2) = Worksheets("reverse").Cells(i - 1, 9)
                    Worksheets("report").Range("c" & l_report + 4).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!I" & tmp_first_line_ts & ":I" & i - 1
            
            
            
            
            
            Worksheets("visu").Cells(l_visu, 1) = tmp_ticker
            Worksheets("visu").Range("b" & l_visu).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!E" & tmp_first_line_ts & ":E" & i - 1
            Worksheets("visu").Range("c" & l_visu).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!J" & tmp_first_line_ts & ":J" & i - 1
            Worksheets("visu").Range("d" & l_visu).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!K" & tmp_first_line_ts & ":K" & i - 1
            Worksheets("visu").Range("e" & l_visu).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!I" & tmp_first_line_ts & ":I" & i - 1
            
            l_visu = l_visu + 1
            
            
            Worksheets("vert").Cells(5, c_vert) = tmp_ticker
            Worksheets("vert").Range(Worksheets("visu").Cells(6, c_vert).Address).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!E" & tmp_first_line_ts & ":E" & i - 1
            
            Worksheets("vert").Cells(7, c_vert) = Worksheets("reverse_clean").Cells(i - 1, 10)
            Worksheets("vert").Cells(8, c_vert) = Worksheets("reverse_clean").Cells(i - 1, 11)
            Worksheets("vert").Cells(9, c_vert) = Worksheets("reverse_clean").Cells(i - 1, 9)
            
            Worksheets("vert").Range(Worksheets("visu").Cells(11, c_vert).Address).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!J" & tmp_first_line_ts & ":J" & i - 1
            Worksheets("vert").Range(Worksheets("visu").Cells(12, c_vert).Address).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!K" & tmp_first_line_ts & ":K" & i - 1
            Worksheets("vert").Range(Worksheets("visu").Cells(13, c_vert).Address).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse!I" & tmp_first_line_ts & ":I" & i - 1
            
            
            
            
            Worksheets("vert_clean").Cells(5, c_vert) = tmp_ticker
            Worksheets("vert_clean").Range(Worksheets("vert_clean").Cells(6, c_vert).Address).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse_clean!E" & tmp_first_line_ts & ":E" & i - 1
            
            Worksheets("vert_clean").Cells(7, c_vert) = Worksheets("reverse_clean").Cells(i - 1, 10)
            Worksheets("vert_clean").Cells(8, c_vert) = Worksheets("reverse_clean").Cells(i - 1, 11)
            Worksheets("vert_clean").Cells(9, c_vert) = Worksheets("reverse_clean").Cells(i - 1, 9)
            
            Worksheets("vert_clean").Range(Worksheets("vert_clean").Cells(11, c_vert).Address).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse_clean!J" & tmp_first_line_ts & ":J" & i - 1
            Worksheets("vert_clean").Range(Worksheets("vert_clean").Cells(12, c_vert).Address).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse_clean!K" & tmp_first_line_ts & ":K" & i - 1
            Worksheets("vert_clean").Range(Worksheets("vert_clean").Cells(13, c_vert).Address).SparklineGroups.Add Type:=xlSparkLine, SourceData:="reverse_clean!I" & tmp_first_line_ts & ":I" & i - 1
            
            
            If Worksheets("reverse").Cells(i - 1, 9) < 650 Then
                Worksheets("vert_clean").Cells(7, c_vert) = Worksheets("reverse_clean").Cells(i - 1, 10)
                Worksheets("vert_clean").Cells(8, c_vert) = Worksheets("reverse_clean").Cells(i - 1, 11)
                Worksheets("vert_clean").Cells(9, c_vert) = Worksheets("reverse_clean").Cells(i - 1, 9)
            End If
            
            
            c_vert = c_vert + 1
            
            
            
            Dim ticker As String
            If tmp_ticker <> "EEM" Then
                ticker = tmp_ticker & " INDEX"
            Else
                ticker = tmp_ticker & " US EQUITY"
            End If
            
            dic_ticker.Add ticker, l_report
            
            
            
            
            l_report = l_report + 7
            
            
            
            tmp_ticker = Worksheets(1).Cells(i, 1)
            tmp_start_point = i
        End If
    
    End If
    
Next i




Dim vec_ticker() As Variant
k = 0
Dim tmp_entry As Variant
For Each tmp_entry In dic_ticker.Keys
    ReDim Preserve vec_ticker(k)
    vec_ticker(k) = tmp_entry
    k = k + 1
Next

Dim data_bbg As Variant
data_bbg = oBBG.bdp(vec_ticker, Array("name"), output_format.of_vec_without_header)


For i = 0 To UBound(data_bbg, 1)
    Worksheets("report").Cells(dic_ticker.Item(vec_ticker(i)), 1) = data_bbg(i)(0)
Next i


Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True


End Sub


