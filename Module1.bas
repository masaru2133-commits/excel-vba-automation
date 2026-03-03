Attribute VB_Name = "Module1"
Sub ImportAndSummarize()

    Dim wsData As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim i As Long
    Dim key As Variant
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' データシート作成
    On Error Resume Next
    Set wsData = Sheets("Data")
    If wsData Is Nothing Then
        Set wsData = Sheets.Add
        wsData.Name = "Data"
    End If
    On Error GoTo 0
    
    ' CSV取込（ファイル選択）
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "CSV Files", "*.csv"
        If .Show = -1 Then
            wsData.Cells.Clear
            With wsData.QueryTables.Add(Connection:="TEXT;" & .SelectedItems(1), Destination:=wsData.Range("A1"))
                .TextFileCommaDelimiter = True
                .Refresh
            End With
        Else
            Exit Sub
        End If
    End With
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' 集計処理（安全版）
Dim arr As Variant
Dim qty As Double, defect As Double

For i = 2 To lastRow
    key = Trim(CStr(wsData.Cells(i, 2).Value))        ' product
    qty = CDbl(wsData.Cells(i, 3).Value)              ' qty
    defect = CDbl(wsData.Cells(i, 4).Value)           ' defect

    If Not dict.Exists(key) Then
        dict.Add key, Array(0#, 0#)                   ' (TotalQty, TotalDefect)
    End If

    arr = dict(key)                                   ' ← 取り出す
    arr(0) = arr(0) + qty
    arr(1) = arr(1) + defect
    dict(key) = arr                                   ' ← 戻す（ここが重要）
Next i
    ' 集計処理
    'For i = 2 To lastRow
        'key = wsData.Cells(i, 2).Value
        
        'If Not dict.exists(key) Then
            'dict.Add key, Array(0, 0)
        'End If
        
        'dict(key)(0) = dict(key)(0) + wsData.Cells(i, 3).Value
        'dict(key)(1) = dict(key)(1) + wsData.Cells(i, 4).Value
    'Next i
    
    ' 集計シート作成
    On Error Resume Next
    Set wsSummary = Sheets("Summary")
    If wsSummary Is Nothing Then
        Set wsSummary = Sheets.Add
        wsSummary.Name = "Summary"
    End If
    On Error GoTo 0
    
    wsSummary.Cells.Clear
    wsSummary.Range("A1:D1") = Array("Product", "Total Qty", "Total Defect", "Defect Rate")
    
    i = 2
    For Each key In dict.keys
        wsSummary.Cells(i, 1) = key
        wsSummary.Cells(i, 2) = dict(key)(0)
        wsSummary.Cells(i, 3) = dict(key)(1)
        wsSummary.Cells(i, 4) = dict(key)(1) / dict(key)(0)
        i = i + 1
    Next key
    
    MsgBox "集計完了しました"
    
    CreateSummaryChart wsSummary

End Sub
Sub CreateSummaryChart(ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim chartObj As ChartObject
    Dim co As ChartObject
    Dim rngCat As Range, rngQty As Range, rngRate As Range

    ' 最終行（Productが入っている列Aで判定）
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then Exit Sub  ' データが少なすぎる場合は終了

    Set rngCat = ws.Range("A2:A" & lastRow) ' Product
    Set rngQty = ws.Range("B2:B" & lastRow) ' Total Qty
    Set rngRate = ws.Range("D2:D" & lastRow) ' Defect Rate

    ' 既存の同名グラフがあれば削除（作り直し）
    For Each co In ws.ChartObjects
        If co.Name = "SummaryChart" Then
            co.Delete
            Exit For
        End If
    Next co

    ' グラフを配置（位置は適宜。E2あたりから右側に置く）
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("F2").Left, _
        Top:=ws.Range("F2").Top, _
        Width:=420, _
        Height:=260)
    chartObj.Name = "SummaryChart"

    With chartObj.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Production Summary (Qty & Defect Rate)"

        ' いったんクリア
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        ' 棒：Total Qty
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Total Qty"
            .XValues = rngCat
            .Values = rngQty
            .ChartType = xlColumnClustered
            .AxisGroup = xlPrimary
        End With

        ' 折れ線：Defect Rate（第2軸）
        .SeriesCollection.NewSeries
        With .SeriesCollection(2)
            .Name = "Defect Rate"
            .XValues = rngCat
            .Values = rngRate
            .ChartType = xlLineMarkers
            .AxisGroup = xlSecondary
        End With

        ' 第2軸を％表示（不良率が 0.02 のような小数なら％に）
        On Error Resume Next
        .Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0.00%"
        On Error GoTo 0

        ' 凡例
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With

End Sub

