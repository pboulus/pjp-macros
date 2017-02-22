Sub UpdateCharts()

'Define input variables to switch chart
Dim seriesCombos = {{"Building", "Timber"},
                            {"GDP", "Timber"},
                            {"Unemployment", "Concrete"}}

Dim sld As Slide
Dim shp As Shape

For Each sld In ActivePresentation.Slides
For Each shp In sld.Shapes
    'Check whether is chart
    If shp.HasChart Then
        With shp.Chart
            .ChartData.Activate
            For r = 1 To .ChartData.Workbook.Names.Count
                'Change country name
                If .ChartData.Workbook.Names.Item(r).Name = "Series1" Then
                    .ChartData.Workbook.Worksheets(1).Range("Series1").Value = seriesCombos(1,j)
                End If
            Next r
            Start = Timer
            While Timer < Start + 0.1
                DoEvents
            Wend
            .ChartData.Workbook.Close
        End With
    End If
    DoEvents
Next shp
Next sld
End Sub
