Sub SplitFile()

    Dim lSlidesPerFile As Long
    Dim lTotalSlides As Long
    Dim oSourcePres As Presentation
    Dim otargetPres As Presentation
    Dim sFolder As String
    Dim sExt As String
    Dim sBaseName As String
    Dim lCounter As Long
    Dim lPresentationsCount As Long     ' how many will we split it into
    Dim x As Long
    Dim lWindowStart As Long
    Dim lWindowEnd As Long
    Dim sSplitPresName As String

    On Error GoTo ErrorHandler

    Set oSourcePres = ActivePresentation
    If Not oSourcePres.Saved Then
        MsgBox "Please save your presentation then try again"
        Exit Sub
    End If

    lSlidesPerFile = CLng(InputBox("How many slides per file?", "Split Presentation"))
    lTotalSlides = oSourcePres.Slides.Count
    sFolder = ActivePresentation.Path & "\"
    sExt = Mid$(ActivePresentation.Name, InStr(ActivePresentation.Name, ".") + 1)
    sBaseName = Mid$(ActivePresentation.Name, 1, InStr(ActivePresentation.Name, ".") - 1)

    If (lTotalSlides / lSlidesPerFile) - (lTotalSlides \ lSlidesPerFile) > 0 Then
        lPresentationsCount = lTotalSlides \ lSlidesPerFile + 1
    Else
        lPresentationsCount = lTotalSlides \ lSlidesPerFile
    End If

    If Not lTotalSlides > lSlidesPerFile Then
        MsgBox "There are fewer than " & CStr(lSlidesPerFile) & " slides in this presentation."
        Exit Sub
    End If

    For lCounter = 1 To lPresentationsCount

        ' which slides will we leave in the presentation?
        lWindowEnd = lSlidesPerFile * lCounter
        If lWindowEnd > oSourcePres.Slides.Count Then
            ' odd number of leftover slides in last presentation
            lWindowEnd = oSourcePres.Slides.Count
            lWindowStart = ((oSourcePres.Slides.Count \ lSlidesPerFile) * lSlidesPerFile) + 1
        Else
            lWindowStart = lWindowEnd - lSlidesPerFile + 1
        End If

        ' Make a copy of the presentation and open it
        sSplitPresName = sFolder & sBaseName & _
               "_" & CStr(lWindowStart) & "-" & CStr(lWindowEnd) & "." & sExt
        oSourcePres.SaveCopyAs sSplitPresName, ppSaveAsDefault
        Set otargetPres = Presentations.Open(sSplitPresName, , , True)

        With otargetPres
            For x = .Slides.Count To lWindowEnd + 1 Step -1
                .Slides(x).Delete
            Next
            For x = lWindowStart - 1 To 1 Step -1
                .Slides(x).Delete
            Next
            .Save
            .Close
        End With

    Next    ' lpresentationscount

NormalExit:
    Exit Sub
ErrorHandler:
    MsgBox "Error encountered"
    Resume NormalExit
End Sub
Sub DeletePJPFooter()
   Dim objCht As Chart
   Dim shp As Shape
   Dim sld As Slide
   For Each sld In ActivePresentation.Slides
   For Each shp In sld.Shapes
      If shp.Name Like "Footer Placeholder*" Then
      shp.Delete
      ElseIf shp.Name Like "Date Placeholder*" Then
      shp.Delete
      ElseIf shp.Name Like "Slide Number*" Then
      shp.Delete
      End If
   Next shp
   Next sld
End Sub

Sub JumpToSlide()
Dim JumpToIndex As Integer
JumpToIndex = InputBox("Jump to slide number?")
ActiveWindow.View.GotoSlide (JumpToIndex)
End Sub
Sub RemoveTextBoxMargins()
    With ActiveWindow.Selection.ShapeRange.TextFrame
        .MarginBottom = 0
        .MarginTop = 0
        .MarginLeft = 0
        .MarginRight = 0
    End With
End Sub

Sub MoveSlideToEnd()
currentSlide = ActiveWindow.Selection.SlideRange.SlideIndex
ActivePresentation.Slides(currentSlide).MoveTo (ActivePresentation.Slides.Count)
ActiveWindow.View.GotoSlide (currentSlide)
End Sub
Sub RemoveShadowAndBorder()
On Error Resume Next
    With ActiveWindow.Selection.ShapeRange
        If .HasTextFrame Then
        If .TextFrame.HasText Then
        .TextFrame.TextRange.Font.Shadow = msoFalse
        .Shadow.Visible = msoFalse
        .Line.Visible = msoFalse
        
        End If
        End If
    End With
End Sub
Sub UpdateDriverTree()

'Choose input variable to switch chart
Dim CountryChoice
CountryChoice = InputBox("Choose a country", "Country Select", "Thailand")

'Cycle through all shapes on slide
For Each thisshape In Application.ActiveWindow.View.Slide.Shapes
    
    'Check whether is chart
    If thisshape.HasChart Then
        With thisshape.Chart
            
            .ChartData.Activate
            
            For r = 1 To .ChartData.Workbook.Names.Count
                'Change country name
                If .ChartData.Workbook.Names.Item(r).Name = "Country" Then
                    .ChartData.Workbook.Worksheets(1).Range("Country").Value = CountryChoice
                End If
                
                'Get CAGR value
                If .ChartData.Workbook.Names.Item(r).Name = "CAGR" Then
                    CAGRData = .ChartData.Workbook.Worksheets(1).Range("CAGR").Value
                End If
                
                'Get Variable name
                If .ChartData.Workbook.Names.Item(r).Name = "Variable" Then
                    CAGRVariableName = .ChartData.Workbook.Worksheets(1).Range("Variable").Value
                End If
            Next r
                    
            Start = Timer
            While Timer < Start + 0.1
                DoEvents
            Wend
            .ChartData.Workbook.Close
        End With
        
        'Find CAGR bubble and update with formatted value
        ShapeName = CAGRVariableName & " CAGR"
        PrintCAGRValue = Format(CAGRData, "##0%")
        Application.ActiveWindow.View.Slide.Shapes(ShapeName).TextFrame.TextRange.Text = PrintCAGRValue
                            
    End If
    
    DoEvents
    
Next thisshape

End Sub
Sub SetAxesToZero()

For Each thisshape In Application.ActiveWindow.View.Slide.Shapes
    If thisshape.HasChart Then
        'Set the named range value to GroupName
        With thisshape.Chart
            With .Axes(xlValue)
            '.MaximumScale = 0
            .MinimumScale = 0
            
         End With
        End With
                            
    End If
                  
    'Replace CAGR
    
    
    DoEvents
    
Next thisshape

End Sub
Sub ModifyAxisRange_Selective()
   Dim objCht As Chart
   Dim shp As Shape
   For Each shp In Application.ActiveWindow.View.Slide.Shapes
      If shp.HasChart Then
      'If shp.Name Like "FIX_*" Then '<<<Use this condition to only affect charts with a particular label
      With shp.Chart
         ' Value (Y) Axis
            With .Axes(xlValue)
                .MaximumScale = 340
                .MinimumScale = 0
                '.MajorUnit = 5
             End With
        End With
        End If
      'End If
   Next shp
End Sub
Sub AddWeightedAverageToContributionCurve()
   Dim objCht As Chart
   Dim shp As Shape
   For Each shp In Application.ActiveWindow.Selection.ShapeRange
      If shp.HasChart Then
      'If shp.Name Like "FIX_*" Then '<<<Use this condition to only affect charts with a particular label
      With shp.Chart
         .SeriesCollection.NewSeries
            With .SeriesCollection(.SeriesCollection.Count)
                .ChartType = xlXYScatterLines
                .XValues = "=Input_Data!$J$35:$J$36"
                .Values = "=Input_Data!$K$35:$K$36"
                .Name = "=Input_Data!$J$34"
                .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
                .Format.Line.Weight = 1
                .Format.Line.DashStyle = msoLineLongDash
                .MarkerStyle = xlMarkerStyleNone
            End With
        .ChartData.Activate
        .ChartData.Workbook.Worksheets(1).Range("J34").Value = "Average line"
        .ChartData.Workbook.Worksheets(1).Range("J35").Value = 0
        .ChartData.Workbook.Worksheets(1).Range("J36").Value = 100 'Adust this to length of X-axis for contribution curve
        .ChartData.Workbook.Worksheets(1).Range("K35").Value = "=SUMPRODUCT(F8:F307, G8:G307)/SUM(F8:F307)" 'Calculated weighted average in chart data
        .ChartData.Workbook.Worksheets(1).Range("K36").Value = "=K35"
        .ChartData.Workbook.Worksheets(1).Range("K35:K36").NumberFormat = "###,###"
        .ChartData.Workbook.Close
        .SeriesCollection(.SeriesCollection.Count).Points(2).ApplyDataLabels Type:=xlDataLabelsShowValue
        .Refresh
        End With
        End If
      'End If
   DoEvents
   Next
End Sub
Sub moveobjects()

    Dim sld As Slide
    Dim shp As Shape
    Dim sr As Series
    Dim chrt As Chart

        For Each sld In ActivePresentation.Slides
            For Each shp In sld.Shapes

                If shp.Name Like "PRESENTSTAMP" Then
                    With shp
                      .Top = 0
                      .Left = -160
                      '.Left = 0
                      .Height = 15
                      .Width = 150
                    End With

                End If

    Next shp
    Next sld

End Sub
Sub getpixels()
MsgBox (ActiveWindow.Selection.ShapeRange.Left)
End Sub

Sub ChangeFontsOnPage()

    On Error Resume Next
    
    'Disabled: Make sure something is selected before loading the form
    'myTmp = ActiveWindow.Selection.ShapeRange.Count
    'If Err <> -2147188160 Then
        Load frm_ObjectFontSize
        frm_ObjectFontSize.Show
    'End If

End Sub
Sub hello()
MsgBox (ActiveWindow.Selection.ShapeRange(1).HasChart = msoTrue)
End Sub

Sub MakePresentationVersion()
    Dim RetainSlides As New Collection
   For Each sld In ActivePresentation.Slides
    i = 0
    Dim objCht As Chart
    Dim shp As Shape
    For Each shp In sld.Shapes
         If shp.Name = "PRESENTSTAMP" Then '<<<Use this condition to only affect charts with a particular label
             i = i + 1
         End If
    Next shp
    If i = 0 Then RetainSlides.Add sld.SlideIndex
   Next sld
   
   RetainSlidesArray = toArray(RetainSlides)
   ActivePresentation.Slides.Range(RetainSlidesArray).Delete
   
End Sub

Function toArray(col As Collection)
  Dim arr() As Variant
  ReDim arr(1 To col.Count) As Variant
  For i = 1 To col.Count
      arr(i) = col(i)
  Next
  toArray = arr
End Function
