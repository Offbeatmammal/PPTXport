Attribute VB_Name = "Module1"
Dim outlines() As String

Sub PPTXport()
Dim pres As Presentation
Dim oPPApp As Object
Set oPPApp = New PowerPoint.Application

    Set msoTypes = New Scripting.Dictionary
    msoTypes(1) = "AutoShape"
    msoTypes(3) = "Chart"
    msoTypes(9) = "Line"
    msoTypes(13) = "Picture"
    msoTypes(17) = "Text Box"
    msoTypes(28) = "Graphic"
    Set msoAutoShapes = New Scripting.Dictionary
    msoAutoShapes(1) = "Rectangle"
    msoAutoShapes(10) = "Hexagon"
    Set msoDashStyles = New Scripting.Dictionary
    msoDashStyles(1) = "solid"
    msoDashStyles(-2) = "lgDashDot"
    msoDashStyles(4) = "lgDash"
    Set msoParAligns = New Scripting.Dictionary
    msoParAligns(1) = "left"
    msoParAligns(2) = "center"

    ' removes all other presentations
    ' (DON'T RUN THIS UNLESS YOU'VE SAVED EVERYTHING!!)
    For Each m In Application.Presentations
        If m.Name <> ActivePresentation.Name And LCase(m.Name) <> "pptxport.pptm" Then
            m.Close
        End If
    Next
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Filters.Add "Powerpoint Files", "*.pptx;*.pptm", 1
        .Show
    End With
    Set pres = oPPApp.Presentations.Open(FileName:=fd.SelectedItems.Item(1), ReadOnly:=msoTrue, WithWindow:=msoFalse)

    ReDim Preserve outlines(0 To 0)
    Debug.Print (Now & ": Processing...")

    pushLine ("<html>")
    pushLine ("<head>")
    pushLine ("<script src=""https://cdnjs.cloudflare.com/ajax/libs/jquery/3.4.1/jquery.min.js""></script>")
    pushLine ("<script src=""https://cdnjs.cloudflare.com/ajax/libs/jszip/3.1.5/jszip.min.js""></script>")
    pushLine ("<script src=""https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@latest/dist/pptxgen.min.js""></script>")
    pushLine ("</head>")
    pushLine ("<body>")
    pushLine ("<script>")
    pushLine ("let pptx = new PptxGenJS();")
    pushLine ("pptx.layout = 'LAYOUT_WIDE'")

    For Each osl In pres.Slides
   
        pushLine ("slide = pptx.addSlide();")
        ' go through each shape
        For Each ob In osl.Shapes
            If msoTypes.Exists(ob.Type) Then
                Select Case ob.Type
                Case 1 'AutoShape
                    Select Case ob.AutoShapeType
                    Case 1 ' Rectangle
                        st = "slide.addShape(pptx.ShapeType.rect,{line:'" + toRGB(ob.Line.ForeColor.RGB) + "',lineDash:'" & msoDashStyles(ob.Line.DashStyle) & "',x:" + Str(round2(ob.Left, 2)) + ",y:" + Str(round2(ob.Top, 2)) + ",w:" + Str(round2(ob.Width, 2)) + ",h:" + Str(round2(ob.Height, 2)) + ",rotate:" + Str(ob.Rotation) + ", fill:{ type:'solid', color:'" + toRGB(ob.Fill.ForeColor.RGB) + "',zorder:" + Str(ob.ZOrderPosition) + " } })"
                        pushLine (st)
                    Case 10 'Hexagon
                        st = "slide.addShape(pptx.ShapeType.hexagon,{x:" + Str(round2(ob.Left, 2)) + ",y:" + Str(round2(ob.Top, 2)) + ",w:" + Str(round2(ob.Width, 2)) + ",h:" + Str(round2(ob.Height, 2)) + ",rotate:" + Str(ob.Rotation) + ", fill:{ type:'solid', color:'" + toRGB(ob.Fill.ForeColor.RGB) + "' } })"
                        pushLine (st)
                    End Select
                Case 3 ' Chart
                    dataChartAreaLine = "let dataChartAreaLine= ["
                    Dim c As Object
                    Dim chartColors As String
                
                    For temp = 1 To ob.Chart.SeriesCollection.Count
                        With ob.Chart.SeriesCollection.Item(temp)
                            chartColors = chartColors + "'" + toRGB(.Fill.ForeColor.RGB) + "',"
                            Z = .XValues
                            If temp > 1 Then
                                dataChartAreaLine = dataChartAreaLine + ","
                            End If
                            dataChartAreaLine = dataChartAreaLine + "{ name:'" + .Name + "',"
                            dataChartAreaLine = dataChartAreaLine + "labels: ["
                            For Each xv In Z
                                dataChartAreaLine = dataChartAreaLine + "'" + xv + "',"
                            Next xv
                            dataChartAreaLine = Left(dataChartAreaLine, Len(dataChartAreaLine) - 1) ' remove last comma

                            dataChartAreaLine = dataChartAreaLine + "], values: ["
                            Z = .Values
                            For Each v In Z
                                dataChartAreaLine = dataChartAreaLine + "'" & v & "',"
                            Next v
                            dataChartAreaLine = Left(dataChartAreaLine, Len(dataChartAreaLine) - 1) ' remove last comma
    
                            dataChartAreaLine = dataChartAreaLine + "]"
                            dataChartAreaLine = dataChartAreaLine + "}"
                        End With
                    Next temp
                    dataChartAreaLine = dataChartAreaLine + "]"
                    chartColors = "[" + Left(chartColors, Len(chartColors) - 1) + "]" ' remove last comma
                    Select Case ob.Chart.ChartType
                    Case 57 ' Clustered Bar
                        chartColors = chartColors + ",barDir:'bar'"
                    Case 51 ' Clustered Column
                        chartColors = chartColors + ",barDir:'col'"
                    End Select
                    pushLine (dataChartAreaLine)
                    st = "slide.addChart(pptx.ChartType.bar,dataChartAreaLine,"
                    st = st + "{chartColors: " + chartColors + ",x:" + Str(round2(ob.Left, 2)) + ",y:" + Str(round2(ob.Top, 2)) + ",w:" + Str(round2(ob.Width, 2)) + ",h:" + Str(round2(ob.Height, 2))
                    st = st + "} )"
                    pushLine (st)
                Case 17 ' TextBox
                    st = "slide.addText("
                    st = st + "["
                    rcomma = False
                    For Each r In ob.TextFrame.TextRange.Runs
                        If Not rcomma Then
                            rcomma = True
                        Else
                            st = st + ","
                        End If
                        st = st + ("{text:'" + Replace(Replace(Replace(r.Text, vbCrLf, ""), vbCr, ""), vbLf, "") + "', options:{fontName:'" + r.Font.Name + "',fontSize:" + Str(r.Font.Size) + ", color:'" + toRGB(r.Font.Color.RGB) + "'")
                        st = st + "} }"
                    Next r
                    st = st + "], {align:'" + msoParAligns(ob.TextFrame.TextRange.ParagraphFormat.Alignment) + "'"
                    Select Case ob.TextFrame.AutoSize
                    Case 1
                        st = st + ",autofit: 'true'"
                    Case 2
                        st = st + ",shrinkTextL: 'true'"
                    End Select
                    st = st + ",zorder:" + Str(ob.ZOrderPosition) + ",x:" + Str(round2(ob.Left, 2)) + ",y:" + Str(round2(ob.Top, 2)) + ",w:" + Str(round2(ob.Width, 2)) + ",h:" + Str(round2(ob.Height, 2)) + ",rotate:" + Str(ob.Rotation)
                    If ob.Fill.Visible Then
                        st = st + ", fill:{ type:'solid', color:'" + toRGB(ob.Fill.ForeColor.RGB) + "' }"
                    End If
                    st = st + "} )"
                    pushLine (st)
                Case 13, 28
                    iFN = "images/__" + ob.Name + ".png"
                    Call ob.Export(iFN, ppShapeFormatPNG)
                    st = "slide.addImage({x:" + Str(round2(ob.Left, 2)) + ",y:" + Str(round2(ob.Top, 2)) + ",w:" + Str(round2(ob.Width, 2)) + ",h:" + Str(round2(ob.Height, 2)) + ",rotate:" + Str(ob.Rotation) + ", path:'" + iFN + "' })"
                    pushLine (st)

                End Select
            Else
                Debug.Print ("**" + Str(ob.Type) + "**:" + ob.Name)
            End If
        Next
    Next
    
    pres.Close
    
    pushLine ("pptx.writeFile();")
    pushLine ("</script>")
    pushLine ("</body>")
    pushLine ("</html>")

' export the result to the HTML file
    Open "ppt.html" For Output As #1
    For Each ln In outlines
        Print #1, ln
    Next ln
    Close #1
    
    Debug.Print (Now & ": Complete!")

End Sub
Private Function round2(n As Single, f) As Single
    round2 = ((n / 72) + 0)
End Function
Private Function toRGB(c)
Dim retval(3), ii

For ii = 0 To 2
  retval(ii) = c Mod 256
  c = (c - retval(ii)) / 256
Next

toRGB = Right("00" + Hex(retval(0)), 2) + Right("00" + Hex(retval(1)), 2) + Right("00" + Hex(retval(2)), 2)

End Function

Private Sub pushLine(l)
ReDim Preserve outlines(0 To UBound(outlines) + 1)
    outlines(UBound(outlines)) = l
End Sub

