'#Reference {32AC4EE4-094A-4225-8F82-51277730B675}#4.0#0#C:\Program Files (x86)\Common Files\Kofax\Components\CscForms.6.0.dll#Kofax Cascade Forms Processing 4.0#CSCFORMSLib

Public Function GetVLinesCount(ByVal pXDoc As CASCADELib.CscXDocument, Optional ByRef pLocator As CASCADELib.CscXDocField = Nothing) As Integer
   ' You need to include a reference to "Kofax Cascade Forms Processing"
   ' c:\Program Files (x86)\Common Files\Kofax\Components\CscForms2.dll
   Dim pLD As CscLinesDetection
   Dim hL As CscLineInfo
   Dim vL As CscLineInfo
   Dim pImage As CscImage
   Dim yEndAverage, yDelta As Integer
   Dim verticalLines, badLines As Integer
   Dim alt As CscXDocFieldAlternative
   Dim XRes, YRes As Integer
   Dim i As Integer

   Set pImage = pXDoc.CDoc.Pages(0).GetImage()
   Set pLD = New CscLinesDetection

   pLD.DetectHorCombs = False
   pLD.DetectHorDotLines = False
   pLD.DetectHorLines = False ' Do not detect horizontal lines
   pLD.DetectVerLines = True
   XRes = pXDoc.CDoc.Pages.ItemByIndex(0).XRes ' for horizontal resolution, using xres of the first page
   YRes = pXDoc.CDoc.Pages.ItemByIndex(0).YRes ' for vertical resolution, , using yres of the first page
   pLD.MinVerLineLenMM = pImage.Height / YRes * 25.4 * 0.07 ' width/dpi*(mm)*7% - 7% is perfect to skip barcodes

   ' Detecting lines in specified area (also trying to cut off black lines on scan edges here)
   pLD.DetectLines pImage, 10, 10, pImage.Width - 30, pImage.Height - 30

   ' Calculating average yEnd attribute (the end of each line)
   For i = 0 To pLD.VerLineCount - 1
      Set hL = pLD.GetVerLine(i)
      yEndAverage = yEndAverage + hL.EndY
   Next
   yEndAverage = yEndAverage / pLD.VerLineCount
   yDelta = pImage.Height / 10 ' Delta of yEnd from average to accept the line

   ' Remove from average estimate "unique" lines
   Dim tempAverage As Long
   Dim removedCount As Long
   For i = 0 To pLD.VerLineCount - 1
      Set hL = pLD.GetVerLine(i)
      If Abs(hL.EndY - yEndAverage) > yDelta Then
         removedCount = removedCount + 1
      Else
         tempAverage = tempAverage + hL.EndY
      End If
   Next
   tempAverage = tempAverage / (pLD.VerLineCount - removedCount)
   yEndAverage = tempAverage

   ' Now, when average is for the table lines only, we can force more strict limit to separate excess lines
   yDelta = yDelta * 0.75

   Dim addLine As Boolean
   For i = 0 To pLD.VerLineCount - 1
      Set hL = pLD.GetVerLine(i)
      addLine = False

      If Abs(hL.EndY - yEndAverage) <= yDelta Then ' Line ends somewhere near the table's bottom
         addLine = True
      ElseIf hL.EndY > yEndAverage And hL.StartY < yEndAverage Then ' But if the line CROSSES table's bottom - we are accepting it
         addLine = True
      End If

      If addLine Then
         verticalLines = verticalLines + 1
      Else
         badLines = badLines + 1
      End If

         ' Add locator element
      If Not(pLocator Is Nothing) Then
         Set alt = pLocator.Alternatives.Create
         alt.Left = hL.StartX - 3
         alt.Top = hL.StartY
         alt.Width = Abs(hL.EndX - hL.StartX) + 5
         alt.Height = hL.EndY - hL.StartY
         alt.PageIndex = 0
         'alt.Text = "Line " & Str(i + 1) & " From: " & CStr(hL.StartY) & " To: " & CStr(hL.EndY) + " Len: " & CStr(alt.Height)
         alt.Text = IIf(addLine, "Line " & Str(i + 1), "Line " & Str(i + 1) & " *** REMOVED")
         alt.Confidence = 1
      End If
   Next

   Debug.Print(pXDoc.FileName & " V:" & CStr(verticalLines) & " (" & CStr(badLines) & ")")

   GetVLinesCount = verticalLines
End Function
