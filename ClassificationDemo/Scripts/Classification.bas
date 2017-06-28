Public Function GetVLinesCount(ByVal pXDoc As CASCADELib.CscXDocument) As Byte
   'You need to include a reference to "Kofax Cascade Forms Processing"
   'c:\Program Files (x86)\Common Files\Kofax\Components\CscForms2.dll
   Dim pLD As CscLinesDetection
   Dim hL As CscLineInfo
   Dim vL As CscLineInfo
   Dim pImage As CscImage
   Dim xLeft As Long
   Dim xWidth As Long
   Dim yTop As Long
   Dim yHeight As Long
   Dim yEndAverage As Long
   Dim yDelta As Long
   Dim yStartAverage As Long
   Dim verticalLines As Long
   Dim badLines As Long
   Dim alt As CscXDocFieldAlternative
   Dim i, j, k As Integer
   Dim XRes, YRes As Integer
   Dim haveLongLine As Boolean

   Set pImage = pXDoc.CDoc.Pages(0).GetImage()
   Set pLD = New CscLinesDetection

   pLD.DetectHorCombs = False
   pLD.DetectHorDotLines = False
   pLD.DetectHorLines = True
   pLD.DetectVerLines = True
   XRes = pXDoc.CDoc.Pages.ItemByIndex(0).XRes ' for horizontal resolution, using xres of the first page
   YRes = pXDoc.CDoc.Pages.ItemByIndex(0).YRes ' for vertical resolution, , using yres of the first page
'   pLD.MinHorLineLenMM = pImage.Width / XRes * 25.4 * 0.5 ' width/dpi*(mm)*50%
   pLD.MinHorLineLenMM = pImage.Width / XRes * 25.4 * 1.5 ' width/dpi*(mm)*150% - just don't accept any horizontal line
   pLD.MinVerLineLenMM = pImage.Height / YRes * 25.4 * 0.07 ' width/dpi*(mm)*7% - 7% is perfect to skip barcodes

   'SET SEARCHING AREA
   xLeft = 10
   xWidth = pImage.Width - 20
   yTop = 10
   yHeight = pImage.Height - 20

   'DETECT LINES
   pLD.DetectLines pImage, xLeft, yTop, xWidth, yHeight

'   Debug.Print(pXDoc.FileName & " H:" & CStr(pLD.HorLineCount) & " V:" & CStr(pLD.VerLineCount))

   For i = 0 To pLD.VerLineCount - 1
      Set hL = pLD.GetVerLine(i)
      yEndAverage = yEndAverage + hL.EndY
   Next
   yEndAverage = yEndAverage / pLD.VerLineCount
   yDelta = pImage.Height / 10

   ' Remove from average estimate "unique" lines
   Dim tempAverage As Long
   Dim removedCount As Long
   tempAverage = 0
   removedCount = 0
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

   yStartAverage = 0
   For i = 0 To pLD.VerLineCount - 1
      Set hL = pLD.GetVerLine(i)
      yStartAverage = yStartAverage + hL.StartY
   Next
   yStartAverage = yStartAverage / pLD.VerLineCount

   ' Remove from average estimate "unique" lines
   tempAverage = 0
   removedCount = 0
   For i = 0 To pLD.VerLineCount - 1
      Set hL = pLD.GetVerLine(i)
      If Abs(hL.StartY - yStartAverage) > yDelta Then
         removedCount = removedCount + 1
      Else
         tempAverage = tempAverage + hL.StartY
      End If
   Next
   tempAverage = tempAverage / (pLD.VerLineCount - removedCount)
   yStartAverage = tempAverage

   yDelta = yDelta * 0.75 ' More hard limit

   verticalLines = 0
   badLines = 0
   haveLongLine = False
   For i = 0 To pLD.VerLineCount - 1
      Set hL = pLD.GetVerLine(i)
      If Abs(hL.EndY - yEndAverage) > yDelta Or Abs(hL.StartY - yStartAverage) > yDelta Then
         'Debug.Print(pXDoc.FileName & " Line " & Str(i + 1) & " From: " & CStr(hL.StartY) & " To: " & CStr(hL.EndY) + " skipped")
         badLines = badLines + 1
         If (i = 2 Or i = 3 Or i = 4) Then 'And hL.EndY > yEndAverage Then
            haveLongLine = True
         End If
      Else
         verticalLines = verticalLines + 1
      End If
   Next

'   Debug.Print(pXDoc.FileName & " V:" & CStr(verticalLines) & " yEndAverage:" & CStr(yEndAverage) & " yDelta:" & CStr(yDelta))
   Debug.Print(pXDoc.FileName & " V:" & CStr(verticalLines) & " (" & CStr(badLines) & ")" & " " & CStr(haveLongLine))
   Debug.Clear

   GetVLinesCount = verticalLines
End Function