Option Explicit

'#Uses "C:\Projects\Kofax\ClassificationDemo\Scripts\Classification.bas"

' Project Script

Private Sub Application_InitializeBatch(ByVal pXRootFolder As CASCADELib.CscXFolder)
   Debug.Clear
End Sub

Private Sub Document_BeforeClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument, ByRef bSkip As Boolean)
   Dim VertLinesCount As Long
   VertLinesCount = pXDoc.Locators.ItemByName("SL_VertLinesCount").Alternatives(0).LongTag

   If VertLinesCount >= 14 And VertLinesCount <= 15 Then
      pXDoc.Reclassify("Invoice", 0.8)
   ElseIf VertLinesCount >= 16 And VertLinesCount <= 18 Then
       pXDoc.Reclassify("UPD",0.7)
   Else
       ' Continue normal Classification'
   End If

End Sub

Private Sub SL_VertLinesCount_LocateAlternatives(ByVal pXDoc As CASCADELib.CscXDocument, ByVal pLocator As CASCADELib.CscXDocField)

   Dim vLinesCount As Integer
   vLinesCount = GetVLinesCount(pXDoc, pLocator)

   ' Set the first alternative's LongTag attribute to lines count
   Dim alt As CscXDocFieldAlternative
   If pLocator.Alternatives.Count = 0 Then ' In case there are no alternatives created by Classification script
      Set alt = pLocator.Alternatives.Create
      alt.Confidence = 1
   Else
      Set alt = pLocator.Alternatives(0)
   End If
   alt.LongTag = vLinesCount

End Sub
