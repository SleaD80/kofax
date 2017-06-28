Option Explicit

'#Uses "C:\Projects\Kofax\ClassificationDemo\Scripts\Classification.bas"

' Project Script

Private Sub Document_BeforeClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument, ByRef bSkip As Boolean)
   Dim VertLinesCountText As String
   VertLinesCountText=pXDoc.Locators.ItemByName("SL_VertLinesCount").Alternatives(0).Text

   If VertLinesCountText="16" Then
      pXDoc.Reclassify("UPD",0.8)
      ElseIf VertLinesCountText="15" Then
         pXDoc.Reclassify("UPD",0.7)
         ElseIf VertLinesCountText="14" Then
            pXDoc.Reclassify("Invoice",0.8)
            Else
            'Continue normal Classification'
   End If

End Sub

Private Sub SL_VertLinesCount_LocateAlternatives(ByVal pXDoc As CASCADELib.CscXDocument, ByVal pLocator As CASCADELib.CscXDocField)

   Dim vLinesCount As Byte
   vLinesCount=GetVLinesCount(pXDoc)

   'Create an output alternative for the script locator'
   With pLocator.Alternatives.Create
      .Text=CStr(vLinesCount)
      .Confidence=1

   End With

End Sub
