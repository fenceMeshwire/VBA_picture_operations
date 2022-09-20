Option Explicit

Sub delete_all_images_in_worksheet()

Dim shpImage As Shape
Dim wksSheet as Worksheet
  
Set wksSheet = Sheet1
  
For Each shpImage In wksSheet.Shapes
  shpImage.Delete
Next shpImage

End Sub
