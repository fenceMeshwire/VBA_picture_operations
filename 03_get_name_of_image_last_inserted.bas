Option Explicit

Sub get_name_of_image_last_inserted()
  
Dim shpImage as Shape

With Sheet1
  Set shpImage = .Shapes(.Shapes.Count)
End With

Debug.Print shpImage.Name
    
End Sub
