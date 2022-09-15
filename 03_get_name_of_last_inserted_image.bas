Option Explicit

Sub get_name_of_last_inserted_image()
  
Dim shpImage as Shape

With Sheet1
  Set shpImage = .Shapes(.Shapes.Count)
End With

Debug.Print shpImage.Name
    
End Sub
