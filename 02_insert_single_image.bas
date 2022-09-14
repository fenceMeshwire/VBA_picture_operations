Sub insert_single_image()

Dim strDir, strPath As String
Dim dblLeft, dblTop As Double
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

strDir = Thisworkbook.Path & "\"
strPath = strDir & "image.jpg"
dblLeft = wksSheet.Cells(2, 2).Left
dblTop = wksSheet.Cells(2, 2).Top

wksSheet.Shapes.AddPicture _
    fileName:=strPath, _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, _
    Left:=dblLeft, _
    Top:=dblTop, _
    Width:=-1, Height:=-1 ' Original picture dimensions
 
End Sub
