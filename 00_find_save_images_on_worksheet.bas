Option Explicit

Sub find_save_images_on_worksheet()

' Situation: Images are stored in the same row as the Partnumber and the Identifier information.
'            The images need to be stored with the "encoded" information on a specific directory.

Dim shShape As Shape
Dim shPicture As Shape
Dim chart_object As ChartObject

Dim lngRow As Long
Dim strPartnumber, strIdentifier, strEncoding As String
Dim strSaveDir, strSavePath, strSaveFile As String

Dim varAddress As Variant
Dim wksSheet As Worksheet

strSaveDir = "C:\Users\user\...\"

Set wksSheet = ActiveSheet

For Each shShape In wksSheet.Shapes

    If shShape.Type = msoPicture Then

      Set shPicture = ActiveSheet.Shapes(shShape.name)
      Set chart_object = ActiveSheet.ChartObjects.Add(0, 0, shPicture.Width, shPicture.Height)
      
      varAddress = Split(shPicture.TopLeftCell.Cells.Address, "$")
      lngRow = varAddress(2)
      
      strPartnumber = wksSheet.Cells(lngRow, 1).Value
      strIdentifier = wksSheet.Cells(lngRow, 2).Value
      strEncoding = strPartnumber & "m" & CStr(Len(strBezWi))
      strSavePath = strSaveDir & strEncoding
      strSaveFile = strSavePath & ".jpg"
      
      Application.Wait Now() + TimeValue("00:00:01") ' Waiting time for Copy method.
      shPicture.Copy
      chart_object.Chart.ChartArea.Select
      chart_object.Chart.Paste
      chart_object.Chart.Export strSaveFile
      chart_object.Delete
      strSavePath = strSaveDir
      
    End If

Next shShape

End Sub
