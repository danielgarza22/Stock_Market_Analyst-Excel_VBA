Attribute VB_Name = "Module2"
Sub Clear_Cells()

Range("I:P").Value = ""
Range("I:P").Font.Bold = False
Range("I:P").ColumnWidth = 10.38
Range("I:N").Interior.ColorIndex = False

End Sub
