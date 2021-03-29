Attribute VB_Name = "Reset"
Sub Totalreset()

For Each ws In Worksheets

ws.Range("I:Q").Clear

Next ws

MsgBox ("Content cleared")

End Sub

