VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Refresh_Click()

Dim FinanSheet As Worksheet: Set FinanSheet = ActiveSheet
Dim Last As Integer: Last = FinanSheet.Range("A1000").End(xlUp).Row
If Last = 1 Then Exit Sub

Dim Symbols As String
Dim i As Integer
For i = 2 To Last
Symbols = Symbols & FinanSheet.Range("A" & i).Value & "+"
Next i
Symbols = Left(Symbols, Len(Symbols) - 1)



Dim URL As String: URL = "http://finance.yahoo.com/d/quotes.csv?s=" & Symbols & "&f=snl1hgop"
Dim Http As New WinHttpRequest
Http.Open "GET", URL, False
Http.Send


Dim Repsp As String: Resp = Http.ResponseText
Dim Lines As Variant: Lines = Split(Resp, vbNewLine)
Dim sLine As String
Dim Values As Variant

For i = 0 To UBound(Lines)
sLine = Lines(i)
If InStr(sLine, ",") > 0 Then
Values = Split(sLine, ",")
FinanSheet.Cells(i + 2, 2).Value = Split(Split(sLine, Chr(34) & "," & Chr(34))(1), Chr(34))(0)
FinanSheet.Cells(i + 2, 3).Value = Values(UBound(Values) - 4)
FinanSheet.Cells(i + 2, 4).Value = Values(UBound(Values) - 3)
FinanSheet.Cells(i + 2, 5).Value = Values(UBound(Values) - 2)
FinanSheet.Cells(i + 2, 6).Value = Values(UBound(Values) - 1)
FinanSheet.Cells(i + 2, 7).Value = Values(UBound(Values))
End If
Next i

FinanSheet.Cells.Columns.AutoFit

End Sub
