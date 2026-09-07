Private Sub Workbook_Open()
    Range("MacroFlag").Value = True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Range("MacroFlag").Value = False
End Sub
