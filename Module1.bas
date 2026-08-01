Attribute VB_Name = "Module1"
'原本シートを非表示にする
Public Sub 原本シートを隠す()

    ThisWorkbook.Worksheets("原本").Visible = xlSheetHidden

End Sub


'原本シートを再表示する
Public Sub 原本シートを再表示する()

    ThisWorkbook.Worksheets("原本").Visible = xlSheetVisible

End Sub
