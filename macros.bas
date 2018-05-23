Attribute VB_Name = "NewMacros4"
Sub ”далениеѕробелов()
'
' ”далениеѕробелов Macro
' удаление лишних пробелов м создание таблицы в нЄм
'
Selection.WholeStory
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
    End With
    a = Selection.Find.Execute
    Do While a = True
        a = Selection.Find.Execute
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.HomeKey Unit:=wdStory
    Loop
    Selection.MoveDown Unit:=wdLine, Count:=1
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=5, NumColumns:= _
        5, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    Selection.Tables(1).Select
    Selection.SelectCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
End Sub

