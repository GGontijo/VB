'Esta macro serve para transferir um conjunto de células para um outro conjunto de células com filtro
'This macro is used to transfer a group of cells to a filtered group of cells

Sub Copiar_para_celulas_filtradas()
    Set from = Selection
    Set too = Application.InputBox("Selecione o intervalo de células do destino", Type:=8)
    For Each Cell In from
        Cell.Copy
        For Each thing In too
            If thing.EntireRow.RowHeight > 0 Then
                thing.PasteSpecial
                Set too = thing.Offset(1).Resize(too.Rows.Count)
                Exit For
            End If
        Next
    Next
End Sub


