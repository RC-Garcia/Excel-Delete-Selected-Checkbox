Sub DeleteCheckBoxes()
    Dim selectedCell As Range
    Dim checkBox As CheckBox
    
    ' Loop through each selected cell
    For Each selectedCell In Selection
        ' Check if the cell contains a checkbox
        For Each checkBox In selectedCell.CheckBoxes
            ' Delete the checkbox
            checkBox.Delete
        Next checkBox
    Next selectedCell
End Sub
