Sub DeleteCheckboxes()
    Dim chkBox As CheckBox
    Dim selectedRange As Range
    
    ' Check if any cells are selected
    If TypeName(Selection) = "Range" Then
        Set selectedRange = Selection
        
        ' Loop through each cell in the selected range
        For Each cell In selectedRange
            ' Check if the cell contains a checkbox
            For Each chkBox In ActiveSheet.CheckBoxes
                If Not Intersect(chkBox.TopLeftCell, cell) Is Nothing Then
                    ' Delete the checkbox
                    chkBox.Delete
                    Exit For ' Exit the loop after deleting one checkbox per cell
                End If
            Next chkBox
        Next cell
    Else
        MsgBox "Please select a range of cells containing checkboxes.", vbExclamation
    End If
End Sub
