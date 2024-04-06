Sub DeleteLinkedCheckboxes()
    Dim chkBox As CheckBox
    Dim linkedCellAddress As String
    Dim linkedCellRange As Range
    
    ' Loop through each checkbox in the active sheet
    For Each chkBox In ActiveSheet.CheckBoxes
        ' Get the address of the linked cell
        linkedCellAddress = chkBox.LinkedCell
        ' Check if the linked cell is not empty
        If linkedCellAddress <> "" Then
            ' Convert the linked cell address to a range object
            Set linkedCellRange = Range(linkedCellAddress)
            ' Check if the linked cell intersects with the selected range
            If Not Intersect(linkedCellRange, Selection) Is Nothing Then
                ' Delete the checkbox
                chkBox.Delete
            End If
        End If
    Next chkBox
End Sub
