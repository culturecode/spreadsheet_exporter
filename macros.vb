' This Visual Basic for Applications code can be added to the generated
' XLSX files to gain some data entry UX improvements.
'
' Specifically:
' - the ability to select multiple values for a single column with a data source
' - clearing dependent columns values when the column they are dependent_on changes
'
' Limitations:
' - only one parent col is supported and the dependent_on child columns must be
'   all beside each other so they can be selected as a range
'
' You will need to change `ParentTypeCol`, `DependentTypeStartCol` and `DependentTypeEndCol`
' below before adding the code to your Excel file.

Option Explicit

' https://stackoverflow.com/a/48375276/559596
Public Function ExistsInCollection(col As Collection, key As Variant) As Boolean
    On Error GoTo err
    ExistsInCollection = True
    IsObject(col.item(key))
    Exit Function
err:
    ExistsInCollection = False
End Function

' https://stackoverflow.com/a/47500463/559596
Public Function CollectionToArray(myCol As Collection) As Variant
    Dim result  As Variant
    Dim cnt     As Long

    ReDim result(myCol.Count - 1)

    For cnt = 0 To myCol.Count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt

    CollectionToArray = result
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
	Dim existingValue As String
	Dim toggledValue As String
	Dim tokenArr() As String
	Dim tokenCollection As Collection
	Dim token as Variant

	' The column containing the parent type ("site type") - we watch this for changes and then update
	' the dependent columns accordingly
	Dim ParentTypeCol As String
	ParentTypeCol = "AS"

	' All the columns that are dependent on the parent type ("site type")
	Dim DependentTypeStartCol As String
	Dim DependentTypeEndCol As String
	DependentTypeStartCol = "AT"
	DependentTypeEndCol = "AZ"

	Dim DependentTypeRangeSelector As String
	DependentTypeRangeSelector = DependentTypeStartCol & Target.Row & ":" & DependentTypeEndCol & Target.Row

	If Intersect(Target, Range(ParentTypeCol & ":" & ParentTypeCol & "," & DependentTypeRangeSelector)) Is Nothing Then Exit Sub

	' If an error occurs, enable events and quit the code
	On Error GoTo Quit

	Application.EnableEvents = False
	Application.ScreenUpdating = False

	' If we change anything in the parent-type col then clear all the dependent types
	If Not Intersect(Target, Range(ParentTypeCol & ":" & ParentTypeCol)) Is Nothing Then
		Debug.Print "in col for clearing... " & DependentTypeRangeSelector
		Range(DependentTypeRangeSelector).ClearContents
		GoTo Quit
	End If

	' Handle pick-list changes
    If Not Intersect(Target, Range(DependentTypeRangeSelector)) Is Nothing Then
		' If user deletes the dropdown cell's data do nothing
        If Target.Value = "" Then GoTo Quit

	    ' If we already have a comma we assume this is the result of copy-and-pasting
		' and we bail early
		toggledValue = Target.Value
		If InStr(toggledValue, ",") > 0 Then GoTo Quit

		Application.Undo

		existingValue = Target.Value

		tokenArr() = Split(existingValue, ",")
		Set tokenCollection = New Collection
		For Each token in tokenArr
			tokenCollection.Add Trim(token), Trim(token)
		Next


		If ExistsInCollection(tokenCollection, toggledValue) Then
			tokenCollection.Remove toggledValue
		Else
			tokenCollection.Add Trim(toggledValue), Trim(toggledValue)
		End If

		Target.Value = Join(CollectionToArray(tokenCollection), ",")
	End If

Quit:
	Application.EnableEvents = True
	Application.ScreenUpdating = True
End Sub
