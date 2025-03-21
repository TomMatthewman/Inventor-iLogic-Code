''' <summary>
''' Allows an iLogic user to remove broken "pink" unreferenced sketch entities.
''' </summary>
Sub Main()
	Dim i, j As Integer
	If TypeOf(ThisApplication.ActiveDocument) Is PartDocument Then
		Dim doc As PartDocument = ThisApplication.ActiveDocument

		For Each sk As Sketch In doc.ComponentDefinition.Sketches
			For Each ske As SketchEntity In sk.SketchEntities
				If ske.Reference And ske.ReferencedEntity Is Nothing Then
					i = i + 1
					On Error Resume Next
					ske.Delete
					If Err.Number = 0 Then
						j = j + 1
					End If
					On Error GoTo 0
				End If
			Next ske
		Next sk
		Messagebox.Show(Str(i) & " broken links were found in the sketches. " & Str(j) & " got deleted.")
	End If
End Sub