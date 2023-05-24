Dim oCSchemes As ColorSchemes = ThisApplication.ColorSchemes
If ThisApplication.ActiveColorScheme Is oCSchemes("Light") Then
	oCSchemes("Presentation").Activate
ElseIf ThisApplication.ActiveColorScheme Is oCSchemes("Presentation") Then
	oCSchemes("Light").Activate
Else 'set the default, for when neither color scheme was active
	oCSchemes("Light").Activate
End If