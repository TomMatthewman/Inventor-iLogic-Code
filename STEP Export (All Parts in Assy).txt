Sub Main
	If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
	    MessageBox.Show("This Rule " & iLogicVb.RuleName & " only works on Assembly Files.", "WRONG DOCUMENT TYPE", MessageBoxButtons.OK, MessageBoxIcon.Error)
	    Return
	End If
	Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument
	oAsmDoc = ThisApplication.ActiveDocument
	oAsmName = Left(oAsmDoc.DisplayName, Len(oAsmDoc.DisplayName) -4)

	Dim RUsure = MessageBox.Show("This will create a STEP file for all components." _
	& vbLf & " " _
	& vbLf & "Are you sure you want to create STEP Drawings for all of the assembly components?" _
	& vbLf & "This could take a while.", "iLogic - Batch Output STEPs ", MessageBoxButtons.YesNo)
	If RUsure = vbNo Then Return

	Dim oPath As String = System.IO.Path.GetDirectoryName(oAsmDoc.FullFileName)
	oFolder = oPath & "\" & oAsmName & " STEP Files"
	If Not System.IO.Directory.Exists(oFolder) Then
	    System.IO.Directory.CreateDirectory(oFolder)
	End If

	Dim oRefDocs As DocumentsEnumerator = oAsmDoc.AllReferencedDocuments
	For Each oRefDoc As Document In oRefDocs
		'if is not a Part, skip to next referenced document
		'If oRefDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then Continue For
		Dim oCurFile As Document = ThisApplication.Documents.Open(oRefDoc.FullFileName, True)
		Dim oCurFileName = oCurFile.FullFileName
		Dim ShortName = IO.Path.GetFileNameWithoutExtension(oCurFileName)
		Dim oPN As String = oCurFile.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
		'might want to ckeck if oPN = "" or not
		Try
			oCurFile.SaveAs(oFolder & "\" & oPN & ".stp", True)
		Catch
			MessageBox.Show("Error processing " & oCurFileName, "ilogic")
		End Try
		oCurFile.Close()
 
	Next
	MessageBox.Show("New Files Created in: " & vbLf & oFolder, "iLogic")
End Sub