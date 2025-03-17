'define the active document as an assembly file
Dim oAsmDoc As AssemblyDocument
oAsmDoc = ThisApplication.ActiveDocument
oAsmName = Left(oAsmDoc.DisplayName, Len(oAsmDoc.DisplayName) -4)

'check that the active document is an assembly file
If ThisApplication.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
MessageBox.Show("Please run this rule from the assembly file.", "iLogic")
Exit Sub
End If

'get user input
RUsure = MessageBox.Show ( _
"This will create a DXF file for all of the asembly components that have drawings files." _
& vbLf & "This rule expects that the drawing file shares the same name and location as the component." _
& vbLf & " " _
& vbLf & "Are you sure you want to create DXF Drawings for all of the assembly components?" _
& vbLf & "This could take a while.", "iLogic  - Batch Output DXFs ",MessageBoxButtons.YesNo)

If RUsure = vbNo Then
Return
Else
End If

oPath = ThisDoc.Path

'get DXF target folder path
oFolder = oPath & "\" & oAsmName & " DXF Files"

'Check for the DXF folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder) Then
    System.IO.Directory.CreateDirectory(oFolder)
End If



'[ DXF setup

' Get the DXF translator Add-In.
Dim DXFAddIn As TranslatorAddIn
DXFAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
'Set a reference to the active document (the document to be published).
Dim oDocument As Document
oDocument = ThisApplication.ActiveEditDocument
Dim oContext As TranslationContext
oContext = ThisApplication.TransientObjects.CreateTranslationContext
oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
' Create a NameValueMap object
Dim oOptions As NameValueMap
oOptions = ThisApplication.TransientObjects.CreateNameValueMap
' Create a DataMedium object
Dim oDataMedium As DataMedium
oDataMedium = ThisApplication.TransientObjects.CreateDataMedium

' Check whether the translator has 'SaveCopyAs' options
If DXFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
	Dim strIniFile As String
	strIniFile = "C:\temp\dxfout.ini"
	' Create the name-value that specifies the ini file to use.
	oOptions.Value("Export_Acad_IniFile") = strIniFile
End If

'] end of DXF setup


'[ ComponentDrawings 
	'look at the files referenced by the assembly
	Dim oRefDocs As DocumentsEnumerator
	oRefDocs = oAsmDoc.AllReferencedDocuments
	Dim oRefDoc As Document
	
	'work the the drawing files for the referenced models
	'this expects that the model has a drawing of the same path and name 
For Each oRefDoc In oRefDocs
	idwPathName = Left(oRefDoc.FullDocumentName, Len(oRefDoc.FullDocumentName) - 3) & "idw"
	
	'check to see that the model has a drawing of the same path and name 
	If(System.IO.File.Exists(idwPathName)) Then
			Dim oDrawDoc As DrawingDocument
		oDrawDoc = ThisApplication.Documents.Open(idwPathName, True)
		oFileName = Left(oRefDoc.DisplayName, Len(oRefDoc.DisplayName) -3)
	
		On Error Resume Next ' if DXF exists and is open or read only, resume next
		'Set the DXF target file name
		oDataMedium.FileName = oFolder & "\" & oFileName & "DXF"
		'Write out the DXF
		'Set the destination file name
		oDataMedium.FileName = oFolder & "\" & oFileName & "dxf"
		'Publish document.
		Call DXFAddIn.SaveCopyAs(oDrawDoc, oContext, oOptions, oDataMedium)
		'close the file
		oDrawDoc.Close
	Else
	'If the model has no drawing of the same path and name - do nothing
	End If
Next
'] End of ComponentDrawings 



'[ Top Level Drawing 
	oAsmDrawing = ThisDoc.ChangeExtension(".idw")
	oAsmDrawingDoc = ThisApplication.Documents.Open(oAsmDrawing, True)
	oAsmDrawingName = Left(oAsmDrawingDoc.DisplayName, Len(oAsmDrawingDoc.DisplayName) -3)
	'write out the DXF for the Top Level Assembly Drawing file
	On Error Resume Next ' if DXF exists and is open or read only, resume next
	'Set the DXF target file name
	oDataMedium.FileName = oFolder & "\" & oAsmDrawingName & "dxf"
	'Write out the DXF
	Call DXFAddIn.SaveCopyAs(oAsmDrawingDoc, oContext, oOptions, oDataMedium)
	'Close the top level drawing
	oAsmDrawingDoc.Close
'] Top Level Drawing 

MessageBox.Show("New Files Created in: " & vbLf & oFolder, "iLogic")
'open the folder where the new ffiles are saved
Shell("explorer.exe " & oFolder,vbNormalFocus)