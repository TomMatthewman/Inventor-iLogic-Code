Sub Main()
    Dim oDoc As Document
    oDoc = ThisDoc.Document
    oDocName = System.IO.Path.GetDirectoryName(oDoc.FullFileName) & "\" & System.IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName)
    
    If Not (oDoc.DocumentType = kAssemblyDocumentObject Or oDoc.DocumentType = kDrawingDocumentObject) Then
        MessageBox.Show("Please run this rule from the assembly or drawing files.", "iLogic")
        Exit Sub
    End If
    
    'get user input
    If MessageBox.Show ( _
        "This will create a PDF file for all of the files referenced by this document that have drawings files." _
        & vbLf & "This rule expects that the drawing file shares the same name and location as the component." _
        & vbLf & " " _
        & vbLf & "Are you sure you want to create PDF Drawings for all of the referenced documents?" _
        & vbLf & "This could take a while.", "iLogic  - Batch Output PDFs ",MessageBoxButtons.YesNo) = vbNo Then
        Exit Sub
    End If
        
    Dim PDFAddIn As TranslatorAddIn
    Dim oContext As TranslationContext
    Dim oOptions As NameValueMap
    Dim oDataMedium As DataMedium
    
    Call ConfigurePDFAddinSettings(PDFAddIn, oContext, oOptions, oDataMedium)
    
    oFolder = oDocName & " PDF Files"
    If Not System.IO.Directory.Exists(oFolder) Then
        System.IO.Directory.CreateDirectory(oFolder)
    End If
    
    '- - - - - - - - - - - - -Component Drawings - - - - - - - - - - - -
    Dim oRefDoc As Document
    Dim oDrawDoc As DrawingDocument
    
    For Each oRefDoc In oDoc.AllReferencedDocuments
        oBaseName = System.IO.Path.GetFileNameWithoutExtension(oRefDoc.FullFileName)
        oPathAndName = System.IO.Path.GetDirectoryName(oRefDoc.FullFileName) & "\" & oBaseName
        If(System.IO.File.Exists(oPathAndName & ".idw")) Then
            oDrawDoc = ThisApplication.Documents.Open(oPathAndName & ".idw", True)
			InventorVb.DocumentUpdate()
            oDataMedium.FileName = oFolder & "\" & oBaseName & ".pdf"
            Call PDFAddIn.SaveCopyAs(oDrawDoc, oContext, oOptions, oDataMedium)
            oDrawDoc.Close
        Else
            oNoDwgString = oNoDwgString & vbLf & idwPathName
        End If
    Next
    '- - - - - - - - - - - - -
    
    '- - - - - - - - - - - - -Top Level Drawing - - - - - - - - - - - -
    oBaseName = System.IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName)
    oPathAndName = System.IO.Path.GetDirectoryName(oDoc.FullFileName) & "\" & oBaseName
    oDataMedium.FileName = oFolder & "\" & oBaseName & ".pdf"
    
    If oDoc.DocumentType = kAssemblyDocumentObject Then
        oDrawDoc = ThisApplication.Documents.Open(oPathAndName & ".idw", True)
		InventorVb.DocumentUpdate()
        Call PDFAddIn.SaveCopyAs(oDrawDoc, oContext, oOptions, oDataMedium)
        oDrawDoc.Close
    ElseIf oDoc.DocumentType = kDrawingDocumentObject Then
        Call PDFAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)    
    End If
    '- - - - - - - - - - - - -
    
    MessageBox.Show("New Files Created in: " & vbLf & oFolder, "iLogic")
    MsgBox("Files found without drawings: " & vbLf & oNoDwgString, )
    Shell("explorer.exe " & oFolder,vbNormalFocus)
End Sub

Sub ConfigurePDFAddinSettings(ByRef PDFAddIn As TranslatorAddIn, ByRef oContext As TranslationContext, ByRef oOptions As NameValueMap, ByRef oDataMedium As DataMedium)

    PDFAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
    oContext = ThisApplication.TransientObjects.CreateTranslationContext
    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
        
    oOptions = ThisApplication.TransientObjects.CreateNameValueMap
    oOptions.Value("All_Color_AS_Black") = 1
    oOptions.Value("Remove_Line_Weights") = 0
    oOptions.Value("Vector_Resolution") = 400
    oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
    oOptions.Value("Custom_Begin_Sheet") = 1
    oOptions.Value("Custom_End_Sheet") = 1

    oDataMedium = ThisApplication.TransientObjects.CreateDataMedium
End Sub
