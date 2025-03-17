Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Timers

Sub Main()
    ' Ensure the script is run from an assembly file
    If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("This Rule only works on Assembly Files.", "WRONG DOCUMENT TYPE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return
    End If

    ' Get the active assembly document
    Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument
    Dim oAsmName As String = Left(oAsmDoc.DisplayName, Len(oAsmDoc.DisplayName) - 4) ' Remove the file extension

    ' Ask for user confirmation
    Dim RUsure = MessageBox.Show("This will create a STEP file for all components." _
    & vbLf & vbLf & "Are you sure you want to create STEP Drawings for all of the assembly components?" _
    & vbLf & "This could take a while.", "iLogic - Batch Output STEPs", MessageBoxButtons.YesNo)
    
    If RUsure = vbNo Then Return ' Exit if user selects 'No'

    ' Define output folder on the M: drive for STEP files
    Dim oFolder As String = "M:\Exported STEP\" & oAsmName & " Exported STEP"
    
    ' Create the output folder if it doesn't exist
    If Not System.IO.Directory.Exists(oFolder) Then
        System.IO.Directory.CreateDirectory(oFolder)
    End If

    ' Ask user if assemblies should also be saved as STEP
    Dim Assy = MessageBox.Show("Do you want assemblies as well as parts?" _
    & vbLf & "If you click 'No', only parts will be saved as STEP.", "iLogic - Batch Output STEPs", MessageBoxButtons.YesNo)
    
    If Assy = vbNo Then
        ' PART ONLY STEP SAVE
        ExportPartsOnlyAsSTEP(oAsmDoc, oFolder)
    ElseIf Assy = vbYes Then
        ' PARTS + ASSEMBLIES STEP SAVE
        ExportPartsAndAssembliesAsSTEP(oAsmDoc, oFolder)
    End If

    ' Open the output folder
    Shell("explorer.exe " & oFolder, vbNormalFocus)
End Sub

' Export only part files as STEP
Sub ExportPartsOnlyAsSTEP(oAsmDoc As AssemblyDocument, oFolder As String)
    Dim oRefDocs As DocumentsEnumerator = oAsmDoc.AllReferencedDocuments
    For Each oRefDoc As Document In oRefDocs
        ' Skip if the referenced document is not a part (.ipt)
        If oRefDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then Continue For
        
        ' Open the referenced part file
        Dim oCurFile As Document = ThisApplication.Documents.Open(oRefDoc.FullFileName, True)
        Dim oCurFileName As String = oCurFile.FullFileName
        Dim ShortName As String = IO.Path.GetFileNameWithoutExtension(oCurFileName)
        Dim oPN As String = oCurFile.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
        
        ' Save the part as a STEP file
        Try
            oCurFile.SaveAs(oFolder & "\" & oPN & ".stp", True)
        Catch
            MessageBox.Show("Error processing " & oCurFileName, "iLogic")
        End Try
        
        ' Close the document after saving
        oCurFile.Close()
    Next
    MessageBox.Show("(Part Only) New Files Created in: " & vbLf & oFolder, "iLogic")
End Sub

' Export both part and assembly files as STEP
Sub ExportPartsAndAssembliesAsSTEP(oAsmDoc As AssemblyDocument, oFolder As String)
    Dim oRefDocs As DocumentsEnumerator = oAsmDoc.AllReferencedDocuments
    For Each oRefDoc As Document In oRefDocs
        ' Open each referenced document (part or assembly)
        Dim oCurFile As Document = ThisApplication.Documents.Open(oRefDoc.FullFileName, True)
        Dim oCurFileName As String = oCurFile.FullFileName
        Dim ShortName As String = IO.Path.GetFileNameWithoutExtension(oCurFileName)
        Dim oPN As String = oCurFile.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
        
        ' Save the document as a STEP file
        Try
            oCurFile.SaveAs(oFolder & "\" & oPN & ".stp", True)
        Catch
            MessageBox.Show("Error processing " & oCurFileName, "iLogic")
        End Try
        
        ' Close the document after saving
        oCurFile.Close()
    Next

    ' Save the top-level assembly as a STEP file
    If oAsmDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        Dim oCurFile As Document = ThisApplication.Documents.Open(oAsmDoc.FullFileName, True)
        Dim oCurFileName As String = oCurFile.FullFileName
        Dim oPN As String = oCurFile.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
        Try
            oCurFile.SaveAs(oFolder & "\" & oPN & ".stp", True)
        Catch
            MessageBox.Show("Error processing " & oCurFileName, "iLogic")
        End Try
    End If

    MessageBox.Show("(+ Assembly +Base Assembly) New Files Created in: " & vbLf & oFolder, "iLogic")
End Sub
