  Dim Title As String = "Excitech iLogic"

        ' Attempt to get the active assembly document
        Dim oDoc As AssemblyDocument = Nothing
        Try
            oDoc = ThisApplication.ActiveEditDocument
        Catch
            MessageBox.Show("This rule must be run from an assembly", Title)
            Exit Sub
        End Try

        Dim oADoc As AssemblyDocument = Nothing
        Dim oPDoc As PartDocument = Nothing
        Dim TotCount As Integer = oDoc.AllReferencedDocuments.Count

        Dim oDef As AssemblyComponentDefinition = oDoc.ComponentDefinition
        Dim oRefDoc As Document = Nothing
        Dim oPartDef As PartComponentDefinition = Nothing
        Dim oAsmDef As AssemblyComponentDefinition = Nothing
        Dim FailCount As Integer = 0

        Dim Count As Integer = 1

        ' Top level clear override command
        oDef.ClearAppearanceOverrides()

        ' Loop through all the documents referenced by this assembly document...
        For Each oRefDoc In oDoc.AllReferencedDocuments
            Try
                ThisApplication.StatusBarText = Count & " of " & TotCount & " components processed."

                ' Is it a part document?
                If oRefDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    ' Get the component definition.
                    oPartDef = oRefDoc.ComponentDefinition

                    ' First set the top level part appearance to be the same as the material 
                    oRefDoc.AppearanceSourceType = AppearanceSourceTypeEnum.kMaterialAppearance

                    ' Try a top level 'clear appearance overrides' command first
                    oPartDef.ClearAppearanceOverrides()

                    ' Clear the override on all the override objects found....
                    oPartDef.ClearAppearanceOverrides(ObjColl)

                    ThisApplication.ActiveView.Update()

                    ' Is it an assembly document?
                ElseIf oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    ' Get the assembly definition
                    oAsmDef = oRefDoc.ComponentDefinition
                    ' Run top level 'clear appearances' command on this assembly
                    ThisApplication.StatusBarText = Count & " of " & TotCount & " components processed. Clearing assembly overrides..."
                    oAsmDef.ClearAppearanceOverrides()
                End If
                Count += 1
            Catch
                FailCount += 1
            End Try
        Next

        ThisApplication.ActiveView.Update()

        If FailCount > 0 Then
            MsgBox("All Appearance overrides removed." & vbLf & vbLf & _
        "Operation failed on " & FailCount & " component(s) - these may be read-only.", , Title)
        Else
            MsgBox("All Appearance overrides removed.", , Title)
        End If