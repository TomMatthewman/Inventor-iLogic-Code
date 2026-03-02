' Ensure we are in an assembly
If ThisApplication.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
    MessageBox.Show("Please run this rule from the Top-Level Assembly window.", "iLogic")
    Return
End If

Dim oAsmDoc As AssemblyDocument = ThisDoc.Document

' 1. Get all appearances from the active library to build the drop-down list
Dim oAppearances As New ArrayList
For Each oAsset In ThisApplication.ActiveAppearanceLibrary.AppearanceAssets
    oAppearances.Add(oAsset.DisplayName)
Next
oAppearances.Sort()

' 2. Show the drop-down (InputListBox)
Dim selectedApp As String = InputListBox("Select the appearance to apply to all sub-parts:", _
                            oAppearances, oAppearances.Item(0), "Appearance Selector", "Library Appearances")

' Exit if the user cancels
If String.IsNullOrEmpty(selectedApp) Then Return

' 3. Get the actual Asset from the library
Dim oLibAsset As Asset = ThisApplication.ActiveAppearanceLibrary.AppearanceAssets.Item(selectedApp)

' 4. Deep Dive: Apply to every part file (.ipt)
For Each oDoc In oAsmDoc.AllReferencedDocuments
    If oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        Dim oPartDoc As PartDocument = oDoc
        Try
            ' Copy asset to the part if it doesn't already exist there
            Dim oLocalAsset As Asset
            Try
                oLocalAsset = oPartDoc.Assets.Item(selectedApp)
            Catch
                oLocalAsset = oLibAsset.CopyTo(oPartDoc)
            End Try

            ' Apply to the base part file
            oPartDoc.ActiveAppearance = oLocalAsset
        Catch
            ' Skip read-only or locked files
        End Try
    End If
Next

' 5. Clear assembly-level overrides to make the new part colors visible
For Each oOcc As ComponentOccurrence In oAsmDoc.ComponentDefinition.Occurrences.AllLeafOccurrences
    oOcc.AppearanceSourceType = AppearanceSourceTypeEnum.kPartAppearance
Next

iLogicVb.UpdateWhenDone = True
ThisApplication.ActiveView.Update()
