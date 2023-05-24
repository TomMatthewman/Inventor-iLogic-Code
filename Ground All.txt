'start of ilogic code
' set a reference to the assembly component definintion.
' This assumes an assembly document is open.
Dim oAsmCompDef As AssemblyComponentDefinition
oAsmCompDef = ThisApplication.ActiveDocument.ComponentDefinition

'set the Master LOD active
Dim oLODRep As LevelofDetailRepresentation
oLODRep = oAsmCompDef.RepresentationsManager.LevelOfDetailRepresentations.Item("Master")
oLODRep.Activate

 'Iterate through all of the occurrences and ground them.
Dim oOccurrence As ComponentOccurrence
For Each oOccurrence In oAsmCompDef.Occurrences
'check for and skip virtual components
If Not TypeOf oOccurrence.Definition Is VirtualComponentDefinition Then
oOccurrence.Grounded = True
else
end if
Next
'end of ilogic code
