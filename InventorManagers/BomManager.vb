Imports Inventor
Public Class BomManager
    Public Sub AssemblyCount()
        ' Set reference to active document.
        ' This assumes the active document is an assembly
        Dim oDoc As Inventor.AssemblyDocument
        oDoc = ThisApplication.ActiveDocument

        ' Get assembly component definition
        Dim oCompDef As Inventor.ComponentDefinition
        oCompDef = oDoc.ComponentDefinition

        Dim sMsg As String
        Dim iLeafNodes As Long
        Dim iSubAssemblies As Long

        ' Get all occurrences from component definition for Assembly document
        Dim oCompOcc As ComponentOccurrence
        For Each oCompOcc In oCompDef.Occurrences
            ' Check if it's child occurrence (leaf node)
            If oCompOcc.SubOccurrences.Count = 0 Then
                Debug.Print oCompOcc.Name
            iLeafNodes = iLeafNodes + 1
            Else
                Debug.Print oCompOcc.Name
            iSubAssemblies = iSubAssemblies + 1
                Call processAllSubOcc(oCompOcc,
                                    sMsg,
                                    iLeafNodes,
                                    iSubAssemblies) ' subassembly
            End If
        Next

        Debug.Print "No of leaf nodes    : " + CStr(iLeafNodes)
    Debug.Print "No of sub assemblies: " + CStr(iSubAssemblies)
End Sub

    ' This function is called for processing sub assembly.  It is called recursively
    ' to iterate through the entire assembly tree.
    Private Sub processAllSubOcc(ByVal oCompOcc As ComponentOccurrence,
                                 ByRef sMsg As String,
                                 ByRef iLeafNodes As Long,
                                 ByRef iSubAssemblies As Long)

        Dim oSubCompOcc As ComponentOccurrence
        For Each oSubCompOcc In oCompOcc.SubOccurrences
            ' Check if it's child occurrence (leaf node)
            If oSubCompOcc.SubOccurrences.Count = 0 Then
                Debug.Print oSubCompOcc.Name
            iLeafNodes = iLeafNodes + 1
            Else
                sMsg = sMsg + oSubCompOcc.Name + vbCr
                iSubAssemblies = iSubAssemblies + 1

                Call processAllSubOcc(oSubCompOcc,
                                      sMsg,
                                      iLeafNodes,
                                      iSubAssemblies)
            End If
        Next
    End Sub
End Class
