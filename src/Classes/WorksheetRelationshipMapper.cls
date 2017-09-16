VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetRelationshipMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_tables As WorksheetTableCollection


Private Sub Class_Initialize()
    Set m_tables = New WorksheetTableCollection
End Sub


Public Function All(sheetName As String) As JsonArray
    Set All = New JsonArray

    Dim table As WorksheetTable
    Set table = m_tables.FindTable(sheetName)
    
    For Each entry In table.Entries
        Dim obj As JsonObject
        Set obj = New JsonObject
        
        For Each column In table.Columns
            Dim val As JsonValue
            Set val = New JsonValue
            
            val.value = entry(column)
            
            obj.SetProperty column, val
        Next column
        
        All.AddItem MapEntryRelationships(obj)
    Next entry
End Function


Private Function MapEntryRelationships(entry As JsonObject) As JsonObject
    Set MapEntryRelationships = entry
    
    For Each property In entry.PropertyNames
        Dim val As JsonValue
        Set val = entry.GetProperty(property)
        
        If property Like "fk_*" Then
            Dim foreignTableName As String
            foreignTableName = StringExtensions.SubString(property, 3)
            
            Dim newVal As IJson
            
            If val.ContainsNull Then
                Set newVal = val
            Else
                Dim foreignObject As JsonObject
                Set foreignObject = New JsonObject
                
                Dim foreignTable As WorksheetTable
                Set foreignTable = m_tables.FindTable(foreignTableName)
                
                For Each foreignEntry In foreignTable.Entries
                    If foreignEntry(foreignTable.PrimaryKey) = val.value Then
                        Set foreignObject = New JsonObject
                        
                        For Each foreignColumn In foreignTable.Columns
                            Dim foreignVal As JsonValue
                            Set foreignVal = New JsonValue
                            foreignVal.value = foreignEntry(foreignColumn)
                            
                            foreignObject.SetProperty foreignColumn, foreignVal
                        Next foreignColumn
                        
                        Set foreignObject = MapEntryRelationships(foreignObject)
                    End If
                Next foreignEntry
                
                If foreignObject Is Nothing Then
                    Err.Raise StatusCode.ErrorInvalidForeignKeyUsed
                End If
                
                Set newVal = foreignObject
            End If
            
            entry.RemoveProperty property
            entry.SetProperty foreignTableName, newVal
        End If
    Next property
End Function


Public Function Find(sheetName As String, key) As JsonObject
    Dim table As WorksheetTable
    Set table = m_tables.FindTable(sheetName)
    
    Dim allValues As JsonArray
    Set allValues = All(sheetName)
    
    Dim i As Integer
    For i = 0 To allValues.Count
        Dim value As JsonObject
        Set value = allValues.GetItem(i)
        
        If value.GetProperty(table.PrimaryKey) = key Then
            Set Find = value
            Exit Function
        End If
    Next i
End Function


Public Sub Insert(sheetName As String, item As JsonObject)

End Sub


Public Sub Update(sheetName As String, item As JsonObject)

End Sub


Public Sub Remove(sheetName As String, item As JsonObject)

End Sub
