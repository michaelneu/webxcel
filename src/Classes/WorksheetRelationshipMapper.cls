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


Public Function PrimaryKeyName(sheetName As String) As String
    Dim table As WorksheetTable
    Set table = m_tables.FindTable(sheetName)
    
    PrimaryKeyName = table.PrimaryKey
End Function


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
            foreignTableName = StringExtensions.Substring(property, 3)
            
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


Public Function Find(sheetName As String, key) As IJson
    Dim table As WorksheetTable
    Set table = m_tables.FindTable(sheetName)
    
    Dim allValues As JsonArray
    Set allValues = All(sheetName)
    
    Dim i As Integer
    For i = 0 To allValues.count - 1
        Dim value As JsonObject
        Set value = allValues.GetItem(i)
        
        Dim valueAsJsonValue As JsonValue
        Set valueAsJsonValue = value.GetProperty(table.PrimaryKey)
        
        If valueAsJsonValue.value = key Then
            Set Find = value
            Exit Function
        End If
    Next i
    
    Dim nullValue As JsonValue
    Set nullValue = New JsonValue
    nullValue.value = Null
    
    Set Find = nullValue
End Function


Public Sub Insert(sheetName As String, item As JsonObject)
    Dim table As WorksheetTable
    Set table = m_tables.FindTable(sheetName)
    
    Dim itemToInsert
    Set itemToInsert = CreateObject("Scripting.Dictionary")
    
    Dim pkFound As Boolean
    pkFound = False
    
    For Each property In item.PropertyNames
        If property = table.PrimaryKey Then
            pkFound = True
        End If
        
        Dim propertyIsForeignKey As Boolean
        propertyIsForeignKey = False
        
        For Each column In table.Columns
            If "fk_" & property = column Then
                propertyIsForeignKey = True
                Exit For
            End If
        Next column
        
        If propertyIsForeignKey Then
            Dim foreignPropertyName As String
            foreignPropertyName = "fk_" & property
            
            Dim foreignTableName As String
            foreignTableName = property
            
            Dim foreignTable As WorksheetTable
            Set foreignTable = m_tables.FindTable(foreignTableName)
            
            Dim foreignItem As IJson
            Set foreignItem = item.GetProperty(property)
            
            If TypeName(foreignItem) = "JsonValue" Then
                Dim foreignItemAsValue As JsonValue
                Set foreignItemAsValue = foreignItem
                
                If foreignItemAsValue.ContainsNull Then
                    itemToInsert.Add foreignPropertyName, ""
                Else
                    Err.Raise StatusCode.ErrorInvalidForeignKeyUsed
                End If
            ElseIf TypeName(foreignItem) = "JsonObject" Then
                Dim foreignItemAsObject As JsonObject
                Set foreignItemAsObject = foreignItem
                
                Dim foreignItemContainsPk As Boolean
                foreignItemContainsPk = False
                
                For Each foreignProperty In foreignItemAsObject.PropertyNames
                    If foreignProperty = foreignTable.PrimaryKey Then
                        Dim foreignPkValue As JsonValue
                        Set foreignPkValue = foreignItemAsObject.GetProperty(foreignProperty)
                        
                        itemToInsert.Add foreignPropertyName, foreignPkValue.value
                        foreignItemContainsPk = True
                    End If
                Next foreignProperty
                
                If foreignItemContainsPk Then
                    Insert foreignTableName, foreignItemAsObject
                Else
                    Err.Raise StatusCode.ErrorInvalidForeignKeyUsed
                End If
            Else
                Err.Raise StatusCode.ErrorInvalidForeignKeyUsed
            End If
        Else
            Dim value As JsonValue
            Set value = item.GetProperty(property)
            
            If value.ContainsNull Then
                itemToInsert.Add property, ""
            Else
                itemToInsert.Add property, value.value & ""
            End If
        End If
    Next property
    
    If Not pkFound Then
        Err.Raise StatusCode.ErrorDataDoesntContainPrimaryKey
        Exit Sub
    End If
    
    table.Insert itemToInsert
End Sub


Public Function Delete(sheetName As String, key As String) As JsonObject
    Dim table As WorksheetTable
    Set table = m_tables.FindTable(sheetName)
    
    Dim entry
    Set entry = table.Delete(key)
    
    If Not entry Is Nothing Then
        Dim deletedValue As JsonObject
        Set deletedValue = New JsonObject
            
        For Each column In table.Columns
            Dim val As JsonValue
            Set val = New JsonValue
            
            val.value = entry(column)
            
            deletedValue.SetProperty column, val
        Next column
        
        Set Delete = MapEntryRelationships(deletedValue)
    End If
End Function


Public Function Update(sheetName As String, key As String, item As JsonObject) As Boolean
    Dim table As WorksheetTable
    Set table = m_tables.FindTable(sheetName)
    
    Dim values
    Set values = CreateObject("Scripting.Dictionary")
    
    For Each property In item.PropertyNames
        Dim val As JsonValue
        Set val = item.GetProperty(property)
        
        If IsNull(val.value) Then
            values(property) = ""
        Else
            values(property) = val.value
        End If
    Next property
    
    Update = table.Update(key, values)
End Function
