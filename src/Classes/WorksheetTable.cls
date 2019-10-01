VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type ColumnHeader
    name As String
    index As Integer
End Type


Private m_sheet
Private m_columns As Collection
Private m_primaryKey As Integer


Public Property Get PrimaryKey() As String
    For Each column In m_columns
        If column.index = m_primaryKey Then
            PrimaryKey = column.name
            Exit Property
        End If
    Next column
End Property


Public Property Get Columns() As Collection
    Set Columns = New Collection
    
    For Each column In m_columns
        Columns.Add column.name
    Next column
End Property


Public Property Get Entries() As Collection
    Set Entries = New Collection
    
    Dim i As Long
    For i = 2 To m_sheet.Rows.count
        Dim row
        Set row = m_sheet.Rows(i)
        
        Dim pkText
        pkText = Trim(row.Cells(m_primaryKey))
        
        If pkText = "" Then
            Exit For
        End If
        
        Dim entry
        Set entry = CreateObject("Scripting.Dictionary")
        
        For Each column In m_columns
            Dim val
            val = Trim(row.Cells(column.index).text)
            
            If val = "" Then
                entry(column.name) = Null
            Else
                entry(column.name) = val
            End If
        Next column
        
        Entries.Add entry
    Next i
End Property


Public Sub Insert(value)
    Dim insertionRow As Long
    
    insertionRow = m_sheet.Cells(1, m_primaryKey).End(xlDown).row + 1
    
    If insertionRow = 1048577 Then
        For insertionRow = 2 To 1048576
            Dim text
            text = m_sheet.Cells(insertionRow, m_primaryKey).text
            If Trim(text) = "" Then
                Exit For
            End If
        Next insertionRow
    End If
    
    SetRowValues insertionRow, value
End Sub


Public Function Delete(key As String)
    Set Delete = Nothing
    
    Dim i As Long
    For i = 2 To m_sheet.Rows.count
        Dim row
        Set row = m_sheet.Rows(i)
        
        Dim pkText
        pkText = Trim(row.Cells(m_primaryKey))
        
        If pkText = key Then
            Dim entry
            Set entry = CreateObject("Scripting.Dictionary")
            
            For Each column In m_columns
                Dim val
                val = Trim(row.Cells(column.index).text)
                
                If val = "" Then
                    entry(column.name) = Null
                Else
                    entry(column.name) = val
                End If
            Next column
            
            row.Delete
            Set Delete = entry
            
            Exit Function
        End If
    Next i
End Function


Public Function Update(key As String, values) As Boolean
    Update = False
    
    Dim updateRow As Long
    For updateRow = 2 To m_sheet.Rows.count
        Dim row
        Set row = m_sheet.Rows(updateRow)
        
        Dim pkText
        pkText = Trim(row.Cells(m_primaryKey))
        
        If pkText = key Then
            SetRowValues updateRow, values
            Update = True
            
            Exit Function
        End If
    Next updateRow
End Function


Private Sub SetRowValues(row As Long, values)
    For Each column In m_columns
        If values.Exists(column.name) Then
            Dim insertionCell
            Set insertionCell = m_sheet.Cells(row, column.index)
            
            insertionCell.value = values(column.name)
        End If
    Next column
End Sub


Public Sub Initialize(worksheet)
    m_primaryKey = -1
    
    Set m_sheet = worksheet
    Set m_columns = New Collection
    
    Dim firstRow
    Set firstRow = m_sheet.Rows(1)
    
    Dim i As Integer
    For i = 1 To firstRow.Cells.count
        Dim column
        Set column = firstRow.Cells(i)
        
        Dim text As String
        text = Trim(column.text)
        
        If text = "" Then
            Exit For
        End If
        
        Dim header As WorksheetTableColumnHeader
        Set header = New WorksheetTableColumnHeader
        header.name = text
        header.index = i
        
        m_columns.Add header
        
        If column.Font.Bold And m_primaryKey = -1 Then
            m_primaryKey = i
        End If
    Next i
    
    If m_primaryKey = -1 Then
        Err.Raise StatusCode.ErrorNoPrimaryKeyDefined
    End If
End Sub
