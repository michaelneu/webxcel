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
    Name As String
    Index As Integer
End Type


Private m_sheet
Private m_columns As Collection
Private m_primaryKey As Integer


Public Property Get PrimaryKey() As String
    For Each column In m_columns
        If column.Index = m_primaryKey Then
            PrimaryKey = column.Name
        End If
    Next column
End Property


Public Property Get Columns() As Collection
    Set Columns = New Collection
    
    For Each column In m_columns
        Columns.Add column.Name
    Next column
End Property


Public Property Get Entries() As Collection
    Set Entries = New Collection
    
    Dim i As Long
    For i = 2 To m_sheet.Rows.Count
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
            val = Trim(row.Cells(column.Index).text)
            
            If val = "" Then
                entry(column.Name) = Null
            Else
                entry(column.Name) = val
            End If
        Next column
        
        Entries.Add entry
    Next i
End Property


Public Sub Initialize(worksheet)
    m_primaryKey = -1
    
    Set m_sheet = worksheet
    Set m_columns = New Collection
    
    Dim firstRow
    Set firstRow = m_sheet.Rows(1)
    
    Dim i As Integer
    For i = 1 To firstRow.Cells.Count
        Dim column
        Set column = firstRow.Cells(i)
        
        Dim text As String
        text = Trim(column.text)
        
        If text = "" Then
            Exit For
        End If
        
        Dim header As WorksheetTableColumnHeader
        Set header = New WorksheetTableColumnHeader
        header.Name = text
        header.Index = i
        
        m_columns.Add header
        
        If column.Font.Bold And m_primaryKey = -1 Then
            m_primaryKey = i
        End If
    Next i
    
    If m_primaryKey = -1 Then
        Err.Raise StatusCode.ErrorNoPrimaryKeyDefined
    End If
End Sub
