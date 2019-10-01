VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetTableCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_tables


Private Sub Class_Initialize()
    Set m_tables = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    For i = 1 To Worksheets.count
        Dim Sheet
        Set Sheet = Worksheets(i)
        
        Dim sheetName As String
        sheetName = LCase(Trim(Sheet.name))
        
        Dim table As WorksheetTable
        Set table = New WorksheetTable
        table.Initialize Sheet
        
        m_tables.Add sheetName, table
    Next i
End Sub


Public Function FindTable(tableName As String)
    tableName = LCase(Trim(tableName))
    
    If Not m_tables.Exists(tableName) Then
        Err.Raise StatusCode.ErrorSheetNotFound
    End If
    
    Set FindTable = m_tables(tableName)
End Function
