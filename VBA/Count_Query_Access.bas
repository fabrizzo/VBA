Attribute VB_Name = "Count_Query_Access"
Option Compare Database
Option Explicit
Function Count_Query()
Dim CountSelectSQL As Integer
Dim Connect As ADODB.Connection
Dim rsCount As ADODB.Recordset
Set Connect = New ADODB.Connection
With Connect
    .ConnectionString _
    "Provider = Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=" & CurrentPath() & "\Real Estate.accdb"
    .Open
End With
Set rsCount = New ADODB.Recordset
With rsCount
    .Source = SQLText
    .ActiveConnection = Connect
    .CursorType = adOpenKeyset
    .Open
End With
CountSelectSQL = rsCount.RecordCount
Set rsCount = Nothing
Set Connect = Nothing
CountQuery = CountSelectSQL
End Function

