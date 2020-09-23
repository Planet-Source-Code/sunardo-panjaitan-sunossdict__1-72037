Attribute VB_Name = "mdlConnection"
Public con As ADODB.Connection
Public cmd As ADODB.Command
Public rs, rsi, recordEng, recordInd, recordSemantik As ADODB.Recordset
Public query As String
Public cari As String

Public Sub Connect()
Set con = New ADODB.Connection
    con.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & App.Path & "\sunossdict.b;" & _
        "Mode=ReadWrite"
    con.CursorLocation = adUseClient
    con.Open
End Sub

Public Sub rst()
    Call Connect
    Set rs = New ADODB.Recordset
    query = "select * from English  where English like " & "'%" & cari & "%' order by English asc "
    rs.Open query, con, adOpenDynamic, adLockOptimistic
End Sub

Public Sub recEng()
    Call Connect
    Set recordEng = New ADODB.Recordset
    query = "select * from English where English like " & "'%" & cari & "%'"
    recordEng.Open query, con, adOpenStatic, adLockOptimistic
End Sub
Public Sub rsti()
    Call Connect
    Set rsi = New ADODB.Recordset
    query = "select * from Indonesia where Indonesia like " & "'" & cari & "%' order by Indonesia asc"
    rsi.Open query, con, adOpenDynamic, adLockOptimistic
End Sub
Public Sub recInd()
    Call Connect
    Set recordInd = New ADODB.Recordset
    query = "select * from Indonesia where Indonesia like " & "'" & cari & "%'"
    recordInd.Open query, con, adOpenStatic, adLockOptimistic
End Sub



