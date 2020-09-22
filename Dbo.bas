Attribute VB_Name = "Dbo"
Option Explicit
Public Function SysConn()
    Dim File As String
    File = App.Path & "\smscenter.mdb"
    SysConn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & File & ";"
End Function

Public Function DBExecute(ByVal sql As String)
    Dim db As New ADODB.Connection
    
    db.CursorLocation = adUseClient
    db.Open SysConn, "Admin", ""
    Set DBExecute = db.Execute(sql)
    'db.Close
    Set db = Nothing
    
End Function
