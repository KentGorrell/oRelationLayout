Attribute VB_Name = "oVersion"
Option Compare Database
Option Explicit

' Download https://github.com/KentGorrell/oRelationLayout
' module: oVersion

Public Function Version_Number() As String
'...260818 db, CodeDb
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim sql As String

    sql = "SELECT LAST(Version_Number) AS LastVersion_Number" _
            & " FROM oSysVersion"
    Set db = CodeDb()
    Set rst = db.OpenRecordset(sql, dbReadOnly)
    With rst
        If Not .EOF Then
            Version_Number = !LastVersion_Number
        End If
    End With
    Set rst = Nothing
    Set db = Nothing
End Function
