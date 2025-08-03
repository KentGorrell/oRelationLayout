Attribute VB_Name = "oFrameStartup"
Option Compare Database
Option Explicit

Public Function Startup()
    Call Addin_Load
    DoCmd.OpenForm "oFrameRelationLayout"
End Function
Public Sub Addin_Load()
On Error GoTo Error_Handler
Dim str As String
Dim strTableName As String
Dim qdf As QueryDef
Dim tdf As TableDef
Dim strFullPath As String

    strTableName = "oSysRelationLayout"
    strFullPath = CurrentProject.FullName
    ' check that the table exists in the current project
    On Error Resume Next
        Set tdf = CurrentDb.TableDefs(strTableName)
        If Not Err = 0 Then
            oTable.LayoutTable_Add
            Err.Clear
            CurrentDb.TableDefs.Refresh
            Set tdf = CurrentDb.TableDefs(strTableName)
            If Not Err = 0 Then
                MsgBox "Relation Layout Table is missing", vbCritical, "Relation Layout"
                GoTo Exit_Procedure
            End If
        Else
            oTable.Sample_Add
        End If
    On Error GoTo 0

    Set qdf = CodeDb.QueryDefs("oFrameRelationLayout_Name")
    str = "SELECT DISTINCT RelationLayout_Name FROM oSysRelationLayout"
    str = str & " IN '" & strFullPath & "'"
    str = str & " ORDER BY RelationLayout_Name;"
    qdf.SQL = str

    Set qdf = CodeDb.QueryDefs("oFrameRelationLayoutTable")
    str = "SELECT RelationLayout_Name" _
                & ", Window_Name" _
                & ", Window_Left" _
                & ", Window_Top" _
                & ", Window_Right" _
                & ", Window_Bottom" _
                & ", [Window_Right]-[Window_Left] AS Width" _
                & ", [Window_Bottom]-[Window_Top] AS Height" _
            & " FROM oSysRelationLayout"
    str = str & " IN '" & strFullPath & "'"
    str = str & " ORDER BY Window_Left, Window_Top;"
    qdf.SQL = str
    

Exit_Procedure:
    On Error Resume Next
    Exit Sub
Error_Handler:
    MsgBox Err.Number & vbNewLine _
            & Err.Description _
            , vbCritical, "Addin_Load"
    Resume Exit_Procedure
    Resume
End Sub


