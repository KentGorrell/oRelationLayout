Attribute VB_Name = "oTable"
Option Compare Database
Option Explicit

Public Function LayoutTable_Add() As Boolean
On Error GoTo Error_Handler
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim tdf As DAO.TableDef
Dim fld As DAO.Field
Dim idx As DAO.Index
Dim prp As DAO.Property
Dim strTableName As String

    strTableName = "oSysRelationLayout"
    Set db = CurrentDb
    On Error Resume Next
        Set tdf = db.TableDefs(strTableName)
        If Err <> 0 Then
            ' create the table
            Set tdf = db.CreateTableDef(strTableName)
            Err.Clear
            With tdf
                Set fld = .CreateField("RelationLayout_Name", dbText, 100)
                    fld.AllowZeroLength = False
                    fld.Required = True
                .Fields.Append fld
                Set fld = .CreateField("Window_Name", dbText, 100)
                    fld.AllowZeroLength = False
                    fld.Required = True
                .Fields.Append fld
                Set idx = .CreateIndex("PrimaryKey")
                    Set fld = idx.CreateField("RelationLayout_Name")
                        idx.Fields.Append fld
                    Set fld = idx.CreateField("Window_Name")
                        idx.Fields.Append fld
                        idx.Primary = True
                .Indexes.Append idx
            End With
            db.TableDefs.Append tdf
        Else
            
        End If
    On Error GoTo 0
On Error GoTo Error_Handler
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(strTableName)
    With tdf
        Set fld = .CreateField("Window_Left", dbLong)
            fld.DefaultValue = 0
        .Fields.Append fld
        Set fld = .CreateField("Window_Top", dbLong)
            fld.DefaultValue = 0
        .Fields.Append fld
        Set fld = .CreateField("Window_Right", dbLong)
            fld.DefaultValue = 0
        .Fields.Append fld
        Set fld = .CreateField("Window_Bottom", dbLong)
            fld.DefaultValue = 0
        .Fields.Append fld
        
    End With
    Set rst = db.OpenRecordset(strTableName)
    With rst
        If .EOF Then
            .AddNew
                !RelationLayout_Name = "Layout"
                !Window_Name = "sample"
            .Update
        End If
    End With

Exit_Procedure:
    On Error Resume Next
    Set db = Nothing
    Set tdf = Nothing
    Set fld = Nothing
    Exit Function
Error_Handler:
    If Err = 3191 Then
        ' column exists
        Resume Next
    End If
    MsgBox Err.Number & vbNewLine _
            & Err.Description _
            , vbCritical, "CreateLocalDataTable"
    Resume Exit_Procedure
    Resume
End Function


Public Sub Sample_Add()
Dim rst As DAO.Recordset
Dim strTableName As String

    strTableName = "oSysRelationLayout"

    Set rst = CurrentDb.OpenRecordset(strTableName)
    With rst
        If .EOF Then
            .AddNew
                !RelationLayout_Name = "Layout"
                !Window_Name = "sample"
            .Update
        End If
    End With
    Set rst = Nothing
End Sub
