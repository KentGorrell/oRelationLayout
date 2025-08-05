VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_oFrameRelationLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 2000 VBA
'
'Copyright: Stephen Lebans - Lebans Holdings 1999 Inc.
'           Please feel free to use this code within your own projects,
'           both private and commercial, with no obligation.
'           You may not resell this code by itself or as part of a collection.
'
'
'Name:      Save/Restore Relationship Window view
'
'Version:   1.3
'
'Purpose:   1) To allow the saving of the layout of the Relationship window to a table.
'           2) To allow the restoration of the layout of the Relationship Window from a table.
'           3) To allow the saving/restoration of multiple Relationship Window views.
'           4) To allow the importing of the layout of the Relationship Window from an external accdb.
' 
'Requires: This Form must be bound to the table named oSysRelationLayout
'
'Author:    Stephen Lebans

' Modified by Kent 2025


Option Compare Database
Option Explicit

Private Type POINTAPI
   x As Long
   Y As Long
End Type

Private Type tRect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type


Private Declare PtrSafe Function ScreenToClient Lib "user32" _
                                                (ByVal hwnd As LongPtr _
                                                 , lpPoint As POINTAPI _
                                                 ) As Long

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
                                               (ByVal hWnd1 As LongPtr _
                                                , ByVal hWnd2 As LongPtr _
                                                , ByVal lpsz1 As String _
                                                , ByVal lpsz2 As String _
                                                ) As LongPtr

Private Declare PtrSafe Function GetWindow Lib "user32" _
                                                (ByVal hwnd As LongPtr _
                                                , ByVal wCmd As Long _
                                                ) As LongPtr

Private Declare PtrSafe Function GetWindowRect Lib "user32" _
                                               (ByVal hwnd As LongPtr _
                                                , lpRect As tRect _
                                                ) As Long

Private Declare PtrSafe Function SetWindowPos Lib "user32" _
                                               (ByVal hwnd As LongPtr _
                                                , ByVal hWndInsertAfter As LongPtr _
                                                , ByVal x As Long _
                                                , ByVal Y As Long _
                                                , ByVal cx As Long _
                                                , ByVal cy As Long _
                                                , ByVal wFlags As Long _
                                                ) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                                                (ByVal hwnd As LongPtr _
                                                , ByVal lpString As String _
                                                , ByVal cch As Long _
                                                ) As Long


'Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'               (ByRef destination As Any, ByRef SOURCE As Any, ByVal Length As Long)

' SetWindowPos() Constants
Private Const SWP_SHOWWINDOW = &H40

' GetWindow() Constants
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5


Public Sub Layout_Save(strID As String)

    Dim rst As DAO.Recordset
    Dim lngRet As Long
    Dim hWndMDI As LongPtr
    Dim hWndRel As LongPtr
    Dim hWndODsk As LongPtr
    Dim hWndTemp As LongPtr
    Dim rc As tRect
    Dim pt As POINTAPI
    Dim strWindowName As String

    If Len(strID) = 0 Then
        MsgBox "You Must Enter a Key to identify this Relationship View"
        Exit Sub
    Else
        ' Delete everything (for this layout) so we can start from scratch
        CurrentDb.Execute "DELETE *" _
                            & " FROM oSysRelationLayout" _
                            & " WHERE RelationLayout_Name = '" & strID & "'"
    End If

OpenRelation_Window:
    DoCmd.RunCommand acCmdRelationships

    ' The Relationships Window is a child of the MDI Client window
Find_MDIClient:
    hWndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)

GetRelationship_Window:
    hWndRel = FindWindowEx(hWndMDI, 0&, "OSysRel", "Relationships")

    If hWndRel = 0 Then
        MsgBox "The Relationships Window is not open.", vbCritical, "Critical Error"
        Exit Sub
    End If

    ' The first child window is of class ODsk
First_TableWindow:
    hWndODsk = FindWindowEx(hWndRel, 0&, "ODsk", vbNullString)

    ' Let'strWindowName get first Child Window of the ODsk window
    hWndTemp = GetWindow(hWndODsk, GW_CHILD)
    If hWndTemp = 0 Then
        MsgBox "Their are no Relationships!", vbCritical, "Critical Error"
        Exit Sub
    Else
        Set rst = CurrentDb.OpenRecordset("oSysRelationLayout")
        lngRet = GetWindowRect(hWndTemp, rc)

        ' MUST store coords as Client Not Screen coords
        ' Upper Left corner of window first
Top_Left:
        pt.x = rc.Left
        pt.Y = rc.Top
        lngRet = ScreenToClient(hWndRel, pt)
        rc.Left = pt.x
        rc.Top = pt.Y

        ' Lower right corner of window second
Bottom_Right:
        pt.x = rc.Right
        pt.Y = rc.Bottom
        lngRet = ScreenToClient(hWndRel, pt)
        rc.Right = pt.x
        rc.Bottom = pt.Y
        ' Get the Name for this Window
First_WindowName:
        strWindowName = Space(256)
        lngRet = GetWindowText(hWndTemp, strWindowName, 256)
        strWindowName = Left$(strWindowName, lngRet)
        With rst
            .AddNew
            !RelationLayout_Name = strID
            !Window_Name = strWindowName
            !Window_Left = rc.Left
            !Window_Right = rc.Right
            !Window_Top = rc.Top
            !Window_Bottom = rc.Bottom
            .Update
        End With
    End If

    ' Let'strWindowName walk through every sibling window
    Do

        ' Let'strWindowName get the NEXT SIBLING Window
        hWndTemp = GetWindow(hWndTemp, GW_HWNDNEXT)
        ' Do we have a valid Window handle
        If hWndTemp <> 0 Then
            lngRet = GetWindowRect(hWndTemp, rc)
            ' MUST store coords as Client Not Screen coords
            ' Upper Left corner of window first
            pt.x = rc.Left
            pt.Y = rc.Top
            lngRet = ScreenToClient(hWndRel, pt)
            rc.Left = pt.x
            rc.Top = pt.Y

            ' Lower right corner of window second
            pt.x = rc.Right
            pt.Y = rc.Bottom
            lngRet = ScreenToClient(hWndRel, pt)
            rc.Right = pt.x
            rc.Bottom = pt.Y

            ' Get the Name for this Window
            strWindowName = Space(256)
            lngRet = GetWindowText(hWndTemp, strWindowName, 256)

            strWindowName = Left$(strWindowName, lngRet)
            With rst
                .AddNew
                    !RelationLayout_Name = strID
                    !Window_Name = strWindowName
                    !Window_Left = rc.Left
                    !Window_Right = rc.Right
                    !Window_Top = rc.Top
                    !Window_Bottom = rc.Bottom
                .Update
            End With
        End If

    ' Let'strWindowName Start the process from the Top again.
    ' End this loop if no more Windows.
    Loop While hWndTemp <> 0

End Sub


Public Function Layout_Restore(strID As String) As Boolean
On Error GoTo Err_Restore
 
     ' Create a recordset of the selected records
    Dim rst As DAO.Recordset
    Dim strSQL As String

    Dim lngRet As Long
    Dim hWndMDI As LongPtr
    Dim hWndRel As LongPtr
    Dim hWndODsk As LongPtr
    Dim hWndTemp

    Dim strTableName As String
    Dim strWindowName As String
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngAdjust As Long

    Dim i As Long

    ' Collection for Window Text and hWnd
    Dim colWindows As New Collection

    ' Ensure we are zero
    hWndTemp = 0

Open_RelationWindow:
    DoCmd.RunCommand acCmdRelationships

    ' Have they selected a previously saved Relationship View from our ListBox
    If Len(strID) = 0 Then
        MsgBox "You Must Select an Existing Relationship View"
        Layout_Restore = False
        Exit Function
    End If


    ' The Relationships Window is a child of the MDI Client window
MDIClient_Find:
    hWndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)

RelationWindow_Find:
    hWndRel = FindWindowEx(hWndMDI, 0&, "OSysRel", "Relationships")

    If hWndRel = 0 Then  ' Do we have a valid Window handle?
        MsgBox "The Relationships Window is not open.", vbCritical, "Critical Error"
        Layout_Restore = False
        Exit Function
    End If

First_Window:
    ' The first child window is of class ODsk
    hWndODsk = FindWindowEx(hWndRel, 0&, "ODsk", vbNullString)

    strSQL = "SELECT *" _
            & " FROM oSysRelationLayout" _
            & " WHERE RelationLayout_Name = '" & strID & "'"
    Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

    ' Let's get first Child Window of the ODsk window
    hWndTemp = GetWindow(hWndODsk, GW_CHILD)

    If hWndTemp = 0 Then
        MsgBox "There are no Relationships!", vbCritical, "Critical Error"
        Layout_Restore = False
        Exit Function
    Else
        ' Add this window to our collection
        ' Grab the Windows Text
        strTableName = Space(256)
        lngRet = GetWindowText(hWndTemp, strTableName, 256)
        strTableName = Left$(strTableName, lngRet)
        colWindows.Add hWndTemp, strTableName
    End If

    ' Loop through the rest of the sibling windows adding them to our collection
    Do
        hWndTemp = GetWindow(hWndTemp, GW_HWNDNEXT)
        If hWndTemp = 0 Then Exit Do
        ' Grab the Windows Text
        strTableName = Space(256)
        lngRet = GetWindowText(hWndTemp, strTableName, 256)
        strTableName = Left$(strTableName, lngRet)
        colWindows.Add hWndTemp, strTableName
        
    Loop

    ' Walk through our windows collection and match Saved hWnd with current record
    ' Move our recordset to the first row
    With rst
        .MoveFirst
        ' Loop through all of the rows
        Do Until .EOF
            On Error Resume Next
            strWindowName = !Window_Name
            ' Get the Window handle by using the Window's Text as the Key
            hWndTemp = colWindows.Item(strWindowName)
            On Error GoTo 0
            If hWndTemp <> 0 Then
                ' Copy the window coords into our temp vars
                lngLeft = .Fields("Window_Left")
'                If !Window_Visible Then
                    lngRight = .Fields("Window_Right")
'                Else
'                    lngRight = lngLeft
'                End If
                lngTop = .Fields("Window_Top")
                lngBottom = .Fields("Window_Bottom")
                If lngTop < 0 Then
                    lngAdjust = Abs(lngTop)
                    lngTop = 0
                    lngBottom = lngBottom + lngAdjust
                End If
                If lngLeft < 0 Then
                    lngAdjust = Abs(lngLeft)
                    lngLeft = 0
                    lngRight = lngRight + lngAdjust
                End If
                ' Move and size the window to its original values
                lngRet = SetWindowPos(hWndTemp _
                                        , 0& _
                                        , lngLeft _
                                        , lngTop _
                                        , lngRight - lngLeft _
                                        , lngBottom - lngTop _
                                        , SWP_SHOWWINDOW)
            End If
            .MoveNext
            hWndTemp = 0
        Loop ' eof
    End With
    ' Return Success
    Layout_Restore = True


 ' All done
Exit_Restore:
    Set rst = Nothing
    ' Delete our Collection object
    ' Remove the first object each time  through the loop until there are
    ' no objects left in the collection.
    Dim x As Long
    For x = 1 To colWindows.Count
        colWindows.Remove 1
    Next x
Exit Function

Err_Restore:
    MsgBox Err.Description
    Resume Exit_Restore

End Function


Private Sub cmdClose_Click()
    DoCmd.Close
End Sub


Private Sub cmdDelete_Click()
Dim strID As String

    strID = Me.lstRelationLayout_Name
    CurrentDb.Execute "DELETE *" _
                        & " FROM oSysRelationLayout" _
                        & " WHERE RelationLayout_Name = '" & strID & "'" _
                        , dbFailOnError + dbSeeChanges
    oTable.Sample_Add ' add a record if there aren't any otherwise the form won't show controls
    Me.oFrameRelationLayoutTable.Requery
    Me.lstRelationLayout_Name.Requery
End Sub


Private Sub cmdNew_Click()
    Me.txtRelationLayout_Name = Null
    Me.txtRelationLayout_Name.SetFocus
'    Call txtRelationLayout_Name_AfterUpdate
End Sub


Private Sub cmdRestore_Click()
    Layout_Restore (Me.lstRelationLayout_Name)
End Sub


Private Sub cmdSave_Click()
    Dim strLayoutName As String

    strLayoutName = Me.txtRelationLayout_Name & ""
    If strLayoutName = "" Then strLayoutName = Me.lstRelationLayout_Name & ""

    If strLayoutName = "" Then
        MsgBox "Select an Existing Layout or Enter a New Name", vbOKOnly
    Else
        Layout_Save strLayoutName
'        Me.lstRelationLayout_Name.Requery
        If Not Me.txtRelationLayout_Name & "" = "" Then
            Me.txtRelationLayout_Name = Null
            Me.lstRelationLayout_Name.Requery
        End If
    End If
End Sub


Private Sub Form_Open(Cancel As Integer)
'    Addin_Load
'    Me.oFrameRelationLayoutTable.SourceObject = "oFrameRelationLayoutTable"
    Call lstRelationLayout_Name_AfterUpdate
End Sub


Private Sub lblTitle_DblClick(Cancel As Integer)
    DoCmd.Close
End Sub


Private Sub lstRelationLayout_Name_AfterUpdate()
    Me.txtRelationLayout_Name = Me.lstRelationLayout_Name
    Call txtRelationLayout_Name_AfterUpdate
End Sub


Private Sub lstRelationLayout_Name_Click()
    Me.txtRelationLayout_Name = Me.lstRelationLayout_Name
End Sub


Private Sub txtRelationLayout_Name_AfterUpdate()
    FindMe
    Me.cmdSave.Enabled = Len(Me.txtRelationLayout_Name & "") > 0
    Me.cmdRestore.Enabled = Len(Me.lstRelationLayout_Name & "") > 0
    Me.lblRestore.Caption = "Restore " & Me.lstRelationLayout_Name & vbNewLine _
                                & "and " & vbNewLine _
                                & "Open Relationship Window"
    Me.cmdDelete.Enabled = Len(Me.lstRelationLayout_Name & "") > 0
    Me.lblDelete.Caption = "Delete " & Me.lstRelationLayout_Name
End Sub


Private Sub txtRelationLayout_Name_Change()
    Me.cmdSave.Enabled = Len(Me.txtRelationLayout_Name.Text) > 0
End Sub


Public Function FindMe()
Dim rst As DAO.Recordset
    Set rst = Me.RecordsetClone
    With rst
        .FindFirst "RelationLayout_Name = '" & Me.lstRelationLayout_Name & "'"
        If Not .NoMatch Then Me.Bookmark = .Bookmark
    End With
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
        ' we need to transfer a copy or???
        MsgBox "Table is missing", vbCritical, "Relation Layout"
        GoTo Exit_Procedure
    End If
On Error GoTo 0
    If Not CodeDb Is CurrentDb Then
        strFullPath = CurrentProject.FullName
    End If
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

