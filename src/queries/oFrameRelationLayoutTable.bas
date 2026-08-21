dbMemo "SQL" ="SELECT RelationLayout_Name, Window_Name, Window_Left, Window_Top, Window_Right, "
    "Window_Bottom, [Window_Right]-[Window_Left] AS Width, [Window_Bottom]-[Window_To"
    "p] AS Height\015\012FROM oSysRelationLayout IN 'E:\\App\\DPP\\Desktop Promotions"
    " TABLES.accdb'\015\012ORDER BY Window_Left, Window_Top;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Width"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Window_Right"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Window_Bottom"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RelationLayout_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Window_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Window_Left"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Window_Top"
        dbLong "AggregateType" ="-1"
    End
End
