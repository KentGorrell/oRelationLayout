Operation =1
Option =0
Begin InputTables
    Name ="USysRegInfo"
End
Begin OutputColumns
    Expression ="USysRegInfo.*"
End
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
        dbText "Name" ="USysRegInfo.Subkey"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="USysRegInfo.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="USysRegInfo.ValName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="USysRegInfo.Value"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =248
    Top =44
    Right =1541
    Bottom =832
    Left =-1
    Top =-1
    Right =1277
    Bottom =430
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =41
        Top =20
        Right =206
        Bottom =193
        Top =0
        Name ="USysRegInfo"
        Name =""
    End
End
