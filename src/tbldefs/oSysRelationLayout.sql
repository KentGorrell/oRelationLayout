CREATE TABLE [oSysRelationLayout] (
  [RelationLayout_Name] VARCHAR (50),
  [Window_Name] VARCHAR (255),
  [Window_Left] LONG,
  [Window_Top] LONG,
  [Window_Right] LONG,
  [Window_Bottom] LONG,
  [Window_Visible] BIT,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([RelationLayout_Name], [Window_Name])
)
