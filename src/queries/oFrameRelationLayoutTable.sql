SELECT
  RelationLayout_Name,
  Window_Name,
  Window_Left,
  Window_Top,
  Window_Right,
  Window_Bottom,
  [Window_Right] - [Window_Left] AS Width,
  [Window_Bottom] - [Window_Top] AS Height
FROM
  oSysRelationLayout IN 'E:\App\DPP\Desktop Promotions TABLES.accdb'
ORDER BY
  Window_Left,
  Window_Top;
