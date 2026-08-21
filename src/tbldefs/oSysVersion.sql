CREATE TABLE [oSysVersion] (
  [Version_ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Version_Date] DATETIME,
  [Version_Number] VARCHAR (20),
  [Version_Detail] LONGTEXT
)
