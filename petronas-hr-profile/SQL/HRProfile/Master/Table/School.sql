CREATE TABLE [Master].[School]
(
	[SchoolID] INT IDENTITY(1, 1) NOT NULL,
	[SchoolName] VARCHAR(500) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_School_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_School_SchoolID] PRIMARY KEY ([SchoolID]),
	CONSTRAINT [UC_School_SchoolName] UNIQUE ([SchoolName])
)