CREATE TABLE [dbo].[PositionDegree]
(
	[PositionID] INT NOT NULL,
	[DegreeID] INT NOT NULL,
	[Major] VARCHAR(500),
	[Required] BIT,
	[SchoolID] INT,
	[CountryID] INT,
	[StudyAreaID] INT,
	[CreatedBy] NVARCHAR(200),
	[CreatedTimestamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_PositionDegree_PositionID_DegreeID] PRIMARY KEY ([PositionID], [DegreeID], [StudyAreaID]),
	CONSTRAINT [FK_Position_PositionDegree_PositionID] FOREIGN KEY ([PositionID]) REFERENCES [DBO].[Position]([PositionID]),
	CONSTRAINT [FK_StudyArea_PositionDegree_StudyAreaID] FOREIGN KEY ([StudyAreaID]) REFERENCES [Master].[StudyArea]([StudyAreaID]),
	CONSTRAINT [FK_Country_PositionDegree_CountryID] FOREIGN KEY ([CountryID]) REFERENCES [Master].[Country]([CountryID]),
	CONSTRAINT [FK_School_PositionDegree_SchoolID] FOREIGN KEY ([SchoolID]) REFERENCES [Master].[School]([SchoolID])
)
