CREATE TABLE [dbo].[SPURDegree]
(
	[SPURID] INT NOT NULL,
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
	CONSTRAINT [PK_SPURDegree_SPURID_DegreeID] PRIMARY KEY ([SPURID], [DegreeID], [StudyAreaID]),
	CONSTRAINT [FK_SPUR_SPURDegree_SPURID] FOREIGN KEY ([SPURID]) REFERENCES [DBO].[SPUR]([SPURID]),
	CONSTRAINT [FK_StudyArea_SPURDegree_StudyAreaID] FOREIGN KEY ([StudyAreaID]) REFERENCES [Master].[StudyArea]([StudyAreaID]),
	CONSTRAINT [FK_Country_SPURDegree_CountryID] FOREIGN KEY ([CountryID]) REFERENCES [Master].[Country]([CountryID]),
	CONSTRAINT [FK_School_SPURDegree_SchoolID] FOREIGN KEY ([SchoolID]) REFERENCES [Master].[School]([SchoolID])
)
