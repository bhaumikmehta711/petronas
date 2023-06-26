CREATE TABLE [dbo].[PositionLanguage]
(
	[PositionID] INT NOT NULL,
	[LanguageID] INT NOT NULL,
	[ReadingProficiency] VARCHAR(10) NULL,
	[WritingProficiency] VARCHAR(10) NULL,
	[SpeakingProficiency] VARCHAR(10) NULL,
	[Required] BIT,
	[CreatedBy] NVARCHAR(200),
	[CreatedTimestamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_PositionLanuage_PositionID_LanguageID] PRIMARY KEY ([PositionID], [LanguageID]),
	CONSTRAINT [FK_Position_PositionLanguage_PositionID] FOREIGN KEY ([PositionID]) REFERENCES [DBO].[Position]([PositionID]),
	CONSTRAINT [FK_Language_PositionLanguage_LanguageID] FOREIGN KEY ([LanguageID]) REFERENCES [Master].[Language]([LanguageID])
)
