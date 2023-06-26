CREATE TABLE [dbo].[SPURLanguage]
(
	[SPURID] INT NOT NULL,
	[LanguageID] INT NOT NULL,
	[ReadingProficiency] VARCHAR(10) NULL,
	[WritingProficiency] VARCHAR(10) NULL,
	[SpeakingProficiency] VARCHAR(10) NULL,
	[Required] BIT,
	[CreatedBy] NVARCHAR(200),
	[CreatedTimestamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_SPURLanuage_SPURID_LanguageID] PRIMARY KEY ([SPURID], [LanguageID]),
	CONSTRAINT [FK_SPUR_SPURLanguage_SPURID] FOREIGN KEY ([SPURID]) REFERENCES [DBO].[SPUR]([SPURID]),
	CONSTRAINT [FK_Language_SPURLanguage_LanguageID] FOREIGN KEY ([LanguageID]) REFERENCES [Master].[Language]([LanguageID])
)
