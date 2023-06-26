CREATE TABLE [Master].[Language]
(
	[LanguageID] INT IDENTITY(1, 1) NOT NULL,
	[LanguageName] VARCHAR(500) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_Language_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_Language_LanguageID] PRIMARY KEY ([LanguageID]),
	CONSTRAINT [UC_Language_LanguageName] UNIQUE ([LanguageName])
)