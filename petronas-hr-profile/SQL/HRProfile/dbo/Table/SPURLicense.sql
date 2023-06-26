CREATE TABLE [dbo].[SPURLicense]
(
	[SPURID] INT NOT NULL,
	[LicenseID] INT NOT NULL,
	[CountryID] INT,
	[State] VARCHAR(500),
	[Title] VARCHAR(500),
	[Required] BIT,
	[Importance] VARCHAR(50) NULL,
	[JG] VARCHAR(50) NULL,
	[CreatedBy] NVARCHAR(200),
	[CreatedTimestamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_SPURLicense_SPURID_LicenseID] PRIMARY KEY ([SPURID], [LicenseID]),
	CONSTRAINT [FK_SPUR_SPURLicense_SPURID] FOREIGN KEY ([SPURID]) REFERENCES [DBO].[SPUR]([SPURID]),
	CONSTRAINT [FK_License_SPURLicense_LicenseID] FOREIGN KEY ([LicenseID]) REFERENCES [Master].[License]([LicenseID]),
	CONSTRAINT [FK_Country_SPURLicense_CountryID] FOREIGN KEY ([CountryID]) REFERENCES [Master].[Country]([CountryID])
)
