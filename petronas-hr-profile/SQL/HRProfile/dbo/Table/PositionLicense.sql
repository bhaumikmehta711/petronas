CREATE TABLE [dbo].[PositionLicense]
(
	[PositionID] INT NOT NULL,
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
	CONSTRAINT [PK_PositionLicense_PositionID_LicenseID] PRIMARY KEY ([PositionID], [LicenseID]),
	CONSTRAINT [FK_Position_PositionLicense_PositionID] FOREIGN KEY ([PositionID]) REFERENCES [DBO].[Position]([PositionID]),
	CONSTRAINT [FK_License_PositionLicense_LicenseID] FOREIGN KEY ([LicenseID]) REFERENCES [Master].[License]([LicenseID]),
	CONSTRAINT [FK_Country_PositionLicense_CountryID] FOREIGN KEY ([CountryID]) REFERENCES [Master].[Country]([CountryID])
)
