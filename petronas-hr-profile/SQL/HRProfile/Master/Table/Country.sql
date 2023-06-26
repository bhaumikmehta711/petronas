CREATE TABLE [Master].[Country]
(
	[CountryID] INT IDENTITY(1, 1) NOT NULL,
	[CountryName] VARCHAR(500) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_Country_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_Country_CountryID] PRIMARY KEY ([CountryID]),
	CONSTRAINT [UC_Country_CountryName] UNIQUE ([CountryName])
)