CREATE TABLE [Master].[License]
(
	[LicenseID] INT IDENTITY(1, 1) NOT NULL,
	[LicenseName] VARCHAR(500) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_License_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_License_LicenseID] PRIMARY KEY ([LicenseID]),
	CONSTRAINT [UC_License_LicenseName] UNIQUE ([LicenseName])
)