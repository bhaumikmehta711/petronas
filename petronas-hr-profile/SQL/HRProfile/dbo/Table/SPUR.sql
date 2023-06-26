﻿CREATE TABLE [dbo].[SPUR]
(
	[SPURID] INT IDENTITY(1, 1) NOT NULL,
	[SPURCode] VARCHAR(20) NOT NULL,
	[SPURName] VARCHAR(100),
	[JobCode] VARCHAR(50),
	[EffectiveStartDate] DATETIME,
	[EffectiveEndDate] DATETIME, 
	[RoleLevel] VARCHAR(100),
	[MinimumExperienceRequiredInYear] NUMERIC(3, 2),
	[DesiredExperienceInYear] NUMERIC(3, 2),
	[Industry] VARCHAR(100),
	[Domain] VARCHAR(100),
	[ContentItem] VARCHAR(100),
	[AreaOfStudy] VARCHAR(100),
	[OtherAreaOfStudy] VARCHAR(100),
	[LicenseAndCertificate] VARCHAR(100),
	[OtherLicnseAndCertificate] VARCHAR(100),
	[Membership] VARCHAR(100),
	[OtherMembership] VARCHAR(100),
	[BatchID] INT,
	[SPURFilePath] VARCHAR(500),
	[PurposeAndAccountability] NVARCHAR(MAX),
	[Challenge] NVARCHAR(MAX),
	[Experience] NVARCHAR(MAX),
	[KPI] NVARCHAR(MAX),
	[ActiveFlag] BIT,
	[SubmittedTimeStamp] DATETIME,
	[EndUserCreatedBy] NVARCHAR(200),
	[EndUserCreatedTimestamp] DATETIME,
	[EndUserModifiedBy] NVARCHAR(200),
	[EndUserModifiedTimestamp] DATETIME,
	[BackendUserModifiedBy] NVARCHAR(200),
	[BackendUserModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_SPUR_SPURID] PRIMARY KEY ([SPURID]),
	CONSTRAINT [FK_Batch_SPUR_BatchID] FOREIGN KEY ([BatchID]) REFERENCES [DBO].[Batch]([BatchID])
)
