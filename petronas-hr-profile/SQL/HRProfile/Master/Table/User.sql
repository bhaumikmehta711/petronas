﻿CREATE TABLE [Master].[User]
(
	[UserID] INT IDENTITY(1, 1) NOT NULL,
	[EmailID] NVARCHAR(200) NOT NULL,
	[FirstName] NVARCHAR(50) NOT NULL,
	[LastName] NVARCHAR(50) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_User_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_User_UserID] PRIMARY KEY ([UserID]),
	CONSTRAINT [UC_User_EmailID] UNIQUE ([EmailID])
)
