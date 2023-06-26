CREATE TABLE [Master].[Role]
(
	[RoleID] INT IDENTITY(1, 1) NOT NULL,
	[RoleName] VARCHAR(50) NOT NULL,
	[RoleCode] VARCHAR(20) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_Role_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_Role_RoleID] PRIMARY KEY ([RoleID]),
	CONSTRAINT [UC_Role_RoleCode] UNIQUE ([RoleCode])
)
