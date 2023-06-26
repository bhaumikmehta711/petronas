CREATE TABLE [Master].[UserRole]
(
	[UserID] INT NOT NULL,
	[RoleID] INT NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_UserRole_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_UserRole_UserID_RoleID] PRIMARY KEY ([UserID],[RoleID]),
	CONSTRAINT [FK_User_UserRole_UserID] FOREIGN KEY ([UserID]) REFERENCES [Master].[User]([UserID]),
	CONSTRAINT [FK_Role_UserRole_RoleID] FOREIGN KEY ([RoleID]) REFERENCES [Master].[Role]([RoleID])
)
