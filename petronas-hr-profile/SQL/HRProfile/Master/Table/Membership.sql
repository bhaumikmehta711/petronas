CREATE TABLE [Master].[Membership]
(
	[MembershipID] INT IDENTITY(1, 1) NOT NULL,
	[MembershipName] VARCHAR(500) NOT NULL,
	[ActiveFlag] BIT CONSTRAINT [DC_Membership_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_Membership_MembershipID] PRIMARY KEY ([MembershipID]),
	CONSTRAINT [UC_Membership_MembershipName] UNIQUE ([MembershipName])
)