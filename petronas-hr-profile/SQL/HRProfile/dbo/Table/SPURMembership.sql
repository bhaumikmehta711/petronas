CREATE TABLE [dbo].[SPURMembership]
(
	[SPURID] INT NOT NULL,
	[MembershipID] INT NOT NULL,
	[Title] VARCHAR(500),
	[Establishment] VARCHAR(500),
	[Required] BIT,
	[Importance] VARCHAR(50) NULL,
	[JG] VARCHAR(50) NULL,
	[CreatedBy] NVARCHAR(200),
	[CreatedTimestamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_SPURMembership_SPURID_MembershipID] PRIMARY KEY ([SPURID], [MembershipID]),
	CONSTRAINT [FK_SPUR_SPURMembership_SPURID] FOREIGN KEY ([SPURID]) REFERENCES [DBO].[SPUR]([SPURID]),
	CONSTRAINT [FK_Language_SPURMembership_MembershipID] FOREIGN KEY ([MembershipID]) REFERENCES [Master].[Membership]([MembershipID])
)
