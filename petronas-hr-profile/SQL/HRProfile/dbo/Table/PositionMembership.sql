CREATE TABLE [dbo].[PositionMembership]
(
	[PositionID] INT NOT NULL,
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
	CONSTRAINT [PK_PositionMembership_PositionID_MembershipID] PRIMARY KEY ([PositionID], [MembershipID]),
	CONSTRAINT [FK_Position_PositionMembership_PositionID] FOREIGN KEY ([PositionID]) REFERENCES [DBO].[Position]([PositionID]),
	CONSTRAINT [FK_Language_PositionMembership_MembershipID] FOREIGN KEY ([MembershipID]) REFERENCES [Master].[Membership]([MembershipID])
)
