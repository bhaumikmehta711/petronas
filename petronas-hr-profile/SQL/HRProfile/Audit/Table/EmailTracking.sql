CREATE TABLE [Audit].[EmailTracking]
(
	[EmailTrackingID] BIGINT IDENTITY(1, 1) NOT NULL,
	[EmailTemplateID] INT,
	[Category] VARCHAR(50) NOT NULL,
	[CategoryID] INT,
	[From] NVARCHAR(200),
	[To] NVARCHAR(MAX),
	[CC] NVARCHAR(MAX),
	[Status] VARCHAR(20) CONSTRAINT [DC_Audit_EmailTracking] DEFAULT('Pending') NOT NULL,
	[SentTimeStamp] DATETIME,
	[MessageID] UNIQUEIDENTIFIER,
	[CreatedBy] NVARCHAR(200),
	[CreatedTimeStamp] DATETIME,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimeStamp] DATETIME,
	CONSTRAINT [PK_EmailTracking_EmailTrackingID] PRIMARY KEY ([EmailTrackingID]),
	CONSTRAINT [FK_EmailTemplate_EmailTracking_EmailTemplateID] FOREIGN KEY ([EmailTemplateID]) REFERENCES [Master].[EmailTemplate]([EmailTemplateID])
)
