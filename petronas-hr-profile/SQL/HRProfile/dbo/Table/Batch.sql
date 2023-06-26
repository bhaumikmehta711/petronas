CREATE TABLE [dbo].[Batch]
(
	[BatchID] INT IDENTITY(1, 1) NOT NULL,
	[BatchName] VARCHAR(50) NOT NULL,
	[SubmittedBy] NVARCHAR(200) NOT NULL,
	[SubmittedTimeStamp] DATETIME NOT NULL,
	[BatchType] VARCHAR(20) NOT NULL,
	[BatchStatus] VARCHAR(20),
	[BatchProcessedStatus] VARCHAR(20) CONSTRAINT [DC_Batch_BatchProcessedStatus] DEFAULT('Pending') NOT NULL,
	[TicketNumber] VARCHAR(50),
	[Approver] NVARCHAR(200),
	[EndUserCreatedBy] NVARCHAR(200),
	[EndUserCreatedTimestamp] DATETIME,
	[EndUserModifiedBy] NVARCHAR(200),
	[EndUserModifiedTimestamp] DATETIME,
	[BackendUserModifiedBy] NVARCHAR(200),
	[BackendUserModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_Batch_BatchID] PRIMARY KEY ([BatchID])
)
