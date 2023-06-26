CREATE TABLE [Master].[EmailTemplate]
(
	[EmailTemplateID] INT IDENTITY(1, 1) NOT NULL,
	[EmailTemplateName] VARCHAR(50) NOT NULL,
	[EmailTemplateCode] VARCHAR(50) NOT NULL,
	[EmailSubject] NVARCHAR(1000) NOT NULL,
	[EmailBody] NVARCHAR(MAX) NOT NULL,
	[EmailSeverity] VARCHAR(20),
	[ActiveFlag] BIT CONSTRAINT [DC_EmailTemplate_ActiveFlag] DEFAULT(1) NOT NULL,
	[CreatedBy] NVARCHAR(200) NOT NULL,
	[CreatedTimestamp] DATETIME NOT NULL,
	[ModifiedBy] NVARCHAR(200),
	[ModifiedTimestamp] DATETIME,
	CONSTRAINT [PK_EmailTemplate_UserID] PRIMARY KEY ([EmailTemplateID]),
	CONSTRAINT [UC_EmailTemplate_EmailTemplateCode] UNIQUE ([EmailTemplateCode])
)
