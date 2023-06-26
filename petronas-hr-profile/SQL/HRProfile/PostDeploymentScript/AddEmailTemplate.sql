-- Load email template
SELECT * INTO #EmailTemplate FROM 
(
	SELECT 
		'Send email to job SPUR requestor' [EmailTemplateName],
		'SendEmailToSPURBatchRequestor' [EmailTemplateCode],
		'SPUR submission is processed' [EmailSubject],
		'<p>Hi<p><p>Batch {{BatchName}} has been moved to {{BatchStatus}} status.</p><p>Please refer <a href="https://apps.powerapps.com/play/e/default-d5d09639-dce2-4298-a34e-404cca1d324a/a/46da03d9-1a61-4093-9198-80d22c35eaf9?tenantId=d5d09639-dce2-4298-a34e-404cca1d324a&BatchID={{batchId}}&source=portal#">batch</a> to take required actions further.<p>' [EmailBody]
)A


-- Insert email template
DBCC CHECKIDENT ('[Master].[EmailTemplate]', RESEED, 1);
INSERT INTO [Master].[EmailTemplate]
(
	[EmailTemplateName],
	[EmailTemplateCode],
	[EmailSubject],
	[EmailBody],
	[EmailSeverity],
	[CreatedBy],
	[CreatedTimestamp]
)
SELECT 
	A.[EmailTemplateName],
	A.[EmailTemplateCode],
	A.[EmailSubject],
	A.[EmailBody],
	A.[EmailSeverity],
	'Admin',
	GETUTCDATE()
FROM #EmailTemplate A
LEFT JOIN [Master].[EmailTemplate] B ON B.EmailTemplateCode = A.EmailTemplateCode
WHERE B.EmailTemplateID IS NULL