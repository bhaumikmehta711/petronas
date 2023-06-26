CREATE PROCEDURE [audit].[uspAddEmail]
	@pBatchID INT
AS
BEGIN
	SET NOCOUNT ON;

	INSERT INTO [Audit].[EmailTracking]
	(
		[EmailTemplateID],
		[Category],
		[CategoryID],
		[To],
		[CC],
		[Status],
		[CreatedBy],
		[CreatedTimeStamp]
	)
	SELECT DISTINCT
		C.EmailTemplateID,
		'Batch',
		A.BatchID,
		A.Approver,
		null,
		'Pending',
		CURRENT_USER,
		GETUTCDATE()
	FROM [dbo].[Batch] A
	LEFT JOIN [dbo].[SPUR] B ON A.BatchID = B.BatchID AND A.BatchType = 'SPUR' AND B.ActiveFlag = 1
	INNER JOIN [Master].[EmailTemplate] C ON ((C.EmailTemplateCode = 'SendEmailToSPURBatchRequestor' AND A.BatchStatus = 'Approved') OR (C.EmailTemplateCode = 'SendEmailToSPURBatchApprover' AND A.BatchStatus = 'Pending Submit')) AND C.ActiveFlag = 1
	WHERE A.BatchID = @pBatchID
END