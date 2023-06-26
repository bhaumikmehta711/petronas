CREATE PROCEDURE [audit].[uspGetEmailData]
	@pEmailTrackingID BIGINT
AS
BEGIN
	SET NOCOUNT ON;

	SELECT 
		A.[To],
		A.CC,
		B.EmailSubject,
		REPLACE(REPLACE(REPLACE(B.EmailBody, '{{BatchName}}', C.BatchName), '{{BatchStatus}}', C.BatchStatus), '{{BatchID}}', C.BatchID) EmailBody
	FROM Audit.EmailTracking A
	INNER JOIN Master.EmailTemplate B ON A.EmailTemplateID = B.EmailTemplateID
	INNER JOIN DBO.Batch C ON C.BatchID = A.CategoryID AND A.Category = 'Batch'
	WHERE A.EmailTrackingID = @pEmailTrackingID AND A.Status = 'Pending'
END