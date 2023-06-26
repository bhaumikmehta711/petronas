	CREATE PROCEDURE [dbo].[uspGetBatchForProcessing]
AS
BEGIN
	SET NOCOUNT ON;

	SELECT 
		BatchID,
		BatchName
	INTO #tmpBatch
	FROM [dbo].[Batch] 
	WHERE [BatchProcessedStatus] = 'Pending' 
	
	UPDATE A
	SET 
		[BatchProcessedStatus] = 'Processing'
	FROM [dbo].[Batch] A
	INNER JOIN #tmpBatch B ON A.BatchID = B.BatchID

	SELECT * FROM #tmpBatch
END