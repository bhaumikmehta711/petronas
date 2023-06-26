SELECT * INTO #Award FROM (SELECT 'Long Service Award - 15 years' [AwardName] UNION ALL
SELECT 'Long Service Award - 10 years' UNION ALL
SELECT 'Long Service Award - 35 years' UNION ALL
SELECT 'Service Award' UNION ALL
SELECT 'Long Service Award - 25 years' UNION ALL
SELECT 'Long Service Award - 40 years' UNION ALL
SELECT 'Long Service Award - Retiree' UNION ALL
SELECT 'Outstanding Contributor Award' UNION ALL
SELECT 'Long Service Award - 30 years' UNION ALL
SELECT 'Dean''s List' UNION ALL
SELECT 'Long Service Award - 20 years' UNION ALL
SELECT 'Board of Directors Appointment') A

DBCC CHECKIDENT ('[Master].[Award]', RESEED, 1);
INSERT INTO [Master].[Award]
(
	[AwardName],
	[CreatedBy],
	[CreatedTimestamp]
)
SELECT 
	A.[AwardName],
	'Admin',
	GETUTCDATE()
FROM #Award A
LEFT JOIN [Master].[Award] B ON B.AwardName = A.AwardName
WHERE B.AwardID IS NULL