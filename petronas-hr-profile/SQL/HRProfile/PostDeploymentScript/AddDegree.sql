SELECT * INTO #Degree FROM (SELECT 'Bachelor''s' [DegreeName] UNION ALL
SELECT 'Diploma' UNION ALL
SELECT 'Doctorate' UNION ALL
SELECT 'High School' UNION ALL
SELECT 'Other' UNION ALL
SELECT 'Post-Doctorate' UNION ALL
SELECT 'Pre-University' UNION ALL
SELECT 'Certificate' UNION ALL
SELECT 'Master''s') A

DBCC CHECKIDENT ('[Master].[Degree]', RESEED, 1);
INSERT INTO [Master].[Degree]
(
	[DegreeName],
	[CreatedBy],
	[CreatedTimestamp]
)
SELECT 
	A.[DegreeName],
	'Admin',
	GETUTCDATE()
FROM #Degree A
LEFT JOIN [Master].[Degree] B ON B.DegreeName = A.DegreeName
WHERE B.DegreeID IS NULL