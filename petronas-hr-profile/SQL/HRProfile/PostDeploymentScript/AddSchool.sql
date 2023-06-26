SELECT * INTO #School FROM (SELECT 'Universiti Teknologi PETRONAS (UTP) PETRONAS University of Technology' SchoolName UNION ALL
SELECT 'MS University' UNION ALL
SELECT 'Universiti Teknologi PETRONAS') A

DBCC CHECKIDENT ('[Master].[School]', RESEED, 1);
INSERT INTO [Master].[School]
(
	[SchoolName],
	[CreatedBy],
	[CreatedTimestamp]
)
SELECT 
	A.SchoolName,
	'Admin',
	GETUTCDATE()
FROM #School A
LEFT JOIN [Master].[School] B ON B.SchoolName = A.SchoolName
WHERE B.SchoolID IS NULL