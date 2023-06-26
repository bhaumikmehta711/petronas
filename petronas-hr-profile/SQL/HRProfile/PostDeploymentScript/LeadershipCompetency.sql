SELECT * INTO #LeadershipCompetency FROM (SELECT 'Energise: Interpersonal Effectiveness (E)' LeadershipCompetencyName UNION ALL
SELECT 'Energise: Foster Collaboration & Teamwork (E)' UNION ALL
SELECT 'Decide: Set Goals & Drive Directions (E)' UNION ALL
SELECT 'Decide: Analysis & Problem Solving (E)' UNION ALL
SELECT 'Grow: Lead Change & Innovation (E)' UNION ALL
SELECT 'Grow: Commitment to Learning & Development (E)' UNION ALL
SELECT 'Execute: Deliver Performance (E)' UNION ALL
SELECT 'Execute: Professionalism & Expertise (E)' UNION ALL
SELECT 'Energise: Interpersonal Effectiveness (M)' UNION ALL
SELECT 'Energise: Foster Collaboration & Teamwork (M)' UNION ALL
SELECT 'Decide: Set Goals & Drive Directions (M)' UNION ALL
SELECT 'Decide: Analysis & Problem Solving (M)' UNION ALL
SELECT 'Grow: Lead Change & Innovation (M)' UNION ALL
SELECT 'Grow: Commitment to Learning & Development (M)' UNION ALL
SELECT 'Execute: Deliver Performance (M)' UNION ALL
SELECT 'Execute: Professionalism & Expertise (M)' UNION ALL
SELECT 'Energise: Interpersonal Effectiveness (SM)' UNION ALL
SELECT 'Energise: Foster Collaboration & Teamwork (SM)' UNION ALL
SELECT 'Decide: Set Goals & Drive Directions (SM)' UNION ALL
SELECT 'Decide: Analysis & Problem Solving (SM)' UNION ALL
SELECT 'Grow: Lead Change & Innovation (SM)' UNION ALL
SELECT 'Grow: Commitment to Learning & Development (SM)' UNION ALL
SELECT 'Execute: Deliver Performance (SM)' UNION ALL
SELECT 'Execute: Professionalism & Expertise (SM)' UNION ALL
SELECT 'Energise: Interpersonal Effectiveness (GM)' UNION ALL
SELECT 'Energise: Foster Collaboration & Teamwork (GM)' UNION ALL
SELECT 'Decide: Set Goals & Drive Directions (GM)' UNION ALL
SELECT 'Decide: Analysis & Problem Solving (GM)' UNION ALL
SELECT 'Grow: Lead Change & Innovation (GM)' UNION ALL
SELECT 'Grow: Commitment to Learning & Development (GM)' UNION ALL
SELECT 'Execute: Deliver Performance (GM)' UNION ALL
SELECT 'Execute: Professionalism & Expertise (GM)' UNION ALL
SELECT 'Energise: Interpersonal Effectiveness (SGM)' UNION ALL
SELECT 'Energise: Foster Collaboration & Teamwork (SGM)' UNION ALL
SELECT 'Decide: Set Goals & Drive Directions (SGM)' UNION ALL
SELECT 'Decide: Analysis & Problem Solving (SGM)' UNION ALL
SELECT 'Grow: Lead Change & Innovation (SGM)' UNION ALL
SELECT 'Grow: Commitment to Learning & Development (SGM)' UNION ALL
SELECT 'Execute: Deliver Performance (SGM)' UNION ALL
SELECT 'Execute: Professionalism & Expertise (SGM)' UNION ALL
SELECT 'Energise: Interpersonal Effectiveness (S)' UNION ALL
SELECT 'Energise: Foster Collaboration & Teamwork (S)' UNION ALL
SELECT 'Decide: Set Goals & Drive Directions (S)' UNION ALL
SELECT 'Decide: Analysis & Problem Solving (S)' UNION ALL
SELECT 'Grow: Lead Change & Innovation (S)' UNION ALL
SELECT 'Grow: Commitment to Learning & Development (S)' UNION ALL
SELECT 'Execute: Deliver Performance (S)' UNION ALL
SELECT 'Execute: Professionalism & Expertise (S)' UNION ALL
SELECT 'Energise: Interpersonal Effectiveness (P)' UNION ALL
SELECT 'Energise: Foster Collaboration & Teamwork (P)' UNION ALL
SELECT 'Decide: Set Goals & Drive Directions (P)' UNION ALL
SELECT 'Decide: Analysis & Problem Solving (P)' UNION ALL
SELECT 'Grow: Lead Change & Innovation (P)' UNION ALL
SELECT 'Grow: Commitment to Learning & Development (P)' UNION ALL
SELECT 'Execute: Deliver Performance (P)' UNION ALL
SELECT 'Execute: Professionalism & Expertise (P)' UNION ALL
SELECT 'Energise: Interpersonal Effectiveness (C)' UNION ALL
SELECT 'Energise: Foster Collaboration & Teamwork (C)' UNION ALL
SELECT 'Decide: Set Goals & Drive Directions (C)' UNION ALL
SELECT 'Decide: Analysis & Problem Solving (C)' UNION ALL
SELECT 'Grow: Lead Change & Innovation (C)' UNION ALL
SELECT 'Grow: Commitment to Learning & Development (C)' UNION ALL
SELECT 'Execute: Deliver Performance (C)' UNION ALL
SELECT 'Execute: Professionalism & Expertise (C)') A

DBCC CHECKIDENT ('[Master].[LeadershipCompetency]', RESEED, 1);
INSERT INTO [Master].[LeadershipCompetency]
(
	[LeadershipCompetencyName],
	[CreatedBy],
	[CreatedTimestamp]
)
SELECT 
	A.[LeadershipCompetencyName],
	'Admin',
	GETUTCDATE()
FROM #LeadershipCompetency A
LEFT JOIN [Master].[LeadershipCompetency] B ON B.LeadershipCompetencyName = A.LeadershipCompetencyName
WHERE B.LeadershipCompetencyID IS NULL