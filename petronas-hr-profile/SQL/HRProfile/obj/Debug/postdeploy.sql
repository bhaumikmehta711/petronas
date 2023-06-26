-- Load users along with their roles
SELECT EmailID, FirstName, LastName, RoleName, RoleCode INTO #UserRole FROM 
(SELECT 'nasruldeen.anua@petronas.com.my' EmailID, 'Nasruldeen' FirstName, 'Anuar' LastName, 'Requestor' RoleName, 'Requestor' RoleCode UNION ALL
SELECT 'sitimarina.abdrahi@petronas.com.my' EmailID, 'Siti' FirstName, 'Marina' LastName, 'Requestor' RoleName, 'Requestor' RoleCode UNION ALL
SELECT 'linaarina.mohamadaf@petronas.com.my' EmailID, 'Lina' FirstName, 'Arina' LastName, 'Requestor' RoleName, 'Requestor' RoleCode UNION ALL
SELECT 'mfathullah.amiruddin@petronas.com.my' EmailID, 'M Fathullah' FirstName, 'Amiruddin' LastName, 'Requestor' RoleName, 'Requestor' RoleCode UNION ALL
SELECT 'ainakhalishah.murshi@petronas.com.my' EmailID, 'Aina' FirstName, 'Khalishah' LastName, 'Requestor' RoleName, 'Requestor' RoleCode UNION ALL
SELECT 'danielakmal.anuar@petronas.com.my' EmailID, 'Daniel' FirstName, 'Akmal' LastName, 'Requestor' RoleName, 'Requestor' RoleCode UNION ALL
SELECT 'nursyakirah.zainal@petronas.com.my' EmailID, 'Nursyakirah' FirstName, 'Zainal' LastName, 'Requestor' RoleName, 'Requestor' RoleCode UNION ALL
SELECT 'hazilah_alba@petronas.com.my' EmailID, 'Hazilah' FirstName, 'Alba' LastName, 'Approver' RoleName, 'Approver' RoleCode UNION ALL
SELECT 'atikah.mohamaalimoh@petronas.com.my' EmailID, 'Atikah' FirstName, 'Mohama Ali' LastName, 'Approver' RoleName, 'Approver' RoleCode UNION ALL
SELECT 'sitimarina.abdrahi@petronas.com.my' EmailID, 'Siti' FirstName, 'Marina' LastName, 'Approver' RoleName, 'Approver' RoleCode UNION ALL
SELECT 'linaarina.mohamadaf@petronas.com.my' EmailID, 'Lina' FirstName, 'Arina' LastName, 'Approver' RoleName, 'Approver' RoleCode UNION ALL
SELECT 'anand.karuppaiah@petronas.com' EmailID, 'Anandaraju' FirstName, 'Karuppaiah' LastName, 'Technical Support' RoleName, 'TechnicalSupport' RoleCode UNION ALL
SELECT 'haniyasmin.zaki@petronas.com' EmailID, 'Hani' FirstName, 'Zaki' LastName, 'Technical Support' RoleName, 'TechnicalSupport' RoleCode UNION ALL
SELECT 'noorhidayah.hashim@petronas.com' EmailID, 'Noor' FirstName, 'Hidayah' LastName, 'Technical Support' RoleName, 'TechnicalSupport' RoleCode) A

-- Insert users
DBCC CHECKIDENT ('[Master].[User]', RESEED, 1);
INSERT INTO [Master].[User] 
(
	EmailID,
	FirstName,
	LastName,
	CreatedBy,
	CreatedTimestamp
)
SELECT DISTINCT
	A.EmailID,
	A.FirstName,
	A.LastName,
	'Admin',
	GETUTCDATE()
FROM #UserRole A
LEFT JOIN [Master].[User] B ON B.EmailID = A.EmailID
WHERE B.UserID IS NULL
ORDER BY A.EmailID

-- Insert roles
DBCC CHECKIDENT ('[Master].[Role]', RESEED, 1);
INSERT INTO [Master].[Role] 
(
	RoleName,
	RoleCode,
	CreatedBy,
	CreatedTimestamp
)
SELECT DISTINCT
	A.RoleName,
	A.RoleCode,
	'Admin',
	GETUTCDATE()
FROM #UserRole A
LEFT JOIN [Master].[Role] B ON B.RoleCode = A.RoleCode
WHERE B.RoleID IS NULL
ORDER BY A.RoleCode

-- Insert user-roles
INSERT INTO [Master].[UserRole] 
(
	UserID,
	RoleID,
	CreatedBy,
	CreatedTimestamp
)
SELECT 
	B.UserID,
	C.RoleID,
	'Admin',
	GETUTCDATE()
FROM #UserRole A
INNER JOIN [Master].[User] B ON B.EmailID = A.EmailID
INNER JOIN [Master].[Role] C ON C.RoleCode = A.RoleCode
LEFT JOIN [Master].[UserRole] D ON D.UserID = B.UserID AND D.RoleID = C.RoleID
WHERE D.UserID IS NULL
ORDER BY B.UserID, C.RoleID
GO
