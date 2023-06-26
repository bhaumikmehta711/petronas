CREATE VIEW [Master].[vUserRole]
AS 
	SELECT 
		A.UserID, 
		B.EmailID, 
		B.FirstName, 
		B.LastName, 
		C.RoleName, 
		C.RoleCode
	FROM [Master].[UserRole] A
	INNER JOIN [Master].[User] B ON B.UserID = A.UserID AND A.ActiveFlag = 1 AND B.ActiveFlag = 1
	INNER JOIN [Master].[Role] C ON C.RoleID = A.RoleID AND C.ActiveFlag = 1