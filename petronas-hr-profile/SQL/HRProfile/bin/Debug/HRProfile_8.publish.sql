﻿/*
Deployment script for HRProfile

This code was generated by a tool.
Changes to this file may cause incorrect behavior and will be lost if
the code is regenerated.
*/

GO
SET ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, CONCAT_NULL_YIELDS_NULL, QUOTED_IDENTIFIER ON;

SET NUMERIC_ROUNDABORT OFF;


GO
:setvar DatabaseName "HRProfile"
:setvar DefaultFilePrefix "HRProfile"
:setvar DefaultDataPath ""
:setvar DefaultLogPath ""

GO
:on error exit
GO
/*
Detect SQLCMD mode and disable script execution if SQLCMD mode is not supported.
To re-enable the script after enabling SQLCMD mode, execute the following:
SET NOEXEC OFF; 
*/
:setvar __IsSqlCmdEnabled "True"
GO
IF N'$(__IsSqlCmdEnabled)' NOT LIKE N'True'
    BEGIN
        PRINT N'SQLCMD mode must be enabled to successfully execute this script.';
        SET NOEXEC ON;
    END


GO
/*
The type for column BatchType in table [dbo].[Batch] is currently  VARCHAR (25) NOT NULL but is being changed to  VARCHAR (20) NOT NULL. Data loss could occur and deployment may fail if the column contains data that is incompatible with type  VARCHAR (20) NOT NULL.
*/

IF EXISTS (select top 1 1 from [dbo].[Batch])
    RAISERROR (N'Rows were detected. The schema update is terminating because data loss might occur.', 16, 127) WITH NOWAIT

GO
PRINT N'Dropping Foreign Key [dbo].[FK_Batch_SPUR_BatchID]...';


GO
ALTER TABLE [dbo].[SPUR] DROP CONSTRAINT [FK_Batch_SPUR_BatchID];


GO
PRINT N'Starting rebuilding table [dbo].[Batch]...';


GO
BEGIN TRANSACTION;

SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;

SET XACT_ABORT ON;

CREATE TABLE [dbo].[tmp_ms_xx_Batch] (
    [BatchID]                      INT            IDENTITY (1, 1) NOT NULL,
    [BatchName]                    VARCHAR (50)   NOT NULL,
    [SubmittedBy]                  NVARCHAR (200) NOT NULL,
    [SubmittedTimeStamp]           DATETIME       NOT NULL,
    [BatchType]                    VARCHAR (20)   NOT NULL,
    [BatchStatus]                  VARCHAR (20)   NULL,
    [EndUserCreatedBy]             NVARCHAR (200) NULL,
    [EndUserCreatedTimestamp]      DATETIME       NULL,
    [EndUserModifiedBy]            NVARCHAR (200) NULL,
    [EndUserModifiedTimestamp]     DATETIME       NULL,
    [BackendUserModifiedBy]        NVARCHAR (200) NULL,
    [BackendUserModifiedTimestamp] DATETIME       NULL,
    CONSTRAINT [tmp_ms_xx_constraint_PK_Batch_BatchID1] PRIMARY KEY CLUSTERED ([BatchID] ASC)
);

IF EXISTS (SELECT TOP 1 1 
           FROM   [dbo].[Batch])
    BEGIN
        SET IDENTITY_INSERT [dbo].[tmp_ms_xx_Batch] ON;
        INSERT INTO [dbo].[tmp_ms_xx_Batch] ([BatchID], [BatchName], [SubmittedBy], [SubmittedTimeStamp], [BatchType], [EndUserCreatedBy], [EndUserCreatedTimestamp], [EndUserModifiedBy], [EndUserModifiedTimestamp], [BackendUserModifiedBy], [BackendUserModifiedTimestamp])
        SELECT   [BatchID],
                 [BatchName],
                 [SubmittedBy],
                 [SubmittedTimeStamp],
                 [BatchType],
                 [EndUserCreatedBy],
                 [EndUserCreatedTimestamp],
                 [EndUserModifiedBy],
                 [EndUserModifiedTimestamp],
                 [BackendUserModifiedBy],
                 [BackendUserModifiedTimestamp]
        FROM     [dbo].[Batch]
        ORDER BY [BatchID] ASC;
        SET IDENTITY_INSERT [dbo].[tmp_ms_xx_Batch] OFF;
    END

DROP TABLE [dbo].[Batch];

EXECUTE sp_rename N'[dbo].[tmp_ms_xx_Batch]', N'Batch';

EXECUTE sp_rename N'[dbo].[tmp_ms_xx_constraint_PK_Batch_BatchID1]', N'PK_Batch_BatchID', N'OBJECT';

COMMIT TRANSACTION;

SET TRANSACTION ISOLATION LEVEL READ COMMITTED;


GO
PRINT N'Creating Foreign Key [dbo].[FK_Batch_SPUR_BatchID]...';


GO
ALTER TABLE [dbo].[SPUR] WITH NOCHECK
    ADD CONSTRAINT [FK_Batch_SPUR_BatchID] FOREIGN KEY ([BatchID]) REFERENCES [dbo].[Batch] ([BatchID]);


GO
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
DBCC CHECKIDENT ('[Master].[User]', RESEED, 0);
INSERT INTO [Master].[User] 
(
	EmailID,
	FirstName,
	LastName,
	CreatedBy,
	CreatedTimestamp,
	ModifiedBy,
	ModifiedTimestamp
)
SELECT DISTINCT
	A.EmailID,
	A.FirstName,
	A.LastName,
	'Admin',
	GETUTCDATE(),
	'Admin',
	GETUTCDATE()
FROM #UserRole A
LEFT JOIN [Master].[User] B ON B.EmailID = A.EmailID
WHERE B.UserID IS NULL
ORDER BY A.EmailID

-- Insert roles
DBCC CHECKIDENT ('[Master].[Role]', RESEED, 0);
INSERT INTO [Master].[Role] 
(
	RoleName,
	RoleCode,
	CreatedBy,
	CreatedTimestamp,
	ModifiedBy,
	ModifiedTimestamp
)
SELECT DISTINCT
	A.RoleName,
	A.RoleCode,
	'Admin',
	GETUTCDATE(),
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
	CreatedTimestamp,
	ModifiedBy,
	ModifiedTimestamp
)
SELECT 
	B.UserID,
	C.RoleID,
	'Admin',
	GETUTCDATE(),
	'Admin',
	GETUTCDATE()
FROM #UserRole A
INNER JOIN [Master].[User] B ON B.EmailID = A.EmailID
INNER JOIN [Master].[Role] C ON C.RoleCode = A.RoleCode
LEFT JOIN [Master].[UserRole] D ON D.UserID = B.UserID AND D.RoleID = C.RoleID
WHERE D.UserID IS NULL
ORDER BY B.UserID, C.RoleID
GO

GO
