/*
USE [master];
EXEC sp_configure 'clr enabled', 1;
RECONFIGURE;

USE [SPT.MSGraphCLR];
ALTER DATABASE [SPT.MSGraphCLR] SET TRUSTWORTHY ON;
*/


/*
DECLARE @hash VARBINARY(64);

SELECT @hash = HASHBYTES('SHA2_256', BulkColumn) + CAST(REPLICATE(0x00, 32) AS VARBINARY(32))
FROM OPENROWSET(BULK 'C:\MSGraphClr\SPT.MSGraphCLR.dll', SINGLE_BLOB) AS x;

EXEC sp_add_trusted_assembly 
    @hash = @hash, 
    @description = N'SPT.MSGraphCLR';

SELECT @hash = HASHBYTES('SHA2_256', BulkColumn) + CAST(REPLICATE(0x00, 32) AS VARBINARY(32))
FROM OPENROWSET(BULK 'C:\MSGraphClr\Microsoft.IdentityModel.Abstractions.dll', SINGLE_BLOB) AS x;

EXEC sp_add_trusted_assembly 
    @hash = @hash, 
    @description = N'Microsoft.IdentityModel.Abstractions';

SELECT @hash = HASHBYTES('SHA2_256', BulkColumn) + CAST(REPLICATE(0x00, 32) AS VARBINARY(32))
FROM OPENROWSET(BULK 'C:\MSGraphClr\Microsoft.Identity.Client.dll', SINGLE_BLOB) AS x;

EXEC sp_add_trusted_assembly 
    @hash = @hash, 
    @description = N'Microsoft.Identity.Client';

-- Get 32-byte SHA256 hash, then pad with 32 zero bytes to make 64 bytes total
SELECT @hash = HASHBYTES('SHA2_256', BulkColumn) + CAST(REPLICATE(0x00, 32) AS VARBINARY(32))
FROM OPENROWSET(BULK 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Runtime.Serialization.dll', SINGLE_BLOB) AS x;

EXEC sp_add_trusted_assembly 
    @hash = @hash, 
    @description = N'System.Runtime.Serialization';


SELECT @hash = HASHBYTES('SHA2_256', BulkColumn) + CAST(REPLICATE(0x00, 32) AS VARBINARY(32))
FROM OPENROWSET(BULK 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Net.Http.dll', SINGLE_BLOB) AS x;

EXEC sp_add_trusted_assembly 
    @hash = @hash, 
    @description = N'System.Net.Http';


SELECT @hash = HASHBYTES('SHA2_256', BulkColumn) + CAST(REPLICATE(0x00, 32) AS VARBINARY(32))
FROM OPENROWSET(BULK 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Windows.Forms.dll', SINGLE_BLOB) AS x;

EXEC sp_add_trusted_assembly 
    @hash = @hash, 
    @description = N'System.Windows.Forms';

SELECT @hash = HASHBYTES('SHA2_256', BulkColumn) + CAST(REPLICATE(0x00, 32) AS VARBINARY(32))
FROM OPENROWSET(BULK 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.IdentityModel.dll', SINGLE_BLOB) AS x;

EXEC sp_add_trusted_assembly 
    @hash = @hash, 
    @description = N'System.IdentityModel';
    */

/*


CREATE ASSEMBLY [System.Runtime.Serialization]
FROM 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Runtime.Serialization.dll'
WITH PERMISSION_SET = UNSAFE;

CREATE ASSEMBLY [System.Windows.Forms]
FROM 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Windows.Forms.dll'
WITH PERMISSION_SET = UNSAFE;

CREATE ASSEMBLY [System.Net.Http]
FROM 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Net.Http.dll'
WITH PERMISSION_SET = UNSAFE;

CREATE ASSEMBLY [System.IdentityModel]
FROM 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.IdentityModel.dll'
WITH PERMISSION_SET = UNSAFE;

CREATE ASSEMBLY [Microsoft.IdentityModel.Clients.ActiveDirectory]
FROM 'C:\MSGraphClr\Microsoft.IdentityModel.Clients.ActiveDirectory.dll'
WITH PERMISSION_SET = UNSAFE;

CREATE ASSEMBLY [Newtonsoft.Json]
FROM 'C:\MSGraphClr\Newtonsoft.Json.dll'
WITH PERMISSION_SET = UNSAFE;


CREATE ASSEMBLY [Microsoft.Identity.Client]
FROM 'C:\MSGraphClr\Microsoft.Identity.Client.dll'
WITH PERMISSION_SET = UNSAFE;

CREATE ASSEMBLY [Microsoft.IdentityModel.Abstractions]
FROM 'C:\MSGraphClr\Microsoft.IdentityModel.Abstractions.dll'
WITH PERMISSION_SET = UNSAFE;

--SELECT * FROM sys.assemblies order by create_date desc;
 

DROP PROC IF EXISTS [GEN_MSGraphGetEmailSp];
DROP PROC IF EXISTS [GEN_MSGraphSendEmailSp];
DROP PROC IF EXISTS [GEN_MSGraphSendTeamsMessageSp];
DROP ASSEMBLY IF EXISTS [SPT.MSGraphCLR];

/*
use master;
CREATE MASTER KEY ENCRYPTION BY PASSWORD = 'B98DF984-2E04-4823-B109-F1DE362B5ADE';
CREATE ASYMMETRIC KEY MSGraphCLRKey
FROM FILE = 'C:\MSGraphClr\MSGraphCLRKey.snk';
 
CREATE LOGIN MSGraphCLRKeyLogin
FROM ASYMMETRIC KEY MSGraphCLRKey;

GRANT UNSAFE ASSEMBLY TO MSGraphCLRKeyLogin;
*/


CREATE ASSEMBLY [SPT.MSGraphCLR]
FROM 'C:\MSGraphClr\SPT.MSGraphCLR.dll'
WITH PERMISSION_SET = UNSAFE;


CREATE PROCEDURE GEN_MSGraphGetEmailSp
    @userEmail NVARCHAR(100),
    @subjectFilter nvarchar(3800)
AS EXTERNAL NAME [SPT.MSGraphCLR].[StoredProcedures].GEN_MSGraphGetEmailSp;

CREATE PROCEDURE GEN_MSGraphSendEmailSp
    @fromUser NVARCHAR(3800),
    @toUser NVARCHAR(3800),
    @subject NVARCHAR(3800),
    @body NVARCHAR(3800),
    @tableRowPointer NVARCHAR(3800)
AS EXTERNAL NAME [SPT.MSGraphCLR].[StoredProcedures].GEN_MSGraphSendEmailSp;

CREATE PROCEDURE GEN_MSGraphSendTeamsMessageSp
    @teamId NVARCHAR(100),
    @channelId NVARCHAR(100),
    @messageText NVARCHAR(3800)
AS EXTERNAL NAME [SPT.MSGraphCLR].[StoredProcedures].GEN_MSGraphSendTeamsMessageSp;
 
CREATE PROCEDURE dbo.GEN_MSGraphSendEmailAndReturnInfoSp
    @fromUser NVARCHAR(256),
    @toUser NVARCHAR(MAX),
    @ccUser NVARCHAR(MAX),
    @bccUser NVARCHAR(MAX),
    @subject NVARCHAR(500),
    @body NVARCHAR(MAX),
    @tableRowPointer NVARCHAR(3800),
    @messageId NVARCHAR(500) OUTPUT,
    @conversationId NVARCHAR(100) OUTPUT,
    @internetMessageId NVARCHAR(500) OUTPUT,
    @createdDateTime NVARCHAR(50) OUTPUT,
    @sentDateTime NVARCHAR(50) OUTPUT,
    @toRecipients NVARCHAR(MAX) OUTPUT,
    @ccRecipients NVARCHAR(MAX) OUTPUT,
    @bccRecipients NVARCHAR(MAX) OUTPUT,
    @hasAttachments BIT OUTPUT,
    @isReadReceiptRequested BIT OUTPUT,
    @bodyPreview NVARCHAR(MAX) OUTPUT
AS EXTERNAL NAME [SPT.MSGraphCLR].[StoredProcedures].GEN_MSGraphSendEmailAndReturnInfoSp;

CREATE PROCEDURE dbo.GEN_MSGraphReplyEmailSp
    @fromUser NVARCHAR(256),
    @originalMessageId NVARCHAR(100),
    @replyBody NVARCHAR(MAX),
    @replyAll BIT,
    @replyMessageId NVARCHAR(100) OUTPUT,
    @conversationId NVARCHAR(100) OUTPUT,
    @internetMessageId NVARCHAR(255) OUTPUT,
    @createdDateTime NVARCHAR(50) OUTPUT,
    @sentDateTime NVARCHAR(50) OUTPUT
AS EXTERNAL NAME [SPT.MSGraphCLR].[StoredProcedures].[GEN_MSGraphReplyEmailSp]
*/

 
--EXEC GEN_MSGraphGetEmailSp @userEmail = 'gennovas@sambopiping.co.th', @subjectFilter = 'แจ้งปัญหา'
EXEC GEN_MSGraphSendEmailSp @fromUser = 'gennovas@sambopiping.co.th', @toUser = 'gennovas@sambopiping.co.th;sarit@gennovas.com', @subject = 'Purchase Order Requisition Approvval', @body = 'Purchase Order Requisition Approvval #0000000000', @tableRowPointer='3BB4B1F1-5635-4908-B958-4880C9B0360D'
--EXEC GEN_MSGraphSendTeamsMessageSp @teamId = 'bd4b2933-c709-4695-be4a-69cc4922ccf1', @channelId = '19:KiU4tNsnKFeMFxj3J6uoqe3YZ2kAOsQqwcG73WD52K41@thread.tacv2', @messageText = 'test'


/*
USE master;
GO

-- Step 1: Create asymmetric key from the dependency DLL
CREATE ASYMMETRIC KEY ADAL_Key
FROM EXECUTABLE FILE = 'C:\MSGraphClr\Microsoft.IdentityModel.Clients.ActiveDirectory.dll';
GO

-- Step 2: Create a login from that key
CREATE LOGIN ADAL_Login FROM ASYMMETRIC KEY ADAL_Key;
GO

-- Step 3: Grant UNSAFE ASSEMBLY permission
GRANT UNSAFE ASSEMBLY TO ADAL_Login;
GO
*/

/*
USE master;
GO

-- Step 1: Create asymmetric key from the dependency DLL
CREATE ASYMMETRIC KEY NewtonSoftKey
FROM EXECUTABLE FILE = 'C:\MSGraphClr\Newtonsoft.Json.dll';
GO

-- Step 2: Create a login from that key
CREATE LOGIN NewtonSoft_Login FROM ASYMMETRIC KEY NewtonSoftKey;
GO

-- Step 3: Grant UNSAFE ASSEMBLY permission
GRANT UNSAFE ASSEMBLY TO NewtonSoft_Login;
GO
*/
