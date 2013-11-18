USE [master]
GO
/****** Object:  Database [AUMPWLetterStoreV2AGPL]    Script Date: 10/25/2013 15:18:15 ******/
CREATE DATABASE [AUMPWLetterStoreV2AGPL] ON  PRIMARY 
( NAME = N'AUMPWLetterStoreAGPL', FILENAME = N'F:\MSSQL\DB\AUMPWLetterStoreAGPL.mdf' , SIZE = 10240KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'AUMPWLetterStoreAGPL_log', FILENAME = N'G:\MSSQL\TR\AUMPWLetterStoreAGPL.ldf' , SIZE = 1964480KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET COMPATIBILITY_LEVEL = 100
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [AUMPWLetterStoreV2AGPL].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET ANSI_NULL_DEFAULT OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET ANSI_NULLS OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET ANSI_PADDING OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET ANSI_WARNINGS OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET ARITHABORT OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET AUTO_CLOSE OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET AUTO_CREATE_STATISTICS ON
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET AUTO_SHRINK OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET AUTO_UPDATE_STATISTICS ON
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET CURSOR_CLOSE_ON_COMMIT OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET CURSOR_DEFAULT  GLOBAL
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET CONCAT_NULL_YIELDS_NULL OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET NUMERIC_ROUNDABORT OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET QUOTED_IDENTIFIER OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET RECURSIVE_TRIGGERS OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET  DISABLE_BROKER
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET AUTO_UPDATE_STATISTICS_ASYNC OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET DATE_CORRELATION_OPTIMIZATION OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET TRUSTWORTHY ON
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET ALLOW_SNAPSHOT_ISOLATION OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET PARAMETERIZATION SIMPLE
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET READ_COMMITTED_SNAPSHOT OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET HONOR_BROKER_PRIORITY OFF
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET  READ_WRITE
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET RECOVERY FULL
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET  MULTI_USER
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET PAGE_VERIFY CHECKSUM
GO
ALTER DATABASE [AUMPWLetterStoreV2AGPL] SET DB_CHAINING OFF
GO
EXEC sys.sp_db_vardecimal_storage_format N'AUMPWLetterStoreV2AGPL', N'ON'
GO
USE [AUMPWLetterStoreV2AGPL]
GO
/****** Object:  User [AGPLTestUser]    Script Date: 10/25/2013 15:18:15 ******/
CREATE USER [AGPLTestUser] FOR LOGIN [AGPLTestUser] WITH DEFAULT_SCHEMA=[dbo]
GO
/****** Object:  UserDefinedTableType [dbo].[typ_DepDataRights]    Script Date: 10/25/2013 15:18:15 ******/
CREATE TYPE [dbo].[typ_DepDataRights] AS TABLE(
	[Department] [char](1) NULL,
	[DataRight] [varchar](10) NULL
)
GO
/****** Object:  Table [dbo].[tPWHistory]    Script Date: 10/25/2013 15:18:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tPWHistory](
	[TimeStamp] [datetime] NOT NULL,
	[UniqueID] [int] NOT NULL,
	[KurzZeichen] [varchar](8) NOT NULL,
	[Departement] [char](1) NOT NULL,
	[DepDescription] [varchar](100) NOT NULL,
	[PasswordEncrypted] [varbinary](256) NOT NULL,
 CONSTRAINT [PK_tPWHistory] PRIMARY KEY CLUSTERED 
(
	[TimeStamp] ASC,
	[UniqueID] ASC,
	[KurzZeichen] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tPWAllStudInClassesLocal]    Script Date: 10/25/2013 15:18:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tPWAllStudInClassesLocal](
	[UniqueID] [int] NOT NULL,
	[KurzZeichen] [varchar](8) NOT NULL,
	[Departement] [char](1) NOT NULL,
	[DepDescription] [varchar](100) NOT NULL,
	[AnlassEvento] [varchar](100) NOT NULL,
 CONSTRAINT [PK_tPWAllStudInClassesLocal] PRIMARY KEY CLUSTERED 
(
	[UniqueID] ASC,
	[KurzZeichen] ASC,
	[DepDescription] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tPWAllADAccounts]    Script Date: 10/25/2013 15:18:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tPWAllADAccounts](
	[WhenCreated] [datetime] NOT NULL,
	[PwdLastSet] [datetime] NULL,
	[LastLogon] [datetime] NULL,
	[UniqueID] [int] NOT NULL,
	[KurzZeichen] [varchar](8) NOT NULL,
	[PersKat] [varchar](20) NOT NULL,
	[Nachname] [varchar](100) NULL,
	[Vorname] [varchar](100) NULL,
	[Geschlecht] [char](1) NOT NULL,
	[EMail1] [varchar](50) NULL,
	[EMail2] [varchar](50) NULL,
	[Departement] [char](1) NOT NULL,
	[DepDescription] [varchar](100) NULL,
 CONSTRAINT [PK_tPWAllADAccounts] PRIMARY KEY CLUSTERED 
(
	[WhenCreated] ASC,
	[UniqueID] ASC,
	[KurzZeichen] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  UserDefinedFunction [dbo].[uf_PWAllStudInClasses]    Script Date: 10/25/2013 15:18:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[uf_PWAllStudInClasses] 
()
RETURNS TABLE 
AS
RETURN 
(
	SELECT * FROM OPENQUERY([SRV-DB-V-004],
	'
		SELECT		EventoIDPerson AS UniqueID, 
					KurzZStudent AS KurzZeichen, 
					CASE WHEN Departement IS NULL THEN LEFT(AnlassNr, 1) ELSE DEPARTEMENT END AS Departement, 
					AnlassNr AS DepDescription,
					AnlassNrFull AS AnlassEvento
		FROM        [SRV-DB-V-004].InterfaceConnect.dbo.[vITS-S2-ZHAWStudentAllAccount]
		WHERE		(AnlassNr IS NOT NULL)
	')
)
GO
/****** Object:  View [dbo].[tPWAllStudentInClasses]    Script Date: 10/25/2013 15:18:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[tPWAllStudentInClasses]
AS
SELECT		EventoIDPerson, 
			KurzZStudent, CASE WHEN Departement IS NULL THEN LEFT(AnlassNr, 1) ELSE DEPARTEMENT END AS Departement, 
			AnlassNr
FROM        [SRV-DB-V-004].InterfaceConnect.dbo.[vITS-S2-ZHAWStudentAllAccount]
WHERE		(AnlassNr IS NOT NULL)
GO
/****** Object:  UserDefinedFunction [dbo].[uf_SplitCharListToTable]    Script Date: 10/25/2013 15:18:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[uf_SplitCharListToTable]
                 (@list      nvarchar(MAX),
                  @delimiter nchar(1) = N',')
      RETURNS @tbl TABLE (listpos int IDENTITY(1, 1) NOT NULL,
                          str     varchar(4000)      NOT NULL,
                          nstr    nvarchar(2000)     NOT NULL) AS

BEGIN
   DECLARE @endpos   int,
           @startpos int,
           @textpos  int,
           @chunklen smallint,
           @tmpstr   nvarchar(4000),
           @leftover nvarchar(4000),
           @tmpval   nvarchar(4000)

   SET @textpos = 1
   SET @leftover = ''
   WHILE @textpos <= datalength(@list) / 2
   BEGIN
      SET @chunklen = 4000 - datalength(@leftover) / 2
      SET @tmpstr = @leftover + substring(@list, @textpos, @chunklen)
      SET @textpos = @textpos + @chunklen

      SET @startpos = 0
      SET @endpos = charindex(@delimiter COLLATE Slovenian_BIN2, @tmpstr)

      WHILE @endpos > 0
      BEGIN
         SET @tmpval = ltrim(rtrim(substring(@tmpstr, @startpos + 1,
                                             @endpos - @startpos - 1)))
         INSERT @tbl (str, nstr) VALUES(@tmpval, @tmpval)
         SET @startpos = @endpos
         SET @endpos = charindex(@delimiter COLLATE Slovenian_BIN2,
                                 @tmpstr, @startpos + 1)
      END

      SET @leftover = right(@tmpstr, datalength(@tmpstr) / 2 - @startpos)
   END

   INSERT @tbl(str, nstr)
      VALUES (ltrim(rtrim(@leftover)), ltrim(rtrim(@leftover)))
   RETURN
END
GO
/****** Object:  UserDefinedFunction [dbo].[uf_PWHistoryNewest]    Script Date: 10/25/2013 15:18:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[uf_PWHistoryNewest] 
()
RETURNS TABLE 
AS
RETURN 
(
	SELECT * FROM dbo.tPWHistory tOut WHERE 
			TimeStamp IN (	SELECT MAX(TimeStamp) 
							FROM tPWHistory tIn 
							WHERE	tOut.KurzZeichen = tIn.KurzZeichen
						  )
)
GO
/****** Object:  StoredProcedure [dbo].[usp_PWStoreTruncateData]    Script Date: 10/25/2013 15:18:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[usp_PWStoreTruncateData]
WITH EXECUTE AS OWNER
AS
BEGIN
	TRUNCATE TABLE dbo.tPWAllADAccounts;
	TRUNCATE TABLE dbo.tPWAllStudInClassesLocal;
	INSERT dbo.tPWAllStudInClassesLocal SELECT * FROM dbo.uf_PWAllStudInClasses()
	
END
GO
/****** Object:  UserDefinedFunction [dbo].[uf_ListPWDecrypted]    Script Date: 10/25/2013 15:18:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[uf_ListPWDecrypted] 
(	
	@passphrase varchar(50)
)
RETURNS TABLE 
AS
RETURN 
(
	SELECT
			[TimeStamp] = CASE WHEN (pwd.TimeStamp IS NOT NULL) THEN pwd.TimeStamp ELSE acc.WhenCreated END,
			acc.UniqueID,
			acc.KurzZeichen,
			acc.PersKat,
			acc.Nachname,
			acc.Vorname,
			acc.Geschlecht,
			acc.Email1,
			acc.Email2,
			acc.Departement,
			acc.DepDescription,
			DepDescriptionSub = cls.DepDescription,
			AnlassEvento,
			IsMain = CASE WHEN (RIGHT(acc.DepDescription, LEN(cls.DepDescription)) = cls.DepDescription) OR (cls.DepDescription IS NULL) THEN 1 ELSE 0 END,
			PWIsValid = CASE WHEN (DATEDIFF(hour,pwd.[TimeStamp],acc.PwdLastSet) < 4) AND (PasswordEncrypted IS NOT NULL) THEN 1 ELSE 0 END,
			DATEDIFFT = DATEDIFF(hour,pwd.[TimeStamp],acc.PwdLastSet),
			AccountType = CASE acc.PersKat
					WHEN '9810' THEN 'LE'
					WHEN '9830' THEN 'LE'
					WHEN '9930' THEN 'STAFF'
					WHEN '9820' THEN 'LE'
					WHEN '9920' THEN 'STAFF'						
					WHEN '#weiter#' THEN 'WB'
					ELSE 'STAFF'
				  END,
			pwd.PasswordEncrypted,
			PasswordDecrypted = CONVERT(varchar(50),DECRYPTBYPASSPHRASE(@passphrase,[PasswordEncrypted])),
			acc.WhenCreated,
			acc.PwdLastSet,
			acc.LastLogon			
	FROM	dbo.tPWAllADAccounts acc
			left outer join dbo.tPWAllStudInClassesLocal cls on acc.UniqueID=cls.UniqueID AND acc.KurzZeichen=cls.KurzZeichen
			left outer join dbo.uf_PWHistoryNewest() pwd on acc.UniqueID=pwd.UniqueID AND acc.KurzZeichen=pwd.KurzZeichen					  
	WHERE   (acc.Departement <> '') AND (acc.DepDescription <> '')
)
GO
/****** Object:  StoredProcedure [dbo].[usp_PWStoreGetUserData]    Script Date: 10/25/2013 15:18:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[usp_PWStoreGetUserData] 
(	
	@User varchar(50),
	@Key varchar(50)
)
AS
BEGIN
	SET NOCOUNT ON;
	
	SELECT * FROM dbo.uf_ListPWDecrypted(@Key) WHERE KurzZeichen = @User
	
END
GO
/****** Object:  StoredProcedure [dbo].[usp_PWStoreGetData]    Script Date: 10/25/2013 15:18:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[usp_PWStoreGetData] 
(	
	@startDate varchar(50)=null,
	@endDate varchar(50)=null,
	@Rights dbo.typ_DepDataRights READONLY,
	@Department varchar(50)=null,
	@Type varchar(50)=null,
	@Key varchar(50)=null,
	@Filter varchar(50)=null
)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @result int
	DECLARE @startDateLocal datetime
	DECLARE @endDateLocal datetime
	
	-- Check the date and make sure we can handle both formats 04.11.2013 and 11/04/2013
	IF (@startDate IS NOT NULL) BEGIN
		IF (PATINDEX('%.%', @startdate) = 0) BEGIN
			SET @startDateLocal = CONVERT(datetime, @startDate)
		END	ELSE BEGIN
			SET @startDateLocal = CONVERT(datetime, @startDate, 104)
		END
	END
	
	IF (@endDate IS NULL) BEGIN
		SET @endDateLocal = getdate()
	END ELSE BEGIN
		IF (PATINDEX('%.%', @endDate) = 0) BEGIN
			SET @endDateLocal = CONVERT(datetime, @endDate)
		END ELSE BEGIN
			SET @endDateLocal = CONVERT(datetime, @endDate, 104)
		END
	END

	IF (@startDateLocal = @endDateLocal) BEGIN
		SET @endDateLocal = DATEADD(dd,1,@startDateLocal)
	END

	SET @Filter = '%'+@Filter+'%'
		
	SELECT * 
	FROM dbo.uf_ListPWDecrypted(@Key) tOut WHERE
	    -- Check time span
		((TimeStamp >= @startDateLocal) OR (@startDate IS NULL)) AND 
		((TimeStamp <= @endDateLocal) OR (@endDate IS NULL)) AND
		-- Check rights
		(EXISTS(SELECT * FROM @Rights r WHERE tOut.Departement like r.Department AND tOut.AccountType like r.DataRight) OR ((SELECT Count(*) FROM @Rights) = 0)) AND		
		-- Check departement. @Department can be a comma delimited list
		((@Department like '%' + Departement + '%') OR (@Department IS NULL)) AND
		-- Check typ LE, WB or STAFF. @Type is comma delimited list
		((@Type like '%' + AccountType + '%') OR (@Type IS NULL)) AND
		-- Implement filter, currently not used
		((KurzZeichen like @Filter) OR (Nachname like @Filter) OR (Vorname like @Filter) OR (Departement + ' ' + AccountType + '-' + DepDescriptionSub like @Filter) OR (AnlassEvento like @Filter) OR (@Filter IS NULL))
	ORDER BY LEFT(DepDescription, 1), AccountType DESC, DepDescription, Nachname, Vorname
END
GO
/****** Object:  Default [DF_tPWAllStudInClassesLocal_AnlassEvento]    Script Date: 10/25/2013 15:18:15 ******/
ALTER TABLE [dbo].[tPWAllStudInClassesLocal] ADD  CONSTRAINT [DF_tPWAllStudInClassesLocal_AnlassEvento]  DEFAULT ('') FOR [AnlassEvento]
GO
USE [master]
GO
GRANT CONNECT TO [AGPLTestUser] AS [dbo]
GO
USE [AUMPWLetterStoreV2AGPL]
GO
GRANT CONTROL ON TYPE::[dbo].[typ_DepDataRights] TO [public] AS [dbo]
GO
GRANT REFERENCES ON TYPE::[dbo].[typ_DepDataRights] TO [public] AS [dbo]
GO
GRANT EXECUTE ON [dbo].[usp_PWStoreGetUserData] TO [public] AS [dbo]
GO
GRANT EXECUTE ON [dbo].[usp_PWStoreGetData] TO [public] AS [dbo]
GO
