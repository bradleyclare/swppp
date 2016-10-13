SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_UpdateCommissions]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'


CREATE       proc [dbo].[sp_UpdateCommissions]
	@projectID int,
      @userID  int,
	@phase1 money,
	@phase2 money,
	@phase3 money,
	@phase4 money,
	@phase5 money
as
/* check to see if this record already exists */
SELECT * FROM Commissions
WHERE userID=@userID AND projectID=@projectID
IF @@ROWCOUNT=0 
/* if it doesn''t exist, create it */
	BEGIN
	INSERT INTO Commissions (projectID,userID, phase1, phase2, phase3, phase4, phase5) 
	VALUES (@projectID, @userID, @phase1, @phase2, @phase3, @phase4, @phase5) 
	END
ELSE
/* we may need to update it */
	BEGIN
	UPDATE Commissions SET 
	phase1=@phase1, phase2=@phase2, phase3=@phase3, phase4=@phase4, phase5=@phase5
	WHERE projectID=@projectID AND userID=@userID 
	END
/* otherwise, the update and the existing record match... don''t update */

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_qb1_old]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE        proc [dbo].[sp_qb1_old]
		@firstDay smalldateTime,
		@lastDay smalldateTime,
		@billCycle tinyint
as
SELECT 	i.inspecID, i.inspecDate, i.projectID, i.compName, i.compAddr, 
	i.compAddr2, i.compCity, i.compState, i.compZip, i.projectCity, i.projectState,
	p.projectName, p.projectPhase, p.inspecCost
FROM 	
	Inspections as i 
	LEFT JOIN Projects as p 
		ON i.projectID = p.projectID
WHERE 	
	i.inspecDate BETWEEN @firstDay  AND @lastDay 
	AND p.billCycle=@billCycle
ORDER BY 
	p.projectName ASC, 
	p.projectPhase ASC, 
	inspecDate DESC

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_GetInspectionData]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[sp_GetInspectionData](@inspecID as int)
as
SELECT * FROM Inspections WHERE inspecID=@inspecID

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_getAllReportsforProject]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[sp_getAllReportsforProject](@projectID as int)
as
SELECT DISTINCT inspecID, inspecDate, p.projectName
FROM Projects as p, Inspections as i
WHERE i.projectID=p.projectID
AND i.projectID=@projectID
ORDER BY inspecDate DESC

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_getAllProjects]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[sp_getAllProjects]
as
SELECT * FROM Projects Order by projectName ASC

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_GetAllInspectionsforProject]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'



CREATE  procedure [dbo].[sp_GetAllInspectionsforProject](@projectID as int)
as

SELECT DISTINCT inspecID, inspecDate, p.projectName
FROM Projects as p, Inspections as i
WHERE i.projectID=p.projectID
AND i.projectID=@projectID
ORDER BY inspecDate desc




' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_arcTempTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[sp_arcTempTable]
as
CREATE TABLE ##tempArchive (
	arcID int PRIMARY KEY CLUSTERED,
	arcFileName char(60) NOT NULL,
	arcSrcPath char(250) NOT NULL,
	arcSrcDest char(250) NOT NULL
)

' 
END
GO
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fnGetcPID]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
BEGIN
execute dbo.sp_executesql @statement = N'

CREATE  function [dbo].[fnGetcPID] (@cPID int)
RETURNS @listofProjectIDs table (
	projectID int )
AS
BEGIN
IF @cPID=-1
Begin 
	INSERT @listofProjectIDs
	SELECT projectID FROM Projects 
End
ELSE
Begin
	INSERT @listofProjectIDs
	SELECT projectID FROM Projects WHERE projectID=@cPID
End
return 
END


' 
END

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fnGetFullName]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
BEGIN
execute dbo.sp_executesql @statement = N'

CREATE  function [dbo].[fnGetFullName] (@userID int)
RETURNS char(100)
AS
BEGIN
DECLARE @fullName char(100)
SET @fullName= (SELECT LTRIM(RTRIM(lastName)) +'', ''+ LTRIM(RTRIM(firstName)) FROM users WHERE userID=@userID)
return (@fullName)
END


' 
END

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_getAllProjectsActions]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'



CREATE   procedure [dbo].[sp_getAllProjectsActions] @userID int, @cPID int, @highestRights char(5)
as

IF @highestRights=''admin''
BEGIN
	SELECT 	DISTINCT p.projectID,
		IsNull(LTRIM(RTRIM(p.projectName)),'''') as projectName,  
		IsNull(LTRIM(RTRIM(p.projectPhase)),'''') as projectPhase,
		a.orig_actionDate, a.actionText, u.firstName, u.lastName
	FROM 	Projects p inner join ProjectsUsers pu on p.projectID=pu.projectID 
		inner join Actions a on p.projectID=a.projectID
		inner join Users u on a.orig_userID=u.userID
	WHERE   p.projectID IN (SELECT * FROM dbo.fnGetcPID(@cPID))
	ORDER BY 
		p.projectName asc, p.projectPhase asc, a.orig_actionDate desc
END
IF @highestRights=''dir''
BEGIN
	SELECT 	DISTINCT p.projectID,
		IsNull(LTRIM(RTRIM(p.projectName)),'''') as projectName,  
		IsNull(LTRIM(RTRIM(p.projectPhase)),'''') as projectPhase,
		a.orig_actionDate, a.actionText, u.firstName, u.lastName
	FROM 	Projects p inner join ProjectsUsers pu on p.projectID=pu.projectID 
		inner join Actions a on p.projectID=a.projectID
		inner join Users u on a.orig_userID=u.userID
	WHERE 	pu.userID=@userID AND pu.rights IN (''director'', ''action'', ''user'')
		AND p.projectID IN (SELECT * FROM dbo.fnGetcPID(@cPID))
	ORDER BY 
		p.projectName asc, p.projectPhase asc, a.orig_actionDate desc
END
IF @highestRights=''ins''
BEGIN
	SELECT 	DISTINCT p.projectID,
		IsNull(LTRIM(RTRIM(p.projectName)),'''') as projectName,  
		IsNull(LTRIM(RTRIM(p.projectPhase)),'''') as projectPhase,
		a.orig_actionDate, a.actionText, u.firstName, u.lastName
	FROM 	Projects p inner join ProjectsUsers pu on p.projectID=pu.projectID 
		inner join Actions a on p.projectID=a.projectID
		inner join Users u on a.orig_userID=u.userID
	WHERE 	pu.userID=@userID AND pu.rights=''inspector''
		AND p.projectID IN (SELECT * FROM dbo.fnGetcPID(@cPID))
	ORDER BY 
		p.projectName asc, p.projectPhase asc, a.orig_actionDate desc
END
IF @highestRights=''user''
BEGIN
	SELECT 	DISTINCT p.projectID,
		IsNull(LTRIM(RTRIM(p.projectName)),'''') as projectName,  
		IsNull(LTRIM(RTRIM(p.projectPhase)),'''') as projectPhase,
		a.orig_actionDate, a.actionText, u.firstName, u.lastName
	FROM 	Projects p inner join ProjectsUsers pu on p.projectID=pu.projectID 
		inner join Actions a on p.projectID=a.projectID
		inner join Users u on a.orig_userID=u.userID
	WHERE 	pu.userID=@userID AND pu.rights IN (''user'',''action'')
		AND p.projectID IN (SELECT * FROM dbo.fnGetcPID(@cPID))
	ORDER BY 
		p.projectName asc, p.projectPhase asc, a.orig_actionDate desc		
END



' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spAEDCoordinate]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'


/*this script will create the Add Edit Delete stored procedure*/
CREATE PROCEDURE [dbo].[spAEDCoordinate] 
	@_iCOID int, 
	@_iDelFlag smallint = 0, 
	@_inspecID int = 0, 
	@_iCoordinates char(150) = '''', 
	@_icorrectiveMods char(1000) = '''', 
	@_iOrderBy int = 0
AS
BEGIN
	--validate all of the inputs. If any of them are not submitted then abort
	If (@_inspecID = 0) OR (@_iCoordinates = '''') OR (@_icorrectiveMods = '''')
	Begin
		Return
	End
	--If @_iCOID is 0 then we need to add the record
	If @_iCOID = 0
	Begin
--		Set @_iCoordinates = Replace(@_iCoordinates, ''--'', ''–'')
--		Set @_icorrectiveMods = Replace(@_icorrectiveMods, ''--'', ''–'')
		Insert Into Coordinates (inspecID, coordinates, correctiveMods, orderby)
		Values (@_inspecID, @_iCoordinates, @_icorrectiveMods, @_iOrderBy)
	End
	Else Begin
		-- Check the @_iDelFlag to see if we need to delete this record
		If @_iDelFlag = 1
		Begin
			Delete From Coordinates Where coID = @_iCOID
		End
		Else Begin
			Update Coordinates Set
				coordinates = @_iCoordinates,
				correctiveMods = @_icorrectiveMods,
				orderby = @_iOrderBy
			Where coID = @_iCOID
		End
	End
END



' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_getProjectsPhases]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'

create procedure [dbo].[sp_getProjectsPhases] @userID int, @highestRights char(5)
as

IF @highestRights=''admin''
BEGIN
	SELECT 	DISTINCT projectID, 
		IsNull(LTRIM(RTRIM(projectName)),'''') as projectName,  
		IsNull(LTRIM(RTRIM(projectPhase)),'''') as projectPhase
	FROM Projects 
	Order by projectName, projectPhase ASC
END
IF @highestRights=''dir''
BEGIN
	SELECT 	DISTINCT p.projectID, 
		IsNull(LTRIM(RTRIM(p.projectName)),'''') as projectName,  
		IsNull(LTRIM(RTRIM(p.projectPhase)),'''') as projectPhase
	FROM Projects p, ProjectsUsers pu
	WHERE p.projectID=pu.projectID
		AND pu.userID=@userID
		AND pu.rights IN (''director'', ''action'', ''user'')
	Order by projectName, projectPhase ASC
END
IF @highestRights=''ins''
BEGIN
	SELECT 	DISTINCT p.projectID, 
		IsNull(LTRIM(RTRIM(p.projectName)),'''') as projectName,  
		IsNull(LTRIM(RTRIM(p.projectPhase)),'''') as projectPhase
	FROM Projects p, ProjectsUsers pu
	WHERE p.projectID=pu.projectID
		AND pu.userID=@userID
		AND pu.rights=''inspector''
	Order by projectName, projectPhase ASC
END
IF @highestRights=''user''
BEGIN
	SELECT 	DISTINCT p.projectID, 
		IsNull(LTRIM(RTRIM(p.projectName)),'''') as projectName,  
		IsNull(LTRIM(RTRIM(p.projectPhase)),'''') as projectPhase
	FROM Projects p, ProjectsUsers pu
	WHERE p.projectID=pu.projectID
		AND pu.userID=@userID
		AND pu.rights IN (''user'',''action'')
	Order by projectName, projectPhase ASC
END

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_UpdateInsertPUEmail]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE   proc [dbo].[sp_UpdateInsertPUEmail]
        @userID  int,
	@projectID int,
	@rights char(20),
	@email bit
as
/* check to see if this record already exists */
SELECT * FROM ProjectsUsers 
WHERE userID=@userID 
AND projectID=@projectID 
AND rights=@rights
IF @@ROWCOUNT=0 
/* if it doesn''t exist, create it */
	BEGIN
	INSERT INTO ProjectsUsers (userID, projectID, rights, emailReport) 
	VALUES (@userID, @projectID, @rights, 1) 
	END
ELSE
/* we just need to update it */
	BEGIN
	UPDATE ProjectsUsers SET emailReport=@email
	WHERE userID=@userID 
	AND projectID=@projectID
	AND rights=@rights
	END

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_InsertPU]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'

CREATE     proc [dbo].[sp_InsertPU]
        @userID  int,
	@projectID int,
	@rights char(20)
as
/* check to see if this record already exists */
SELECT * FROM ProjectsUsers 
WHERE userID=@userID 
AND projectID=@projectID 
AND rights=@rights
IF @@ROWCOUNT=0 
/* if it doesn''t exist, create it */
BEGIN
	INSERT INTO ProjectsUsers (userID, projectID, rights) 
	VALUES (@userID, @projectID, @rights) 
END


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_GetOptionalImages]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'



CREATE   procedure [dbo].[sp_GetOptionalImages](@inspecID as int)
as
SELECT oi.*, oit.oitDesc, oit.oitName 
FROM OptionalImages oi INNER JOIN OptionalImagesTypes oit on
	oi.oitID=oit.oitID	
WHERE inspecID=@inspecID ORDER BY oit.oitSortByVal asc, oi.oOrder asc



' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_oImagesByType]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'


CREATE    proc [dbo].[sp_oImagesByType]
        @inspecID  int,
	@oitID int
as
	SELECT oi.*, oit.oitName, oit.oitDesc 
	FROM OptionalImages oi
	INNER JOIN OptionalImagesTypes oit
	ON oi.oitID=oit.oitID
	WHERE inspecID=@inspecID
		AND oi.oitID=@oitID
	ORDER BY oOrder, oImageFileName



' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_oImagesByInspecID]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create proc [dbo].[sp_oImagesByInspecID]
        @inspecID  int
as
	SELECT * FROM OptionalImages WHERE inspecID=@inspecID

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_AddOptImage]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'


CREATE    proc [dbo].[sp_AddOptImage]
	@oImageName char(50),
	@oIamgeDesc char(50),
        @inspecID  int,
	@oitID int,
	@oImageFileName char(50),
	@oOrder smallint
as
	SELECT * FROM OptionalImages
	WHERE inspecID=@inspecID 
		AND oitID=@oitID
		AND oImageFileName=@oImageFileName
	IF @@ROWCOUNT=0
		INSERT INTO OptionalImages
			(oImageName,oImageDesc,oImageFileName,oitID,inspecID,oOrder)
		VALUES 
			(@oImageName,@oIamgeDesc,@oImageFileName,@oitID,@inspecID,@oOrder)



' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_qb1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'

CREATE           proc [dbo].[sp_qb1]
		@firstDay smalldateTime,
		@lastDay smalldateTime,
		@billCycle tinyint
as
SELECT 	i.compName, i.compAddr, i.compAddr2, i.compCity, 
	i.compState, i.compZip, i.projectCity, i.projectState,
	v.projectName, v.inspecCost, v.inspecDate, v.reportType, v.reportType, v.invoiceMemo
FROM 	vQB1 as v LEFT JOIN
	Inspections as i 
		ON (v.projectName=i.projectName AND v.inspecDate=i.inspecDate)
WHERE 	
	v.inspecDate BETWEEN @firstDay AND @lastDay
	AND v.billCycle = @billCycle
--	AND i.reportType=''weekly''
GROUP BY i.compName, i.compAddr, i.compAddr2, i.compCity, 
	i.compState, i.compZip, i.projectCity, i.projectState,
	v.projectName, v.inspecCost, v.inspecDate, v.reportType, v.invoiceMemo
ORDER BY 
	v.projectName ASC, v.inspecDate ASC


' 
END
