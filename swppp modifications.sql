/* The following Script will populate the remaining optional file types and then resort them as directed */
INSERT INTO [dbo].[OptionalImagesTypes]([oitSortbyVal],[oitName],[oitDesc]) VALUES (12, 'Report', 'Report')
--INSERT INTO [dbo].[OptionalImagesTypes]([oitSortbyVal],[oitName],[oitDesc]) VALUES (13, 'CSN', 'CSN')
INSERT INTO [dbo].[OptionalImagesTypes]([oitSortbyVal],[oitName],[oitDesc]) VALUES (14, 'MS4Submittal', 'MS4 Submittal')
INSERT INTO [dbo].[OptionalImagesTypes]([oitSortbyVal],[oitName],[oitDesc]) VALUES (15, 'MS4SubmittalReceipt', 'MS4 Submittal Receipt')
INSERT INTO [dbo].[OptionalImagesTypes]([oitSortbyVal],[oitName],[oitDesc]) VALUES (16, 'ReleasesTraining', 'Releases'' Training')
INSERT INTO [dbo].[OptionalImagesTypes]([oitSortbyVal],[oitName],[oitDesc]) VALUES (17, 'StabilizationLog', 'Stabilization Log')
INSERT INTO [dbo].[OptionalImagesTypes]([oitSortbyVal],[oitName],[oitDesc]) VALUES (18, 'ActionsTaken', 'Actions Taken')
INSERT INTO [dbo].[OptionalImagesTypes]([oitSortbyVal],[oitName],[oitDesc]) VALUES (19, 'LocationMap', 'Location Map')
INSERT INTO [dbo].[OptionalImagesTypes]([oitSortbyVal],[oitName],[oitDesc]) VALUES (20, 'RegulatoryReports', 'Regulatory Reports')
INSERT INTO [dbo].[OptionalImagesTypes]([oitSortbyVal],[oitName],[oitDesc]) VALUES (12, 'OperatorCertification', 'Operator Certification')
INSERT INTO [dbo].[OptionalImagesTypes]([oitSortbyVal],[oitName],[oitDesc]) VALUES (9, 'CSNDI', 'CSNDI')
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 1 Where oitName = 'Report'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 2 Where oitName = 'sitemap'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 3 Where oitName = 'SWPPP'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 4 Where oitName = 'NOI'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 5 Where oitName = 'NOC'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 6 Where oitName = 'NOT'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 7 Where oitName = 'Permit'
Update [dbo].[OptionalImagesTypes] Set oitName = 'CSN', oitDesc = 'CSN', oitSortByVal = 8 Where oitID = 5
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 10 Where oitName = 'MS4Submittal'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 11 Where oitName = 'MS4SubmittalReceipt'
Update [dbo].[OptionalImagesTypes] Set oitName = 'DelegationLetter', oitDesc = 'Delegation Letter', oitSortByVal = 12 Where oitID = 1
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 13 Where oitName = 'OperatorCertification'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 14 Where oitName = 'SubcontractorCertification'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 15 Where oitName = 'SoilData'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 16 Where oitName = 'ReleasesTraining'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 17 Where oitName = 'OperatorForm'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 18 Where oitName = 'StabilizationLog'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 19 Where oitName = 'ActionsTaken'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 20 Where oitName = 'LocationMap'
Update [dbo].[OptionalImagesTypes] Set oitSortByVal = 21 Where oitName = 'RegulatoryReports'

/*We will need to update all of the existing inspections to reflect the changes to these defaults */
Alter Table Inspection Alter Column reportType char(30)
Update Inspections set
	reportType = case 
		when reportType = 'Initial' then 'Inital Inspection'
		when reportType = 'Weelky' then 'Weekly Inspection'
		when reportType = 'Bi-Weekly' then 'Bi-Weekly Inspection'
		when reportType = 'Complaint' then 'Complaint Inspection'
--		when reportType = 'Storm' then 'Storm Event Inspection'
		when reportType = 'Monthly' then 'Monthly Inspection'
		else reportType
	end

/*The Following Script will populate the Monthly Inspectio Report Value */
Insert Into [dbo].[ReportTypes] (reportType, priority) Values ('Monthly Inspections', 0)

/*The following Script will modify the ReporTypes table to set the correct orderby*/
Alter Table dbo.ReportTypes Alter Column reportType char(30)
Update dbo.ReportTypes Set ReportType = 'Initiial Inspection', priority = 9 where reportTypeID = 1
Update dbo.ReportTypes Set ReportType = 'Weekly Inspection', priority = 8 where reportTypeID = 2
Update dbo.ReportTypes Set ReportType = 'Bi-Weekly Inspection', priority = 7 where reportTypeID = 3
Update dbo.ReportTypes Set priority = -1 where reportTypeID = 4
Update dbo.ReportTypes Set priority = 5 where reportTypeID = 5
Update dbo.ReportTypes Set ReportType = 'Complaint Inspection', priority = 4 where reportTypeID = 6
Update dbo.ReportTypes Set ReportType = 'Storm Event Inspection', priority = 2 where reportTypeID = 7
Update dbo.ReportTypes Set ReportType = 'Monthly Inspection', priority = 6 Where reportTypeID = 9

/*The Following script will modify the inspec cost defaults for new projects */
Alter Table dbo.Projects Alter Column initInspecCost Set Default(80)
Alter Table dbo.Projects Alter Column inspecCost Set Default (80)
USE [SWPPP]
GO
/****** Object:  Table [dbo].[OPForms]    Script Date: 02/19/2009 20:14:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[OPForms](
	[OPFormID] [int] IDENTITY(1,1) NOT NULL,
	[projectID] [int] NOT NULL,
	[orig_UserID] [int] NOT NULL,
	[edit_userID] [int] NOT NULL,
	[createDate] [smalldatetime] NOT NULL,
	[editDate] [smalldatetime] NOT NULL,
	[OPFormSection] [varchar](30) NULL,
	[OPFormText] [varchar](100) NOT NULL,
	[SectionSortby] [smallint] NOT NULL,
	[SectionSequence] [smallint] NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
/*this script will create the Add Edit Delete stored procedure*/
CREATE PROCEDURE dbo.spAEDCoordinate 
	@_iCOID int, 
	@_iDelFlag smallint = 0, 
	@_inspecID int = 0, 
	@_iCoordinates char(150) = '', 
	@_icorrectiveMods char(255) = '', 
	@_iOrderBy int = -1
AS
BEGIN
	--validate all of the inputs. If any of them are not submitted then abort
	If (@_inspecID = 0) OR (@_iCoordinates = '') OR (@_icorrectiveMods = '') OR (@_iOrderBy = -1)
	Begin
		Return
	End
	--If @_iCOID is 0 then we need to add the record
	If @_iCOID = 0
	Begin
--		Set @_iCoordinates = Replace(@_iCoordinates, '--', '–')
--		Set @_icorrectiveMods = Replace(@_icorrectiveMods, '--', '–')
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
GO

--KenB: 20121105 : the following script was used to change the project names for 2 projects

--begin tran
--	update projects set ProjectName = 'Woodland Enclave' where ProjectID = 955
--	update projects set ProjectName = 'Santa Fe Trails'where ProjectID = 900
--rollback tran
--commit tran