Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("STAR_ITWFeed")> 
<Assembly: AssemblyDescription("McKesson STAR_ITWFeed.20130514-added financial table entry routine.20130520-recompiled for production.20130522-added department A01,4,8.20130523-added corpNo extraction.20130524-removed patnum insert.20130525-fix financial insert routine.20130528-added onset date from UB1_16_2.20130529-fixed department.20130531-add physConsult raw data20130603-add panum if orphan encountered.20130606-added datetime orphan added.20130617-use status codes  based on HL7 record processed.20130619-capture class in episode table with translation.2013020-capture raw class from STAR and calcStatus function.20130808 - fixed race description error for A01,4.20140103 - changd ITW connection string. 20140124 - fixed A01 problem with inpatients also added department to A01 and 4 update to A05. 20140202 - mods for wave3 testing on cscsysfeed5 20140205 - replace event log with text log. 20140212 - modified convert date functions. 20140213 -  added new process for hospital service based on department for outpatients. 20140215 - A31 processing for Allergies.20140303 - expanded Allergy prcessing with additional fields. 20140306-added global error processing. 20140310 - added status update to A08. 20140310 - corporate number now in PID_2 - JHHS mrnum. 20140310 - do same as A03 for A11. 20140311 -added A13 code to calSTARStatus.2014038 - calcstarstatus same for A03 and A11. 20140321-removed region from the mrnum for AL1 processing.20140414-A07 process for LOA fixed Main Routine Error Handling.20140423- fixed corpNo handling.20140424-added A08 Pat status Case. '20140528 Added AuditNotes from ZIN_9. 20140603 - code to put and use the star plancode in the iplancode field of 03insurer. 20140811 - Mods for AccDate in 08Accidents Table. 20140817- mods for W3 Production. 20140904 - return to integer fclass.20140908 - removed ZMI. 20140912 - added star_region to A08 update process.20140915 - modified search in processAL1.'20140916 - capture all AL1 data. '20140916 - add update of Admindate on A08.'20140917 - zero out gblAl1Count before counting.  '20140918 - accept more than 9 AL1 segments.20141002 - fixed ARO. 20141021 - process A34 records. 20141022 - add onset processing back in. 20141023 - new onsetdate procedure for multiple occodes. 201410231102 - added process to populate 075coourrence table. 20141117 - fixed redundant A34 processing call. 20141206 - added dictNVP.Clear() if no panum20141209 - updated processA34 code. 20150330 - for A03 - if patient class is O then don't set dcdate.20150413 - VS2013 version. '20150506 - added A12 processing. Same as A02' 20150717 - remove orphan processing.20150909 - put orphan lagging back in and remove admission date processing for A08 records. 20150914 - new criteria to update admission date on an A08.20150929 - added A44 processing to production ITW feed.20150930 - main routine cleanup.20151130 - added update admin date going from IP to IA.20151215 - add InsuredDOB and InsuredSex to 03Insurer table.'20160125 - fixed calcStarStatus routine. 20160204 - remove call to addOrphanPanum in processError. '20160415 - handle contract (COQ) patients with a department. '20160419 - Check if jhhs mrnum is null.  Allow nulls on patient Type COQ.Removed...all COQ should have a jhhs mrnum number 8/26/2016. '20160503 - Insert blank if patient type is COQ - contract patients. '20160829 - Set status to OA if patient type is COQ. '20160829 - Insert record into 001WeeFim Table. '20160830 - do not discharge is Patient Type = COQ.  Bypass due to nightly discharge.'20170110 - Always Update Admission Date on Pre-Reg. If Admission Date is missing use expected date.***Decided not to use 1/18/2017")> 
<Assembly: AssemblyCompany("Systemax Corp")> 
<Assembly: AssemblyProduct("STAR_ITWFeed")> 
<Assembly: AssemblyCopyright("Copyright ©  2014")> 
<Assembly: AssemblyTrademark("")> 

<Assembly: ComVisible(False)>

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("ba0a843f-3b22-40f0-80bb-22522f26675f")> 

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers 
' by using the '*' as shown below:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("2011.01.05.0")> 
<Assembly: AssemblyFileVersion("2016.04.19.0")> 
