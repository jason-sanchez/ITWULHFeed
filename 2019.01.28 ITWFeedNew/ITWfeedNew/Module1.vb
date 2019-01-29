
'20130401 STAR_ITWFeed.exe

'20100105 - ITWNewFeed.exe
'20100610 -  added ZMI Processing code
'201211010 - setup for McKesson Testing as ITWFeed_MC
'20121206 - add processing for Corporate Number
'20130129 - production for McKesson
'20130524 - remove patnum from A05 insert to 001episode
'20130531 - capture consulting physician raw data from PV1_9 in physconsult field of the episode table.
'20130603 - add the panum to the episode table if an orphan
'20130617 - added code to set the value of global string variable visitStatus based on HL7 record type.
'20130619 - capture class in episode table from "Patient Class" map field
'           added xlateClass function to translate the value of PV1_2 based on Trimap
'20130620 - capture raw star class value from PV1_2 into STAR_class
'20130620 - calculate the status based on raw Patient Class data from STAR in calcStatus Function.
'20130808 - fixed issue with raceDesc during A01,4 insert in the upateEpisode routine.

'20130103 - changed ITW connection string. Old: 10.48.10.246,1433.  New: 10.48.242.249,1433.
'20140124 - added department to update on A01 and A04 Also fixed A01 update to show IA
'20140202 - mods for wave3 testing on cscsysfeed5
'20140205 - modify to use a text logfile instead of the event viewer.
'20140212 - modified convert date functions
'20140213 - added new process for hospital service based on department for outpatients.
'20140215 - started to add A31 processing in for Allergies.
'20140216 - use TriggerEventID instead of Event Type Code for A31 identification.
'20140219 - start calc status based on 130PatientStatus table
'20140220 - only mod status for A01458 if present status is IP or OP or new patient.
'20140220 - use ExpectedDate (PV2_8) for admindate if patient class = p
'20140303 - added severity, reaction and IDDate to Allergy processing.
'20140310 - added status update to A08 in updateepisode routine.
'20140310 - corporate number now in PID_2 - JHHS mrnum
'20140310 - do same as A03 for A11
'20140311 - add A13, cancel discharge.
'20140318 - calcSTARStatus same for A03 and A11
'20140321 - added extractMrnum and extractPanum routines from Mcare because A31's follow Cerner which affixes region to these numbers.
'20140321 - Added code to use extracted  mrnum in processA31 routine.
'20140415 - add A06 process for Leave of Absence (LOA) in UpdateEpisode. Fixed Main Routine Error.
'20140423 - fixed estractCorpNo routine to return "0" instead of "" 
'20140424 - added A08 case for calcStarStatus
'20140528  - Added AuditNotes from ZIN_9
'20140603 - change to put star plancode in the iplancode field for 03insurer.
'20140811 - Mods for AccDate in 08Accidents Table.
'20150817 - Mods for W3 Production
'20140908 - remove ZMI processes.
'20140912 - added star_region to A08 update process.
'20140915 - modified search in processAL1
'20140916 - capture all AL1 data
'20140916 - add update of Admindate on A08
'20140917 - zero out gblAl1Count before counting in checkAL1
'20140918 - accept more than 9 AL1 segments
'20141002 - fixed ARO.
'20141021 - write A34 Information to PatientGlobal database, table = A34Queue
'20141022 - add anset processing back in.
'20141023 - new process ti get onsetdate. procesOnSetDate()
'201410231102 - added process to populate 075occurrence table
'20141206 - added dictNVP.Clear() if no panum 
'20141209 - updated processA34 code. set boolConsultingExists to constant false.
'20150330 - for A03 - if patient class is O then don't set dcdate. Also separated A03 and A11 processes.

'20150413 - VS2013 version
'20150506 - added A12 processing. Same as A02
'20150717 - remove orphan logging and backup tp orphan s directory.

'20150909 - put orphan lagging back in and remove admission date processing for A08 records.
'20150914 - new criteria to determine when to update admission date on an A08 to handle the new re-reg process
'           this will add a new global boolean value that will be calculated in the calcSTARStatus routine. Update admission date on A08 based on this new boolean value.


'20150929 - added A44 processing to production ITW feed.

'20150929 - start cleanup of main routine to deal with all the different trigger events that were added over time.
'           Use a Select Case statement to seperate the different major groups of trigger events - A31, A34, A44 and everything else.


'20151215 - add InsuredDOB and InsuredSex to 03Insurer table.
'20160125 - fixed calcStarStatus routine.

'20160204 - remove call to addOrphanPanum in processError.
'20160415 - do not update department if patient type is COQ - contract patients.
'20160419 - Check if jhhs mrnum is null.  Allow nulls on patient Type COQ. Removed...all COQ should have a jhhs mrnum number 8/26/2016
'20160503 - Insert blank if patient type is COQ - contract patients.
'20160829 - Set status to OA if patient type is COQ.
'20160829 - Insert record into 001WeeFim Table.
'20160830 - do not discharge is Patient Type = COQ.  Bypass due to nightly discharge.
'20170110 - Always Update Admission Date on Pre-Reg. If Admission Date is missing use expected date.***Decided not to use 1/18/2017    
'1/7/2019 - set to false so there is a string of orphans followed by an a31 they are not sent to orphans

Imports System
Imports System.IO
Imports System.Collections
Imports System.Data.SqlClient
Imports System.Diagnostics


Module Module1
    '3/15/2006 - add processing for PCPNo
    Dim gblBoolUpdateAdmDate As Boolean = False '20150914

    Dim connectionString As String
    Dim dictNVP As New Hashtable
    'Dim dictNVP
    Dim sql As String
    Dim sql2 As String
    Dim myfile As StreamReader
    Dim dir As String
    Dim gblORIGINAL_PA_NUMBER As String
    Dim globalError As Boolean = False
    Dim bolDebug As Boolean
    Dim visitPaNumExists As Boolean
    Dim bolProcessInsurer As Boolean
    Dim theError As String
    Dim gblInsCount As Integer
    Dim consultingArray() As String
    Dim strConsultingText As String
    Dim boolConsultingExists As Boolean
    Dim gblEPNum As Integer = 0
    Dim gblLogString As String = ""
    '4/26/2007 Error Handler Global Variables
    Dim functionError As Boolean
    Dim dbError As Boolean
    Dim continueProcessing As Boolean
    Dim orphanFound As Boolean
    '4/26/2007===========================

    '20130617 - added visitstatus as global string variable to use Mcare Status processing
    Dim visitStatus As String = ""

    Private fullinipath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\..\Configs\ULH\HL7Mapper.ini")) ' New test
    'Private fullinipath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\..\..\..\..\..\..\..\Configs\ULH\HL7Mapper.ini")) 'local
    Public objIniFile As New INIFile(fullinipath) '20140817 - New Test
    'Public objIniFile As New INIFile("d:\W3Production\HL7Mapper.ini") '20140817 - Prod
    'Public objIniFile As New INIFile("C:\KY1 Test Environment\HL7Mapper.ini") '20140817 - Local
    'Public objIniFile As New INIFile("C:\W3Feeds\HL7Mapper.ini") '20140817 - Test

    Private fullconinipath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\..\Configs\ULH\ConnProd.ini")) 'New Test
    'Private fullconinipath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\..\..\..\..\..\..\..\Configs\ULH\ConnProd.ini")) 'local
    Public conIniFile As New INIFile(fullconinipath) '20140805 New Test
    'Public conIniFile As New INIFile("d:\W3Production\KY1ConnProd.ini") '20140805 Prod
    'Public conIniFile As New INIFile("C:\KY1 Test Environment\KY1ConnDev.ini") 'Local
    'Public conIniFile As New INIFile("C:\W3Feeds\KY1ConnTest.ini") 'Test

    Dim strInputDirectory As String = ""
    Dim strOutputDirectory As String = ""
    '20140205 - add log file location
    Dim strLogDirectory As String = ""
    Public thefile As FileInfo
    Dim strMapperFile As String = ""
    '20100610 -
    Dim gblZMICount As Integer = 0

    '20121206 - Add Corp No
    Dim gblCorporateNumber As Integer = 0

    Dim gblAL1Count As Integer = 0 '20140215



    Sub Main()
        bolDebug = False

        'declarations for split function
        Dim delimStr As String = "="
        Dim delimiter As Char() = delimStr.ToCharArray()

        'declarations for stream reader
        Dim strLine As String
        Dim sql As String = ""
        Dim strServer As String = ""
        
        Try
            Dim dir As String = objIniFile.GetString("Settings", "directory", "(none)") & ":\"
            Dim parent As String = objIniFile.GetString("Settings", "parentDir", "(none)") & "\"

            '20140205 - add logfile location
            strLogDirectory = dir & parent & objIniFile.GetString("Settings", "logs", "(none)")

            strOutputDirectory = dir & parent & objIniFile.GetString("ITW", "ITWoutputdirectory", "(none)") 'd:\feeds\nvp\itworks\
            'strOutputDirectory = "C:\KY1 Test Environment\W3Feeds\NVP\ITW\"
            strMapperFile = dir & parent & objIniFile.GetString("ITW", "ITWmapper", "(none)")
            'strServer = objIniFile.GetString("Settings", "server", "(none)")

            'setup directory
            '20121013 - use NVP files
            'Dim dirs As String() = Directory.GetFiles(strOutputDirectory, "LTW.*")
            Dim dirs As String() = Directory.GetFiles(strOutputDirectory, "NVP.*")

            '20121013 - new connection string to test server
            'connectionString = "server=10.48.64.6\sqlexpress;database=ITW_MC;uid=sysmax;pwd=Condor!"
            'connectionString = "server=" & strServer & ";database=ITW_MC;uid=sysmax;pwd=Condor!"
            'connectionString = "server=10.48.64.5\sqlexpress;database=ITWTest;uid=sysmax;pwd=Condor!"
            'connectionString = "server=10.48.242.249,1433;database=ITW;uid=sysmax;pwd=Condor!" '20140817
            'connectionString = "server=HPLAPTOP;database=STAR_ITW;uid=sa;pwd=b436328"
            'connectionString = "server=(localdb)\myInstance;database=ITW;uid=sa;pwd=password"
            connectionString = conIniFile.GetString("Strings", "ITW", "(none)")

            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            Dim updatecommand As New SqlCommand
            updatecommand.Connection = myConnection
            objCommand.Connection = myConnection

            For Each dir In dirs

                orphanFound = False '1/7/2019 - set to false so there is a string of orphans followed by an a31 they are not sent to orphans
                functionError = False '1/29/2019 - Set false so other messages are not tagged this way.  Same as orphans.
                dbError = False '1/29/2019 - Set false so other messages are not tagged this way.  Same as orphans.
                globalError = False '1/29/2019 - Set false so other messages are not tagged this way.  Same as orphans.

                thefile = New FileInfo(dir)
                If thefile.Extension <> ".$#$" Then
                    '1.set up the streamreader to get a file
                    myfile = File.OpenText(dir)
                    'and read the first line
                    'strLine = myfile.ReadLine()

                    '20100119 - Catch a problem if the NVP file is messes up
                    Try
                        'Do While Not strLine Is Nothing
                        Do While Not myfile.EndOfStream
                            Dim myArray As String() = Nothing
                            strLine = myfile.ReadLine()
                            If strLine <> "" Then
                                myArray = strLine.Split(delimiter, 2)
                                'add array key and item to hashtable
                                Try
                                    dictNVP.Add(myArray(0), myArray(1))
                                Catch
                                End Try
                            End If
                        Loop
                    Catch ex As Exception
                        'make copy in the problems directory delete any previous ones with same name
                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "problems\" & thefile.Name)
                        fi2.Delete()
                        thefile.CopyTo(strOutputDirectory & "problems\" & thefile.Name)

                        gblLogString = gblLogString & "Dictionary Error" & " - " & thefile.Name & vbCrLf
                        gblLogString = gblLogString & ex.Message & vbCrLf
                        writeTolog(gblLogString, 1)
                        'get rid of the file so it doesn't mess up the next run.
                        myfile.Close()
                        If thefile.Exists Then
                            thefile.Delete()
                            Exit Sub
                        End If
                    End Try
                    '20100119 - Catch a problem if the NVP file is messes up

                    myfile.Close()

                    Select Case dictNVP.Item("TriggerEventID")
                        '20180305 - Move files that do not have an observation datetime to the NO OBS Datetime Folder
                        Case "Z47"
                            Call checkZ47(thefile, myfile)

                        Case "A44"
                            Call ProcessA44(dictNVP)

                        Case "A31"
                            gblCorporateNumber = extractCorpNo(dictNVP.Item("fullpid"))
                            gblORIGINAL_PA_NUMBER = extractPanum(dictNVP.Item("panum"))
                            gblAL1Count = 0 '20140917
                            Call checkAL1(dictNVP) '20130828
                            Call processAL1(dictNVP) '20130828

                        Case "A34"
                            gblCorporateNumber = dictNVP.Item("JHHS mrnum")
                            Call ProcessA34(dictNVP)

                        Case Else

                            '20160419 - Check if jhhs mrnum is null.  Allow nulls on patient Type COQ
                            'If dictNVP("Patient Type") = "COQ" Then
                            'gblCorporateNumber = checkcorpnum(dictNVP.Item("JHHS mrnum"))
                            'End If
                            gblCorporateNumber = dictNVP.Item("JHHS mrnum")

                            gblORIGINAL_PA_NUMBER = ""
                            If (dictNVP.Item("panum") <> "") Then

                                gblORIGINAL_PA_NUMBER = extractPanum(dictNVP.Item("panum"))

                                gblInsCount = 0
                                '20100610 ->
                                gblZMICount = 0
                                '===================================================================================================
                                'call subdirectories here
                                globalError = False


                                Call processError(dictNVP)
                                If continueProcessing Then
                                    Call checkIN1(dictNVP)
                                    '20100610 ->
                                    'Call checkZMI(dictNVP) '20140908 - removed ZMI

                                    Call updateEpisode(dictNVP)
                                    Call updateEpisodeSupplement(dictNVP)
                                    Call updateContact(dictNVP)
                                    Call UpdateInsurer(dictNVP)
                                    '20130514 added the financial routine of 20130429
                                    Call updateFinancial(dictNVP)
                                    '20100610 ->
                                    'Call processZMI(dictNVP) '20140908 - removed ZMI

                                    Call UpdatePPS(dictNVP)
                                    Call processOccurrenceCodes(dictNVP, gblORIGINAL_PA_NUMBER) '20141023


                                End If 'If continueProcessing 

                            End If 'If (dictNVP.Item("panum") <> "") Then
                    End Select


                    '===================================================================================================
                    dictNVP.Clear()
                    If functionError Then

                        gblLogString = "Function Error - " & thefile.Name & vbCrLf & gblLogString
                        writeTolog(gblLogString, 2)
                        gblLogString = ""

                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "backup\" & thefile.Name)
                        fi2.Delete()
                        thefile.CopyTo(strOutputDirectory & "backup\" & thefile.Name)
                        thefile.Delete()



                    ElseIf dbError Then

                        gblLogString = "dbError Error - " & thefile.Name & vbCrLf & gblLogString
                        writeTolog(gblLogString, 2)
                        gblLogString = ""

                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "reprocess\" & thefile.Name)
                        fi2.Delete()
                        thefile.CopyTo(strOutputDirectory & "reprocess\" & thefile.Name)
                        thefile.Delete()

                    ElseIf orphanFound Then


                        writeTolog(gblLogString, 3)
                        gblLogString = ""

                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "orphans\" & thefile.Name)
                        fi2.Delete()
                        thefile.CopyTo(strOutputDirectory & "orphans\" & thefile.Name)
                        thefile.Delete()

                    ElseIf globalError Then
                        gblLogString = "Global Error - " & thefile.Name & vbCrLf & gblLogString
                        writeTolog(gblLogString, 2)
                        gblLogString = ""

                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "backup\" & thefile.Name)
                        fi2.Delete()
                        thefile.CopyTo(strOutputDirectory & "backup\" & thefile.Name)
                        thefile.Delete()

                    Else
                        '20121210 - make a backup copy for mckesson testing
                        'Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "mckesson\" & thefile.Name)
                        'fi2.Delete()
                        'thefile.CopyTo(strOutputDirectory & "mckesson\" & thefile.Name)
                        thefile.Delete()

                    End If

                Else
                    dictNVP.Clear() '20141206 added
                    thefile.Delete()


                End If 'If theFile.Extension <> ".$#$"
            Next

        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "Main Routine Error: " & thefile.Name & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            writeTolog(gblLogString, 1)
            gblLogString = ""

            Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "backup\" & thefile.Name)
            fi2.Delete()
            thefile.CopyTo(strOutputDirectory & "backup\" & thefile.Name)
            thefile.Delete()

            '20091117 - get rid of the problem file if it exists
            'If thefile.Exists Then
            'thefile.Delete()
            'End If

            Exit Sub
        End Try
    End Sub
    Public Sub insertInsurer(ByVal dictNVP As Hashtable)
        '20140528 Added AuditNotes from ZIN_9
        '20140603 - change to put star plancode in the iplancode field
        '20140904 - put back integer fclass routine
        '20151215 - add Insuredsex and InsuredDOB processing to 03Insurer table
        Try
            Dim i As Integer = 0
            Dim strSubName As String = ""
            Dim strPolicyIssueDate As String = ""
            Dim strAuthServices As String = ""
            Dim tempStr As String = ""
            Dim strFClass As String = ""
            Dim intFClass As Integer = 0 '20140904
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            Dim updatecommand As New SqlCommand
            updatecommand.Connection = myConnection
            '3/23/2006
            Dim iplancodeExists As Boolean = False

            objCommand.Connection = myConnection
            Dim dataReader As SqlDataReader
            Dim STAR_Plancode As String = ""


            intFClass = 0
            sql = "select id from [104finclass] where finclass = '" & dictNVP.Item("insurer.fClass") & "' and inactive = 0" '20140904
            objCommand.CommandText = sql

            myConnection.Open()
            dataReader = objCommand.ExecuteReader()
            While dataReader.Read()
                intFClass = dataReader.GetInt32(0)
            End While
            'strFClass = dictNVP.Item("insurer.fClass") '20130410

            myConnection.Close()
            dataReader.Close()

            i = 0
            For i = 1 To gblInsCount
                If i = 1 Then
                    tempStr = ""
                End If
                If i > 1 Then
                    tempStr = "_000" & i
                End If

                '20130429 - Added STAR_Plancode to [03insurer] table
                STAR_Plancode = Trim(Replace(dictNVP("iplancode2" & tempStr), "'", "''")) & Trim(Replace(dictNVP("iplancode" & tempStr), "'", "''"))

                '==============================================================================
                '3/23/2006
                'don't insert if iplancode exists
                sql = "SELECT epnum FROM [03Insurer] where epnum = " & gblEPNum & " "
                sql = sql & "AND star_plancode = '" & STAR_Plancode & "'"
                objCommand.CommandText = sql
                myConnection.Open()
                dataReader = objCommand.ExecuteReader()

                If dataReader.HasRows Then
                    iplancodeExists = True
                Else
                    iplancodeExists = False
                End If
                myConnection.Close()
                dataReader.Close()
                '==============================================================================
                If Not iplancodeExists Then '3/23/2006
                    sql = ""
                    strAuthServices = Replace(dictNVP.Item("authservices1" & tempStr), "'", "''") & " " & Replace(dictNVP.Item("authservices2" & tempStr), "'", "''")
                    If Len(dictNVP.Item("PolicyIssueDate" & tempStr)) > 0 Then
                        strPolicyIssueDate = ConvertDate(dictNVP.Item("PolicyIssueDate" & tempStr))
                    Else
                        strPolicyIssueDate = ""
                    End If

                    strSubName = Replace(dictNVP.Item("Insured First Name" & tempStr), "'", "''")
                    strSubName = strSubName & " " & Replace(dictNVP.Item("Insured Middle Name" & tempStr), "'", "''")
                    strSubName = strSubName & " " & Replace(dictNVP("Insured Last Name" & tempStr), "'", "''")
                    'remove these later
                    gblORIGINAL_PA_NUMBER = dictNVP.Item("panum")

                    sql = "Insert [03insurer] "
                    '20120429 - added star_plancode,conum and company name
                    '20140528 Added AuditNotes from ZIN_9
                    sql = sql & "(epnum, iPlanCode, iplancode2, coNum, STAR_Plancode, coname, subname, policyNum, "
                    sql = sql & "authNum1, theGroup, PIssue, aprimary, Fclass, AuditNotes, reqCert, " '20140528
                    'sql = sql & "theGroup, PIssue, aprimary, Fclass, AuditNotes, reqCert, " 'the will return for authnum7
                    '20151215 - add insuredSex and InsuredDOB
                    sql = sql & "InsuredDOB, InsuredSex,"

                    sql = sql & "updated) "

                    sql = sql & "VALUES ("
                    sql = sql & gblEPNum & ", "

                    '20140603 - change to put star plancode in the iplancode field
                    'insertString(Replace(dictNVP("iplancode" & tempStr), "'", "''"))
                    insertString(STAR_Plancode)
                    '======================================================================
                    insertString(Replace(dictNVP("iplancode2" & tempStr), "'", "''"))
                    '20130502 - insert iplancode in company number as an int
                    insertNumber(Replace(dictNVP("iplancode2" & tempStr), "'", "''"))

                    '20120429 - added star_plancode and company name
                    insertString(STAR_Plancode)
                    insertString(Replace(dictNVP("CompanyName" & tempStr), "'", "''"))

                    insertString(strSubName)
                    insertString(dictNVP.Item("PolicyNumber" & tempStr))
                    '20170510 - Removed AuthNum Process using Process IN1_14 for authcode project.  Leave for ULHT.
                    insertString(Replace(dictNVP.Item("AuthNum" & tempStr), "'", "''"))
                    insertString(Replace(dictNVP("group" & tempStr), "'", "''"))

                    insertString(strPolicyIssueDate)

                    '20140528
                    '20140528 Added AuditNotes from ZIN_9
                    If dictNVP("COBPriority" & tempStr) = "1" Then
                        sql = sql & "1, " & intFClass & ", '" & Replace(dictNVP("InsNotes"), "'", "''") & "', " '20140528 '20140904
                    Else
                        sql = sql & "0, NULL, NULL, " '20140528
                    End If

                    sql = sql & "1, "

                    '20151215 - add InsuredDOB and InsuredSex
                    If Len(dictNVP("Insured DOB")) > 0 Then
                        sql = sql & "'" & ConvertDate(dictNVP("Insured DOB")) & "', "
                    Else
                        sql = sql & "NULL, "
                    End If

                    If Len(dictNVP("Insured Sex")) > 0 Then
                        sql = sql & "'" & dictNVP("Insured Sex") & "', "
                    Else
                        sql = sql & "NULL, "
                    End If
                    '20151215 end

                    sql = sql & "'" & DateTime.Now & "') "

                    updatecommand.CommandText = sql
                    myConnection.Open()
                    updatecommand.ExecuteNonQuery()
                    myConnection.Close()
                    '20170623 Process IN1_14 and ZGI for multiple authcodes
                    ProcessIN1_14(dictNVP, tempStr)
                End If 'If Not iplancodeExists
                '20170510 - Removed AuthNum Process using Process IN1_14

            Next 'i to gblInsCount


        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "Insert Insurer Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

            Exit Sub
        End Try
    End Sub

    Public Sub updateContact(ByVal dictNVP As Hashtable)
        Try
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            Dim updatecommand As New SqlCommand
            updatecommand.Connection = myConnection
            Dim dataReader As SqlDataReader
            objCommand.Connection = myConnection
            Dim boolContactExists As Boolean = False
            Dim strID As String = ""
            Dim sql As String

            '8/16/2001 - UpdateContact:
            'This updates the [03contact] Table if gblEPNum <> 0
            'Assumes one contact per episode
            '1. Search for records where [03contact].epnum = glbEPNum.
            '
            '2. If found, update [03contact] records per spreadsheet.
            '
            '6/13/2002: run only if A01, A04, A05 and A08
            Select Case dictNVP.Item("Event Type Code")
                Case "A01", "A04", "A05", "A08"

                    If gblEPNum <> 0 Then

                        If Len(dictNVP.Item("03contact.lastName")) > 0 Then
                            'Console.WriteLine(gblEPNum)
                            'Console.ReadLine()
                            sql = "SELECT ID from [03contact] where epnum = " & gblEPNum
                            objCommand.CommandText = sql
                            myConnection.Open()
                            dataReader = objCommand.ExecuteReader()
                            If dataReader.HasRows Then
                                boolContactExists = True
                                dataReader.Read()
                                strID = dataReader.Item("ID")
                            Else
                                boolContactExists = False
                            End If

                            myConnection.Close()
                            dataReader.Close()

                            If boolContactExists Then
                                sql = "UPDATE [03contact] "
                                sql = sql & "SET updated = '" & DateTime.Now & "'"
                                '20130502\

                                If Len(dictNVP.Item("03contact.lastName")) > 2 Then
                                    If dictNVP.Item("03contact.firstName") = """""" Then dictNVP.Item("03contact.firstName") = ""
                                    If dictNVP.Item("03contact.lastName") = """""" Then dictNVP.Item("03contact.lastName") = ""
                                    sql = sql & ", name = '" & Replace(dictNVP.Item("03contact.firstName"), "'", "''") & " " & Replace(dictNVP.Item("03contact.lastName"), "'", "''") & "' "
                                End If

                                'sql = sql & ", relation = '" & Replace(dictNVP.Item("nok.code"), "'", "''") & "' "
                                'If Len(dictNVP.Item("nok.code")) > 0 Then
                                sql = sql & STARupdateString("relation", dictNVP.Item("nok.code"))
                                'End If

                                sql = sql & ", relationDesc = '" & Replace(dictNVP.Item("nok.desc"), "'", "''") & "' "

                                'sql = sql & ", ph1 = '" & dictNVP.Item("nok.phone") & "' "
                                'If Len(dictNVP.Item("nok.phone")) > 0 Then
                                sql = sql & STARupdateString("ph1", dictNVP.Item("nok.phone"))
                                'End If

                                'sql = sql & ", ph2 = '" & dictNVP.Item("nok.businessphone") & "' "
                                'If Len(dictNVP.Item("nok.businessphone")) > 0 Then
                                sql = sql & STARupdateString("ph2", dictNVP.Item("nok.businessphone"))
                                'End If

                                sql = sql & ", emergency = 1"
                                sql = sql & " WHERE ID = " & strID

                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()
                            Else
                                sql = "INSERT [03contact] (epnum, name, relation, relationDesc, ph1, ph2, emergency, updated) "
                                sql = sql & "VALUES ("
                                sql = sql & gblEPNum & ", "

                                sql = sql & "'" & Replace(dictNVP.Item("03contact.firstName"), "'", "''") & " " & Replace(dictNVP.Item("03contact.lastName"), "'", "''") & "', "
                                sql = sql & "'" & Replace(dictNVP.Item("nok.code"), "'", "''") & "', "
                                sql = sql & "'" & Replace(dictNVP.Item("nok.desc"), "'", "''") & "', "
                                sql = sql & "'" & dictNVP.Item("nok.phone") & "', "
                                sql = sql & "'" & dictNVP.Item("nok.businessphone") & "', "
                                sql = sql & "1, "
                                sql = sql & "'" & DateTime.Now & "') "

                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()
                            End If ' check if contact record exists for epnum
                        End If

                    End If 'if gblEPNum <> 0

            End Select ' Case "A01", "A04", "A05", "A08"

        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "Update Contact Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

            Exit Sub
        End Try
    End Sub

    Public Sub updateEpisode(ByVal dictNVP As Hashtable)
        '============================================================
        '3/15/2006 - add processing for PCPNo
        '04/17/2007 added length check
        '20130522 - added department A01,4,8
        '20130528 - add onsetdate from occurrence codes in UB1_16_1 and 2 if UB1_16_1 is 11
        '20140912 - added star_region to A08.
        '============================================================
        Try
            Dim boolRecordExists As Boolean = False
            '12/12/2002 - added accident boolean for accident table processing
            Dim boolAccidentExists As Boolean = False
            Dim intIntakeFacility As Long = 0
            Dim strRoomBed As String = ""
            Dim tempstr As String = ""
            '12/07/2004
            Dim strAro As String = ""
            Dim strAdvanceDir As String = ""
            Dim strAllergies As String = ""
            Dim strReferralSourceID As String = ""
            Dim strJHHS_mrnum As String = ""
            '5/9/2005
            Dim strPV1_19 As String = ""
            '2/8/2006
            Dim j As Integer = 0
            Dim i As Integer = 0
            '2/17/2006

            Dim visitStatus As String = ""
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            Dim updatecommand As New SqlCommand
            updatecommand.Connection = myConnection

            objCommand.Connection = myConnection
            Dim dataReader As SqlDataReader

            Dim addit As Boolean = False
            Dim updateit As Boolean = False
            Dim strEventTypeCode As String = dictNVP.Item("Event Type Code")

            boolConsultingExists = False
            Dim star_plancode As String = ""
            '20130502 - added star_region to episode table. Handle on A01, 4 and 5 insert. 20140912 - added to A08 update also
            Dim star_region As String = UCase(dictNVP("Sending Facility"))

            '20090803 deal with phone numbers=======================================================
            Dim primaryPhone As String = ""
            Dim altPhone As String = ""
            Dim businessPhone As String = ""

            '20120528
            Dim strOcDate As String = ""
            'If InStr(Replace(dictNVP.Item("01patient.patPhone"), "'", "''"), "~") Then
            'Dim phoneArray() As String
            'phoneArray = Split(Replace(dictNVP.Item("01patient.patPhone"), "'", "''"), "~")
            'primaryPhone = phoneArray(0)
            'altPhone = phoneArray(1)
            'Else
            primaryPhone = Replace(dictNVP.Item("01patient.patPhone STAR"), "'", "''")
            'End If
            businessPhone = Replace(dictNVP.Item("Patient Business Phone STAR"), "'", "''")
            '20090803==================================================================================

            sql = ""
            'code here to combine the separate room and bed fields from the generic feed
            strRoomBed = Replace(dictNVP.Item("01visit.room"), "'", "''") & Replace(dictNVP.Item("01visit.bed"), "'", "''")


            If Len(dictNVP.Item("Consulting Physician")) > 5 Then
                boolConsultingExists = True ' will use this later when updating database
                strConsultingText = dictNVP.Item("Consulting Physician")
                consultingArray = Split(strConsultingText, "~")
            End If

            boolConsultingExists = False '20141209
            '=====================================================================================================
            'check accident table and set boolean for processing

            sql = "select * from [08accidents] where panum = '" & dictNVP.Item("panum") & "'"
            objCommand.CommandText = sql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()
            If dataReader.HasRows Then
                boolAccidentExists = True
            Else
                boolAccidentExists = False
            End If

            myConnection.Close()
            dataReader.Close()

            '=====================================================================================================
            'Try
            'see if panum exixts
            sql = "select epnum from [001Episode] where panum = '" & gblORIGINAL_PA_NUMBER & "'"
            objCommand.CommandText = sql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()

            If dataReader.HasRows Then
                boolRecordExists = True
                dataReader.Read()
                gblEPNum = dataReader.Item("epnum")
            Else
                boolRecordExists = False
                gblEPNum = 0
            End If

            myConnection.Close()
            dataReader.Close()
            '=====================================================================================================
            '20130502 - get intake facility for ITW regions as below:
            intIntakeFacility = 0
            'End If
            Select Case UCase(dictNVP("Sending Facility"))
                Case "Q", "R" 'Q = frazier out; R = frazier in
                    intIntakeFacility = 200
                Case "H" 'SIRH
                    intIntakeFacility = 300
                Case "T" '20170623 - ULH
                    intIntakeFacility = 400
            End Select
            '=====================================================================================================
            '12/07/2004 add processing for aro, jhhs mrnum, allergies, advance directive, referral source ID
            'fields added after mrnum
            If dictNVP.Item("AROFieldName") = "ARO" Then
                strAro = Replace(dictNVP("ARODataField"), "'", "''")
            Else
                strAro = ""
            End If

            If Len(dictNVP.Item("AdvDirective")) >= 6 And dictNVP.Item("AdvDirective") <> """""" Then
                strAdvanceDir = Left$(dictNVP.Item("AdvDirective"), 6)
            Else
                strAdvanceDir = ""
            End If

            If dictNVP("Allergy Description") <> "" Then
                strAllergies = Replace(dictNVP("Allergy Description"), "'", "''")
            Else
                strAllergies = ""
            End If

            If dictNVP("Referral Source ID") <> "" Then
                strReferralSourceID = Replace(dictNVP("Referral Source ID"), "'", "''")
            Else
                strReferralSourceID = ""
            End If

            If dictNVP("JHHS mrnum") <> "" Then
                strJHHS_mrnum = Replace(dictNVP("JHHS mrnum"), "'", "''")
            Else
                strJHHS_mrnum = ""
            End If

            '5/9/2005
            If dictNVP("PV1_19") <> "" Then
                strPV1_19 = dictNVP("PV1_19")
            Else
                strPV1_19 = "NOT SENT"
            End If

            '20130528 - set the dictionary value
            '20141023 - use new function, processOnSetDate to extract the date if code is 11
            'If dictNVP.Item("occode") = "11" Then
            'dictNVP("OnsetDate") = dictNVP("ocdate")
            'End If
            dictNVP("OnsetDate") = processOnSetDate(dictNVP.Item("UB1string"))

            '20130617 - added code to set the value of global string variable visitStatus based on HL7 record type.
            'visitStatus = ""
            'If dictNVP.Item("Event Type Code") = "A01" Then visitStatus = "IP"
            'If dictNVP.Item("Event Type Code") = "A03" Then visitStatus = "OEC"
            'If dictNVP.Item("Event Type Code") = "A04" Then visitStatus = "OP"
            'If dictNVP.Item("Event Type Code") = "A05" Then visitStatus = "PRE"
            'If dictNVP.Item("Event Type Code") = "A06" Then visitStatus = "IP"
            'If dictNVP.Item("Event Type Code") = "A07" Then visitStatus = "OP"

            '20130620 - use the calcStatus function to generate the visitStatus value
            visitStatus = ""

            '20140219 - use new process
            'visitStatus = calcStatus(dictNVP)
            visitStatus = calcSTARStatus(dictNVP)

            '20140213 - new process for hospital service
            '***********************************************************
            '***********************************************************
            Dim tmpHospitalService As String = ""
            Dim tmpPatientType As String = ""
            Dim boolContinueProcessing As Boolean = False
            Dim departmentCode As String = ""

            tmpPatientType = xlateClass(dictNVP)

            If UCase(tmpPatientType) = "O" Then
                departmentCode = dictNVP.Item("department")
                tmpHospitalService = departmentCode

                '20160503 - Insert blank department and hservice if patient type is COQ - contract patients.
                If dictNVP.Item("Patient Type") = "COQ" Then
                    boolContinueProcessing = True
                    tmpHospitalService = ""
                    departmentCode = ""

                ElseIf tmpHospitalService <> "" Then
                    boolContinueProcessing = True

                End If

            ElseIf UCase(tmpPatientType) = "I" Then

                tmpHospitalService = Replace(dictNVP.Item("Hospital Service"), "'", "''")
                boolContinueProcessing = True

            Else
                boolContinueProcessing = True
            End If
            '***********************************************************
            '***********************************************************

            If boolContinueProcessing Then '20140213

                Select Case dictNVP.Item("Event Type Code")
                    '======================================================================================================
                    Case "A01", "A04"
                        '12/13/2001 - added room processing for A01 and A04
                        '3/14/2002 - added code to handle gender, race and marital status
                        '12/02/2002 - added hservice and patient_type after DOB
                        '12/06/2002 - fixed gender, added height and weight
                        '8/14/2003 added dcDisp and county fields before status
                        '11/17/2004 - added religion
                        '12/07/2004 add processing for aro, jhhs mrnum, allergies, advance directive, referral source ID
                        'fields added after mrnum



                        '2. If record does not exist add it
                        If Not boolRecordExists Then
                            '5/9/2005 - handle pv1_19 after mrnum
                            '3/14/2007 - add onset date handling after pv1_19
                            '20121206 - add corpNo after mrmun
                            '20130502 - add star_region after onsetDate
                            '20130619 - add class (nvarchar(10) to episode table using "Patient Class", PV1_2. After corpNo.
                            '20130620 - add STAR_class (nvarchar(10) to episode table after class 
                            sql = "INSERT into [001Episode] (mrnum, corpNo, class, STAR_class, PV1_19, onsetDate, STAR_region, department, aro, allergies, jhhs_mrnum, advanceDir, ReferralSourceID, "
                            sql = sql & "dcDisp, county, status, panum, room, lname, fname, mname, socsec, dob, hService, patient_type, "
                            '20090803 - add files for AltPhone and BusinessPhone after phone
                            sql = sql & "addr1, addr2, city, state, zip, phone, altPhone, businessPhone, physRefer, physAdmit, physAttend, physconsult, AdminDate, "
                            '2'8'2006
                            sql = sql & "consult1, consult2, consult3, consult4, consult5, "
                            '3/15/2006 - added PCPNo after consult10
                            sql = sql & "consult6, consult7, consult8, consult9, consult10, PCPNo, "



                            If Len(dictNVP("Patient Religion")) > 0 Then
                                sql = sql & "diagnosis, religionID, primaryIns, race, raceDesc, gender, MaritalStatus, IntakeFacility, created, feedstarted, active, preAdmit) "
                            Else
                                sql = sql & "diagnosis, primaryIns, race, raceDesc, gender, MaritalStatus, IntakeFacility, created, feedstarted, active, preAdmit) "
                            End If


                            '9/4/2003 - using isnumeric to check mrnum
                            If IsNumeric(dictNVP("mrnum")) Then
                                sql = sql & "VALUES (" & dictNVP("mrnum") & ", "
                            Else
                                sql = sql & "VALUES (0, "
                            End If

                            '20121206 - add corpNo
                            sql = sql & gblCorporateNumber & ", "

                            '20130619 - add class from "Patient Class" map with translation.
                            sql = sql & "'" & xlateClass(dictNVP) & "', "

                            '20130620 - add class from "Patient Class" map with translation.
                            sql = sql & "'" & dictNVP("Patient Class") & "', "

                            '5/9/2005
                            sql = sql & "'" & strPV1_19 & "', "
                            '3/14/2007

                            '20141022 - add anset processing back in.
                            If Len(dictNVP("OnsetDate")) > 7 Then
                                sql = sql & "'" & ConvertDate(dictNVP("OnsetDate")) & "', " '20140906
                            Else
                                sql = sql & "NULL" & ", "
                            End If

                            '20130502 - add star_region
                            sql = sql & "'" & star_region & "', "

                            '20120522 - added department
                            '20160503 - Insert blank if patient type is COQ - contract patients.
                            sql = sql & "'" & Replace(departmentCode, "'", "''") & "', "

                            sql = sql & "'" & strAro & "', "
                            sql = sql & "'" & strAllergies & "', "
                            sql = sql & "'" & strJHHS_mrnum & "', "
                            sql = sql & "'" & strAdvanceDir & "', "
                            sql = sql & "'" & strReferralSourceID & "', "

                            '8/14/2003 added dcDisp and county
                            sql = sql & "'" & Replace(dictNVP.Item("Discharge Disposition"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("Patient County/Parish Code"), "'", "''") & "', "

                            '12/02/2002 - changed code above to get the patient status from the feed
                            '20130617 - use visitStatus varible like Mcare
                            'sql = sql & "'" & Replace(dictNVP.Item("Patient Status Code"), "'", "''") & "', "
                            '20130617
                            sql = sql & "'" & visitStatus & "', "

                            sql = sql & "'" & Replace(dictNVP.Item("panum"), "'", "''") & "', "
                            sql = sql & "'" & strRoomBed & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patlast"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patfirst"), "'", "''") & "',"
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patmi"), "'", "''") & "', "

                            '04/17/2007 added length check
                            If Len(dictNVP.Item("01patient.patSS")) >= 9 Then
                                sql = sql & "'" & Replace(dictNVP.Item("01patient.patSS"), "'", "''") & "', "
                            Else
                                sql = sql & "NULL, "
                            End If

                            sql = sql & "'" & ConvertDate(dictNVP.Item("01patient.DOB")) & "', "
                            '12/20/2002 added hservice and patient_type

                            '20140213 
                            'sql = sql & "'" & Replace(dictNVP.Item("Hospital Service"), "'", "''") & "', "
                            sql = sql & "'" & tmpHospitalService & "', "

                            sql = sql & "'" & Replace(dictNVP.Item("Patient Type"), "'", "''") & "', "

                            sql = sql & "'" & Replace(dictNVP.Item("01patient.pataddr1"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.pataddr2"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patcity"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patstate"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patzip"), "'", "''") & "', "
                            '20090803 add processing for altphone and businessPhone
                            If primaryPhone <> "" Then
                                sql = sql & "'" & primaryPhone & "', "
                            Else
                                sql = sql & "NULL, "
                            End If
                            If altPhone <> "" Then
                                sql = sql & "'" & altPhone & "', "
                            Else
                                sql = sql & "NULL, "
                            End If
                            If businessPhone <> "" Then
                                sql = sql & "'" & businessPhone & "', "
                            Else
                                sql = sql & "NULL, "
                            End If
                            '5/30/2002
                            'added code to verify that physician numbers are numeric
                            '12/11/2002 removed length check
                            If IsNumeric(dictNVP.Item("Referring.patPhysNum")) Then
                                sql = sql & dictNVP.Item("Referring.patPhysNum") & ", "
                            Else
                                sql = sql & "0, "
                            End If
                            '12/11/2002 removed length check
                            If IsNumeric(dictNVP.Item("Admitting.patPhysNum")) Then
                                sql = sql & dictNVP.Item("Admitting.patPhysNum") & ", "
                            Else
                                sql = sql & "0, "
                            End If

                            '6/28/2005 added attending and consulting physicians
                            '=============================================================================
                            If IsNumeric(dictNVP.Item("Attending Physician ID")) Then
                                sql = sql & dictNVP.Item("Attending Physician ID") & ", "
                            Else
                                sql = sql & "0, "
                            End If

                            '20130531 - capture the physconsult value from the feed
                            'If IsNumeric(dictNVP.Item("Consulting Physician ID")) Then
                            sql = sql & "'" & Replace(dictNVP.Item("Consulting Physician ID"), "'", "''") & "', "
                            'Else
                            'sql = sql & "0, "
                            'End If
                            '============================================================================

                            '20140220
                            If dictNVP.Item("Patient Class") = "P" Then
                                sql = sql & "'" & ConvertDate(dictNVP.Item("ExpectedDate")) & "', "
                            Else
                                sql = sql & "'" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "', "
                            End If

                            '2/8/2006
                            'If boolConsultingExists Then
                            'For j = 0 To UBound(consultingArray)
                            'If j < 10 Then
                            'insertNumber(consultingArray(j))
                            'End If
                            'Next
                            'If UBound(consultingArray) < 9 Then
                            'For j = (UBound(consultingArray) + 1) To 9
                            'sql = sql & "NULL,"
                            'Next
                            'End If
                            'Else

                            sql = sql & "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,"
                            'End If

                            '3/15/2006 added PCPNo (PD1_4_1)
                            insertNumber(dictNVP.Item("PCPNo"))

                            sql = sql & "'" & Replace(dictNVP.Item("02diag.diagnosis"), "'", "''") & "', "

                            '11/17/2004 - added religion
                            If Len(dictNVP.Item("Patient Religion")) > 0 Then
                                sql = sql & "'" & Replace(dictNVP.Item("Patient Religion"), "'", "''") & "', "
                            End If
                            '===============================================================
                            '11/08/2001
                            'If dictNVP("1.COBPriority") = "1" Then
                            'sql = sql & "'" & Replace(dictNVP.Item("1.iplancode"), "'", "''") & "', "
                            'ElseIf dictNVP("2.COBPriority") = "1" Then
                            'sql = sql & "'" & Replace(dictNVP.Item("2.iplancode"), "'", "''") & "', "
                            'ElseIf dictNVP("3.COBPriority") = "1" Then
                            'sql = sql & "'" & Replace(dictNVP.Item("3.iplancode"), "'", "''") & "', "
                            'Else
                            'sql = sql & "'NNN', "
                            'End If
                            Dim tempPlanCode As String = ""
                            For i = 1 To gblInsCount

                                If i = 1 Then
                                    tempstr = ""
                                End If
                                If i > 1 Then
                                    tempstr = "_000" & i
                                End If
                                '20130429
                                star_plancode = Trim(Replace(dictNVP("iplancode2" & tempstr), "'", "''")) & Trim(Replace(dictNVP("iplancode" & tempstr), "'", "''"))

                                If dictNVP.Item("COBPriority" & tempstr) = "1" Then
                                    'sql = sql & "'" & Replace(dictNVP.Item("iplancode" & tempstr), "'", "''") & "', "
                                    tempPlanCode = star_plancode
                                End If
                            Next
                            If tempPlanCode <> "" Then
                                sql = sql & "'" & tempPlanCode & "', "
                            Else
                                sql = sql & "'NNN', "
                            End If

                            If dictNVP("race code STAR") <> "" Then
                                sql = sql & "'" & dictNVP("race code STAR") & "', "
                            Else
                                sql = sql & "'-', "
                            End If

                            If dictNVP("race description STAR") <> "" Then
                                sql = sql & "'" & dictNVP("race description STAR") & "', "
                            Else
                                sql = sql & "'-', "
                            End If

                            If dictNVP("Patient Sex") <> "" Then
                                sql = sql & "'" & dictNVP("Patient Sex") & "', "
                            Else
                                sql = sql & "'-', "
                            End If

                            If dictNVP("Patient Marital Status") <> "" Then
                                sql = sql & "'" & dictNVP("Patient Marital Status") & "', "
                            Else
                                sql = sql & "'-', "
                            End If

                            sql = sql & intIntakeFacility & ", "
                            '6/26/2002: set the preAdmit flag to false
                            sql = sql & "'" & DateTime.Now & "', 1, 1, 0) "

                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()

                            'get the last epnum added
                            sql = "SELECT max(epnum) AS [lastNumber] from [001episode]"
                            objCommand.CommandText = sql
                            myConnection.Open()
                            dataReader = objCommand.ExecuteReader()

                            If dataReader.HasRows Then
                                dataReader.Read()
                                gblEPNum = dataReader.Item("lastNumber")
                            End If
                            myConnection.Close()
                            dataReader.Close()


                        Else 'If Not boolRecordExixts Then
                            '3. record exixts, see if it needs to be updated

                            sql = "UPDATE [001episode] "
                            sql = sql & "SET modified = '" & DateTime.Now & "'"

                            '20130502 - update star_region on A01 and A04 following an A05
                            If star_region <> "" Then
                                sql = sql & ", star_region = '" & star_region & "'"
                            End If

                            '20130619 - update class from Patient Class if not blank with translation
                            If dictNVP("Patient Class") <> "" Then
                                sql = sql & ", class = '" & xlateClass(dictNVP) & "' "
                            End If

                            '20130620 - update class from Patient Class raw data
                            If dictNVP("Patient Class") <> "" Then
                                sql = sql & ", star_class = '" & dictNVP.Item("Patient Class") & "' "
                            End If
                            '=============================================================================================
                            '20140124 - added department to update on A01 and A04
                            '20160415 - do not update department if patient type is COQ - contract patients.
                            If departmentCode <> "" Then
                                sql = sql & ", department = '" & Replace(departmentCode, "'", "''") & "' "
                            End If
                            '==============================================================================

                            '9/4/2003 - using isnumeric to check mrnum
                            If IsNumeric(dictNVP("mrnum")) Then
                                sql = sql & ", mrnum = " & dictNVP.Item("mrnum") & " "
                            End If

                            '20121206 - update corpNo with A01 or A04 only
                            If IsNumeric(gblCorporateNumber) Then
                                sql = sql & ", corpNo = " & gblCorporateNumber & " "
                            End If

                            '3/15/2006 added PCPNo (PD1_4_1)
                            If IsNumeric(dictNVP.Item("PCPNo")) Then
                                sql = sql & ", PCPNo = " & dictNVP.Item("PCPNo")
                            End If

                            '20141022 - add anset processing back in.
                            If Len(dictNVP("OnsetDate")) > 7 Then
                                sql = sql & ", onsetDate = '" & ConvertDate(dictNVP("OnsetDate")) & "'"
                            Else
                                sql = sql & ", onsetDate = NULL "
                            End If

                            '12/13/2001 - room handling added
                            If Len(strRoomBed) > 1 Then
                                sql = sql & ", room = '" & strRoomBed & "'"
                            End If

                            sql = sql & STARupdateString("lname", dictNVP.Item("01patient.patlast"))

                            sql = sql & STARupdateString("fname", dictNVP.Item("01patient.patfirst"))

                            sql = sql & STARupdateString("mname", dictNVP.Item("01patient.patmi"))


                            '5/10/2005
                            '04/17/2007 - added length check
                            If Len(dictNVP.Item("01patient.patSS")) >= 9 Then

                                sql = sql & STARupdateString("SocSec", dictNVP.Item("01patient.patSS"))

                            End If

                            sql = sql & ", DOB = '" & ConvertDate(dictNVP.Item("01patient.DOB")) & "'"
                            '12/02/2002 - added hservice and patient_type

                            '20140213
                            'sql = sql & ", hService = '" & Replace(dictNVP.Item("Hospital Service"), "'", "''") & "'"
                            sql = sql & ", hService = '" & tmpHospitalService & "'"

                            sql = sql & ", patient_type = '" & Replace(dictNVP.Item("Patient Type"), "'", "''") & "'"

                            '1/2/2007
                            'If Len(dictNVP.Item("01patient.pataddr1")) > 0 Then
                            sql = sql & STARupdateString("addr1", dictNVP.Item("01patient.pataddr1"))
                            'End If
                            'If Len(dictNVP.Item("01patient.pataddr2")) > 0 Then
                            sql = sql & STARupdateString("addr2", dictNVP.Item("01patient.pataddr2"))
                            'End If
                            'If Len(dictNVP.Item("01patient.patcity")) > 0 Then
                            sql = sql & STARupdateString("city", dictNVP.Item("01patient.patcity"))
                            'End If
                            'If Len(dictNVP.Item("01patient.patstate")) > 0 Then
                            sql = sql & STARupdateString("state", dictNVP.Item("01patient.patstate"))
                            'End If
                            '1/2/2007 - end

                            '5/10/2005
                            'If Len(dictNVP.Item("01patient.patzip")) > 0 Then
                            sql = sql & STARupdateString("zip", dictNVP.Item("01patient.patzip"))
                            'End If
                            '20090803==============================================================
                            If primaryPhone <> "" Then
                                sql = sql & ", phone = '" & primaryPhone & "'"
                            End If

                            If altPhone <> "" Then
                                sql = sql & ", altPhone = '" & altPhone & "'"
                            End If

                            If businessPhone <> "" Then
                                sql = sql & ", businessPhone = '" & businessPhone & "'"
                            End If
                            '=====================================================================
                            '8/14/2003 added dcDisp and county
                            'If Len(dictNVP.Item("Discharge Disposition")) > 0 Then
                            sql = sql & STARupdateString("dcDisp", dictNVP.Item("Discharge Disposition"))
                            'End If

                            'If Len(dictNVP.Item("Patient County/Parish Code")) > 0 Then
                            sql = sql & STARupdateString("county", dictNVP.Item("Patient County/Parish Code"))
                            'End If

                            '3/20/2002 add code for gender, sex, marital status and intake facility
                            If Len(dictNVP("race code STAR")) = 1 Then
                                sql = sql & ", race = '" & dictNVP("race code STAR") & "'"
                            End If
                            'If Len(dictNVP.Item("Patient Sex")) > 0 Then
                            sql = sql & STARupdateString("gender", dictNVP.Item("Patient Sex"))

                            'End If

                            If Len(dictNVP("Patient Marital Status")) = 1 Then
                                sql = sql & ", MaritalStatus = '" & dictNVP("Patient Marital Status") & "'"
                            ElseIf dictNVP("Patient Marital Status") = """""" Or dictNVP("Patient Marital Status") = "" Then
                                sql = sql & ", MaritalStatus = NULL "
                            End If

                            '12/07/2004 start==============================================================================
                            If dictNVP.Item("AROFieldName") = "ARO" Then
                                sql = sql & " ,ARO = '" & Replace(dictNVP("ARODataField"), "'", "''") & "'"
                            ElseIf dictNVP.Item("AROFieldName") = """""" Or dictNVP.Item("AROFieldName") = "" Then
                                sql = sql & " ,ARO = NULL "
                            End If

                            'If Len(dictNVP.Item("AdvDirective")) >= 6 Then
                            'sql = sql & " ,AdvanceDir = '" & Left$(dictNVP.Item("AdvDirective"), 6) & "'"
                            sql = sql & STARupdateString("AdvanceDir", Left$(dictNVP.Item("AdvDirective"), 6))
                            'End If

                            'If Len(dictNVP.Item("Allergy Description")) > 0 Then
                            sql = sql & STARupdateString("allergies", dictNVP.Item("Allergy Description"))
                            'End If

                            'If Len(dictNVP.Item("Referral Source ID")) > 0 Then
                            sql = sql & STARupdateString("referralsourceID", dictNVP.Item("Referral Source ID"))
                            'End If

                            'If Len(dictNVP.Item("JHHS mrnum")) > 0 Then
                            sql = sql & STARupdateString("JHHS_mrnum", dictNVP.Item("JHHS mrnum"))
                            'End If

                            '12/07/2004 end================================================================================

                            '11/17/2004 - added religion
                            If Len(dictNVP.Item("Patient Religion")) > 0 Then
                                sql = sql & ", religionID = '" & Replace(dictNVP.Item("Patient Religion"), "'", "''") & "'"
                            ElseIf dictNVP.Item("Patient Religion") = """""" Or dictNVP.Item("Patient Religion") = "" Then
                                sql = sql & ", religionID = NULL "
                            End If

                            sql = sql & ", intakeFacility = " & intIntakeFacility
                            '3/20/2002 - end

                            '11/08/2001 - changed critera from not < 6
                            '12/11/2002 removed length check
                            If IsNumeric(dictNVP.Item("Referring.patPhysNum")) Then
                                sql = sql & ", physRefer = " & dictNVP.Item("Referring.patPhysNum")
                            End If
                            '12/11/2002 removed length check
                            If IsNumeric(dictNVP.Item("Admitting.patPhysNum")) Then
                                sql = sql & ", physAdmit = " & dictNVP.Item("Admitting.patPhysNum")
                            End If

                            '20130531 - update physconsult fields with PV1_9
                            If dictNVP.Item("Consulting Physician") <> "" Then
                                sql = sql & ", physConsult = '" & Replace(dictNVP.Item("Consulting Physician ID"), "'", "''") & "' "
                            End If

                            '6/28/2005 added attending and consulting physicians
                            '=========================================================================================
                            If IsNumeric(dictNVP.Item("Attending Physician ID")) Then
                                sql = sql & ", physAttend = " & dictNVP.Item("Attending Physician ID")
                            End If


                            '2/8/2006
                            'handle consulting physicians
                            If boolConsultingExists Then
                                For j = 0 To UBound(consultingArray)
                                    If j < 10 Then
                                        If Trim(consultingArray(j)) <> "" Or IsDBNull(consultingArray(j)) Then
                                            sql = sql & ",consult" & Trim(Str(j + 1)) & " = " & (consultingArray(j))
                                        Else
                                            sql = sql & ",consult" & Trim(Str(j + 1)) & " = NULL"
                                        End If
                                    End If
                                Next
                                If UBound(consultingArray) < 9 Then
                                    For j = (UBound(consultingArray) + 1) To 9
                                        sql = sql & ",consult" & Trim(Str(j + 1)) & " = NULL"
                                    Next
                                End If
                            End If

                            '=========================================================================================

                            '20140220
                            If dictNVP.Item("Patient Class") = "P" Then
                                sql = sql & ", AdminDate = '" & ConvertDate(dictNVP.Item("ExpectedDate")) & "'"
                            Else
                                sql = sql & ", AdminDate = '" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "'"
                            End If

                            sql = sql & ", diagnosis = '" & Replace(dictNVP.Item("02diag.diagnosis"), "'", "''") & "'"

                            '===============================================================

                            Dim tempCOBPriority As String = ""
                            For i = 1 To gblInsCount
                                If i = 1 Then
                                    tempstr = ""
                                End If
                                If i > 1 Then
                                    tempstr = "_000" & i
                                End If
                                '20130429
                                star_plancode = Trim(Replace(dictNVP("iplancode2" & tempstr), "'", "''")) & Trim(Replace(dictNVP("iplancode" & tempstr), "'", "''"))


                                If dictNVP.Item("COBPriority" & tempstr) = "1" Then
                                    'sql = sql & ", primaryIns = '" & Replace(dictNVP.Item("iplancode" & tempstr), "'", "''") & "' "
                                    tempCOBPriority = star_plancode
                                End If
                            Next
                            If tempCOBPriority <> "" Then
                                sql = sql & ", primaryIns = '" & tempCOBPriority & "' "
                            End If
                            '===============================================================



                            '12/16/2002 - use patient status instead of the above code.
                            'sql = sql & ", status = '" & Replace(dictNVP.Item("Patient Status Code"), "'", "''") & "'"
                            '20130617 - use visitStatus Variable
                            sql = sql & ", status = '" & visitStatus & "'"

                            sql = sql & ", feedstarted = 1"
                            '6/26/2002: set the preAdmit flag to false
                            sql = sql & ", preAdmit = 0"

                            '5/9/2005
                            If strPV1_19 <> "NOT SENT" Then
                                sql = sql & ", PV1_19 = '" & strPV1_19 & "'"
                            End If



                            sql = sql & " Where epnum = " & gblEPNum
                            'Debug.Print "update episode: " & sql
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()
                        End If 'If Not boolRecordExixts Then

                        '3/23/2006
                        boolRecordExists = True

                        '12/12/2002===============================================================
                        'procedure to add entry in the new 08Accidents table for A01 and A04 or update if entry exists
                        '12/16/2002 - added code for accident date and time
                        If boolRecordExists Then

                            If Replace(dictNVP.Item("panum"), "'", "''") <> "" And Not boolAccidentExists Then 'add it

                                sql = "INSERT [08Accidents] (panum, accmemo1, accmemo2, accdate, updated) "
                                sql = sql & "VALUES ("
                                sql = sql & "'" & gblORIGINAL_PA_NUMBER & "', "
                                sql = sql & "'" & Replace(dictNVP.Item("03insurer.accmemo1"), "'", "''") & "', "
                                sql = sql & "'" & Replace(dictNVP.Item("03insurer.accmemo2"), "'", "''") & "', "
                                '20140811=========================================================================
                                If Len(dictNVP.Item("Accident Date/Time")) > 2 Then
                                    sql = sql & "'" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "', "
                                Else
                                    sql = sql & "NULL, "
                                End If
                                '==================================================================================
                                sql = sql & "'" & DateTime.Now & "') "
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()

                            Else ' update the accident record
                                sql = "UPDATE [08accidents] "
                                sql = sql & "SET updated = '" & DateTime.Now & "'"
                                sql = sql & ", accmemo1 = '" & Replace(dictNVP.Item("03insurer.accmemo1"), "'", "''") & "'"
                                sql = sql & ", accmemo2 = '" & Replace(dictNVP.Item("03insurer.accmemo2"), "'", "''") & "'"

                                'sql = sql & ", accdate = '" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "'"
                                '20140811=========================================================================
                                If Len(dictNVP.Item("Accident Date/Time")) > 2 Then
                                    sql = sql & ", accdate = '" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "'"
                                Else
                                    sql = sql & ", accdate = NULL "
                                End If
                                '==================================================================================
                                sql = sql & " Where panum = '" & dictNVP.Item("panum") & "'"
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()
                            End If
                        End If
                        '======================================================================================================
                    Case "A05"
                        'we have a preregistration form so add entry to episode table
                        If Not boolRecordExists Then
                            '12/02/2002 added hService and patient_type after dob
                            '12/06/2002 - added patient height and weight before race
                            '12/07/2004 processing
                            '3/14/2007 - handle onset date after aro
                            '20121206 - add corpNo
                            '20130502 - added star_region after onsetdate
                            '20130619 - add class after corpNO
                            '20130620 - add STAR_class after class to capture the raw clas value
                            sql = "INSERT into [001Episode] (mrnum, corpNo, class, STAR_class, aro, onsetDate, star_region, allergies, jhhs_mrnum, advanceDir, ReferralSourceID, "
                            sql = sql & "panum, lname, fname, mname, socsec, dob, hService, patient_type, "
                            '20090803 - added altPhone and BusinessPhone after phone
                            sql = sql & "addr1, addr2, city, state, zip, phone, altPhone, businessPhone, physRefer, physAdmit, physAttend, physConsult, prevAdmit, "
                            'sql = sql & "addr1, addr2, city, state, zip, phone, altPhone, businessPhone, physRefer, physAdmit, physAttend, physConsult, adminDate, "
                            '3/20/2002
                            sql = sql & "race, gender, MaritalStatus, IntakeFacility, "
                            '3/20/2002 end
                            '6/26/2002: set the preAdmit flag to true

                            '11/27/2001 insert into prediagnosis field for A05
                            '6/26/2002: set the preAdmit flag to true
                            '8/14/2003 added dsDisp and county before status

                            '2'8'2006
                            sql = sql & "consult1, consult2, consult3, consult4, consult5, "
                            '3/15/2006 - added PCPNo after consult10
                            sql = sql & "consult6, consult7, consult8, consult9, consult10, PCPNo, "

                            sql = sql & "preDiagnosis, primaryIns, dcDisp, county, status, created, feedstarted, active, preAdmit) "

                            '9/4/2003 - using isnumeric to check mrnum
                            If IsNumeric(dictNVP("mrnum")) Then
                                sql = sql & "VALUES (" & dictNVP("mrnum") & ", "
                            Else
                                sql = sql & "VALUES (" & "0, "
                            End If

                            '20121206 - add corpNo
                            sql = sql & gblCorporateNumber & ", "

                            '20130619 - add class from Patient Class with translation
                            sql = sql & "'" & xlateClass(dictNVP) & "', "

                            '20130620 - add class from Patient Class raw data
                            sql = sql & "'" & dictNVP("Patient Class") & "', "

                            '12/07/2004
                            sql = sql & "'" & strAro & "', "

                            '3/14/2007
                            '20141022 - add anset processing back in.
                            If Len(dictNVP("OnsetDate")) > 7 Then
                                sql = sql & "'" & ConvertDate(dictNVP("OnsetDate")) & "', " '20140906
                            Else
                                sql = sql & "NULL" & ", "
                            End If

                            sql = sql & "'" & star_region & "', "
                            sql = sql & "'" & strAllergies & "', "
                            sql = sql & "'" & strJHHS_mrnum & "', "
                            sql = sql & "'" & strAdvanceDir & "', "
                            sql = sql & "'" & strReferralSourceID & "', "

                            sql = sql & "'" & Replace(dictNVP.Item("panum"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patlast"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patfirst"), "'", "''") & "',"
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patmi"), "'", "''") & "', "
                            '04/17/2007 - added length check
                            If Len(dictNVP.Item("01patient.patSS")) >= 9 Then
                                sql = sql & "'" & Replace(dictNVP.Item("01patient.patSS"), "'", "''") & "', "
                            Else
                                sql = sql & "NULL, "
                            End If

                            sql = sql & "'" & ConvertDate(dictNVP.Item("01patient.DOB")) & "', "
                            '12/02/2002 added hService, status and patient_type

                            '20140213
                            'sql = sql & "'" & Replace(dictNVP.Item("Hospital Service"), "'", "''") & "', "
                            sql = sql & "'" & tmpHospitalService & "', "


                            sql = sql & "'" & Replace(dictNVP.Item("Patient Type"), "'", "''") & "', "

                            sql = sql & "'" & Replace(dictNVP.Item("01patient.pataddr1"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.pataddr2"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patcity"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patstate"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("01patient.patzip"), "'", "''") & "', "
                            '20090803 add processing for altphone and businessPhone
                            If primaryPhone <> "" Then
                                sql = sql & "'" & primaryPhone & "', "
                            Else
                                sql = sql & "NULL, "
                            End If
                            If altPhone <> "" Then
                                sql = sql & "'" & altPhone & "', "
                            Else
                                sql = sql & "NULL, "
                            End If
                            If businessPhone <> "" Then
                                sql = sql & "'" & businessPhone & "', "
                            Else
                                sql = sql & "NULL, "
                            End If
                            '12/11/2002 removed length check
                            If IsNumeric(dictNVP.Item("Referring.patPhysNum")) Then
                                sql = sql & dictNVP.Item("Referring.patPhysNum") & ", "
                            Else
                                sql = sql & "0, "
                            End If
                            '12/11/2002 removed length check
                            If IsNumeric(dictNVP.Item("Admitting.patPhysNum")) Then
                                sql = sql & dictNVP.Item("Admitting.patPhysNum") & ", "
                            Else
                                sql = sql & "0, "
                            End If

                            '6/28/2005 added attending and consulting physicians
                            '=============================================================================
                            If IsNumeric(dictNVP.Item("Attending Physician ID")) Then
                                sql = sql & dictNVP.Item("Attending Physician ID") & ", "
                            Else
                                sql = sql & "0, "
                            End If

                            ''If IsNumeric(dictNVP.Item("Consulting Physician ID")) Then
                            ''sql = sql & dictNVP.Item("Consulting Physician ID") & ", "
                            ''Else
                            'sql = sql & "0, "
                            ''End If

                            '20130531
                            sql = sql & "'" & Replace(dictNVP.Item("Consulting Physician ID"), "'", "''") & "', "

                            '============================================================================

                            '20140220
                            If dictNVP.Item("Patient Class") = "P" Then
                                sql = sql & "'" & ConvertDate(dictNVP.Item("ExpectedDate")) & "', "
                            Else
                                sql = sql & "'" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "', "
                            End If

                            '20170110 - Always Update Admission Date on Pre-Reg. If Admission Date is missing use expected date.  
                            'On an A08 if class is "P" update Admindate with expected date else if patient is upgrading update Admindate with admission date.

                            'If dictNVP.Item("01visit.AdminDate") = "" Then
                            'sql = sql & "'" & ConvertDate(dictNVP.Item("ExpectedDate")) & "', "
                            'Else
                            'sql = sql & "'" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "', "
                            'End If

                            '3/20/2002 added race, sex, maritalStatus and intake facility
                            'sql = sql & dictNVP.Item("01patient.patSS") & ", "


                            If dictNVP("race code STAR") <> "" Then
                                sql = sql & "'" & dictNVP("race code STAR") & "', "
                            Else
                                sql = sql & "'-', "
                            End If
                            If dictNVP("Patient Sex") <> "" Then
                                sql = sql & "'" & dictNVP("Patient Sex") & "', "
                            Else
                                sql = sql & "'-', "
                            End If

                            If dictNVP("Patient Marital Status") <> "" Then
                                sql = sql & "'" & dictNVP("Patient Marital Status") & "', "
                            Else
                                sql = sql & "'-', "
                            End If

                            sql = sql & intIntakeFacility & ", "
                            '3/20/2002 - end

                            '2/8/2006
                            If boolConsultingExists Then
                                For j = 0 To UBound(consultingArray)
                                    If j < 10 Then
                                        insertNumber(consultingArray(j))
                                    End If
                                Next
                                If UBound(consultingArray) < 9 Then
                                    For j = (UBound(consultingArray) + 1) To 9
                                        sql = sql & "NULL,"
                                    Next
                                End If
                            Else
                                sql = sql & "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,"
                            End If

                            '3/15/2006 added PCPNo (PD1_4_1)
                            insertNumber(dictNVP.Item("PCPNo"))

                            sql = sql & "'" & Replace(dictNVP.Item("02diag.diagnosis"), "'", "''") & "', "

                            '===============================================================

                            Dim tempPlanCode As String = ""
                            For i = 1 To gblInsCount

                                If i = 1 Then
                                    tempstr = ""
                                End If
                                If i > 1 Then
                                    tempstr = "_000" & i
                                End If
                                '20130429
                                star_plancode = Trim(Replace(dictNVP("iplancode2" & tempstr), "'", "''")) & Trim(Replace(dictNVP("iplancode" & tempstr), "'", "''"))


                                If dictNVP.Item("COBPriority" & tempstr) = "1" Then
                                    'sql = sql & "'" & Replace(dictNVP.Item("iplancode" & tempstr), "'", "''") & "', "
                                    tempPlanCode = star_plancode

                                End If
                            Next
                            If tempPlanCode <> "" Then
                                sql = sql & "'" & tempPlanCode & "', "
                            Else
                                sql = sql & "'NNN', "
                            End If
                            '===============================================================

                            '11/08/2001 don't update status on A05 but I will capture the hospital service entry anyway
                            'for testing.
                            '12/12/2002 - update status field for Patient Status Code

                            '8/14/2003 added dcDisp and county
                            sql = sql & "'" & Replace(dictNVP.Item("Discharge Disposition"), "'", "''") & "', "
                            sql = sql & "'" & Replace(dictNVP.Item("Patient County/Parish Code"), "'", "''") & "', "

                            '20130617
                            'If dictNVP.Item("Patient Status Code") <> "" Then
                            'sql = sql & "'" & Replace(dictNVP.Item("Patient Status Code"), "'", "''") & "', "
                            'Else
                            'sql = sql & "'None', "
                            'End If
                            sql = sql & "'" & visitStatus & "', "

                            '20130524 - removed patnum from insert process
                            '12/06/2001
                            '04/17/2007 added length check
                            'If Len(dictNVP.Item("01patient.patSS")) > 9 Then
                            'sql = sql & dictNVP.Item("01patient.patSS") & ", "
                            'Else
                            'sql = sql & "0, "
                            'End If

                            '6/26/2002: set the preAdmit flag to true
                            sql = sql & "'" & DateTime.Now & "', 1, 1,1)"
                            'Debug.Print sql
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()

                            'get the last epnum added
                            sql = "SELECT max(epnum) AS [lastNumber] from [001Episode]"
                            objCommand.CommandText = sql
                            myConnection.Open()
                            dataReader = objCommand.ExecuteReader()

                            If dataReader.HasRows Then
                                dataReader.Read()
                                gblEPNum = dataReader.Item("lastNumber")
                            End If
                            myConnection.Close()
                            dataReader.Close()
                        End If ' if not boolRecordExixts
                            '3/23/2006
                            boolRecordExists = True

                            '12/12/2002===============================================================
                            'procedure to add entry in the new 08Accidents table for A05 or update if entry exists
                            '12/16/2002 - added code for Accident date and time
                            If boolRecordExists Then

                                If Replace(dictNVP.Item("panum"), "'", "''") <> "" And Not boolAccidentExists Then 'add it

                                    sql = "INSERT [08Accidents] (panum, accmemo1, accmemo2, accdate, updated) "
                                    sql = sql & "VALUES ("
                                    sql = sql & "'" & gblORIGINAL_PA_NUMBER & "', "
                                    sql = sql & "'" & Replace(dictNVP.Item("03insurer.accmemo1"), "'", "''") & "', "
                                    sql = sql & "'" & Replace(dictNVP.Item("03insurer.accmemo2"), "'", "''") & "', "

                                    '20140811=========================================================================
                                    If Len(dictNVP.Item("Accident Date/Time")) > 2 Then
                                        sql = sql & "'" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "', "
                                    Else
                                        sql = sql & "NULL, "
                                    End If
                                    '==================================================================================
                                    sql = sql & "'" & DateTime.Now & "') "
                                    updatecommand.CommandText = sql
                                    myConnection.Open()
                                    updatecommand.ExecuteNonQuery()
                                    myConnection.Close()


                                Else ' update the accident record
                                    sql = "UPDATE [08accidents] "
                                    sql = sql & "SET updated = '" & DateTime.Now & "'"
                                    sql = sql & ", accmemo1 = '" & Replace(dictNVP.Item("03insurer.accmemo1"), "'", "''") & "'"
                                    sql = sql & ", accmemo2 = '" & Replace(dictNVP.Item("03insurer.accmemo2"), "'", "''") & "'"

                                    'sql = sql & ", accdate = " & ConvertDate(dictNVP.Item("Accident Date/Time")) & ""
                                    '20140811=========================================================================
                                    If Len(dictNVP.Item("Accident Date/Time")) > 2 Then
                                        sql = sql & ", accdate = '" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "'"
                                    Else
                                        sql = sql & ", accdate = NULL "
                                    End If
                                    '==================================================================================
                                    sql = sql & " Where panum = '" & dictNVP.Item("panum") & "'"
                                    updatecommand.CommandText = sql
                                    myConnection.Open()
                                    updatecommand.ExecuteNonQuery()
                                    myConnection.Close()
                                End If

                            End If
                            '12/12/2002===============================================================

                            '======================================================================================================
                    Case "A02", "A12" '20150506 - added A12 processing. Same as A02
                            If boolRecordExists Then
                                If Len(strRoomBed) > 0 Then
                                    sql = "UPDATE [001episode] "
                                    sql = sql & "SET modified = '" & DateTime.Now & "'"
                                    sql = sql & ", room = '" & strRoomBed & "'"

                                    '20130619 - update class from Patient Class if not blank with translation
                                    If dictNVP("Patient Class") <> "" Then
                                        sql = sql & ", class = '" & xlateClass(dictNVP) & "' "
                                    End If

                                    '20130620 - update class from Patient Class if not blank raw data
                                    If dictNVP("Patient Class") <> "" Then
                                        sql = sql & ", STAR_class = '" & dictNVP("Patient Class") & "' "
                                    End If


                                    sql = sql & " Where epnum = " & gblEPNum
                                    updatecommand.CommandText = sql
                                    myConnection.Open()
                                    updatecommand.ExecuteNonQuery()
                                    myConnection.Close()
                                End If 'If Len(strRoomBed) > 0
                            End If 'if boolRecordExists for A02
                            '======================================================================================================
                    Case "A03"
                            '20140310 - do same as A03 for A11
                            '20150330 - if patient class is O then don't set dcdate. Separated out A11 as separate process.
                            '20160830 - do not discharge is Patient Type = COQ.  Bypass due to nightly discharge.
                            If boolRecordExists And dictNVP.Item("Patient Type") <> "COQ" Then

                                sql = "UPDATE [001episode] "
                                sql = sql & "SET modified = '" & DateTime.Now & "'"
                                If Len(dictNVP("01visit.dischdate")) > 7 And xlateClass(dictNVP) <> "O" Then
                                    sql = sql & ", DCDate = '" & ConvertDate(dictNVP.Item("01visit.dischdate")) & "'"
                                End If

                                '20130619 - update class from Patient Class if not blank with translation
                                If dictNVP("Patient Class") <> "" Then
                                    sql = sql & ", class = '" & xlateClass(dictNVP) & "' "
                                End If

                                '20130620 - update class from Patient Class if not blank with translation
                                If dictNVP("Patient Class") <> "" Then
                                    sql = sql & ", STAR_class = '" & dictNVP("Patient Class") & "' "
                                End If

                                '12/02/2002 - added status
                                '20130617 - don't change status, look at discharged field set to 1.
                                'If Len(dictNVP.Item("Patient Status Code")) > 0 Then
                                'sql = sql & STARupdateString("status", dictNVP.Item("Patient Status Code"))
                                'End If

                                '20130620 - add status update for A03 again
                                sql = sql & ", status = '" & visitStatus & "' "

                                '8/14/2003 added dcDisp
                                'If Len(dictNVP.Item("Discharge Disposition")) > 0 Then
                                sql = sql & STARupdateString("dcDisp", dictNVP.Item("Discharge Disposition"))
                                'End If
                                sql = sql & ", discharged = 1"
                                sql = sql & " Where epnum = " & gblEPNum
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()

                            End If 'if boolRecordExists for A02

                    Case "A11"
                            '20140310 - do same as A03 for A11
                            '20150330 - if patient class is O then don't set dcdate. Don't implement this for A11.
                            If boolRecordExists Then

                                sql = "UPDATE [001episode] "
                                sql = sql & "SET modified = '" & DateTime.Now & "'"
                                If Len(dictNVP("01visit.dischdate")) > 7 Then
                                    sql = sql & ", DCDate = '" & ConvertDate(dictNVP.Item("01visit.dischdate")) & "'"
                                End If

                                '20130619 - update class from Patient Class if not blank with translation
                                If dictNVP("Patient Class") <> "" Then
                                    sql = sql & ", class = '" & xlateClass(dictNVP) & "' "
                                End If

                                '20130620 - update class from Patient Class if not blank with translation
                                If dictNVP("Patient Class") <> "" Then
                                    sql = sql & ", STAR_class = '" & dictNVP("Patient Class") & "' "
                                End If

                                '12/02/2002 - added status
                                '20130617 - don't change status, look at discharged field set to 1.
                                'If Len(dictNVP.Item("Patient Status Code")) > 0 Then
                                'sql = sql & STARupdateString("status", dictNVP.Item("Patient Status Code"))
                                'End If

                                '20130620 - add status update for A03 again
                                sql = sql & ", status = '" & visitStatus & "' "

                                '8/14/2003 added dcDisp
                                'If Len(dictNVP.Item("Discharge Disposition")) > 0 Then
                                sql = sql & STARupdateString("dcDisp", dictNVP.Item("Discharge Disposition"))
                                'End If
                                sql = sql & ", discharged = 1"
                                sql = sql & " Where epnum = " & gblEPNum
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()

                            End If 'if boolRecordExists for A02

                    Case "A07" '20140415 add process for Leave of Absence (LOA)
                            If boolRecordExists Then
                                sql = "UPDATE [001episode] "
                                sql = sql & "SET modified = '" & DateTime.Now & "'"
                                sql = sql & ", room = 'LOA' "
                                sql = sql & " Where epnum = " & gblEPNum
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()
                            End If
                            '======================================================================================================
                    Case "A08"
                            '12/02/2002 - added hservice and status
                            '11/17/2004 - added religion
                            '3/14/2007 - handle onsetdate after mrnum
                            '20140912 - added star_region to A08.
                            If boolRecordExists Then
                                sql = "UPDATE [001episode] "
                                sql = sql & "SET modified = '" & DateTime.Now & "'"

                                '9/4/2003 - using isnumeric to check mrnum
                                If IsNumeric(dictNVP("mrnum")) Then
                                    sql = sql & ", mrnum = " & dictNVP.Item("mrnum") & " "
                                End If

                                '20140912 - added star_region
                                If star_region <> "" Then
                                    sql = sql & ", star_region = '" & star_region & "'"
                                End If

                                '20140213 - added to A08
                                If tmpHospitalService <> "" Then
                                    sql = sql & ", hService = '" & tmpHospitalService & "'"
                                End If

                                '20130619 - update class from Patient Class if not blank with translation
                                If dictNVP("Patient Class") <> "" Then
                                    sql = sql & ", class = '" & xlateClass(dictNVP) & "' "
                                End If

                                '20130620 - update class from Patient Class if not blank raw data
                                If dictNVP("Patient Class") <> "" Then
                                    sql = sql & ", STAR_class = '" & dictNVP("Patient Class") & "' "
                                End If

                                '3/14/2007
                                '20141022 - add anset processing back in.
                                If Len(dictNVP("OnsetDate")) > 7 Then
                                    sql = sql & ", onsetDate = '" & ConvertDate(dictNVP("OnsetDate")) & "'" '20140906

                                Else
                                    sql = sql & ", onsetDate = NULL"
                                End If

                                '3/15/2006 added PCPNo (PD1_4_1)
                                If IsNumeric(dictNVP.Item("PCPNo")) Then
                                    sql = sql & ", PCPNo = " & dictNVP.Item("PCPNo")
                                ElseIf dictNVP.Item("PCPNo") = """""" Or dictNVP.Item("PCPNo") = "" Then
                                    sql = sql & ", PCPNo = NULL "
                                End If
                                '11/17/2004 - religion added
                                If Len(dictNVP.Item("Patient Religion")) > 0 Then
                                    sql = sql & ", religionID = '" & Replace(dictNVP.Item("Patient Religion"), "'", "''") & "'"
                                ElseIf dictNVP.Item("Patient Religion") = """""" Or dictNVP.Item("Patient Religion") = "" Then
                                    sql = sql & ", religionID = NULL "
                                End If

                                'sql = sql & ", panum = '" & Replace(dictNVP.Item("panum"), "'", "''") & "'"

                                '20160415 - do not update department if patient COQ.
                                If departmentCode <> "" Then
                                    '20130522 - added department
                                    sql = sql & STARupdateString("department", departmentCode)
                                End If
                                '20130401 - handle double quotes to zero field
                                '=========================================================================================================
                                'If Len(dictNVP.Item("01patient.patlast")) > 0 Then
                                sql = sql & STARupdateString("lname", dictNVP.Item("01patient.patlast"))
                                'End If

                                'sql = sql & ", fname = '" & Replace(dictNVP.Item("01patient.patfirst"), "'", "''") & "'"
                                'If Len(dictNVP.Item("01patient.patfirst")) > 0 Then
                                sql = sql & STARupdateString("fname", dictNVP.Item("01patient.patfirst"))
                                'End If

                                'sql = sql & ", mname = '" & Replace(dictNVP.Item("01patient.patmi"), "'", "''") & "'"
                                'If Len(dictNVP.Item("01patient.patmi")) > 0 Then
                                sql = sql & STARupdateString("mname", dictNVP.Item("01patient.patmi"))
                                'End If
                                '=========================================================================================================

                                '5/10/2005
                                '4/17/2007 added length check
                                If Len(dictNVP.Item("01patient.patSS")) >= 9 Then
                                    'sql = sql & ", SocSec = '" & Replace(dictNVP.Item("01patient.patSS"), "'", "''") & "'"
                                    'If Len(dictNVP.Item("01patient.patSS")) > 0 Then
                                    sql = sql & STARupdateString("SocSec", dictNVP.Item("01patient.patSS"))
                                    'End If
                                End If

                                '12/02/2002 - added hservice, status and patient_type
                                'sql = sql & ", hService = '" & Replace(dictNVP.Item("Hospital Service"), "'", "''") & "'"

                                '12/07/2004 don't process Patient Status Code if blank
                                'If dictNVP.Item("Patient Status Code") <> "" Then
                                'sql = sql & ", status = '" & Replace(dictNVP.Item("Patient Status Code"), "'", "''") & "'"
                                'End If
                                'If Len(dictNVP.Item("Patient Status Code")) > 0 Then
                                '20130617 - don't process status on A08 because visitStatus is blank for an A08 record.
                                'sql = sql & STARupdateString("status", dictNVP.Item("Patient Status Code"))
                                'End If

                                'sql = sql & ", patient_type = '" & Replace(dictNVP.Item("Patient Type"), "'", "''") & "'"
                                'If Len(dictNVP.Item("Patient Type")) > 0 Then
                                sql = sql & STARupdateString("patient_type", dictNVP.Item("Patient Type"))
                                'End If

                                '8/14/2003 added dcDisp and county
                                'sql = sql & ", dcDisp = '" & Replace(dictNVP.Item("Discharge Disposition"), "'", "''") & "'"
                                If Len(dictNVP.Item("Discharge Disposition")) > 0 Then
                                    sql = sql & STARupdateString("dcDisp", dictNVP.Item("Discharge Disposition"))
                                End If

                                'sql = sql & ", county = '" & Replace(dictNVP.Item("Patient County/Parish Code"), "'", "''") & "'"
                                'If Len(dictNVP.Item("Patient County/Parish Code")) > 0 Then
                                sql = sql & STARupdateString("county", dictNVP.Item("Patient County/Parish Code"))
                                'End If

                                'sql = sql & ", DOB = '" & ConvertDate(dictNVP.Item("01patient.DOB")) & "'"
                                'If Len(dictNVP.Item("01patient.DOB")) > 0 Then
                                sql = sql & STARupdateString("DOB", dictNVP.Item("01patient.DOB"))
                                'End If
                                '1/2/2007
                                'If dictNVP.Item("01patient.pataddr1") <> "" Then
                                'sql = sql & ", addr1 = '" & Replace(dictNVP.Item("01patient.pataddr1"), "'", "''") & "'"
                                'End If
                                'If Len(dictNVP.Item("01patient.pataddr1")) > 0 Then
                                sql = sql & STARupdateString("addr1", dictNVP.Item("01patient.pataddr1"))
                                'End If

                                'If dictNVP.Item("01patient.pataddr2") <> "" Then
                                'sql = sql & ", addr2 = '" & Replace(dictNVP.Item("01patient.pataddr2"), "'", "''") & "'"
                                'End If
                                'If Len(dictNVP.Item("01patient.pataddr2")) > 0 Then
                                sql = sql & STARupdateString("addr2", dictNVP.Item("01patient.pataddr2"))
                                'End If

                                'If dictNVP.Item("01patient.patcity") <> "" Then
                                'sql = sql & ", city = '" & Replace(dictNVP.Item("01patient.patcity"), "'", "''") & "'"
                                'End If
                                'If Len(dictNVP.Item("01patient.patcity")) > 0 Then
                                sql = sql & STARupdateString("city", dictNVP.Item("01patient.patcity"))
                                'End If

                                'If dictNVP.Item("01patient.patstate") <> "" Then
                                'sql = sql & ", state = '" & Replace(dictNVP.Item("01patient.patstate"), "'", "''") & "'"
                                'End If
                                If Len(dictNVP.Item("01patient.patstate")) > 0 Then
                                    sql = sql & STARupdateString("state", dictNVP.Item("01patient.patstate"))
                                End If

                                '1/2/2007 - end


                                '5/10/2005
                                'If dictNVP.Item("01patient.patzip") <> "" Then
                                'sql = sql & ", zip = '" & Replace(dictNVP.Item("01patient.patzip"), "'", "''") & "'"
                                'End If
                                'If Len(dictNVP.Item("01patient.patzip")) > 0 Then
                                sql = sql & STARupdateString("zip", dictNVP.Item("01patient.patzip"))
                                'End If
                                '20090803==============================================================
                                'If primaryPhone <> "" Then
                                'sql = sql & ", phone = '" & primaryPhone & "'"
                                'End If
                                'If Len(dictNVP.Item("01patient.patPhone STAR")) > 0 Then
                                sql = sql & STARupdateString("phone", dictNVP.Item("01patient.patPhone STAR"))
                                'End If

                                If altPhone <> "" Then
                                    ' 20130401 sql = sql & ", altPhone = '" & altPhone & "'"
                                End If

                                'If businessPhone <> "" Then
                                'sql = sql & ", businessPhone = '" & businessPhone & "'"
                                'End If
                                'If Len(dictNVP.Item("Patient Business Phone STAR")) > 0 Then
                                sql = sql & STARupdateString("businessPhone", dictNVP.Item("Patient Business Phone STAR"))
                                'End If
                                '=====================================================================

                                '3/20/2002 add code for gender, sex, marital status and intake facility
                                If Len(dictNVP("race code STAR")) = 1 Then
                                    sql = sql & ", race = '" & dictNVP("race code STAR") & "'"
                                End If



                                'If Len(dictNVP("Patient Sex")) = 1 Then
                                'sql = sql & ", gender = '" & dictNVP("Patient Sex") & "'"
                                'End If
                                'If Len(dictNVP.Item("Patient Sex")) > 0 Then
                                sql = sql & STARupdateString("gender", dictNVP.Item("Patient Sex"))

                                'End If
                                If Len(dictNVP("Patient Marital Status")) = 1 Then
                                    sql = sql & ", MaritalStatus = '" & dictNVP("Patient Marital Status") & "'"
                                ElseIf dictNVP("Patient Marital Status") = """""" Or dictNVP("Patient Marital Status") = "" Then
                                    sql = sql & ", MaritalStatus = NULL "
                                End If

                                sql = sql & ", intakeFacility = " & intIntakeFacility
                                '3/20/2002 - end

                                '11/08/2001 added check to ensure len greater than 3 for update
                                '12/11/2002 removed length check
                                If IsNumeric(dictNVP.Item("Referring.patPhysNum")) Then
                                    sql = sql & ", physRefer = " & dictNVP.Item("Referring.patPhysNum")
                                ElseIf dictNVP.Item("Referring.patPhysNum") = """""" Or dictNVP.Item("Referring.patPhysNum") = "" Then
                                    sql = sql & ", physRefer = NULL "
                                End If
                                '12/11/2002 removed length check
                                If IsNumeric(dictNVP.Item("Admitting.patPhysNum")) Then
                                    sql = sql & ", physAdmit = " & dictNVP.Item("Admitting.patPhysNum")
                                ElseIf dictNVP.Item("Admitting.patPhysNum") = """""" Or dictNVP.Item("Admitting.patPhysNum") = "" Then
                                    sql = sql & ", physAdmit = NULL "
                                End If

                                '6/28/2005 added attending and consulting physicians
                                '=========================================================================================
                                If IsNumeric(dictNVP.Item("Attending Physician ID")) Then
                                    sql = sql & ", physAttend = " & dictNVP.Item("Attending Physician ID")
                                ElseIf dictNVP.Item("Attending Physician ID") = """""" Or dictNVP.Item("Attending Physician ID") = "" Then
                                    sql = sql & ", physAttend = NULL "
                                End If

                                '20130531 - update physconsult fields with PV1_9
                                If dictNVP.Item("Consulting Physician ID") <> "" Then
                                    sql = sql & ", physConsult = '" & Replace(dictNVP.Item("Consulting Physician ID"), "'", "''") & "' "
                                End If

                                '=========================================================================================
                                '1/19/2005 - don't update admission date on A08.
                                '20140916 - add update of Admindate on A08
                                '20150909 - remove admission date processing on A08 again!
                                '201590914 - use gblBoolUpdateAdmDate to determine if we should update the admission date on an A08
                            If gblBoolUpdateAdmDate Then
                                sql = sql & ", AdminDate = '" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "' "
                            End If
                            '20170110 - Always Update Admission Date on Pre-Reg.  If Admission Date is missing use expected date.    
                            'On an A08 if class is "P" update Admindate with expected date else if patient is upgrading update Admindate with admission date.
                            'If dictNVP.Item("Patient Class") = "P" Then
                            'If dictNVP.Item("01visit.AdminDate") = "" Then
                            'sql = sql & ", AdminDate = '" & ConvertDate(dictNVP.Item("ExpectedDate")) & "' "
                            'Else
                            'sql = sql & ", AdminDate = '" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "' "
                            'End If
                            'ElseIf gblBoolUpdateAdmDate Then
                            'sql = sql & ", AdminDate = '" & ConvertDate(dictNVP.Item("01visit.AdminDate")) & "' "
                            'End If
                            '=========================================================================================

                            'sql = sql & ", diagnosis = '" & Replace(dictNVP.Item("02diag.diagnosis"), "'", "''") & "'"
                            'If Len(dictNVP.Item("02diag.diagnosis")) > 0 Then
                            sql = sql & STARupdateString("diagnosis", dictNVP.Item("02diag.diagnosis"))
                            'End If

                            Dim tempCOBPriority As String = ""
                            For i = 1 To gblInsCount
                                If i = 1 Then
                                    tempstr = ""
                                End If
                                If i > 1 Then
                                    tempstr = "_000" & i
                                End If
                                '20130429
                                star_plancode = Trim(Replace(dictNVP("iplancode2" & tempstr), "'", "''")) & Trim(Replace(dictNVP("iplancode" & tempstr), "'", "''"))


                                If dictNVP.Item("COBPriority" & tempstr) = "1" Then
                                    'sql = sql & ", primaryIns = '" & Replace(dictNVP.Item("iplancode" & tempstr), "'", "''") & "' "
                                    tempCOBPriority = star_plancode
                                End If
                            Next
                            If tempCOBPriority <> "" Then
                                sql = sql & ", primaryIns = '" & tempCOBPriority & "' "
                            End If
                            '===============================================================
                            '12/07/2004 start==============================================================================
                            If dictNVP.Item("AROFieldName") = "ARO" Then
                                sql = sql & " ,ARO = '" & Replace(dictNVP("ARODataField"), "'", "''") & "'"
                            ElseIf dictNVP.Item("AROFieldName") = """""" Or dictNVP.Item("AROFieldName") = "" Then
                                sql = sql & " ,ARO = NULL "
                            End If

                            'If Len(dictNVP.Item("AdvDirective")) >= 6 Then
                            'sql = sql & " ,AdvanceDir = '" & Left$(dictNVP.Item("AdvDirective"), 6) & "'"
                            sql = sql & STARupdateString("AdvanceDir", Left$(dictNVP.Item("AdvDirective"), 6))
                            'End If

                            'If dictNVP("Allergy Description") <> "" Then
                            'sql = sql & " ,allergies = '" & Replace(dictNVP("Allergy Description"), "'", "''") & "'"
                            'End If
                            'If Len(dictNVP.Item("Allergy Description")) > 0 Then
                            sql = sql & STARupdateString("allergies", dictNVP.Item("Allergy Description"))
                            'End If

                            'If dictNVP("Referral Source ID") <> "" Then
                            'sql = sql & " ,referralsourceID = '" & Replace(dictNVP("Referral Source ID"), "'", "''") & "'"
                            'End If
                            'If Len(dictNVP.Item("Referral Source ID")) > 0 Then
                            sql = sql & STARupdateString("referralsourceID", dictNVP.Item("Referral Source ID"))
                            'End If

                            'If dictNVP("JHHS mrnum") <> "" Then
                            'sql = sql & " ,JHHS_mrnum = '" & Replace(dictNVP("JHHS mrnum"), "'", "''") & "'"
                            'End If
                            'If Len(dictNVP.Item("JHHS mrnum")) > 0 Then
                            sql = sql & STARupdateString("JHHS_mrnum", dictNVP.Item("JHHS mrnum"))
                            'End If

                            '12/07/2004 end================================================================================

                            '8/28/2001 add processing for A03
                            'If (dictNVP.Item("Event Type Code") = "A03") And UCase(dictNVP("Patient Class")) <> "O" Then '20150410
                            'sql = sql & ", DCDate = '" & ConvertDate(dictNVP.Item("01visit.dischdate")) & "'"
                            'sql = sql & ", discharged = 1"
                            'End If

                            '5/9/2005
                            If strPV1_19 <> "NOT SENT" Then
                                sql = sql & ", PV1_19 = '" & strPV1_19 & "'"
                            End If

                            '2/8/2006
                            'handle consulting physicians
                            If boolConsultingExists Then
                                For j = 0 To UBound(consultingArray)
                                    If j < 10 Then
                                        If IsNumeric(consultingArray(j)) Then '  6/7/2006 - check numeric properties
                                            If Trim(consultingArray(j)) <> "" Or IsDBNull(consultingArray(j)) Then
                                                sql = sql & ",consult" & Trim(Str(j + 1)) & " = " & (consultingArray(j))
                                            Else
                                                sql = sql & ",consult" & Trim(Str(j + 1)) & " = NULL"
                                            End If
                                        End If
                                    End If
                                Next
                                If UBound(consultingArray) < 9 Then
                                    For j = (UBound(consultingArray) + 1) To 9
                                        sql = sql & ",consult" & Trim(Str(j + 1)) & " = NULL"
                                    Next
                                End If
                            End If

                            '20140310 - added status update to A08
                            sql = sql & ", status = '" & visitStatus & "' "
                            'sql = sql & ", corpNo = " & gblCorporateNumber & " "
                            sql = sql & ", feedstarted = 1"
                            sql = sql & " Where epnum = " & gblEPNum
                            'Debug.Print "A08: " & sql
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()

                        End If ' if boolRecordExixts

                        '12/12/2002===============================================================
                        'procedure to add entry in the new 08Accidents table for A08 or update if entry exists
                        '12/16/2002 - added code for Accident Date and time
                        If boolRecordExists Then

                            If Replace(dictNVP.Item("panum"), "'", "''") <> "" And Not boolAccidentExists Then 'add it


                                sql = "INSERT [08Accidents] (panum, accmemo1, accmemo2, accdate, updated) "
                                sql = sql & "VALUES ("
                                sql = sql & "'" & gblORIGINAL_PA_NUMBER & "', "
                                sql = sql & "'" & Replace(dictNVP.Item("03insurer.accmemo1"), "'", "''") & "', "
                                sql = sql & "'" & Replace(dictNVP.Item("03insurer.accmemo2"), "'", "''") & "', "

                                '20140811=========================================================================
                                If Len(dictNVP.Item("Accident Date/Time")) > 2 Then
                                    sql = sql & "'" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "', "
                                Else
                                    sql = sql & "NULL, "
                                End If
                                '==================================================================================
                                sql = sql & "'" & DateTime.Now & "') "
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()

                            Else ' update the accident record
                                sql = "UPDATE [08accidents] "
                                sql = sql & "SET updated = '" & DateTime.Now & "'"
                                sql = sql & ", accmemo1 = '" & Replace(dictNVP.Item("03insurer.accmemo1"), "'", "''") & "'"
                                sql = sql & ", accmemo2 = '" & Replace(dictNVP.Item("03insurer.accmemo2"), "'", "''") & "'"

                                'sql = sql & ", accdate = '" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "'"
                                '20140811=========================================================================
                                If Len(dictNVP.Item("Accident Date/Time")) > 2 Then
                                    sql = sql & ", accdate = '" & ConvertDate(dictNVP.Item("Accident Date/Time")) & "'"
                                Else
                                    sql = sql & ", accdate = NULL "
                                End If
                                '==================================================================================
                                sql = sql & " Where panum = '" & dictNVP.Item("panum") & "'"
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()
                            End If
                        End If

                        '12/12/2002===============================================================
                        '======================================================================================================
                    Case "A13" 'cancel discharge
                            If boolRecordExists Then
                                sql = "UPDATE [001episode] "
                                sql = sql & "SET modified = '" & DateTime.Now & "'"
                                sql = sql & ", DCDate = NULL"
                                '12/02/2002 - added status
                                'If Len(dictNVP.Item("Patient Status Code")) > 0 Then

                                '20130619 - changed to update the class field from Patient Class map field with translation
                                sql = sql & STARupdateString("class", xlateClass(dictNVP))

                                '20130620 - changed to update the class field from Patient Class map field raw data
                                sql = sql & STARupdateString("STAR_class", dictNVP("Patient Class"))

                                'End If
                                '8/14/2003 added dcDisp
                                'If Len(dictNVP.Item("Discharge Disposition")) > 0 Then
                                sql = sql & STARupdateString("dcDisp", dictNVP.Item("Discharge Disposition"))
                                'End If

                                '20130620 - process visitStatus on A13
                                If visitStatus <> "" Then
                                    sql = sql & ", status = '" & visitStatus & "' "
                                End If
                                sql = sql & ", discharged = 0"
                                sql = sql & " Where epnum = " & gblEPNum
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()
                            End If 'If boolRecordExists
                            '======================================================================================================
                    Case "A17"
                            'process to swap rooms for the two patients provided
                            Dim processA17 As Boolean = True
                            Dim strRoomBed1 As String = ""
                            Dim strRoomBed2 As String = ""
                            Dim panum1 As String = ""
                            Dim panum2 As String = ""
                            strRoomBed1 = Replace(dictNVP.Item("01visit.room"), "'", "''") & Replace(dictNVP.Item("01visit.bed"), "'", "''")
                            strRoomBed2 = Replace(dictNVP.Item("01visit.room_0002"), "'", "''") & Replace(dictNVP.Item("01visit.bed_0002"), "'", "''")
                            If (dictNVP.Item("panum") <> "") Then
                                panum1 = dictNVP.Item("panum")
                            Else
                                processA17 = False
                            End If

                            If (dictNVP.Item("panum_0002") <> "") Then
                                panum2 = dictNVP.Item("panum_0002")
                            Else
                                processA17 = False
                            End If
                            If processA17 Then
                                'process first panum
                                sql = "UPDATE [001episode] "
                                sql = sql & "SET modified = '" & DateTime.Now & "'"
                                sql = sql & ", room = '" & strRoomBed2 & "'"
                                sql = sql & " Where panum = '" & panum2 & "'"
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()
                                'process second panum
                                sql = "UPDATE [001episode] "
                                sql = sql & "SET modified = '" & DateTime.Now & "'"
                                sql = sql & ", room = '" & strRoomBed1 & "'"
                                sql = sql & " Where panum = '" & panum1 & "'"
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()

                                myConnection.Close()

                            End If

                            '======================================================================================================

                    Case "A24" ' 20140216 - merge corpNo and mrnum data

                End Select
            End If 'If boolContinueProcessing Then '20140213

        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "Update Episode Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

            Exit Sub
        End Try

    End Sub

    Public Sub checkIN1(ByVal dictNVP As Hashtable)
        Try
            'IN1 count
            Dim IN1Count As Integer
            Dim myEnumerator As IDictionaryEnumerator = dictNVP.GetEnumerator()
            IN1Count = 0

            myEnumerator.Reset()
            While myEnumerator.MoveNext()
                If Left$(myEnumerator.Key, 3) = "IN1" Then
                    gblInsCount = gblInsCount + 1
                    IN1Count = IN1Count + 1
                End If

            End While

        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "Insurer Enumeration Error (checkIN1)" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

            Exit Sub
        End Try

    End Sub

    Public Sub insertString(ByVal theString As String)
        '20130422 - modified to handle double quotes
        Try
            'If theString <> "" Then
            'sql = sql & "'" & Replace(theString, "'", "''") & " ', "
            'Else
            'sql = sql & "NULL, "
            'End If

            Select Case theString
                Case ""
                    sql = sql & "NULL, "
                Case """"""
                    sql = sql & "NULL, "
                Case Else
                    sql = sql & "'" & Replace(theString, "'", "''") & " ', "

            End Select


        Finally
        End Try
    End Sub

    Public Sub insertNumber(ByVal theString As String)
        '20130422 - modified to handle double quotes
        Select Case theString
            Case ""
                sql = sql & "NULL, "
            Case """"""
                sql = sql & "NULL, "
            Case Else
                sql = sql & "'" & Replace(theString, "'", "''") & " ', "

        End Select
    End Sub

    Public Function ConvertDate(ByVal datedata As String) As String
        'convert the hl7 date to a database date
        'hl7 in format: yyyymmdd or yyyymmddhhmm
        'returns now if the string is not in one
        'of the two formats.
        '

        '20130422 - code to handle double quotes
        Try
            If datedata = """""" Then
                ConvertDate = "NULL"
            Else

                Dim strYear As String = ""
                Dim strMonth As String = ""
                Dim strDay As String = ""
                Dim strHour As String = ""
                Dim strMinute As String = ""

                If Len(Trim(datedata)) = 8 Then
                    strYear = Mid$(datedata, 1, 4)
                    strMonth = Mid$(datedata, 5, 2)
                    strDay = Mid$(datedata, 7, 2)
                    ConvertDate = strMonth & "/" & strDay & "/" & strYear

                ElseIf Len(Trim(datedata)) >= 12 Then
                    strYear = Mid$(datedata, 1, 4)
                    strMonth = Mid$(datedata, 5, 2)
                    strDay = Mid$(datedata, 7, 2)
                    strHour = Mid$(datedata, 9, 2)
                    strMinute = Mid$(datedata, 11, 2)

                    If strHour = "24" Then
                        ConvertDate = strMonth & "/" & strDay & "/" & strYear
                    Else
                        ConvertDate = strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute
                    End If


                Else
                    '20140212
                    'ConvertDate = DateTime.Now
                    ConvertDate = "NULL"

                End If

            End If 'If datedata = """""" Then
        Finally
        End Try

    End Function

    Public Function ConvertSupDate(ByVal datedata As String) As String
        'convert the hl7 date to a database date
        'hl7 in format: yyyymmdd or yyyymmddhhmm
        'returns now if the string is not in one
        'of the two formats.
        '
        Try
            Dim strYear As String = ""
            Dim strMonth As String = ""
            Dim strDay As String = ""
            Dim strHour As String = ""
            Dim strMinute As String = ""

            If Len(Trim(datedata)) = 8 Then
                'strYear = Left$(datedata, 4)
                'strMonth = Mid$(datedata, 5, 2)
                'strDay = Mid$(datedata, 7, 2)
                strYear = Mid$(datedata, 5, 4)
                strMonth = Left$(datedata, 2)
                strDay = Mid$(datedata, 3, 2)
                ConvertSupDate = strMonth & "/" & strDay & "/" & strYear



            Else
                '20140212
                'ConvertSupDate = DateTime.Now
                ConvertSupDate = "NULL"

            End If

        Finally
        End Try

    End Function

    Public Function ConvertSOS(ByVal data As String) As String
        'converts a sos number without delimiters to:
        ' sss-ss-ssss
        Try
            If Len(data) = 9 Then
                ConvertSOS = Mid$(data, 1, 3) & "-" & Mid$(data, 4, 2) & "-" & Mid$(data, 6, 4)
            Else
                ConvertSOS = ""
            End If
        Finally
        End Try
    End Function

    Public Sub UpdateInsurer(ByVal dictNVP As Hashtable)
        '20140528 Added AuditNotes from ZIN_9
        '20140603 code to use star plancode in the iplancode field
        '20140904 - return to integer fclass
        Try
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            Dim updatecommand As New SqlCommand
            Dim dataReader As SqlDataReader
            Dim addit As Boolean = False
            Dim updateit As Boolean = False
            Dim tempstr As String = ""
            Dim tempstr2 As String = ""
            Dim sql As String
            Dim strFclass As String = ""
            Dim intFClass As Integer = 0 '20140904
            Dim iPlanCodeExists As Boolean
            Dim bolProcessThis As Boolean = False
            Dim strAuthServices As String = ""
            Dim epNumExists As Boolean = False
            Dim boolEpisodeExits As Boolean = False
            Dim i As Integer = 0
            updatecommand.Connection = myConnection
            objCommand.Connection = myConnection

            '20130429 - added star_plancode
            Dim star_plancode As String = ""
            'epNumExists = False
            'strAuthServices = ""
            'bolProcessThis = False

            If dictNVP.Item("Event Type Code") = "A01" Then bolProcessThis = True
            If dictNVP.Item("Event Type Code") = "A04" Then bolProcessThis = True
            If dictNVP.Item("Event Type Code") = "A05" Then bolProcessThis = True
            If dictNVP.Item("Event Type Code") = "A08" Then bolProcessThis = True

            If (bolProcessThis) Then
                If gblEPNum > 0 Then
                    boolEpisodeExits = True

                End If
                'See if any records exist in the insurer table for this panum
                sql = "select epnum from [03insurer] where epnum = " & gblEPNum
                objCommand.CommandText = sql
                myConnection.Open()
                dataReader = objCommand.ExecuteReader()
                If dataReader.HasRows Then
                    epNumExists = True
                Else
                    epNumExists = False
                End If
                myConnection.Close()
                dataReader.Close()

                intFClass = 0
                sql = "select id from [104finclass] where finclass = '" & dictNVP.Item("insurer.fClass") & "' and inactive = 0" '20140904
                objCommand.CommandText = sql
                myConnection.Open()
                dataReader = objCommand.ExecuteReader()
                While dataReader.Read()
                    intFClass = dataReader.GetInt32(0)
                End While
                myConnection.Close()
                dataReader.Close()

                'addit = False
                'updateit = False ' change this

                'strFclass = dictNVP.Item("insurer.fClass") '20130410 '20140904

                If dictNVP.Item("Event Type Code") = "A01" Then addit = True
                If dictNVP.Item("Event Type Code") = "A04" Then addit = True
                If dictNVP.Item("Event Type Code") = "A05" Then addit = True
                If dictNVP.Item("Event Type Code") = "A08" Then updateit = True

                If ((addit) And (boolEpisodeExits) And Not (epNumExists)) Then
                    Call insertInsurer(dictNVP)
                End If

                If ((updateit) And (boolEpisodeExits)) Then
                    'i = 0
                    For i = 1 To gblInsCount
                        If i = 1 Then
                            tempstr = ""
                        End If
                        If i > 1 Then
                            tempstr = "_000" & i
                        End If
                        '20130429 - Added STAR_Plancode to [03insurer] table
                        star_plancode = Trim(Replace(dictNVP("iplancode2" & tempstr), "'", "''")) & Trim(Replace(dictNVP("iplancode" & tempstr), "'", "''"))


                        If Len(dictNVP.Item("iplancode" & tempstr)) >= 3 Then
                            iPlanCodeExists = False
                            '20130502 - change to return epnum instead of patconum
                            '20140603 - change to use star plancode in the iplancode field
                            sql = "SELECT epnum FROM [03Insurer] where epnum = " & gblEPNum & " "
                            'sql = sql & "AND star_plancode  = '" & star_plancode & "'" '20140603
                            sql = sql & "AND iplancode  = '" & star_plancode & "'" '20140603

                            objCommand.CommandText = sql
                            myConnection.Open()
                            dataReader = objCommand.ExecuteReader()

                            If dataReader.HasRows Then
                                iPlanCodeExists = True
                            Else
                                iPlanCodeExists = False
                            End If
                            myConnection.Close()
                            dataReader.Close()

                            If iPlanCodeExists Then

                                sql = "UPDATE [03Insurer] "
                                sql = sql & "SET updated = '" & Now & "'"

                                '20151215
                                If Len(dictNVP("Insured DOB")) > 0 Then
                                    sql = sql & ", InsuredDOB = '" & ConvertDate(dictNVP("Insured DOB")) & "'"
                                End If

                                If Len(dictNVP("Insured Sex")) > 0 Then
                                    sql = sql & ", InsuredSex = '" & dictNVP("Insured Sex") & "'"
                                End If
                                '20151215 end

                                If Len(dictNVP.Item("group" & tempstr)) > 0 Then
                                    sql = sql & ", theGroup = '" & Replace(dictNVP.Item("group" & tempstr), "'", "''") & "'"
                                End If

                                If Len(dictNVP.Item("PolicyNumber" & tempstr)) > 0 Then
                                    sql = sql & ", policyNum = '" & Replace(dictNVP.Item("PolicyNumber" & tempstr), "'", "''") & "'"
                                End If
                                '20170510 - Removed AuthNum Process using Process IN1_14. Left in of ULHT
                                If Len(dictNVP.Item("AuthNum" & tempstr)) > 0 Then
                                    sql = sql & ", authnum1 = '" & Replace(dictNVP.Item("AuthNum" & tempstr), "'", "''") & "'"
                                End If

                                '20130502 - changed name routine to handle double quotes
                                tempstr2 = ""
                                If dictNVP.Item("Insured First Name" & tempstr) <> """""" And dictNVP.Item("Insured First Name" & tempstr) <> "" Then
                                    tempstr2 = tempstr2 & dictNVP.Item("Insured First Name" & tempstr) & " "
                                End If
                                If dictNVP.Item("Insured Middle Name" & tempstr) <> """""" And dictNVP.Item("Insured Middle Name" & tempstr) <> "" Then
                                    tempstr2 = tempstr2 & dictNVP.Item("Insured Middle Name" & tempstr) & " "
                                End If

                                If dictNVP.Item("Insured Last Name" & tempstr) <> """""" And dictNVP.Item("Insured Last Name" & tempstr) <> "" Then
                                    tempstr2 = tempstr2 & dictNVP.Item("Insured Last Name" & tempstr)
                                End If
                                'tempstr2 = dictNVP.Item("Insured First Name" & tempstr) & _
                                '" " & dictNVP.Item("Insured Middle Name" & tempstr) & _
                                '" " & dictNVP("Insured Last Name" & tempstr)

                                '5/29/2003 - check for a string length > 3 vice zero
                                If Len(tempstr2) > 3 Then
                                    sql = sql & ", subname = '" & Replace(tempstr2, "'", "''") & "'"
                                End If

                                '20130502
                                If (Len(dictNVP.Item("PolicyIssueDate" & tempstr)) > 0 And dictNVP.Item("PolicyIssueDate" & tempstr) <> """""") Then
                                    sql = sql & ", PIssue = '" & ConvertDate(dictNVP.Item("PolicyIssueDate" & tempstr)) & "'"
                                ElseIf dictNVP.Item("PolicyIssueDate" & tempstr) = """""" Then
                                    sql = sql & ", PIssue = NULL "
                                End If

                                If dictNVP("COBPriority" & tempstr) = "1" Then
                                    sql = sql & ", aprimary = 1"
                                    sql = sql & ", fclass = " & intFClass & " " '20140904
                                    sql = sql & ", AuditNotes = '" & Replace(dictNVP.Item("InsNotes"), "'", "''") & "'" '20140528
                                Else
                                    sql = sql & ", aprimary = 0"
                                    sql = sql & ", fclass = NULL "
                                End If


                                sql = sql & ", reqCert = 1"
                                sql = sql & " Where epnum = " & gblEPNum
                                sql = sql & " AND iplancode = '" & star_plancode & "'" '20140603
                                'LogFile.WriteLine(sql)
                                updatecommand.CommandText = sql
                                myConnection.Open()
                                updatecommand.ExecuteNonQuery()
                                myConnection.Close()

                                '20170623 Process IN1_14 and ZGI for multiple authcodes
                                ProcessIN1_14(dictNVP, tempstr)


                            Else                'iPlancode does not exist

                                Call insertInsurer(dictNVP)
                            End If              'If iPlanCodeExists Then

                        End If                  'Len(dictNVP.Item("iplancode" & tempstr)) >= 3

                    Next                        'gblInsCount
                End If                          'If ((updateit) And (visitPaNumExists))

            End If                              'If (bolProcessThis)

        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "Update Insurer Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

            Exit Sub
        End Try
    End Sub

    
    Public Sub UpdatePPS(ByVal dictNVP As Hashtable)
        Try
            '3/13/2002
            'add or update records in PPS table
            Dim sql As String = ""
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            Dim updatecommand As New SqlCommand
            Dim dataReader As SqlDataReader
            Dim boolRecordExists As Boolean
            boolRecordExists = False
            updatecommand.Connection = myConnection
            objCommand.Connection = myConnection

            Select Case dictNVP.Item("Event Type Code")
                '================================================================================================
                Case "A01"
                    '1. check to see if a record exists with the existing panum
                    'if the record does not exist, add it.
                    sql = "select panum from [001pps] where panum = '" & dictNVP.Item("panum") & "'"
                    objCommand.CommandText = sql
                    myConnection.Open()
                    dataReader = objCommand.ExecuteReader()
                    If dataReader.HasRows Then
                        boolRecordExists = True
                    Else
                        boolRecordExists = False
                    End If
                    myConnection.Close()
                    dataReader.Close()
                    '2. if no record then add it
                    If Not boolRecordExists Then
                        sql = "INSERT into [001pps] (epnum, panum, created) "
                        sql = sql & "VALUES (" & gblEPNum & ", "
                        sql = sql & "'" & Replace(dictNVP.Item("panum"), "'", "''") & "', "
                        sql = sql & "'" & Now & "') "
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()

                        '20160829 - Insert record into 001WeeFim Table
                        sql = "INSERT into [001WeeFIM] (epnum,created) "
                        sql = sql & "VALUES (" & gblEPNum & ", "
                        sql = sql & "'" & Now & "') "
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()


                    End If 'If boolRecordExists

                Case "A05"
                    '12/12/2002 - changed to patient status code from patient class
                    If dictNVP.Item("Patient Status Code") = "IP" Then
                        sql = "select panum from [001pps] where panum = '" & dictNVP.Item("panum") & "'"
                        objCommand.CommandText = sql
                        myConnection.Open()
                        dataReader = objCommand.ExecuteReader()
                        If dataReader.HasRows Then
                            boolRecordExists = True
                        Else
                            boolRecordExists = False
                        End If
                        myConnection.Close()
                        dataReader.Close()

                        '2. if no record then add it
                        If Not boolRecordExists Then
                            sql = "INSERT into [001pps] (epnum, panum, created) "
                            sql = sql & "VALUES (" & gblEPNum & ", "
                            sql = sql & "'" & Replace(dictNVP.Item("panum"), "'", "''") & "', "
                            sql = sql & "'" & Now & "') "
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()

                            '20160829 - Insert record into 001WeeFim Table
                            sql = "INSERT into [001WeeFIM] (epnum,created) "
                            sql = sql & "VALUES (" & gblEPNum & ", "
                            sql = sql & "'" & Now & "') "
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()


                        End If 'If boolRecordExists
                    End If 'If dictNVP.Item("Hospital Service") = "IP"

            End Select

        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "Update PPS Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

            Exit Sub
        End Try

    End Sub

    Public Sub processError(ByVal dictNVP As Hashtable)
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        objCommand.Connection = myConnection
        functionError = False
        dbError = False
        continueProcessing = True
        orphanFound = False
        Try
            Dim dataReader As SqlDataReader
            Dim checkThis As Boolean = True
            Dim sql As String = ""
            If dictNVP.Item("Event Type Code") = "A01" Then checkThis = False
            If dictNVP.Item("Event Type Code") = "A04" Then checkThis = False
            If dictNVP.Item("Event Type Code") = "A05" Then checkThis = False

            If checkThis And dictNVP.Item("panum") <> "" Then
                sql = "select panum from [001episode] where panum = '" & dictNVP.Item("panum") & "'"
                objCommand.CommandText = sql
                myConnection.Open()
                dataReader = objCommand.ExecuteReader()
                If dataReader.HasRows Then
                    'do nothing the panum exists

                Else
                    'can't find the panum so send to orpan directory

                    '20130603 - add the orphan panum to the episode table
                    '20160204 - remove call to addOrphanPanum
                    'addOrphanPaNUm(dictNVP)

                    '20150717 - set orphan processing to false always
                    orphanFound = True

                    continueProcessing = False
                    gblLogString = gblLogString & CStr(DateTime.Now) & " - Orphan Found. Panum = " & dictNVP.Item("panum") & vbCrLf
                End If
                dataReader.Close()

            End If


        Catch ex As Exception
            continueProcessing = False
            If Err.Number = 5 Then
                dbError = True
                functionError = False
            Else
                functionError = True
            End If

            gblLogString = gblLogString & "Connection Error (processError): " & Err.Number & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            
            Exit Sub
        Finally
            myConnection.Close()

        End Try
    End Sub
    Public Sub updateEpisodeSupplement(ByVal dictNVP As Hashtable)
        '10/22/2007 add/update supplemental table
        Dim testsql As String = ""
        '20071116: if no data bypass process suing supDataExists
        '20120404 - add records even if no data exists to match the episode table.
        Dim supDataExists As Boolean = True
        'If dictNVP.Item("ZJH_7") <> "" Then supDataExists = True
        'If dictNVP.Item("ZJH_8") <> "" Then supDataExists = True
        'If dictNVP.Item("ZJH_9") <> "" Then supDataExists = True
        'If dictNVP.Item("ZJH_10") <> "" Then supDataExists = True
        'If dictNVP.Item("ZJH_11") <> "" Then supDataExists = True
        'If dictNVP.Item("ZJH_12") <> "" Then supDataExists = True
        'If dictNVP.Item("ZJH_13") <> "" Then supDataExists = True

        Try
            If gblEPNum > 0 And supDataExists Then

                Dim boolEPNUMExists As Boolean = False
                Dim sql As String = ""
                Dim myConnection As New SqlConnection(connectionString)
                Dim objCommand As New SqlCommand
                Dim updatecommand As New SqlCommand
                updatecommand.Connection = myConnection

                objCommand.Connection = myConnection
                Dim dataReader As SqlDataReader

                sql = "select * from [001episodeSup] where epnum = " & gblEPNum
                objCommand.CommandText = sql
                myConnection.Open()
                dataReader = objCommand.ExecuteReader()
                If dataReader.HasRows Then
                    boolEPNUMExists = True
                Else
                    boolEPNUMExists = False
                End If

                myConnection.Close()
                dataReader.Close()

                '======================================================================================================
                Select Case dictNVP.Item("Event Type Code")

                    Case "A01", "A04", "A05", "A08"
                        If boolEPNUMExists Then 'update it
                            sql = "update [001episodeSup] set "
                            If dictNVP.Item("ZJH_7") <> "" Then
                                sql = sql & "cardiacRehab = '" & ConvertSupDate(dictNVP.Item("ZJH_7")) & "'"
                            Else
                                sql = sql & "cardiacRehab = NULL"
                            End If
                            If dictNVP.Item("ZJH_8") <> "" Then
                                sql = sql & ", OTEst = '" & ConvertSupDate(dictNVP.Item("ZJH_8")) & "'"
                            Else
                                sql = sql & ", OTEst = NULL"
                            End If
                            If dictNVP.Item("ZJH_9") <> "" Then
                                sql = sql & ", OTTreatment = '" & ConvertSupDate(dictNVP.Item("ZJH_9")) & "'"
                            Else
                                sql = sql & ", OTTreatment = NULL"
                            End If
                            If dictNVP.Item("ZJH_10") <> "" Then
                                sql = sql & ", PTEst = '" & ConvertSupDate(dictNVP.Item("ZJH_10")) & "'"
                            Else
                                sql = sql & ", PTEst = NULL"
                            End If
                            If dictNVP.Item("ZJH_11") <> "" Then
                                sql = sql & ", PTTreatment = '" & ConvertSupDate(dictNVP.Item("ZJH_11")) & "'"
                            Else
                                sql = sql & ", PTTreatment = NULL"
                            End If
                            If dictNVP.Item("ZJH_12") <> "" Then
                                sql = sql & ", STEst = '" & ConvertSupDate(dictNVP.Item("ZJH_12")) & "'"
                            Else
                                sql = sql & ", STEst = NULL"
                            End If
                            If dictNVP.Item("ZJH_13") <> "" Then
                                sql = sql & ", STTreatment = '" & ConvertSupDate(dictNVP.Item("ZJH_13")) & "'"
                            Else
                                sql = sql & ", STTreatment = NULL"
                            End If

                            '20090920 - add servicing facility

                            If dictNVP.Item("Servicing Facility") <> "" Then
                                sql = sql & ", servFacility = '" & Replace(dictNVP.Item("Servicing Facility"), "'", "''") & "'"
                            Else
                                sql = sql & ", servFacility = NULL"
                            End If

                            '20099020 - end
                            sql = sql & " Where epnum = " & gblEPNum
                            testsql = sql
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()

                        Else 'add it
                            '20090920 - add servicing facility to insert statement
                            sql = "INSERT into [001EpisodeSup] (epnum, cardiacRehab, OTEst, OTTreatment, PTEst, PTTreatment, STEst, STTreatment, servFacility)"
                            sql = sql & "VALUES (" & gblEPNum & ","
                            If dictNVP.Item("ZJH_7") <> "" Then
                                sql = sql & "'" & ConvertSupDate(dictNVP("ZJH_7")) & "', "
                            Else
                                sql = sql & "NULL,"
                            End If
                            If dictNVP.Item("ZJH_8") <> "" Then
                                sql = sql & "'" & ConvertSupDate(dictNVP("ZJH_8")) & "', "
                            Else
                                sql = sql & "NULL,"
                            End If
                            If dictNVP.Item("ZJH_9") <> "" Then
                                sql = sql & "'" & ConvertSupDate(dictNVP("ZJH_9")) & "', "
                            Else
                                sql = sql & "NULL,"
                            End If
                            If dictNVP.Item("ZJH_10") <> "" Then
                                sql = sql & "'" & ConvertSupDate(dictNVP("ZJH_10")) & "', "
                            Else
                                sql = sql & "NULL,"
                            End If
                            If dictNVP.Item("ZJH_11") <> "" Then
                                sql = sql & "'" & ConvertSupDate(dictNVP("ZJH_11")) & "', "
                            Else
                                sql = sql & "NULL,"
                            End If
                            If dictNVP.Item("ZJH_12") <> "" Then
                                sql = sql & "'" & ConvertSupDate(dictNVP("ZJH_12")) & "', "
                            Else
                                sql = sql & "NULL,"
                            End If
                            If dictNVP.Item("ZJH_13") <> "" Then
                                sql = sql & "'" & ConvertSupDate(dictNVP("ZJH_13")) & "', "
                            Else
                                sql = sql & "NULL,"
                            End If

                            '20090920 - add servicing facility
                            If dictNVP.Item("Servicing Facility") <> "" Then
                                sql = sql & "'" & Replace(dictNVP.Item("Servicing Facility"), "'", "''") & "')"
                            Else
                                sql = sql & "NULL)"
                            End If
                            '20090920 - end

                            testsql = sql
                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()

                        End If

                End Select
                '======================================================================================================

            End If ' If gblEPNum > 0 Then
        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "Update Episode Supp Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

            Exit Sub
        End Try
    End Sub
    Public Sub writeToLog2(ByVal logText As String, ByVal eventType As Integer)
        Dim myLog As New EventLog()
        Try
            ' check for the existence of the log that the user wants to create.
            ' Create the source, if it does not already exist.
            If Not EventLog.SourceExists("STAR_ITWfeed") Then
                EventLog.CreateEventSource("STAR_ITWfeed", "STAR_ITWfeed")
            End If

            ' Create an EventLog instance and assign its source.

            myLog.Source = "STAR_ITWfeed"

            ' Write an informational entry to the event log.
            If eventType = 1 Then
                myLog.WriteEntry(logText, EventLogEntryType.Error, 1)
            ElseIf eventType = 2 Then
                myLog.WriteEntry(logText, EventLogEntryType.Warning, 2)
            ElseIf eventType = 3 Then
                myLog.WriteEntry(logText, EventLogEntryType.Information, 3)
            End If


        Finally
            myLog.Close()
        End Try
    End Sub
    Public Sub writeTolog(ByVal strMsg As String, ByVal eventType As Integer)
        '20140205 - use a text file to log errors instead of the event log
        Dim file As System.IO.StreamWriter
        Dim tempLogFileName As String = strLogDirectory & "ITWFeed_log.txt"
        file = My.Computer.FileSystem.OpenTextFileWriter(tempLogFileName, True)
        file.WriteLine(DateTime.Now & " : " & strMsg)
        file.Close()
    End Sub
    Public Sub checkZMI(ByVal dictNVP As Hashtable)
        '20100218 -  procedure to count multiple ZMI segments
        Dim ZMICount As Integer
        Try
            Dim myEnumerator As IDictionaryEnumerator = dictNVP.GetEnumerator()
            ZMICount = 0

            'For Each value In dictNVP.values
            'Console.WriteLine(value)

            'Next
            myEnumerator.Reset()
            While myEnumerator.MoveNext()
                If Left$(myEnumerator.Key, 5) = "ZMI_5" Then
                    gblZMICount = gblZMICount + 1
                    ZMICount = ZMICount + 1
                End If

            End While
            gblLogString = gblLogString & " ZMICount =  "
            gblLogString = gblLogString & ZMICount & vbCrLf
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ZMI Enumeration Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub
    Public Sub processZMI(ByVal dictNVP As Hashtable)
        '20100610 ->
        Dim i As Integer = 0
        Dim sql As String = ""
        Dim tempstr As String = ""
        Dim updateIT As Boolean = False
        Dim recordExists As Boolean = False

        Try
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection
            Dim dataReader As SqlDataReader

            If dictNVP.Item("Event Type Code") = "A01" Then updateIT = True
            If dictNVP.Item("Event Type Code") = "A04" Then updateIT = True
            If dictNVP.Item("Event Type Code") = "A05" Then updateIT = True
            If dictNVP.Item("Event Type Code") = "A08" Then updateIT = True

            If updateIT And gblZMICount > 0 Then
                For i = 1 To gblZMICount
                    If i = 1 Then
                        tempstr = ""
                    End If
                    If i > 1 Then
                        tempstr = "_000" & i
                    End If

                    sql = "select * from [03insurerSupplement] "
                    sql = sql & "where epnum = " & gblEPNum & " and iplancode = '" & dictNVP.Item("ZMI_5" & tempstr) & "'"
                    objCommand.CommandText = sql
                    myConnection.Open()
                    dataReader = objCommand.ExecuteReader()

                    If dataReader.HasRows Then
                        recordExists = True
                    Else
                        recordExists = False

                    End If


                    dataReader.Close()
                    myConnection.Close()
                    If recordExists Then
                        Call ZMIUpdate(dictNVP, tempstr)
                    Else
                        Call ZMIInsert(dictNVP, tempstr)
                    End If



                Next 'gblZMICount
            End If 'If updateIT

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process ZMI2 Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
    Public Sub ZMIInsert(ByVal dictNVP As Hashtable, ByVal tempStr As String)
        '20100610 ->
        Try
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection


            If gblEPNum <> 0 Then
                '20100930 - added code below to insert the episode number into the record.

                sql = "Insert [03InsurerSupplement] "
                sql = sql & "(panum, epnum, iplancode, precertPhone, precertContact, serviceAuth1, ServiceAuth2, coverage1, coverage2) "
                'sql = sql & "precertAmount, benefitPerson, benefitEffectiveDate, benefitDeduction, benefitPhone, benefitMet, "
                'sql = sql & "benefitCoPay, benefitOutOfPocket, coverage1, coverage2, verifiedBy, verifyDate, limitations1, limitations2, "
                'sql = sql & "limitations3, limitations4, "
                'sql = sql & "comments1, comments2, comments3, comments4, comments5, comments6, comments7, comments8, "
                'sql = sql & "comments9, comments10, "
                'sql = sql & "comments11) "
                sql = sql & "VALUES ("
                insertString(gblORIGINAL_PA_NUMBER)
                insertNumber(gblEPNum)
                insertString(dictNVP("ZMI_5" & tempStr))
                insertString(dictNVP("ZMI_8" & tempStr))

                insertString(dictNVP("ZMI_7" & tempStr))
                insertString(dictNVP("ZMI_10" & tempStr))
                insertString(dictNVP("ZMI_11" & tempStr))

                insertString(dictNVP("ZMI_20" & tempStr))
                '20100610 ->
                insertLastString(dictNVP("ZMI_21" & tempStr))

                objCommand.CommandText = sql
                myConnection.Open()

                objCommand.ExecuteNonQuery()
                myConnection.Close()
            End If

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Process ZMI Insert Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
    Public Sub ZMIUpdate(ByVal dictNVP As Hashtable, ByVal tempStr As String)
        '20100610 ->
        Dim sql As String = ""
        Try
            
            Dim myConnection As New SqlConnection(connectionString)
            Dim objCommand As New SqlCommand
            objCommand.Connection = myConnection

            If gblEPNum <> 0 Then

                sql = "UPDATE [03InsurerSupplement] "
                sql = sql & "SET updated = '" & DateTime.Now & "'"

                sql = sql & ", precertPhone = '" & Replace(dictNVP("ZMI_8" & tempStr), "'", "''") & "'"

                sql = sql & ", precertContact = '" & Replace(dictNVP("ZMI_7" & tempStr), "'", "''") & "'"
                sql = sql & ", serviceAuth1 = '" & Replace(dictNVP("ZMI_10" & tempStr), "'", "''") & "'"
                sql = sql & ", serviceAuth2 = '" & Replace(dictNVP("ZMI_11" & tempStr), "'", "''") & "'"

                sql = sql & ", coverage1 = '" & Replace(dictNVP("ZMI_20" & tempStr), "'", "''") & "'"
                sql = sql & ", coverage2 = '" & Replace(dictNVP("ZMI_21" & tempStr), "'", "''") & "'"


                sql = sql & ", panum = '" & gblORIGINAL_PA_NUMBER & "'"

                sql = sql & " Where epnum = " & gblEPNum


                sql = sql & " and iplancode = '" & dictNVP("ZMI_5" & tempStr) & "'"

                objCommand.CommandText = sql
                myConnection.Open()
                objCommand.ExecuteNonQuery()
                myConnection.Close()

            End If

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ZMIUpdate Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub

        End Try
    End Sub
    Public Sub insertLastString(ByVal theString As String)
        '20100610 ->
        '20130422 - code to handle double quotes
        Try
            'If theString <> "" Then
            'sql = sql & "'" & Replace(theString, "'", "''") & "') "
            'Else
            'sql = sql & "NULL) "
            'End If

            Select Case theString
                Case """"""
                    sql = sql & "NULL) "
                Case ""
                    sql = sql & "NULL) "
                Case Else
                    sql = sql & "'" & Replace(theString, "'", "''") & "') "
            End Select


        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Insert Last String Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub
    Public Function STARupdateString(ByVal strVariableName As String, ByVal strVariableValue As String) As String
        STARupdateString = ""

        If Trim(strVariableValue) = """""" Or Trim(strVariableValue) = "" Then
            STARupdateString = ", " & strVariableName & " = NULL "
        Else
            STARupdateString = ", " & strVariableName & " = '" & Replace(strVariableValue, "'", "''") & "'"
        End If
    End Function

    Public Sub updateFinancial(ByVal dictNVP As Hashtable)
        '20130514 - aded this routine to to the star feed program
        '20130429 - add record to 004financial table, if it does not exist. Add epnum only.
        'add on A01, 4 and 5 only
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        Dim dataReader As SqlDataReader
        Dim sql As String = ""

        Dim bolProcessThis As Boolean = False

        Dim epNumExists As Boolean = False
        Dim boolEpisodeExits As Boolean = False
        updatecommand.Connection = myConnection
        objCommand.Connection = myConnection
        Try
            If dictNVP.Item("Event Type Code") = "A01" Then bolProcessThis = True
            If dictNVP.Item("Event Type Code") = "A04" Then bolProcessThis = True
            If dictNVP.Item("Event Type Code") = "A05" Then bolProcessThis = True

            If (bolProcessThis) Then
                If gblEPNum > 0 Then
                    boolEpisodeExits = True

                End If
                'See if any records exist in the insurer table for this panum
                sql = "select epnum from [004Financial] where epnum = " & gblEPNum
                objCommand.CommandText = sql
                myConnection.Open()
                dataReader = objCommand.ExecuteReader()
                If dataReader.HasRows Then
                    epNumExists = True
                Else
                    epNumExists = False
                End If
                myConnection.Close()
                dataReader.Close()

                If Not epNumExists Then
                    sql = "Insert [004financial] "
                    sql = sql & "(epnum) "
                    sql = sql & "VALUES ("
                    sql = sql & gblEPNum & ") "
                    updatecommand.CommandText = sql
                    myConnection.Open()
                    updatecommand.ExecuteNonQuery()
                    myConnection.Close()
                End If

            End If 'If (bolProcessThis) Then

        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "Update Financial Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

            Exit Sub
        End Try



    End Sub
    Public Function extractCorpNo(ByVal pidString As String) As String
        Dim pidArray() As String
        Dim pidItemArray() As String
        Dim tempStr As String = ""
        Dim testData As String = ""
        Dim J As Integer = 0

        Try
            extractCorpNo = "0"

            If Left(pidString, 1) = "~" Then
                pidString = Mid(pidString, 2)
            End If
            extractCorpNo = "0"
            pidArray = Split(pidString, "~")
            For J = 0 To UBound(pidArray)
                tempStr = pidArray(J)
                tempStr = Trim(tempStr)
                'testData = Mid$(tempStr, Len(tempStr) - 1, 2)
                pidItemArray = Split(tempStr, "^")
                If pidItemArray(4) = "PI" Then
                    extractCorpNo = pidItemArray(0)
                End If
                'If testData = "PI" Then
                'Dim pidSubArray() As String
                'pidSubArray = Split(tempStr, "^")
                'extractCorpNo = pidSubArray(0)
                'End If

            Next
        Catch ex As Exception
            extractCorpNo = "0"

        End Try

    End Function

    Public Function processOnSetDate(ByVal UB1string As String) As String
        '20141023 - new function to extract the onsetDate if the code is 11
        Dim UB1Array() As String
        Dim UB1ItemArray() As String
        Dim tempStr As String = ""
        Dim testData As String = ""
        Dim J As Integer = 0

        Try
            processOnSetDate = ""

            If Left(UB1string, 1) = "~" Then
                UB1string = Mid(UB1string, 2)
            End If
            processOnSetDate = ""
            UB1Array = Split(UB1string, "~")
            For J = 0 To UBound(UB1Array)
                tempStr = UB1Array(J)
                tempStr = Trim(tempStr)
                'testData = Mid$(tempStr, Len(tempStr) - 1, 2)
                UB1ItemArray = Split(tempStr, "^")
                If UB1ItemArray(0) = "11" Then
                    processOnSetDate = UB1ItemArray(1)
                End If
                'If testData = "PI" Then
                'Dim pidSubArray() As String
                'pidSubArray = Split(tempStr, "^")
                'extractCorpNo = pidSubArray(0)
                'End If

            Next
        Catch ex As Exception
            processOnSetDate = ""

        End Try

    End Function

    Public Sub processOccurrenceCodes(ByVal dictNVP As Hashtable, ByVal gblORIGINAL_PA_NUMBER As String)
        '20141023 - new procedure to process occurrence codes to 075occurence table
        Dim myConnection As New SqlConnection(connectionString)
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        Dim UB1Array() As String
        Dim UB1ItemArray() As String
        Dim tempStr As String = ""
        Dim testData As String = ""
        Dim J As Integer = 0
        Dim UB1string As String = ""
        Dim sql As String = ""
        Try

            If dictNVP.Item("UB1string") <> "" And gblORIGINAL_PA_NUMBER <> "" Then
                UB1string = dictNVP.Item("UB1string")
                sql = "delete from [075occurrence] where panum = '" & gblORIGINAL_PA_NUMBER & "'"
                updatecommand.CommandText = sql
                myConnection.Open()
                updatecommand.ExecuteNonQuery()
                myConnection.Close()

                If Left(UB1string, 1) = "~" Then
                    UB1string = Mid(UB1string, 2)
                End If

                UB1Array = Split(UB1string, "~")
                For J = 0 To UBound(UB1Array)
                    tempStr = UB1Array(J)
                    tempStr = Trim(tempStr)

                    UB1ItemArray = Split(tempStr, "^")
                    If UB1ItemArray(0) <> "" And UB1ItemArray(1) <> "" Then
                        sql = "insert into [075occurrence] (panum, code, value, created) values("
                        sql = sql & "'" & gblORIGINAL_PA_NUMBER & "', '" & UB1ItemArray(0).ToString & "', '" & UB1ItemArray(1).ToString & "', '" & DateTime.Now & "')"
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()
                    End If

                Next
            End If 'If dictNVP.Item("UB1string") <> ""

        Catch ex As Exception


        End Try
    End Sub

    Public Sub addOrphanPaNUm(ByVal dictNVP As Hashtable)
        '20130603 - routine added to add the panum for an orphan to catch up on previous records in the STAR ITW database
        Dim myConnection As New SqlConnection(connectionString)
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        Dim star_region As String = UCase(dictNVP("Sending Facility"))
        Dim sql As String = ""
        sql = "insert into [001episode] (mrnum, corpNo, star_region, orphanAdded, panum) "
        If IsNumeric(dictNVP("mrnum")) Then
            sql = sql & "VALUES (" & dictNVP("mrnum") & ", "
        Else
            sql = sql & "VALUES (0, "
        End If
        sql = sql & gblCorporateNumber & ", "
        sql = sql & "'" & star_region & "', "
        sql = sql & "'" & DateTime.Now & "', "
        sql = sql & "'" & dictNVP("panum") & "') "
        updatecommand.CommandText = sql
        myConnection.Open()
        updatecommand.ExecuteNonQuery()
        myConnection.Close()
    End Sub

    Public Function xlateClass(ByVal dictNVP As Hashtable) As String
        '20130619 - translate the patient class from STAR
        'PV1_2 = "Patient Class"
        'PV1_18 = "Patient Type"
        'logic from StarToInvision Trimap 2/25/2013 version 1
        xlateClass = " "
        Select Case dictNVP("Patient Class")
            Case "R"
                xlateClass = "O"
            Case "B"
                xlateClass = "E"
            Case "P"
                If UCase(Left(dictNVP("Patient Type"), 2)) = "PZ" Or UCase(Left(dictNVP("Patient Type"), 2)) = "CZ" Then
                    xlateClass = "I"
                Else
                    xlateClass = "O"
                End If
            Case Else
                xlateClass = dictNVP("Patient Class")
        End Select
    End Function

    Public Function calcStatus(ByVal dictNVP As Hashtable) As String
        Dim strFirstCharacter As String = ""
        Dim strSecondCharacter As String = ""
        '20130620 - calculate the status based on raw Patient Class data from STAR and criteria below

        'from AL:
        'If PV1-2 = "R", then set to "O"
        'Else if PV1-2 = "B", then set to "E"
        'Else if PV1-2 = “P”, and:
        '{
        'If PV1-18 starts with “PZ” or “CZ”, then set to"I";
        'Else set to “O” 
        '}
        'otherwise pass thru the Patient Class with no translation.

        'Capture the patient class and add to the 001episode table
        '[001episode].status: 1st character is the Patient Class 
        '2nd character is “A” for A01, A04 & A13
        '2nd character is “P” for A05
        '2nd character is “D” for A03


        calcStatus = ""
        '1. Get the first character using the xlateClass Function
        strFirstCharacter = xlateClass(dictNVP)
        '2. Get the second character based on type for HL7 record

        Select Case dictNVP("Event Type Code")
            Case "A01", "A04", "A13"
                strSecondCharacter = "A"
            Case "A05"
                strSecondCharacter = "P"
            Case "A03"
                strSecondCharacter = "D"
        End Select

        '3. Combine the two calculated characters if the second character was calculated,
        'otherwise send nothing.
        If strSecondCharacter <> "" Then
            calcStatus = strFirstCharacter & strSecondCharacter
        Else
            calcStatus = ""
        End If


    End Function

    Public Sub checkAL1(ByVal dictNVP As Hashtable)
        '20140215 - added to handle A31 messages from Cerner
        '20140917 - zero out gblAl1Count before counting

        Try
            Dim AL1Count As Integer
            Dim myEnumerator As IDictionaryEnumerator = dictNVP.GetEnumerator()
            AL1Count = 0

            '20140917 - zero out gblAl1Count
            gblAL1Count = 0

            myEnumerator.Reset()
            While myEnumerator.MoveNext()
                If Left$(myEnumerator.Key, 3) = "AL1" Then
                    gblAL1Count = gblAL1Count + 1
                    AL1Count = AL1Count + 1
                End If

            End While

        Catch ex As Exception
            globalError = True

            gblLogString = gblLogString & "AL1 Enumeration Error (checkAL1)" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

            Exit Sub
        End Try
    End Sub

    Public Sub processAL1(ByVal dictNVP As Hashtable)
        '20140215 For A31. Note, ther is no panum in the A31 so we need to use the mrnum
        '20140321 - added use of extractMrnum to this function only.
        '20140915 - modified search in processAL1
        '20140916 - capture all AL1 data
        'Dim A31connectionString As String = "server=10.48.242.249,1433;database=PatientGlobal;uid=sysmax;pwd=Condor!"
        'Dim A31connectionString As String = "server=10.48.64.5\sqlexpress;database=PatientGlobal;uid=sysmax;pwd=Condor!"
        'connectionString = "server=HPLAPTOP;database=STAR_ITW;uid=sa;pwd=b436328"
        Dim A31connectionString As String = conIniFile.GetString("Strings", "ITWA31", "(none)")

        Dim myConnection As New SqlConnection(A31connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        objCommand.Connection = myConnection


        Dim boolProsessThis As Boolean
        Dim tempstr As String
        Dim star_region As String = "" '20140514 added to capture region in global table also capture global corporate number
        Try

            Dim i As Integer
            boolProsessThis = False
            If dictNVP.Item("TriggerEventID") = "A31" Then boolProsessThis = True
            '21040514
            star_region = UCase(dictNVP("Sending Facility"))

            If boolProsessThis Then
                updatecommand.Connection = myConnection
                objCommand.Connection = myConnection
                sql = "delete from [Allergies] "
                sql = sql & "where CorporateNumber = " & gblCorporateNumber ' extractMrnum(dictNVP.Item("mrnum"))
                updatecommand.CommandText = sql
                myConnection.Open()
                updatecommand.ExecuteNonQuery()
                myConnection.Close()
                updatecommand.Connection = myConnection
                objCommand.Connection = myConnection

                tempstr = ""
                sql = ""
                i = 0

                For i = 1 To gblAL1Count
                    sql = ""
                    If i = 1 Then
                        tempstr = ""
                    End If
                    '20140918 - accept more than 9 AL1 segments
                    If i > 1 And i < 10 Then
                        tempstr = "_000" & i
                    End If

                    If i >= 10 And i < 100 Then
                        tempstr = "_00" & i
                    End If

                    'If Not IsDBNull(dictNVP("Allergy Code ID" & tempstr)) Then '20140915
                    'If Trim(dictNVP("Allergy Code ID" & tempstr)) <> "" Then '20140916 - capture all AL1 data
                    '20140514 - add star_region and gblCorporateNumber 
                    sql = "Insert [Allergies] "
                    sql = sql & "(mrnum, region, CorporateNumber, type, code_id, description, coding_system, Severity, Reaction, IDDate, "
                    sql = sql & "added) "

                    sql = sql & "VALUES ("
                    sql = sql & extractMrnum(dictNVP.Item("mrnum")) & ", "
                    insertString(star_region)
                    sql = sql & gblCorporateNumber & ", "

                    insertString(dictNVP.Item("Allergy Type" & tempstr))
                    insertString(dictNVP("Allergy Code ID" & tempstr))
                    insertString(dictNVP.Item("Allergy Description" & tempstr))
                    insertString(dictNVP.Item("Allergy Coding System" & tempstr))

                    '20140303 - added severity, reaction and IDDate
                    insertString(dictNVP.Item("Severity" & tempstr))
                    insertString(dictNVP.Item("Reaction" & tempstr))
                    sql = sql & "'" & ConvertDate(dictNVP("IDDate" & tempstr)) & "', "

                    sql = sql & "'" & DateTime.Now & "') "

                    updatecommand.CommandText = sql
                    myConnection.Open()
                    updatecommand.ExecuteNonQuery()
                    myConnection.Close()
                    'End If '20140916 - capture all AL1 data
                Next 'For i = 1 To gblAL1Count
            End If 'If boolProsessThis
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "AL1 Process Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try
    End Sub
    Public Function calcSTARStatus(ByVal dictNVP As Hashtable) As String
        '20140219 - lookup status using 130PatientStatus table.
        '20150914 calculate value for gblBoolUpdateAdmDate at end of function

        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        Dim sql As String = ""
        objCommand.Connection = myConnection
        Dim strCurrentStatus As String = ""
        Dim strNewStatus As String = ""
        Dim dataReader As SqlDataReader

        Dim tempCurrentStatus As String = "" '20150914
        calcSTARStatus = ""

        Try
            'Get the current status if the record exists in the [001episode] table.
            sql = "select Status from [001episode] where panum = '" & dictNVP.Item("panum") & "'"
            objCommand.CommandText = sql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()
            If dataReader.HasRows Then
                While dataReader.Read()
                    strCurrentStatus = dataReader.GetString(0)
                    tempCurrentStatus = strCurrentStatus '20150914
                End While
            Else
                strCurrentStatus = ""
            End If
            myConnection.Close()
            dataReader.Close()



            Select Case dictNVP.Item("TriggerEventID")
                Case "A01", "A04", "A05", "A08"
                    '20140220 - only do for IP, OP or new patient
                    If strCurrentStatus = "IP" Or strCurrentStatus = "OP" Or strCurrentStatus = "" Then

                        'if the department is LJR or LJH don't do anything
                        If dictNVP.Item("department") = "LJH" Or dictNVP.Item("department") = "LJR" Then

                            strNewStatus = strCurrentStatus

                            '20160829 - Set status to OA if patient type is COQ
                        ElseIf dictNVP.Item("Patient Type") = "COQ" Then

                            strNewStatus = "OA"

                        Else
                            'Lookup the the status from the [130PatientStatus] table

                            sql = "select status from [130PatientStatus] where facility = '" & dictNVP.Item("Sending Facility") & "' "
                            sql = sql & "AND department = '" & dictNVP.Item("department") & "' "
                            sql = sql & "AND PatientClass = '" & dictNVP.Item("Patient Class") & "'"
                            objCommand.CommandText = sql
                            myConnection.Open()
                            dataReader = objCommand.ExecuteReader()

                            If dataReader.HasRows Then
                                While dataReader.Read()
                                    strNewStatus = dataReader.GetString(0)
                                End While
                            Else
                                If strCurrentStatus = "" And dictNVP.Item("Patient Class") = "P" Then strNewStatus = "IP"
                                If strCurrentStatus = "" And dictNVP.Item("Patient Class") = "I" Then strNewStatus = "IA"
                                'strNewStatus = "" - 20160125 - removed.
                            End If 'If dataReader.HasRows

                            myConnection.Close()
                            dataReader.Close()


                        End If 'If dictNVP.Item("department") = "LJH" Or dictNVP.Item("department") = "LJR"


                    Else
                        strNewStatus = strCurrentStatus
                    End If 'If strCurrentStatus = "IP" Or strCurrentStatus = "OP" Or strCurrentStatus = ""





                Case "A03", "A11" '20140318
                    'change second character to a "D"
                    strNewStatus = Left(strCurrentStatus, 1) & "D"

                    'Case "A11" ' cancel admission
                    'If dictNVP.Item("Sending Facility") = "Q" And dictNVP.Item("Patient Class") = "O" Then strNewStatus = "OP"
                    'If dictNVP.Item("Sending Facility") = "Q" And dictNVP.Item("Patient Class") = "R" Then strNewStatus = "OP"
                    'If dictNVP.Item("Sending Facility") = "R" And dictNVP.Item("Patient Class") = "I" Then strNewStatus = "IP"
                    'If dictNVP.Item("Sending Facility") = "H" And dictNVP.Item("Patient Class") = "I" Then strNewStatus = "IP"

                    '20140311 - add A13 cancel discharge
                Case "A13"
                    If dictNVP.Item("Sending Facility") = "Q" And dictNVP.Item("Patient Class") = "O" Then strNewStatus = "OA"
                    If dictNVP.Item("Sending Facility") = "Q" And dictNVP.Item("Patient Class") = "R" Then strNewStatus = "OA"
                    '20170623 - Add T for ULHT
                    If dictNVP.Item("Sending Facility") = "T" And dictNVP.Item("Patient Class") = "O" Then strNewStatus = "OA"
                    If dictNVP.Item("Sending Facility") = "T" And dictNVP.Item("Patient Class") = "R" Then strNewStatus = "OA"

                    If dictNVP.Item("Sending Facility") = "R" And dictNVP.Item("Patient Class") = "I" Then strNewStatus = "IA"
                    If dictNVP.Item("Sending Facility") = "H" And dictNVP.Item("Patient Class") = "I" Then strNewStatus = "IA"

                Case "A08" '20140424 - added A08 case
                    If dictNVP.Item("Sending Facility") = "R" And dictNVP.Item("Patient Class") = "P" And dictNVP.Item("department") = "FIPR" Then strNewStatus = "IP"
                    If dictNVP.Item("Sending Facility") = "R" And dictNVP.Item("Patient Class") = "I" And dictNVP.Item("department") = "FIPR" Then strNewStatus = "IA"
                    If dictNVP.Item("Sending Facility") = "R" And dictNVP.Item("Patient Class") = "O" And dictNVP.Item("department") = "LJR" Then strNewStatus = "IA"
                    If dictNVP.Item("Sending Facility") = "H" And dictNVP.Item("Patient Class") = "O" And dictNVP.Item("department") = "LJH" Then strNewStatus = "IA"

                    If dictNVP.Item("Sending Facility") = "H" And dictNVP.Item("Patient Class") = "P" And dictNVP.Item("department") = "ACT" Then strNewStatus = "IP"
                    If dictNVP.Item("Sending Facility") = "H" And dictNVP.Item("Patient Class") = "I" And dictNVP.Item("department") = "ACT" Then strNewStatus = "IA"
            End Select

            calcSTARStatus = strNewStatus

            '20150914 - determine if we should update the admission date on an A08 by computing the value of gblBoolUpdateAdmDate
            If tempCurrentStatus = "OP" And strNewStatus = "OA" Then
                gblBoolUpdateAdmDate = True
            ElseIf tempCurrentStatus = "IP" And strNewStatus = "IA" Then
                gblBoolUpdateAdmDate = True ' 20151130 added
            Else
                gblBoolUpdateAdmDate = False
            End If

        Catch ex As Exception

        End Try

    End Function
    Public Function extractMrnum(ByVal varMrnum As String) As String
        extractMrnum = "0"
        If IsNumeric(Left(varMrnum, 1)) Then
            extractMrnum = varMrnum
        Else
            extractMrnum = Mid(varMrnum, 2)
        End If
    End Function

    Public Function extractPanum(ByVal varPanum As String) As String
        extractPanum = "0"
        If IsNumeric(Left(varPanum, 1)) Then
            extractPanum = varPanum
        Else
            extractPanum = Mid(varPanum, 2)
        End If
    End Function

    Public Sub ProcessA34(ByVal dictNVP As Hashtable)
        '20141021 - write A34 Information to PatientGlobal database, table = A34Queue
        '20141204 - add oldcorpno as MRG_2_1
        'Dim A34connectionString As String = "server=10.48.242.249,1433;database=PatientGlobal;uid=sysmax;pwd=Condor!"
        'Dim A34connectionString As String = "server=10.48.64.5\sqlexpress;database=PatientGlobal;uid=sysmax;pwd=Condor!"
        Dim A34connectionString As String = conIniFile.GetString("Strings", "ITWA34", "(none)")

        Dim myConnection As New SqlConnection(A34connectionString)
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        Dim sql As String = ""
        Dim strCorpNo As String = dictNVP.Item("JHHS mrnum")
        Dim strMrNum As String = dictNVP.Item("mrnum")
        Dim strOldMRNUM As String = dictNVP.Item("oldmrnum")
        Dim strRegion As String = UCase(dictNVP("Sending Facility"))
        Dim strOldCorpNo As String = dictNVP.Item("oldcorpno")
        Try
            If strCorpNo <> "" And strMrNum <> "" And strOldMRNUM <> "" Then
                sql = "Insert [A34Queue] "
                sql = sql & "(facility, originalMRN, NewMRN, originalCorp, NewCorp, RequestedDate) "
                sql = sql & "VALUES ("
                sql = sql & "'" & strRegion & "', " & strOldMRNUM & ", " & strMrNum & ", " & strOldCorpNo & ", " & strCorpNo & ", "
                sql = sql & "'" & DateTime.Now & "') "
                updatecommand.CommandText = sql
                myConnection.Open()
                updatecommand.ExecuteNonQuery()
                myConnection.Close()
            End If
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "A34 Process Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try

    End Sub

    Public Sub ProcessA44(ByVal dictNVP As Hashtable)
        '20150916 - write A44 Information to PatientGlobal database, table = A44Queue
        '20150929 - convert panum fields to nvarchar from bigint
        '20150929 - added to production ITW feed.

        Dim sql As String = ""

        Dim strMrNum As String = dictNVP.Item("mrnum")
        Dim strOldPanum As String = dictNVP.Item("oldPaNum") '20150917 fixed spelling error
        Dim strPanum As String = dictNVP.Item("panum")
        Dim strRegion As String = UCase(dictNVP("Sending Facility"))

        Dim recordexists As Boolean = False

        Try
            Using connect As New SqlConnection(connectionString)
                Dim objdbcommand As New SqlCommand
                Dim dataReader As SqlDataReader
                With objdbcommand

                    .Connection = connect
                    .Connection.Open()

                    'check for duplicate A44
                    sql = " SELECT ID "
                    sql = sql & " FROM [A44Queue] "
                    sql = sql & " WHERE NewPanum = '" & strPanum & "' "
                    sql = sql & " AND OldPanum = '" & strOldPanum & "' "
                    .CommandText = sql
                    dataReader = objdbcommand.ExecuteReader()

                    If dataReader.HasRows Then
                        recordexists = True
                    End If

                End With
            End Using

            If strMrNum <> "" AndAlso strOldPanum <> "" AndAlso strPanum <> "" AndAlso Not recordexists Then

                Using insertconnect As New SqlConnection(connectionString)
                    Dim insertcommand As New SqlCommand
                    With insertcommand
                        .Connection = insertconnect
                        .Connection.Open()

                        sql = "Insert [A44Queue] "
                        sql = sql & "(facility, mrnum, NewPanum, OldPanum,  RequestedDate) "
                        sql = sql & "VALUES ("
                        sql = sql & "'" & strRegion & "', " & strMrNum & ", '" & strPanum & "', '" & strOldPanum & "', "
                        sql = sql & "'" & DateTime.Now & "') "
                        .CommandText = sql
                        .ExecuteNonQuery()

                    End With
                End Using

            End If
        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "A44 Process Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf
            'LogFile.Close()
            Exit Sub
        End Try

    End Sub

    '20170329
    Public Sub ProcessIN1_14(ByVal dictNVP As Hashtable, ByVal tempstr As String)
        'Dim Int As Integer
        'Dim tempStr As String = ""
        'For Int = 1 To gblInsCount
        'If Int = 1 Then
        'tempStr = ""
        'End If
        'If Int > 1 Then
        'tempstr = "_000" & Int
        'End If

        'check to see if records exist
        'Call dbo.smc_InsAuthSelect
        'result contains rows update else insert
        Dim objDBCommand As New SqlCommand
        Dim objDBCommand2 As New SqlCommand
        Dim objDBCommand3 As New SqlCommand
        Dim objDBCommand4 As New SqlCommand
        Dim dreader As SqlDataReader
        'Dim dreader2 As SqlDataReader
        Dim sql As String = ""
        Dim IN114array()
        'STAR_Plancode = Trim(Replace(dictNVP("iplancode2" & tempStr), "'", "''")) & Trim(Replace(dictNVP("iplancode" & tempStr), "'", "''"))

        Dim plancode As String = dictNVP.Item("iplancode" & tempstr)
        Dim plancode2 As String = dictNVP.Item("iplancode2" & tempstr)
        Dim fullcode As String = plancode2 & plancode
        Dim panum As String = dictNVP.Item("panum")
        Dim I As String = ""
        Dim ID As Integer
        Dim insert As Boolean = True

        Using conn As New SqlConnection(connectionString)
            With objDBCommand

                .Connection = conn
                .Connection.Open()

                sql = "Select ID "
                sql += " FROM [03Insurer] i "
                sql += " INNER JOIN [001Episode] e on e.epnum  = i.epnum "
                sql += " WHERE i.iplancode = '" & fullcode & "' "
                sql += " and e.panum = '" & panum & "'"
                .CommandText = sql
                dreader = objDBCommand.ExecuteReader()
                While dreader.Read
                    I = dreader("ID")
                End While

            End With
        End Using
        If I <> "" Then
            Using conn As New SqlConnection(connectionString)
                ID = Convert.ToInt32(I)
                With objDBCommand2
                    .Connection = conn
                    .Connection.Open()

                    sql += "DELETE FROM [03InsAuthReceive] "
                    sql += " WHERE INSID = '" & ID & "' "
                    .CommandText = sql
                    objDBCommand2.ExecuteNonQuery()
                End With
            End Using

            Using conn As New SqlConnection(connectionString)
                With objDBCommand3
                    .Connection = conn
                    .Connection.Open()

                    'Dim position As Integer
                    Dim fromDate = DBNull.Value
                    Dim toDate = DBNull.Value
                    Dim Authcode = DBNull.Value
                    Dim IN114 As String = dictNVP("AuthNum" & tempstr)
                    IN114array = IN114.Split("^")

                    If IN114array(0) <> "" Then
                        Authcode = IN114array(0)
                    End If

                    .CommandText = "dbo.smc_InsAuthUpdateReceive"
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@InsID", ID)
                    .Parameters.AddWithValue("@positionNum", 1)
                    .Parameters.AddWithValue("@AuthCode", Authcode)
                    .Parameters.AddWithValue("@fromDate", fromDate)
                    .Parameters.AddWithValue("@toDate", toDate)
                    .Parameters.AddWithValue("@insert", True)

                    objDBCommand3.ExecuteNonQuery()

                End With

            End Using


            Using conn As New SqlConnection(connectionString)
                With objDBCommand4
                    .Connection = conn
                    .Connection.Open()

                    Dim ZGI2 As String = dictNVP("Additional Auths" & tempstr)
                    If Not ZGI2 Is Nothing Then
                        Dim ZGI() = ZGI2.Split("~")
                        Dim position As Integer = 1

                        Dim valueArray() As String
                        For Each value As String In ZGI
                            Dim fromDate = DBNull.Value
                            Dim toDate = DBNull.Value
                            Dim Authcode = DBNull.Value

                            position += 1

                            If value <> "" Then

                                Dim cnt As Integer = 0
                                For Each c As Char In value
                                    If c = "^" Then
                                        cnt += 1
                                    End If
                                Next


                                If cnt = 2 Then
                                    valueArray = value.Split("^")
                                    If valueArray(0) <> "" Then
                                        Authcode = valueArray(0)
                                    End If
                                    If valueArray(1) <> "" Then
                                        fromDate = valueArray(1)
                                    End If
                                    If valueArray(2) <> "" Then
                                        toDate = valueArray(2)
                                    End If
                                ElseIf cnt = 1 Then
                                    valueArray = value.Split("^")
                                    If valueArray(0) <> "" Then
                                        Authcode = valueArray(0)
                                    End If
                                    If valueArray(1) <> "" Then
                                        fromDate = valueArray(1)
                                    End If
                                ElseIf cnt = 0 Then
                                    valueArray = value.Split("^")
                                    If valueArray(0) <> "" Then
                                        Authcode = valueArray(0)
                                    End If

                                End If

                            End If


                            .CommandText = "dbo.smc_InsAuthUpdateReceive"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.Clear()
                            .Parameters.AddWithValue("@InsID", ID)
                            .Parameters.AddWithValue("@positionNum", position)
                            .Parameters.AddWithValue("@AuthCode", Authcode)
                            .Parameters.AddWithValue("@fromDate", fromDate)
                            .Parameters.AddWithValue("@toDate", toDate)
                            .Parameters.AddWithValue("@insert", True)

                            objDBCommand4.ExecuteNonQuery()

                        Next
                    End If
                End With

            End Using

            'Using conn As New SqlConnection(connectionString)
            '    With objDBCommand3
            '        .Connection = conn
            '        .Connection.Open()

            '        Dim IN114 As String = dictNVP("AuthNum")
            '        Dim IN114array() = IN114.Split("~")
            '        Dim position As Integer
            '        Dim fromDate = DBNull.Value
            '        Dim toDate = DBNull.Value
            '        Dim Authcode = DBNull.Value
            '        Dim valueArray() As String
            '        For Each value As String In IN114array
            '            position += 1


            '            If position = 1 Then
            '                Authcode = Replace(value, "^", "")

            '            Else
            '                valueArray = value.Split("^")
            '                If valueArray(0) <> "" Then
            '                    Authcode = valueArray(0)
            '                End If
            '                If value(1) <> "" Then
            '                    fromDate = valueArray(1)
            '                End If
            '                If value(2) <> "" Then
            '                    toDate = valueArray(2)
            '                End If
            '            End If
            '            .CommandText = "dbo.smc_InsAuthUpdateReceive"
            '            .CommandType = CommandType.StoredProcedure
            '            .Parameters.Clear()
            '            .Parameters.AddWithValue("@InsID", ID)
            '            .Parameters.AddWithValue("@positionNum", position)
            '            .Parameters.AddWithValue("@AuthCode", Authcode)
            '            .Parameters.AddWithValue("@fromDate", fromDate)
            '            .Parameters.AddWithValue("@toDate", toDate)
            '            .Parameters.AddWithValue("@insert", True)

            '            objDBCommand3.ExecuteNonQuery()

            '        Next
            '    End With

            'End Using
        End If
        'Next
    End Sub

    '20180305 - Move files that do not have an observation datetime to the NO OBS Datetime Folder
    Private Sub checkZ47(theFile As FileInfo, myfile As StreamReader)
        Dim newProblemDir As String = strOutputDirectory & "Z47\"

        If Not Directory.Exists(newProblemDir) Then
            Directory.CreateDirectory(newProblemDir)
        End If

        'make copy in the problems directory delete any previous ones with same name
        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "Z47\" & theFile.Name)
        fi2.Delete()
        theFile.CopyTo(strOutputDirectory & "Z47\" & theFile.Name)

        'get rid of the file so it doesn't mess up the next run.
        myfile.Close()
        If theFile.Exists Then
            theFile.Delete()
        End If
    End Sub

End Module
