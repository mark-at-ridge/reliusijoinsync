Imports System.Data.SQLite
Imports System.Text
Imports System.IO

Public Class clsDatabase
    Function getdatatable(sql$) As DataTable



        Dim connectionstring As String = "Data Source=" & My.Settings.ijoindatabase
        Dim msql As String = sql$ '" select *  from fundallocations "
        Dim dt As DataTable = Nothing
        Dim ds As New DataSet
        Try
            Using con As New SQLiteConnection(connectionstring)
                Using cmd As New SQLiteCommand(msql, con)
                    con.Open()
                    Using da As New SQLiteDataAdapter(cmd)
                        da.Fill(ds)
                        dt = ds.Tables(0)
                    End Using
                End Using
            End Using


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return dt
    End Function
    Function doQuery(sql$) As Int32

        Dim rowcount& = 0

        Dim connectionstring As String = "Data Source=" & My.Settings.ijoindatabase
        Dim msql As String = sql$ '" select *  from fundallocations "

        Dim ds As New DataSet
        Try
            Using con As New SQLiteConnection(connectionstring)
                Using cmd As New SQLiteCommand(msql, con)
                    con.Open()

                    rowcount& = cmd.ExecuteNonQuery()
                End Using
            End Using


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return rowcount&
    End Function
    Function impParticipant(strfilename$, strd$, strplan$, b_addtosqlite As Boolean)
        ' change table name 
        Dim intretval& = 0
        Dim startsync As DateTime = Date.Now
        Try
            '  Dim dt As DataTable '= basSQL.getdatatableNetFrame("select * from Census2 where 0 = 1")
            '  Dim strfilename$
            Dim c% = 0

1:          Form1.updatestatus("Importing Participants...", "EXPORT")
            Dim dblhash# = 0

            '   ' change text box name 
2:          '  
            If System.IO.File.Exists(strfilename$) Then
                'The file exists
            Else
                MsgBox("No file " & strfilename$)
                Return 0
                Exit Function
            End If
            Form1.updatestatus("File to Export - " & strfilename$ & " Dated - " & FileDateTime(strfilename$), "EXPORT")
            Dim m_Participants As New SortedList
            Dim c_Participant As clsParticipant
3:          Using Reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(strfilename$)


                '  ' change field widths 
                Reader.TextFieldType =
                     Microsoft.VisualBasic.FileIO.FieldType.Delimited
                Reader.SetDelimiters(",")
                Dim currentRow As String()
4:              currentRow = Reader.ReadFields()

                ' build list since we have to roll these up before inserting


5:              While Not Reader.EndOfData
                    Try
                        'Dim drnew As DataRow = dt.NewRow
                        'drnew(0) = strplan$
                        c% = 0
                        currentRow = Reader.ReadFields()
                        Dim currentField As String = ""
                        Dim strkey$ = currentRow(0).ToString & currentRow(1).ToString & currentRow(3).ToString
10:                     If Not m_Participants.ContainsKey(strkey$) Then
                            c_Participant = New clsParticipant
                            With c_Participant
                                .employerein = currentRow(0)
                                .plancontractnumber = currentRow(1)
                                .planname = currentRow(2)
                                .ssn = currentRow(3)
                                .dob = currentRow(4)
                                .doh = currentRow(5)
                                .firstname = currentRow(6)
                                .lastname = currentRow(7)
                                .middlename = currentRow(8)
                                .address1 = currentRow(9)
                                .address2 = currentRow(10)
                                .city = currentRow(11)
                                .state = currentRow(12)
                                .zip = currentRow(13)
                                .phone = currentRow(14)
                                .phone2 = currentRow(15)
                                .email = currentRow(16)
                                .email2 = currentRow(17)
                                .marital = currentRow(18)
                                .gender = currentRow(19)
                                .contactmethod = currentRow(20)
                                .annualcomp = currentRow(21)
                                .autoIncrRate = currentRow(22)
                                .autoIncrMax = currentRow(23)
                                .preTaxBalance = currentRow(24)
                                .RothBalance = currentRow(25)
                                .afterTaxBalance = currentRow(26)
                                .contribPreTaxRate = currentRow(27)
                                .contribRothRate = currentRow(28)
                                .contribAfterTaxRate = currentRow(29)
                                .pensionAmountMonthly = currentRow(30)
                                .pensionStartage = currentRow(31)
                                .payPeriods = currentRow(32)
                            End With
11:                         m_Participants.Add(strkey$, c_Participant)
                        End If

                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                End While
            End Using
12:         Form1.updatestatus("Found - " & m_Participants.Count & " in File", "EXPORT")

            Dim rows% = 0
13:         Form1.updatestatus("Adding Participants for Plan - " & strplan & " ...", "EXPORT")
            If b_addtosqlite Then
15:             rows% = insertparticipants(m_Participants, strplan$)

16:             ' dt = getdatatable("select * from Participant")
            Else
17:             ' dt = buildparticipants(m_Participants)

            End If
18:     


24:         dblhash# = m_Participants.Count
25:         clsdb.insert_activity(0, startsync, Date.Now, "IMPORT " & strplan$ & " Parts", dblhash#, strplan$, 0)

            Return rows%
        Catch ex As Exception
            MsgBox("census2 " & Erl() & " " & ex.Message)
            Return -1
        End Try
        Form1.updatestatus("Done with Participants", "EXPORT")
    End Function
    Function getparttable() As String

        Dim sql$ = " select employerEIN,planContractNumber,planName,ssn,substr(dob,0) as dob,substr(doh,0),firstName,lastName,middleName,address1,address2,city,state,zip,phone,phone2," _
        & "email,email2,marital,gender,contactmethod,annualComp,autoIncrRate,autoIncrMax,preTaxBalance,rothBalance,afterTaxBalance,contribPreTaxRate, " _
        & "contribRothRate,contribAfterTaxRate,pensionAmountMonthly,pensionStartAge,payPeriods from participant"


        Return sql$
    End Function
    Function writefile(startsync As DateTime, strtable$, strfilename$) As Int16
        Dim intreturn% = 0
        Try
            Form1.updatestatus("Adding " & strtable$ & "...", "EXPORT")
            Dim sql$ = ""
            Select Case strtable$.ToLower
                Case "participant"
                    sql$ = getparttable()

                Case Else
                    sql$ = " select * from " & strtable$
            End Select
            sql$ += " where plancontractnumber in (select planid from plangroups where active = 1)"
16:         Dim dt As DataTable = getdatatable(sql$)

18:         Form1.updatestatus("Exporting " & strtable$ & "...", "EXPORT")
19:         Dim strexportfilename$ = strfilename$

            ' now lets get rid of the planid 

20:
            If File.Exists(strexportfilename) Then
                Debug.Print("")
            End If
            ' append 
            Dim writer As System.IO.StreamWriter
            If File.Exists(strexportfilename) = False Then
                writer = File.CreateText(strfilename)

            Else
                ' writer = File.OpenWrite(strfilename)
                File.Delete(strexportfilename)
                writer = File.CreateText(strfilename)
            End If
            '   writer = New StreamWriter(strexportfilename$, True)

21:         Using writer ' As StreamWriter = New StreamWriter(strexportfilename$, False)
22:             Rfc4180Writer.WriteDataTable(dt, writer, True)
                writer.Close()
            End Using



23:         Debug.Print(dt.Rows.Count)

            dt.TableName = strtable$

            Form1.updatestatus("Created Export File " & strexportfilename$ & "...", "EXPORT")
24:         Dim dblhash# = dt.Rows.Count
25:         clsdb.insert_activity(0, startsync, Date.Now, "Write " & strtable$, dblhash#, strexportfilename$, 0)
            Return dblhash#
        Catch ex As Exception
            MsgBox(ex.Message)
            Return -1
        End Try
    End Function
    Function writeenrollmentfile(startsync As DateTime, strtable$, strfilename$, dt As DataTable) As Int16
        Dim intretval% = 0
        Try
            Form1.updatestatus("Adding " & strtable$ & "...", "IMPORT")


16:

18:         Form1.updatestatus("Exporting " & strtable$ & "...", "IMPORT")
19:         Dim strexportfilename$ = strfilename$ 'Replace(strfilename, ".csv", My.Settings.ijoinclientid.ToString.ToUpper & ".csv")

            ' now lets get rid of the planid 

20:
            If File.Exists(strexportfilename) Then
                Debug.Print("")
            End If
            ' append 
            Dim writer As System.IO.StreamWriter
            If File.Exists(strexportfilename) = False Then
                writer = File.CreateText(strfilename)

            Else
                ' writer = File.OpenWrite(strfilename)
                File.Delete(strexportfilename)
                writer = File.CreateText(strfilename)
            End If
            '   writer = New StreamWriter(strexportfilename$, True)

21:         Using writer ' As StreamWriter = New StreamWriter(strexportfilename$, False)
22:             Rfc4180Writer.WriteDataTable(dt, writer, True)
                writer.Close()
            End Using



23:         Debug.Print(dt.Rows.Count)
            ' change table name 
            dt.TableName = strtable$

            Form1.updatestatus("Created DER " & strexportfilename$ & " for Import", "IMPORT")
24:         Dim dblhash# = dt.Rows.Count
25:         clsdb.insert_activity(0, startsync, Date.Now, "Write " & strtable$, dblhash#, strexportfilename$, 0)
            intretval = dt.Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message)
            intretval = -1
        End Try
        Return intretval
    End Function
    Function impplaninfo(strfilename$, strd$, strplan$, b_addtosqlite As Boolean)
        ' change table name 
        Dim intretval& = 0
1:      Dim startsync As DateTime = Date.Now
        Try
            '  Dim dt As DataTable '= basSQL.getdatatableNetFrame("select * from Census2 where 0 = 1")
            '  Dim strfilename$
            Dim c% = 0

2:          Form1.updatestatus("Importing Plan Info...", "EXPORT")
            Dim dblhash# = 0

            '   ' change text box name 
            '   Dim strfilename$ = UltraTextEditor1.Text & strd$ & "\census.txt"
            If System.IO.File.Exists(strfilename$) Then
                'The file exists
            Else
                MsgBox("No file " & strfilename$)
                Return 0
                Exit Function
            End If
3:          Form1.updatestatus("File to Export - " & strfilename$ & " Dated - " & FileDateTime(strfilename$), "EXPORT")
            Dim m_planinfos As New SortedList
            Dim c_PlanInfo As clsPlaninfo
4:          Using Reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(strfilename$)


                '  ' change field widths 
                Reader.TextFieldType =
                     Microsoft.VisualBasic.FileIO.FieldType.Delimited
                Reader.SetDelimiters(",")
                Dim currentRow As String()
5:              currentRow = Reader.ReadFields()

                ' build list since we have to roll these up before inserting


                While Not Reader.EndOfData
                    Try
                        'Dim drnew As DataRow = dt.NewRow
                        'drnew(0) = strplan$
                        c% = 0
6:                      currentRow = Reader.ReadFields()
                        Dim currentField As String = ""
                        Dim strkey$ = currentRow(0).ToString & currentRow(1).ToString & currentRow(3).ToString & currentRow(7).ToString & currentRow(8).ToString & currentRow(12).ToString & currentRow(10).ToString & currentRow(11).ToString
                        If Not m_planinfos.ContainsKey(strkey$) Then
                            c_PlanInfo = New clsPlaninfo
7:                          With c_PlanInfo
                                .planid = currentRow(0)
                                .yrenddate = currentRow(1)
                                .planname = currentRow(2)
                                .companyname = currentRow(3)
                                .ein = currentRow(4)
                                .plantype = currentRow(5)
                                .planstatus = currentRow(6)
                                .payfreqcd = currentRow(7)
                                .payschednum = currentRow(8)
                                .payschedname = currentRow(9)
                                .begindate = currentRow(10)
                                .enddate = currentRow(11)
                                .payfreqseqnum = currentRow(12)
                                .inactivecd = currentRow(13)

                            End With
10:                         m_planinfos.Add(strkey$, c_PlanInfo)
                        End If

                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                End While
            End Using
11:         Form1.updatestatus("Found - " & m_planinfos.Count & " in File", "EXPORT")

            Dim rows% = 0
            Form1.updatestatus("Adding Plan Infos...", "EXPORT")
            If b_addtosqlite Then
12:             rows% = insertplaninfos(m_planinfos, strplan$)

                '  dt = getdatatable("select * from Participant")
            Else
                '  dt = buildparticipants(m_planinfos)

            End If
        

15:         dblhash# = m_planinfos.Count
16:         clsdb.insert_activity(0, startsync, Date.Now, "IMPORT PlanInfo " & strplan$, dblhash#, strfilename, 0)

            Return rows%
        Catch ex As Exception
            MsgBox("planinfo " & Erl() & " " & ex.Message)
            Return -1
        End Try
        Form1.updatestatus("Done with Planinfo", "EXPORT")
    End Function
    Function impenrollments(strfilename$, strpath$, strd$, b_addtosqlite As Boolean)
        ' change table name 
        Dim intretval& = 0
        Dim startsync As DateTime = Date.Now
        Try
            '  Dim dt As DataTable '= basSQL.getdatatableNetFrame("select * from Census2 where 0 = 1")
            '  Dim strfilename$
            Dim c% = 0

            Form1.updatestatus("Importing Enrollments...", "IMP")
            Dim dblhash#
            '   ' change text box name 
            '   Dim strfilename$ = UltraTextEditor1.Text & strd$ & "\census.txt"
            If System.IO.File.Exists(strpath$ & strfilename$) Then
                'The file exists
            Else
                MsgBox("No file " & strfilename$)
                Return 0
                Exit Function
            End If
            Form1.updatestatus("File to Export - " & strfilename$ & " Dated - " & FileDateTime(strpath$ & strfilename$), "IMPORT")
            Dim m_enrollments As New SortedList
            Dim c_enrollment As clsEnrollment
            Using Reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(strpath$ & strfilename$)


                '  ' change field widths 
                Reader.TextFieldType =
                     Microsoft.VisualBasic.FileIO.FieldType.Delimited
                Reader.SetDelimiters(",")
                Dim currentRow As String()
                currentRow = Reader.ReadFields()

                ' build list since we have to roll these up before inserting


                While Not Reader.EndOfData
                    Try
                        'Dim drnew As DataRow = dt.NewRow
                        'drnew(0) = strplan$
                        c% = 0
                        currentRow = Reader.ReadFields()
                        Dim currentField As String = ""
                        Dim strkey$ = currentRow(0).ToString & currentRow(1).ToString & currentRow(3).ToString
                        If Not m_enrollments.ContainsKey(strkey$) Then
                            c_enrollment = New clsEnrollment
                            With c_enrollment
                                .employerEIN = currentRow(0)
                                .plancontractnumber = currentRow(1)
                                .planname = currentRow(2)
                                .ssn = currentRow(3)
                                .contribpretaxpercent = IIf(currentRow(4).ToString = "", 0, currentRow(4))
                                .autoIncrRate = IIf(currentRow(5).ToString = "", 0, currentRow(5))
                                .autoIncrMax = IIf(currentRow(6).ToString = "", 0, currentRow(6))
                                .contribRothpercent = IIf(currentRow(7).ToString = "", 0, currentRow(7))
                                .contribaftertaxpercent = IIf(currentRow(8).ToString = "", 0, currentRow(8))
                                .rebalance = currentRow(9)
                            End With
                            m_enrollments.Add(strkey$, c_enrollment)
                        End If
                        Form1.updatestatus("Found - " & m_enrollments.Count & " in File", "IMPORT")
                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                End While
            End Using

            Dim rows% = 0
            Form1.updatestatus("Adding Enrollments...", "IMP")
            If b_addtosqlite Then
                rows% = insertenrollments(m_enrollments, strfilename)

                '   dt = getdatatable("select * from enrollments")
            Else
                '   dt = buildfundallocations(m_enrollments)

            End If
           

            dblhash# = m_enrollments.Count
           
            clsdb.insert_activity(0, startsync, Date.Now, "IMPORT Enrollments", dblhash#, strpath$ & strfilename$, 0)

            Return rows%
        Catch ex As Exception
            MsgBox("impenroll " & ex.Message)
            Return -1
        End Try
        Form1.updatestatus("Done with Enrollments", "IMP")
    End Function
    Function buildEnrollmentFile(strplan$) As Int16
        Dim intreturn% = 0
        Try

            Dim dt As DataTable = clsdb.getdatatable(" select employerein,plancontractnumber,planname,ssn,contribpretaxrate,autoincrrate,autoincrmax,contribrothrate,contribaftertaxrate,rebalance , enrollmentid from enrollments where plancontractnumber = '" & strplan$ & "' and exported = 0 ")
            Dim strfilename$ = My.Settings.syncdirectory & "enrollment_" & strplan$ & ".csv"
            Dim strreliusimportfilename$ = My.Settings.reliusimportdirectory & "enrollment_" & strplan$ & ".csv"
            Dim intretval As Int16 = writeenrollmentfile(Date.Now, "enrollments", strfilename$, dt)
            If intretval > 0 Then
                For Each dr As DataRow In dt.Rows
                    Dim sql$ = " update enrollments set exported = 1, exporteddate = '" & Format(Date.Now, "yyyy-MM-dd hh:mm") & "' , xmlcreated = 0 , exportfilename = '" & strreliusimportfilename$ & "'  where enrollmentid = " & dr("enrollmentid")
                    intreturn% = doQuery(sql$)
                Next

            End If
            Return intreturn%
        Catch ex As Exception
            MsgBox("buildenrollment " & strplan$ & " " & ex.Message)
            Return -1
        End Try
    End Function
    Function buildEnrollmentAllocationFile(strplan$) As Int16
        Dim intreturn% = 0
        Try

            Dim dt As DataTable = clsdb.getdatatable(" select * from enrollmentallocations where plancontractnumber = '" & strplan$ & "' and exported = 0 ")
            Dim strfilename$ = My.Settings.syncdirectory & "enrollmentallocations_" & strplan$ & ".csv"
            Dim strreliusimportfilename$ = My.Settings.reliusimportdirectory & "enrollmentallocations_" & strplan$ & ".csv"
            Dim intretval As Int16 = writeenrollmentfile(Date.Now, "enrollmentallocations", strfilename$, dt)
            If intretval > 0 Then
                For Each dr As DataRow In dt.Rows
                    Dim sql$ = " update enrollmentallocations set exported = 1, exporteddate = '" & Format(Date.Now, "yyyy-MM-dd hh:mm") & "' , xmlcreated = 0 , exportfilename = '" & strreliusimportfilename$ & "'  where enrollmentid = " & dr("enrollmentid")
                    intreturn% += doQuery(sql$)
                Next

            End If
            Return intreturn%
        Catch ex As Exception
            MsgBox("buildenrollment " & strplan$ & " " & ex.Message)
            Return intreturn%
        End Try
    End Function
    Function impenrollmentallocations(strfilename$, strpath$, strd$, b_addtosqlite As Boolean) As Int16
        ' change table name 
        Dim intretval& = 0
        Dim startsync As DateTime = Date.Now
        Try
            '  Dim dt As DataTable '= basSQL.getdatatableNetFrame("select * from Census2 where 0 = 1")
            '  Dim strfilename$
            Dim c% = 0

            Form1.updatestatus("Importing Enrollment Allocations...", "IMP")
1:          Dim dblhash#
            '   ' change text box name 
            '   Dim strfilename$ = UltraTextEditor1.Text & strd$ & "\census.txt"
            If System.IO.File.Exists(strpath$ & strfilename$) Then
                'The file exists
            Else
                MsgBox("No file " & strfilename$)
                Return 0
                Exit Function
            End If
2:          Form1.updatestatus("File to Export - " & strfilename$ & " Dated - " & FileDateTime(strpath$ & strfilename$), "IMPORT")
            Dim m_enrollments As New SortedList
            Dim c_enrollmentallocation As clsEnrollmentAllocation
            Using Reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(strpath$ & strfilename$)


                '  ' change field widths 
3:              Reader.TextFieldType =
                     Microsoft.VisualBasic.FileIO.FieldType.Delimited
                Reader.SetDelimiters(",")
                Dim currentRow As String()
                currentRow = Reader.ReadFields()

                ' build list since we have to roll these up before inserting


5:              While Not Reader.EndOfData
                    Try
                        'Dim drnew As DataRow = dt.NewRow
                        'drnew(0) = strplan$
                        c% = 0
                        currentRow = Reader.ReadFields()
                        Dim currentField As String = ""
                        Dim strkey$ = currentRow(0).ToString & currentRow(1).ToString & currentRow(3).ToString & currentRow(4).ToString
6:                      If strkey.ToString.Trim > "" Then
                            If Not m_enrollments.ContainsKey(strkey$) Then
7:                              c_enrollmentallocation = New clsEnrollmentAllocation
                                With c_enrollmentallocation
                                    .employerEIN = currentRow(0)
                                    .plancontractnumber = currentRow(1)
                                    .planname = currentRow(2)
                                    .ssn = currentRow(3)
                                    .ticker = IIf(currentRow(4).ToString = "", 0, currentRow(4))
                                    .allocationpercent = IIf(currentRow(5).ToString = "", 0, currentRow(5))
                                    .datestarted = IIf(currentRow(6).ToString = "", 0, currentRow(6))

                                End With
                                m_enrollments.Add(strkey$, c_enrollmentallocation)
                            End If
                        End If
8:                      Form1.updatestatus("Found - " & m_enrollments.Count & " in File", "IMPORT")
                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                End While
            End Using

10:         Dim rows% = 0
            Form1.updatestatus("Adding Enrollment Allocations...", "IMP")
11:         If b_addtosqlite Then
12:             rows% = insertenrollmentallocations(m_enrollments, strfilename)

                '  dt = getdatatable("select * from enrollments")
            Else
                ' dt = buildfundallocations(m_enrollments)

            End If
            '    Form1.updatestatus("Exporting Fund Allocations...", "EXPORT")
            '    Using writer As StreamWriter = New StreamWriter("d:\ijoin\export\fundallocations.csv")
            ' Rfc4180Writer.WriteDataTable(dt, writer, True)
            ' End Using

            dblhash# = m_enrollments.Count

            clsdb.insert_activity(0, startsync, Date.Now, "IMPORT EnrollmentAllocs", dblhash#, strpath$ & strfilename$, 0)
            '      basSQL.bulkinsert(dt, strplan, "", "ASC", "", "", "", strplan, "")
            Return rows%
        Catch ex As Exception
            MsgBox("impenrollalloc " & Erl() & " " & ex.Message)
            Return -1
        End Try
        Form1.updatestatus("Done with Enrollments", "IMP")
    End Function
    Function impfundallocations(strfilename$, strd$, strplan$, b_addtosqlite As Boolean)
        ' change table name 
        Dim intretval& = 0
        Dim startsync As DateTime = Date.Now
        Try
            '  Dim dt As DataTable '= basSQL.getdatatableNetFrame("select * from Census2 where 0 = 1")
            '  Dim strfilename$
            Dim c% = 0

            Form1.updatestatus("Importing Fund Allocations...", "EXPORT")
            Dim dblhash#
            '   ' change text box name 
            '   Dim strfilename$ = UltraTextEditor1.Text & strd$ & "\census.txt"
            If System.IO.File.Exists(strfilename$) Then
                'The file exists
            Else
                MsgBox("No file " & strfilename$)
                Return 0
                Exit Function
            End If
            Form1.updatestatus("File to Export - " & strfilename$ & " Dated - " & FileDateTime(strfilename$), "EXPORT")
            Dim m_allocations As New SortedList
            Dim c_allocation As clsFundAllocations
            Using Reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(strfilename$)


                '  ' change field widths 
                Reader.TextFieldType =
                     Microsoft.VisualBasic.FileIO.FieldType.Delimited
                Reader.SetDelimiters(",")
                Dim currentRow As String()
                currentRow = Reader.ReadFields()

                ' build list since we have to roll these up before inserting


                While Not Reader.EndOfData
                    Try
                        'Dim drnew As DataRow = dt.NewRow
                        'drnew(0) = strplan$
                        c% = 0
                        currentRow = Reader.ReadFields()
                        Dim currentField As String = ""
                        Dim strkey$ = currentRow(0).ToString & currentRow(1).ToString & currentRow(3).ToString & currentRow(4).ToString
                        If Not m_allocations.ContainsKey(strkey$) Then
                            c_allocation = New clsFundAllocations
                            With c_allocation
                                .employerein = currentRow(0)
                                .plancontractnumber = currentRow(1)
                                .planname = currentRow(2)
                                .ssn = currentRow(3)
                                .ticker = currentRow(4)
                                .currentallocation = currentRow(5)
                                .currentbalance = currentRow(6)
                                .datestarted = currentRow(7)
                            End With
                            m_allocations.Add(strkey$, c_allocation)
                        End If

                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                End While
            End Using
            Form1.updatestatus("Found - " & m_allocations.Count & " in File", "EXPORT")

            Dim rows% = 0
            Form1.updatestatus("Adding Fund Allocations...", "EXPORT")
            If b_addtosqlite Then
                rows% = insertfundallocations(m_allocations, strplan$)

                '      dt = getdatatable("select * from fundallocations")
            Else
                '  dt = buildfundallocations(m_allocations)

            End If
            '    Form1.updatestatus("Exporting Fund Allocations...", "EXPORT")
            '   Dim strexportfilename$ = Replace(strfilename, ".csv", My.Settings.ijoinclientid & ".csv")
            '    Using writer As StreamWriter = New StreamWriter(strexportfilename$, True)
            'Rfc4180Writer.WriteDataTable(dt, writer, True)
            '     End Using

            dblhash# = m_allocations.Count
            '   Debug.Print(dt.Rows.Count)
            ' change table name 

            clsdb.insert_activity(0, startsync, Date.Now, "IMPORT Fund Alloc " & strplan$, dblhash#, strplan$, 0)
            '      basSQL.bulkinsert(dt, strplan, "", "ASC", "", "", "", strplan, "")
            Return rows%
        Catch ex As Exception
            MsgBox("census2 " & ex.Message)
            Return -1
        End Try
        Form1.updatestatus("Done with Fund Allocations", "EXPORT")
    End Function
    Function impBeneficiary(strfilename$, strd$, strplan$, b_addtosqlite As Boolean)
        ' change table name 
        Dim intretval& = 0
        Dim startsync As DateTime = Date.Now
        Try
            '  Dim dt As DataTable '= basSQL.getdatatableNetFrame("select * from Census2 where 0 = 1")
            '  Dim strfilename$
            Dim c% = 0

            Form1.updatestatus("Importing Beneficiary...", "EXPORT")
            Dim dblhash#
            '   ' change text box name 
            '   Dim strfilename$ = UltraTextEditor1.Text & strd$ & "\census.txt"
            If System.IO.File.Exists(strfilename$) Then
                'The file exists
            Else
                MsgBox("No file " & strfilename$)
                Return 0
                Exit Function
            End If
            Form1.updatestatus("File to Export - " & strfilename$ & " Dated - " & FileDateTime(strfilename$), "EXPORT")
            Dim m_Beneficiary As New SortedList
            Dim c_Beneficiary As clsBeneficiary
            Using Reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(strfilename$)


                '  ' change field widths 
                Reader.TextFieldType =
                     Microsoft.VisualBasic.FileIO.FieldType.Delimited
                Reader.SetDelimiters(",")
                Dim currentRow As String()
                currentRow = Reader.ReadFields()

                ' build list since we have to roll these up before inserting


                While Not Reader.EndOfData
                    Try
                        'Dim drnew As DataRow = dt.NewRow
                        'drnew(0) = strplan$
                        c% = 0
                        currentRow = Reader.ReadFields()
                        Dim currentField As String = ""
                        Dim strkey$ = currentRow(0).ToString & currentRow(1).ToString & currentRow(3).ToString & currentRow(10).ToString
                        If Not m_Beneficiary.ContainsKey(strkey$) Then
                            c_Beneficiary = New clsBeneficiary
                            With c_Beneficiary
                                .employerein = currentRow(0)
                                .plancontractnumber = currentRow(1)
                                .planname = currentRow(2)
                                .ssn = currentRow(3)
                                .participantdob = currentRow(4)
                                .firstname = currentRow(5)
                                .lastname = currentRow(6)
                                .relationship = currentRow(7)
                                .percentAllocation = currentRow(8)
                                .benetype = currentRow(9)
                                .ssn = currentRow(10)
                                .address1 = currentRow(11)
                                .address2 = currentRow(12)
                                .city = currentRow(13)
                                .state = currentRow(14)
                                .zip = currentRow(15)
                                .phone = currentRow(16)
                                .dob = currentRow(17)
                            End With
                            m_Beneficiary.Add(strkey$, c_Beneficiary)
                        End If

                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                End While
            End Using
            Form1.updatestatus("Found - " & m_Beneficiary.Count & " in File", "EXPORT")

            Dim rows% = 0
            Form1.updatestatus("Adding Fund Allocations...", "EXPORT")
            If b_addtosqlite Then
                rows% = insertBeneficiary(m_Beneficiary, strplan$)

                '      dt = getdatatable("select * from fundallocations")
            Else
                '  dt = buildfundallocations(m_allocations)

            End If
            

            dblhash# = m_Beneficiary.Count
            '   Debug.Print(dt.Rows.Count)
            ' change table name 

            clsdb.insert_activity(0, startsync, Date.Now, "IMPORT Beneficiary " & strplan$, dblhash#, strplan$, 0)
            '      basSQL.bulkinsert(dt, strplan, "", "ASC", "", "", "", strplan, "")
            Return rows%
        Catch ex As Exception
            MsgBox("census2 " & ex.Message)
            Return -1
        End Try
        Form1.updatestatus("Done with Fund Allocations", "EXPORT")
    End Function
    Sub insert_errorlog(syncid&, errortype$, errormodule$, errordescription$, errornum&)
        Dim sqlstring$ = " insert into errorlog (syncid,errortype, errormodule, errordescription,errornumber, errormachine,errordate) " _
                    & " values (" & syncid& & ",'" & errortype$ & "','" & errormodule$ & "','" & errordescription & "'," & errornum& & ",'" & My.Computer.Name & "','" & Format(Date.Now, "yyyy-MM-dd hh:mm:ss tt") & "')"
        Me.doQuery(sqlstring$)
    End Sub
    Sub insert_activity(syncid&, startsync As DateTime, endsync As DateTime, synctype$, hashtotal As Double, filename$, haserrors As Int16)
        Dim sqlstring$ = " insert into activity (syncid,startsync, endsync, syncmachine,synctype, hashtotal,filename, haserrors) " _
                    & " values (" & syncid& & ",'" & Format(startsync, "yyyy-MM-dd hh:mm:ss tt") & "','" & Format(endsync, "yyyy-MM-dd hh:mm:ss tt") & "','" & My.Computer.Name & "','" & synctype & "'," & hashtotal# & ",'" & filename$ & "'," & haserrors & ")"
        Me.doQuery(sqlstring$)
    End Sub
    Function insertfundallocations(m_allocations As SortedList, strplan$) As Int32
        Dim rows& = 0

        doQuery("delete from fundallocations where plancontractnumber = '" & strplan$ & "'")

        For Each c_allocation In m_allocations.Values
            Dim strinsert As StringBuilder = New StringBuilder("Insert into fundallocations Values (")
            With c_allocation
                strinsert.Append("'" & .employerein & "',")
                strinsert.Append("'" & .plancontractnumber & "',")
                strinsert.Append("'" & .planname & "',")
                strinsert.Append("'" & .ssn & "',")
                strinsert.Append("'" & .ticker & "',")
                strinsert.Append("'" & .currentallocation & "',")
                strinsert.Append("'" & .currentbalance & "',")
                strinsert.Append("'" & Format(.datestarted, "yyyy-MM-dd") & "'")
                strinsert.Append(")")
            End With

            rows& += doQuery(strinsert.ToString)
        Next

        Return rows&

    End Function
    Function insertbeneficiary(m_beneficiary As SortedList, strplan$) As Int32
        Dim rows& = 0

        doQuery("delete from beneficiary where plancontractnumber = '" & strplan$ & "'")
        Dim c_beneficiary As clsBeneficiary
        For Each c_beneficiary In m_beneficiary.Values
            Dim strinsert As StringBuilder = New StringBuilder("Insert into beneficiary Values (")
            With c_beneficiary
                strinsert.Append("'" & .employerein & "',")
                strinsert.Append("'" & .plancontractnumber & "',")
                strinsert.Append("'" & .planname & "',")
                strinsert.Append("'" & .participantssn & "',")
                strinsert.Append("'" & .participantdob & "',")
                strinsert.Append("'" & .firstname & "',")
                strinsert.Append("'" & .lastname & "',")
                strinsert.Append("'" & .relationship & "',")
                strinsert.Append("'" & .percentAllocation & "',")
                strinsert.Append("'" & .benetype & "',")
                strinsert.Append("'" & .ssn & "',")
                strinsert.Append("'" & .address1 & "',")
                strinsert.Append("'" & .address2 & "',")
                strinsert.Append("'" & .city & "',")
                strinsert.Append("'" & .state & "',")
                strinsert.Append("'" & .zip & "',")
                strinsert.Append("'" & .phone & "',")
                strinsert.Append("'" & .dob & "' ")

                strinsert.Append(")")
            End With

            rows& += doQuery(strinsert.ToString)
        Next

        Return rows&

    End Function
    Function insertenrollments(m_enrollments As SortedList, strfilename$) As Int32
        Dim rows& = 0

        doQuery("delete from enrollments where filename = '" & strfilename$ & "'")

        For Each c_enrollment As clsEnrollment In m_enrollments.Values
            Dim strinsert As StringBuilder = New StringBuilder("Insert into enrollments Values (")
            With c_enrollment
                strinsert.Append("'" & .employerEIN & "',")
                strinsert.Append("'" & .plancontractnumber & "',")
                strinsert.Append("'" & .planname & "',")
                strinsert.Append("'" & .ssn & "',")
                strinsert.Append("'" & .contribpretaxpercent & "',")
                strinsert.Append("'" & .autoIncrRate & "',")
                strinsert.Append("'" & .autoIncrMax & "',")
                strinsert.Append("'" & .contribRothpercent & "',")
                strinsert.Append("'" & .contribaftertaxpercent & "',")
                strinsert.Append("'" & .rebalance & "',")
                strinsert.Append("'" & Format(Date.Today, "yyyy-MM-dd") & "',")
                strinsert.Append("'" & strfilename$ & "',")
                strinsert.Append("'" & 0 & "',")
                strinsert.Append("Null,")
                strinsert.Append("'',")
                strinsert.Append("'" & 0 & "',")
                strinsert.Append("Null,")
                strinsert.Append("'" & 0 & "',")
                strinsert.Append("Null,")
                strinsert.Append("'" & 0 & "'")
                strinsert.Append(")")
            End With

            rows& += doQuery(strinsert.ToString)
        Next

        Return rows&

    End Function
    Function insertenrollmentallocations(m_enrollments As SortedList, strfilename$) As Int32
        Dim rows& = 0
        Try

        
        doQuery("delete from enrollmentallocations where filename = '" & strfilename$ & "'")

        For Each c_enrollment As clsEnrollmentAllocation In m_enrollments.Values
            Dim strinsert As StringBuilder = New StringBuilder("Insert into enrollmentallocations (employerein, plancontractnumber, planname, ssn, ticker, allocationpercent, datestarted, filename, exported, exporteddate, xmlcreated,exportfilename ) Values (")
            With c_enrollment
                strinsert.Append("'" & .employerEIN & "',")
                strinsert.Append("'" & .plancontractnumber & "',")
                strinsert.Append("'" & .planname & "',")
                strinsert.Append("'" & .ssn & "',")
                strinsert.Append("'" & .ticker & "',")
                strinsert.Append("'" & .allocationpercent & "',")
                strinsert.Append("'" & .datestarted & "',")
                '   strinsert.Append("'" & Format(Date.Today, "yyyy-MM-dd") & "',")
                strinsert.Append("'" & strfilename$ & "',")

                strinsert.Append("'0',")
                strinsert.Append("Null,")
                strinsert.Append("'0',")
                strinsert.Append("''")
                ' strinsert.Append("'0'")



                strinsert.Append(")")
            End With

            rows& += doQuery(strinsert.ToString)
        Next


        Catch ex As Exception
            MsgBox("insertenrollalloc" & ex.Message)
        End Try
        Return rows&

    End Function
    Function insertparticipants(m_participants As SortedList, strplan$) As Int32
        Dim rows& = 0

        doQuery("delete from participant where plancontractnumber = '" & strplan$ & "'")

        For Each c_participant As clsParticipant In m_participants.Values
            Dim strinsert As StringBuilder = New StringBuilder("Insert into participant Values (")
            With c_participant
                strinsert.Append("'" & .employerein & "',")
                strinsert.Append("'" & .plancontractnumber & "',")
                strinsert.Append("'" & .planname & "',")
                strinsert.Append("'" & .ssn & "',")
                strinsert.Append("'" & .dob & "',")
                strinsert.Append("'" & .doh & "',")
                strinsert.Append("'" & .firstname & "',")
                strinsert.Append("'" & .lastname & "',")
                strinsert.Append("'" & .middlename & "',")
                strinsert.Append("'" & .address1 & "',")
                strinsert.Append("'" & .address2 & "',")
                strinsert.Append("'" & .city & "',")
                strinsert.Append("'" & .state & "',")
                strinsert.Append("'" & .zip & "',")
                strinsert.Append("'" & .phone & "',")
                strinsert.Append("'" & .phone2 & "',")
                strinsert.Append("'" & .email & "',")
                strinsert.Append("'" & .email2 & "',")
                strinsert.Append("'" & .marital & "',")
                strinsert.Append("'" & .gender & "',")
                strinsert.Append("'" & .contactmethod & "',")
                strinsert.Append("'" & .annualcomp & "',")
                strinsert.Append("'" & .autoIncrRate & "',")
                strinsert.Append("'" & .autoIncrMax & "',")
                strinsert.Append("'" & .preTaxBalance & "',")
                strinsert.Append("'" & .RothBalance & "',")
                strinsert.Append("'" & .afterTaxBalance & "',")
                strinsert.Append("'" & .contribPreTaxRate & "',")
                strinsert.Append("'" & .contribRothRate & "',")
                strinsert.Append("'" & .contribAfterTaxRate & "',")
                strinsert.Append("'" & .pensionAmountMonthly & "',")
                strinsert.Append("'" & .pensionStartage & "',")
                strinsert.Append("'" & .payPeriods & "'")
                strinsert.Append(")")
            End With

            rows& += doQuery(strinsert.ToString)
        Next

        Return rows&

    End Function
    Function insertplaninfos(m_planinfos As SortedList, strplan$) As Int32
        Dim rows& = 0

        doQuery("delete from planinfo where planid = '" & strplan$ & "'")

        For Each c_planinfo As clsPlaninfo In m_planinfos.Values
            Dim strinsert As StringBuilder = New StringBuilder("Insert into planinfo Values (")
            With c_planinfo
                strinsert.Append("'" & .planid & "',")
                strinsert.Append("'" & .yrenddate & "',")
                strinsert.Append("'" & .planname & "',")
                strinsert.Append("'" & .companyname & "',")
                strinsert.Append("'" & .ein & "',")
                strinsert.Append("'" & .plantype & "',")
                strinsert.Append("'" & .planstatus & "',")
                strinsert.Append("'" & .payfreqcd & "',")
                strinsert.Append("'" & .payschednum & "',")
                strinsert.Append("'" & .payschedname & "',")
                strinsert.Append("'" & .begindate & "',")
                strinsert.Append("'" & .enddate & "',")
                strinsert.Append("'" & .payfreqseqnum & "',")
                strinsert.Append("'" & .inactivecd & "' ")

                strinsert.Append(")")
            End With

            rows& += doQuery(strinsert.ToString)
        Next

        Return rows&

    End Function
    Function buildfundallocations(m_allocations As SortedList) As DataTable
        Dim rows& = 0

        Dim dt As DataTable = getdatatable("select * from fundallocations where 0 = 1")

        For Each c_allocation In m_allocations.Values

            With c_allocation
              
                Dim newrow As DataRow = dt.NewRow
                newrow("employerein") = .employerein
                newrow("plancontractnumber") = .plancontractnumber
                newrow("planname") = .planname
                newrow("ssn") = .ssn
                newrow("ticker") = .ticker
                newrow("currentallocation") = .currentallocation
                newrow("currentbalance") = .currentbalance
                newrow("datestarted") = Format(.datestarted, "yyyy-MM-dd")

                dt.Rows.Add(newrow)
            End With
          

        Next

        Return dt

    End Function
    Function buildparticipants(m_participants As SortedList) As DataTable
        Dim rows& = 0

        Dim dt As DataTable = getdatatable("select * from participant where 0 = 1")

        For Each c_participant As clsParticipant In m_participants.Values

            With c_participant

                Dim newrow As DataRow = dt.NewRow
                newrow("employerein") = .employerein
                newrow("plancontractnumber") = .plancontractnumber
                newrow("planname") = .planname
                newrow("ssn") = .ssn
                newrow("dob") = .dob

                newrow("doh") = .doh
                newrow("firstname") = .firstname
                newrow("lastname") = .lastname
                newrow("middlename") = .middlename
                newrow("address1") = .address1
                newrow("address2") = .address2
                newrow("city") = .city
                newrow("state") = .state
                newrow("zip") = .zip
                newrow("phone") = .phone
                newrow("phone2") = .phone2
                newrow("email") = .email
                newrow("email2") = .email2
                newrow("marital") = .marital
                newrow("gender") = .gender
                newrow("contactmethod") = .contactmethod
                newrow("annualcomp") = .annualcomp
                newrow("autoIncrRate") = .autoIncrRate
                newrow("autoIncrMax") = .autoIncrMax
                newrow("preTaxBalance") = .preTaxBalance
                newrow("RothBalance") = .RothBalance
                newrow("afterTaxBalance") = .afterTaxBalance
                newrow("contribPreTaxRate") = .contribPreTaxRate
                newrow("contribRothRate") = .contribRothRate
                newrow("contribAfterTaxRate") = .contribAfterTaxRate
                newrow("pensionAmountMonthly") = .pensionAmountMonthly
                newrow("pensionStartage") = .pensionStartage
                newrow("payPeriods") = .payPeriods


                dt.Rows.Add(newrow)
            End With


        Next

        Return dt

    End Function
    Function newdate(olddate) As Date
        Dim mydate As String
        If IsDate(olddate) Then
            Return olddate
        End If
        mydate = Mid(olddate, 1, 2) & "/" & Mid(olddate, 3, 2) & "/" & Mid(olddate, 5, 4)
        If IsDate(mydate) Then
            Return mydate
        Else
            Return Nothing
        End If
    End Function
End Class
