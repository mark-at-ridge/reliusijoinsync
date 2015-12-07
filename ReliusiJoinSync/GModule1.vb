Imports System.IO
Imports System.Xml
Module GModule1
    Public clsdb As New clsDatabase
    Public Sub sendtoftp(syncid&, filename$)

        Dim haserrors% = 0
        Dim startsync As DateTime = Date.Now
        Dim strfilename$ = ""
        Form1.updatestatus("Sending " & filename$ & " Files to FTP Server", "EXPORT")
        Try
            Dim sftp As New Chilkat.SFtp()
            Dim starttime As DateTime = Date.Now
            ' MsgBox("sendtoftp2")
1:

            '  Any string automatically begins a fully-functional 30-day trial.
3:          Dim success As Boolean = sftp.UnlockComponent("RIDGESSSH_zhamu37XmFnW")
            If (success <> True) Then
                Console.WriteLine(sftp.LastErrorText)
                Exit Sub
            End If

            ' port: 2202


            '  Set some timeouts, in milliseconds:
4:          sftp.ConnectTimeoutMs = 5000
            sftp.IdleTimeoutMs = 10000

            '  Connect to the SSH server.
            '  The standard SSH port = 22
            '  The hostname may be a hostname or IP address.
            Dim port As Integer
            Dim hostname As String
            hostname = "ijoin.brickftp.com"
            '   hostname = "ftpes://54.209.231.99"
            port = 22
5:          success = sftp.Connect(hostname, port)
            If (success <> True) Then
                '  Console.WriteLine(sftp.LastErrorText)
                Err.Raise(5, "ftpconnect", sftp.LastErrorText)
                Exit Sub
            End If


            '  Authenticate with the SSH server.  Chilkat SFTP supports
            '  both password-based authenication as well as public-key
            '  authentication.  This example uses password authenication.
6:          ' success = sftp.AuthenticatePw("mark@ridgesolutions.com", "mark4040")
            Dim strpassword$ = decryptpassword(My.Settings.ijoinpassword)
            success = sftp.AuthenticatePw(My.Settings.ijoinuser, strpassword$)
            If (success <> True) Then
                '  Console.WriteLine(sftp.LastErrorText)
                Err.Raise(6, "ftpauthent", sftp.LastErrorText)
                Exit Sub
            End If
            'MsgBox(sftp.ProtocolVersion)
            '  sftp.ProtocolVersion = 3
            '  After authenticating, the SFTP subsystem must be initialized:
7:          success = sftp.InitializeSftp()
            If (success <> True) Then
                ' Console.WriteLine(sftp.LastErrorText)
                Err.Raise(7, "InitializeSftp", sftp.LastErrorText)
                Exit Sub
            End If


            '  Open a file for writing on the SSH server.
            '  If the file already exists, it is overwritten.
            '  (Specify "createNew" instead of "createTruncate" to
            '  prevent overwriting existing files.)
            Dim strzipfilename$ = ""
            '  Dim bZipped As Boolean = zipfile("d:\ijoin\participant.csv", strzipfilename$)
            Dim bZipped As Boolean = zipfiles(My.Settings.syncexport & "process\to_send\", strzipfilename$)

            ' get md5
            Dim md5 As System.Security.Cryptography.MD5CryptoServiceProvider = New System.Security.Cryptography.MD5CryptoServiceProvider
            success = sftp.CreateDir("Participants")
            sftp.OpenDir("Participants")
            Dim f As New FileStream(strzipfilename$, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            f = New FileStream(strzipfilename$, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            md5.ComputeHash(f)
            Dim hash As Byte() = md5.Hash
            Dim buff As System.Text.StringBuilder = New System.Text.StringBuilder
            Dim hashByte As Byte
            For Each hashByte In hash
                buff.Append(String.Format("{0:X1}", hashByte))
            Next
            Dim md5code As String
            md5code = buff.ToString()
            f.Close()
            Dim strplangroup$ = "RATEST"
            Dim strfilenamedate As String = CStr(Format(Date.Now, "yyyyMMdd-HHmmss"))
            Dim strfullfilename$ = "uploading-" & strfilenamedate$ & "-" & md5code & "-" & strplangroup$ & "-" & My.Settings.ijoinclientid & ".zip"
            Dim handle As String
8:          handle = sftp.OpenFile("Participant-imports/" & strfullfilename$, "writeOnly", "openOrCreate")
            If (handle = vbNullString) Then
                Err.Raise(8, "ftpopenfile", sftp.LastErrorText)
                '  Console.WriteLine(sftp.LastErrorText)
                Exit Sub
            End If

            strfilename$ = strzipfilename$

            '  Upload from the local file to the SSH server.
9:          success = sftp.UploadFile(handle, strzipfilename$)
            If (success <> True) Then
                ' Console.WriteLine(sftp.LastErrorText)
                Err.Raise(9, "uploadftp", sftp.LastErrorText)
                Exit Sub
            End If

            '  Close the file.
10:         success = sftp.CloseHandle(handle)
            If (success <> True) Then
                ' Console.WriteLine(sftp.LastErrorText)
                Err.Raise(10, "closehandle", sftp.LastErrorText)
                Exit Sub
            End If

            Dim strnewfilename$ = Replace(strfullfilename$, "uploading-", "")
            '   success = sftp.OpenDir("Participants/")
11:         success = sftp.RenameFileOrDir("Participant-imports/" & strfullfilename$, "Participant-imports/" & strnewfilename$)
            If (success <> True) Then
                ' Console.WriteLine(sftp.LastErrorText)
                Err.Raise(11, "renamefile", sftp.LastErrorText)
                Exit Sub
            End If


            ' go thru zipped files and now rename them to thwe sent directory
            Dim files() As FileInfo

            Dim d As New DirectoryInfo(My.Settings.syncexport & "process\to_send\")
            files = d.GetFiles("*.csv")
            For Each fzip As FileInfo In files

                filename = fzip.FullName

                File.Move(filename, Replace(filename.ToString.ToLower, "to_send", "sent"))
                Dim strmovedfilename$ = Replace(filename.ToString.ToLower, "to_send", "sent")

                Rename(strmovedfilename$, Replace(strmovedfilename$, ".csv", "_" & Format(Date.Now, "yyyyMMddhhmmss") & ".csv"))

                
                
            Next


        
            haserrors = 0
            '   Console.WriteLine("Success.")
        Catch ex As Exception
            MsgBox(ex.Message)
            clsdb.insert_errorlog(0, "FTPSend", "SendtoFTP", ex.Message, Erl)
            haserrors = 1
        End Try

        Form1.updatestatus("Successful send to FTP - " & filename$, "EXPORT")
        clsdb.insert_activity(0, startsync, Date.Now, "TransferFilestoFTP", 0, strfilename$, haserrors)


    End Sub
    Public Function movefiletoprocess(strfilename$) As Int16
        Dim intretval As Int16 = 0
        Try
            Dim m_files As New SortedList
            Dim di As New DirectoryInfo(My.Settings.syncexport)
            Dim difiles As FileInfo() = di.GetFiles

            Dim fi As FileInfo
            For Each fi In difiles
                If InStr(fi.Name.ToUpper, strfilename.ToUpper) Then
                    If Not m_files.Contains(fi.LastWriteTime) Then
                        m_files.Add(fi.LastWriteTime, fi)
                    End If
                End If

            Next
            If m_files.Count > 0 Then

                File.Delete(My.Settings.syncexport.ToUpper & "process\" & strfilename$ & ".csv")
                Dim fi_newest As FileInfo = Nothing
                For Each fi In m_files.Values
                    Debug.Print(fi.LastWriteTime)
                    fi_newest = fi
                Next
                Dim strfiletomove$ = fi_newest.FullName
                Dim strnewfilelocation$ = Replace(strfiletomove$.ToUpper, My.Settings.syncexport.ToUpper, My.Settings.syncexport.ToUpper & "process\")
                Dim strfilemoveto$ = Replace(strnewfilelocation$, fi_newest.Name, strfilename)
                File.Move(strfiletomove$, strfilemoveto$)
                '     File.Delete(My.Settings.syncexport.ToUpper & "process\" & strfilename$ & ".csv")
                Rename(strnewfilelocation$, My.Settings.syncexport.ToUpper & "process\" & strfilename$ & ".csv")
                ' noew rename '
                intretval = 1
            End If
        Catch ex As Exception
            intretval = -1
            MsgBox(ex.Message & vbCrLf & strfilename$)
        End Try

        Return intretval


    End Function
    Public Sub getfromftp(syncid&, filename$, bInsert As Boolean, bProcess As Boolean)

        Dim haserrors% = 0
        Dim startsync As DateTime = Date.Now
        Dim strfilename$ = ""
        Form1.updatestatus("Getting Enrollment Files from FTP Server", "IMP")
        Try
            Dim sftp As New Chilkat.SFtp()
            Dim starttime As DateTime = Date.Now
            ' MsgBox("sendtoftp2")
1:

            '  Any string automatically begins a fully-functional 30-day trial.
3:          Dim success As Boolean = sftp.UnlockComponent("RIDGESSSH_zhamu37XmFnW")
            If (success <> True) Then
                Console.WriteLine(sftp.LastErrorText)
                Err.Raise(5, "ftpconnect", sftp.LastErrorText)
                Exit Sub
            End If

            ' port: 2202


            '  Set some timeouts, in milliseconds:
4:          sftp.ConnectTimeoutMs = 5000
            sftp.IdleTimeoutMs = 10000

            '  Connect to the SSH server.
            '  The standard SSH port = 22
            '  The hostname may be a hostname or IP address.
            Dim port As Integer
            Dim hostname As String
            hostname = "ijoin.brickftp.com"
            '   hostname = "ftpes://54.209.231.99"
            port = 22
5:          success = sftp.Connect(hostname, port)
            If (success <> True) Then
                '  Console.WriteLine(sftp.LastErrorText)
                Err.Raise(5, "ftpconnect", sftp.LastErrorText)
                Exit Sub
            End If


            '  Authenticate with the SSH server.  Chilkat SFTP supports
            '  both password-based authenication as well as public-key
            '  authentication.  This example uses password authenication.
6:          ' success = sftp.AuthenticatePw("mark@ridgesolutions.com", "mark4040")
            Dim strpassword$ = decryptpassword(My.Settings.ijoinpassword)
            success = sftp.AuthenticatePw(My.Settings.ijoinuser, strpassword$)
            If (success <> True) Then
                '  Console.WriteLine(sftp.LastErrorText)
                Err.Raise(6, "ftpauthent", sftp.LastErrorText)
                Exit Sub
            End If
            'MsgBox(sftp.ProtocolVersion)
            '  sftp.ProtocolVersion = 3
            '  After authenticating, the SFTP subsystem must be initialized:
7:          success = sftp.InitializeSftp()
            If (success <> True) Then
                ' Console.WriteLine(sftp.LastErrorText)
                Err.Raise(7, "InitializeSftp", sftp.LastErrorText)
                Exit Sub
            End If




            ' get md5

            Dim strplangroup$ = "RATEST"
            Dim strfilenamedate As String = CStr(Format(Date.Now, "yyyyMMdd-HHmmss"))

            '  Open a directory on the server...
            '  Paths starting with a slash are "absolute", and are relative
            '  to the root of the file system. Names starting with any other
            '  character are relative to the user's default directory (home directory).
            '  A path component of ".." refers to the parent directory,
            '  and "." refers to the current directory.
            Dim handle As String
            '    handle = sftp.OpenDir(My.Settings.ijoinclientid & "\enrollments")
            handle = sftp.OpenDir("completed-enrollments")
            If (handle = vbNullString) Then
                Console.WriteLine(sftp.LastErrorText)
                Err.Raise(8, "ftpopendir", sftp.LastErrorText)
                Exit Sub
            End If


            '  Download the directory listing:
            Dim dirListing As Chilkat.SFtpDir
            dirListing = sftp.ReadDir(handle)
            If (dirListing Is Nothing) Then
                Console.WriteLine(sftp.LastErrorText)
                Err.Raise(8, "ftpreaddir", sftp.LastErrorText)
                Exit Sub
            End If


            '  Iterate over the files.
            Dim i As Integer
            Dim n As Integer = dirListing.NumFilesAndDirs
            Dim filehandle As String = ""
            If (n = 0) Then
                ' Console.WriteLine("No entries found in this directory.")
                Form1.updatestatus("No Files to Download from FTP", "IMP")
            Else
                For i = 0 To n - 1
                    Dim fileObj As Chilkat.SFtpFile
                    fileObj = dirListing.GetFileObject(i)

                    If fileObj.FileType <> "directory" Then
                        Console.WriteLine(fileObj.Filename)
                        Console.WriteLine(fileObj.FileType)
                        Console.WriteLine("Size in bytes: " & fileObj.Size32)
                        Console.WriteLine("----")
                        '     success = sftp.RenameFileOrDir(My.Settings.ijoinclientid & "\enrollments\" & fileObj.Filename, My.Settings.ijoinclientid & "\enrollments\downloading-" & fileObj.Filename)
                        success = sftp.RenameFileOrDir("completed-enrollments\" & fileObj.Filename, "completed-enrollments\downloading-" & fileObj.Filename)

                        If (success <> True) Then
                            ' Console.WriteLine(sftp.LastErrorText)
                            Err.Raise(11, "renamefile", sftp.LastErrorText)
                            Exit Sub
                        End If
                        '  handle = sftp.OpenFile(My.Settings.ijoinclientid & "\enrollments\downloading-" & fileObj.Filename, "readOnly", "openExisting")
                        handle = sftp.OpenFile("completed-enrollments\downloading-" & fileObj.Filename, "readOnly", "openExisting")
                        If (handle = vbNullString) Then
                            Console.WriteLine(sftp.LastErrorText)
                            Err.Raise(10, "ftpopenfile", sftp.LastErrorText)
                            Exit Sub
                        End If




                        '  Download the file:
                        success = sftp.DownloadFile(handle, Replace(My.Settings.syncimport, "\", "/") & fileObj.Filename)
                        If (success <> True) Then
                            Console.WriteLine(sftp.LastErrorText)
                            Err.Raise(8, "ftpdownloadfile", sftp.LastErrorText)
                            Exit Sub
                        End If
                        Form1.updatestatus("Downloaded " & fileObj.Filename & " from FTP", "IMP")

                        '  Close the file.
                        success = sftp.CloseHandle(handle)
                        If (success <> True) Then
                            Console.WriteLine(sftp.LastErrorText)
                            Exit Sub
                        End If

                        '      success = sftp.RenameFileOrDir(My.Settings.ijoinclientid & "\enrollments\downloading-" & fileObj.Filename, My.Settings.ijoinclientid & "\enrollments\" & fileObj.Filename)
                        success = sftp.RenameFileOrDir("completed-enrollments\downloading-" & fileObj.Filename, "completed-enrollments\" & fileObj.Filename)

                        If (success <> True) Then
                            ' Console.WriteLine(sftp.LastErrorText)
                            Err.Raise(14, "renamefileback", sftp.LastErrorText)
                            Exit Sub
                        End If


                        Directory.CreateDirectory(My.Settings.syncimport & Microsoft.VisualBasic.Left(fileObj.Filename, 4))
                        Directory.CreateDirectory(My.Settings.syncimport & Microsoft.VisualBasic.Left(fileObj.Filename, 4) & "\" & Mid(fileObj.Filename, 5, 2))
                        Dim mydir = Directory.CreateDirectory(My.Settings.syncimport & Microsoft.VisualBasic.Left(fileObj.Filename, 4) & "\" & Mid(fileObj.Filename, 5, 2) & "\" & Mid(fileObj.Filename, 7, 2))
                        FileCopy(My.Settings.syncimport & fileObj.Filename, mydir.FullName & "\" & fileObj.Filename)

                        clsdb.insert_activity(0, startsync, Date.Now, "TransferFilesfromFTP", 0, My.Settings.syncimport & fileObj.Filename, haserrors)

                        Dim strdownloadedfilename$ = Replace(fileObj.Filename, "downloading-", "")
                        If bInsert And InStr(fileObj.Filename.ToLower, "enrollments_to_record") Then
                            clsdb.impenrollments(strdownloadedfilename$, My.Settings.syncimport, "", True)
                            If bProcess Then
                                'generate_DER_xmlfile(My.Settings.syncimport & fileObj.Filename)
                            End If
                        End If
                        If bInsert And InStr(fileObj.Filename.ToLower, "enrollments_allocations_to_record") Then
                            clsdb.impenrollmentallocations(strdownloadedfilename$, My.Settings.syncimport, "", True)
                            If bProcess Then
                                'generate_alloc_xmlfile(My.Settings.syncimport & fileObj.Filename)
                            End If
                        End If
                    End If ' this is a directory
                Next
            End If



            '  Close the file.
10:         '  success = sftp.CloseHandle(handle)
            If (success <> True) Then
                ' Console.WriteLine(sftp.LastErrorText)
                '     Exit Sub
            End If

            '    Dim strnewfilename$ = Replace(strfullfilename$, "uploading-", "")
            '   success = sftp.OpenDir("Participants/")
            '    success = sftp.RenameFileOrDir("Participants/" & strfullfilename$, "Participants/" & strnewfilename$)




            haserrors = 0
            '   Console.WriteLine("Success.")
        Catch ex As Exception
            MsgBox(ex.Message)
            clsdb.insert_errorlog(0, "FTPSend", "SendtoFTP", ex.Message, Erl)
            haserrors = 1
        End Try

        Form1.updatestatus("Successful Download from FTP", "IMP")
        clsdb.insert_activity(0, startsync, Date.Now, "CompletedTransferFilesfromFTP", 0, strfilename$, haserrors)


    End Sub
    Function generate_alloc_xmlfile(strfilename$, strplan$) As Int16
        Dim haserrors% = 0
        Try
            Dim strxmlfile$ = ""

            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Indent = True
            settings.Encoding = System.Text.Encoding.ASCII

            Dim dtpayroll As DataTable = clsdb.getdatatable(" select * from planinfo where planid = '" & strplan$ & "' and payschedname like 'ijoin%' ")
            If dtpayroll.Rows.Count = 0 Then
                MsgBox("Can not find payroll schedule with payschedulename." & vbCrLf & "Please check RK System for setup", MsgBoxStyle.Critical)
                haserrors% = 1
                Return haserrors
                Exit Function
            End If
            Dim drpayroll As DataRow = dtpayroll.Rows(0)


            ' settings.Encoding = System.Text.ASCIIEncoding
            Dim dt As DataTable = clsdb.getdatatable("select distinct employerein, plancontractnumber, ssn, datestarted from enrollmentallocations ")
            ' Create XmlWriter.
            Using writer As XmlWriter = XmlWriter.Create(My.Settings.syncimport & "process\enrollments_allocations" & strplan$ & ".xml", settings)
                ' Begin writing.
                writer.WriteStartDocument()
                ' Dim e As XElement = <REQUESTS/>
                '  e.SetAttributeValue("ActionCode", "S")

                writer.WriteStartElement("REQUESTS") ' Root.
                writer.WriteAttributeString("ActionCode", "S")
                writer.WriteStartElement("ADD") ' Root.
                writer.WriteStartElement("REQUEST_ALLOCATION_CHANGE")
                writer.WriteAttributeString("EmployerIdentificationNumber", Replace(dt.Rows(0)(0), "-", ""))
                writer.WriteStartElement("ALLOCATIONS") ' Root.


                For Each dr As DataRow In dt.Rows
                    writer.WriteStartElement("ALLOCATION") ' Root.
                    writer.WriteAttributeString("PlanID", dr(1))
                    writer.WriteElementString("PlanID", dr(1))
                    writer.WriteElementString("YearEndDate", Format(CDate(drpayroll("yrenddate")), "yyyy-MM-dd"))
                    writer.WriteElementString("SSNum", Replace(dr("ssn"), "-", ""))
                    writer.WriteElementString("EffectiveDate", Format(CDate(dr("datestarted")), "yyyy-MM-dd"))
                    writer.WriteElementString("AllocationSetTypeCode", "P")

                    writer.WriteStartElement("SOURCES") ' Root.
                    writer.WriteStartElement("SOURCE") ' Root.
                    writer.WriteAttributeString("SourceID", "0")
                    ' this is where we all up to 100%
                    Dim dtdetail As DataTable = clsdb.getdatatable(" select ticker, allocationpercent from enrollmentallocations where plancontractnumber = '" & dr(1) & "' and ssn = '" & dr("ssn") & "' and datestarted = '" & dr("datestarted") & "'")

                    For Each drdetail As DataRow In dtdetail.Rows
                        writer.WriteStartElement("ALLOCATION_DETAIL")
                        writer.WriteElementString("AllocationPercent", Format(drdetail("AllocationPercent"), "0.00"))
                        writer.WriteElementString("FundID", drdetail("ticker"))

                        writer.WriteEndElement()
                    Next
                    writer.WriteEndElement()
                    writer.WriteEndElement()

                    writer.WriteEndElement()
                    'Next
                Next
                ' End document.
                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.WriteEndDocument()
            End Using


        Catch ex As Exception
            MsgBox(ex.Message)
            clsdb.insert_errorlog(0, "GenerateXML", "AllocFile", ex.Message, Erl)
            haserrors = 1
        End Try
        Return haserrors
    End Function
    Function generate_DER_xmlfile(strfilename$, strplan$) As Int16
        Dim haserrors% = 0
        Try
            Dim strxmlfile$ = ""

            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Indent = True
            settings.Encoding = System.Text.Encoding.ASCII
            ' settings.Encoding = System.Text.ASCIIEncoding
            Dim dtpayroll As DataTable = clsdb.getdatatable(" select * from planinfo where planid = '" & strplan$ & "' and payschedname like 'ijoin%' ")
            If dtpayroll.Rows.Count = 0 Then
                MsgBox("Can not find payroll schedule with payschedulename." & vbCrLf & "Please check RK System for setup", MsgBoxStyle.Critical)
                haserrors% = 1
                Return haserrors
                Exit Function
            End If
            Dim drpayroll As DataRow = dtpayroll.Rows(0)
            Dim dt As DataTable = clsdb.getdatatable("select distinct employerein, plancontractnumber,effectivedate, filename, exportfilename from enrollments where xmlcreated = 0 and exported = 1 and plancontractnumber = '" & strplan$ & "'")
            ' Create XmlWriter.
            If dt.Rows.Count = 0 Then

                MsgBox("No records found. Records expected though.", MsgBoxStyle.Critical)
                haserrors% = 1
                Return haserrors
                Exit Function
            End If

            Using writer As XmlWriter = XmlWriter.Create(My.Settings.syncimport & "process\enrollments_participants" & strplan$ & ".xml", settings)
                ' Begin writing.
                writer.WriteStartDocument()
                ' Dim e As XElement = <REQUESTS/>
                '  e.SetAttributeValue("ActionCode", "S")

                writer.WriteStartElement("REQUESTS") ' Root.
                writer.WriteAttributeString("ActionCode", "P")
                writer.WriteStartElement("IMPORT_PAYROLL") ' Root.
                writer.WriteStartElement("REQUEST_PAYROLL")
                writer.WriteAttributeString("EmployerIdentificationNumber", Replace(dt.Rows(0)(0), "-", ""))
                writer.WriteAttributeString("PlanID", dt.Rows(0)(1))
                writer.WriteStartElement("PAYROLL_PARAMETER_INFO") ' Root.


                '   For Each dr As DataRow In dt.Rows

                writer.WriteElementString("PlanID", dt.Rows(0)(1))
                writer.WriteElementString("YearEndDate", Format(CDate(drpayroll("yrenddate")), "yyyy-MM-dd"))
                writer.WriteElementString("PayPeriodEndDate", Format(CDate(drpayroll("enddate")), "yyyy-MM-dd"))
                writer.WriteElementString("PayPeriodBeginDate", Format(CDate(drpayroll("begindate")), "yyyy-MM-dd"))
                writer.WriteElementString("DERNAME", "IJOINEnroll")
                writer.WriteElementString("DERFileName", dt.Rows(0)("exportfilename"))

                writer.WriteStartElement("DERGenerateNewEmployee")
                writer.WriteAttributeString("tc", "Y")
                writer.WriteEndElement()

                writer.WriteStartElement("DERUpdateExistingEmployee")
                writer.WriteAttributeString("tc", "Y")
                writer.WriteEndElement()

                writer.WriteStartElement("DERValidateImportOnly")
                writer.WriteAttributeString("tc", "N")
                writer.WriteEndElement()

                writer.WriteStartElement("DERCreateEligibility")
                writer.WriteAttributeString("tc", "N")
                writer.WriteEndElement()

                writer.WriteStartElement("SuppressCode")
                writer.WriteAttributeString("tc", "N")
                writer.WriteEndElement()

                writer.WriteStartElement("DERCreatePostTaxMatchTrans")
                writer.WriteAttributeString("tc", "N")
                writer.WriteEndElement()

                writer.WriteStartElement("DERCreatePreTaxMatchTrans")
                writer.WriteAttributeString("tc", "N")
                writer.WriteEndElement()

                writer.WriteElementString("AllocationEffectiveDate", Format(CDate(dt.Rows(0)(2)), "yyyy-MM-dd"))
                writer.WriteElementString("ContributionPercentEffectiveDate", Format(CDate(dt.Rows(0)(2)), "yyyy-MM-dd"))
                writer.WriteElementString("CensusDataEffectiveDate", Format(CDate(dt.Rows(0)(2)), "yyyy-MM-dd"))


                writer.WriteEndElement()

                writer.WriteEndElement()
                'Next
                '   Next
                ' End document.
                writer.WriteEndElement()
                writer.WriteEndElement()

                writer.WriteEndDocument()
            End Using

            clsdb.insert_activity(0, Date.Now, Date.Now, "BuildXML", 0, My.Settings.syncimport & "process\enrollments_participants" & strplan$ & ".xml", haserrors)
            Form1.updatestatus("Built XML File - " & My.Settings.syncimport & "process\enrollments_participants" & strplan$ & ".xml", "IMP")
        Catch ex As Exception
            MsgBox(ex.Message)
            clsdb.insert_errorlog(0, "GenerateXML", "EnrollFile", ex.Message, Erl)
            haserrors = 1
        End Try
        Return haserrors
    End Function
    Public Sub sendtoftptest(syncid&)

        Dim haserrors% = 0
        Dim startsync As DateTime = Date.Now
        Dim strfilename$ = ""
        Try
            Dim sftp As New Chilkat.SFtp()
            Dim starttime As DateTime = Date.Now
            ' MsgBox("sendtoftp2")
1:

            '  Any string automatically begins a fully-functional 30-day trial.
3:          Dim success As Boolean = sftp.UnlockComponent("RIDGESSSH_zhamu37XmFnW")
            If (success <> True) Then
                Console.WriteLine(sftp.LastErrorText)
                Exit Sub
            End If

            ' port: 2202


            '  Set some timeouts, in milliseconds:
4:          sftp.ConnectTimeoutMs = 5000
            sftp.IdleTimeoutMs = 10000

            '  Connect to the SSH server.
            '  The standard SSH port = 22
            '  The hostname may be a hostname or IP address.
            Dim port As Integer
            Dim hostname As String
            hostname = "ijoin.brickftp.com"
            '   hostname = "ftpes://54.209.231.99"
            port = 22
5:          success = sftp.Connect(hostname, port)
            If (success <> True) Then
                '  Console.WriteLine(sftp.LastErrorText)
                Err.Raise(5, "ftpconnect", sftp.LastErrorText)
                Exit Sub
            End If


            '  Authenticate with the SSH server.  Chilkat SFTP supports
            '  both password-based authenication as well as public-key
            '  authentication.  This example uses password authenication.
6:          '   success = sftp.AuthenticatePw("mark@ridgesolutions.com", "mark4040")
            Dim strpassword$ = decryptpassword(My.Settings.ijoinpassword)
            success = sftp.AuthenticatePw(My.Settings.ijoinuser, strpassword$)
            If (success <> True) Then
                '  Console.WriteLine(sftp.LastErrorText)
                Err.Raise(6, "ftpauthent", sftp.LastErrorText)
                Exit Sub
            End If
            'MsgBox(sftp.ProtocolVersion)
            '  sftp.ProtocolVersion = 3
            '  After authenticating, the SFTP subsystem must be initialized:
7:          success = sftp.InitializeSftp()
            If (success <> True) Then
                ' Console.WriteLine(sftp.LastErrorText)
                Err.Raise(7, "InitializeSftp", sftp.LastErrorText)
                Exit Sub
            End If


            '  Open a file for writing on the SSH server.
            '  If the file already exists, it is overwritten.
            '  (Specify "createNew" instead of "createTruncate" to
            '  prevent overwriting existing files.)
            Dim strzipfilename$ = ""
            Dim bZipped As Boolean = zipfile("d:\ijoin\participant.csv", strzipfilename$)

            ' get md5
            Dim md5 As System.Security.Cryptography.MD5CryptoServiceProvider = New System.Security.Cryptography.MD5CryptoServiceProvider
            success = sftp.CreateDir("Participants")
            sftp.OpenDir("Participants")
            Dim f As New FileStream("d:\ijoin\participant.csv", FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            f = New FileStream("d:\ijoin\participant.csv", FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            md5.ComputeHash(f)
            Dim hash As Byte() = md5.Hash
            Dim buff As System.Text.StringBuilder = New System.Text.StringBuilder
            Dim hashByte As Byte
            For Each hashByte In hash
                buff.Append(String.Format("{0:X1}", hashByte))
            Next
            Dim md5code As String
            md5code = buff.ToString()
            f.Close()
            Dim strplangroup$ = "RATEST"
            Dim strfilenamedate As String = CStr(Format(Date.Now, "yyyyMMdd-HHmmss"))
            Dim strfullfilename$ = "uploading-" & strfilenamedate$ & "-" & md5code & "-" & strplangroup$ & "-" & My.Settings.ijoinclientid & ".zip"
            Dim handle As String
8:          handle = sftp.OpenFile("Participants/" & strfullfilename$, "writeOnly", "openOrCreate")
            If (handle = vbNullString) Then
                Err.Raise(8, "ftpopenfile", sftp.LastErrorText)
                '  Console.WriteLine(sftp.LastErrorText)
                Exit Sub
            End If

            strfilename$ = strzipfilename$

            '  Upload from the local file to the SSH server.
9:          success = sftp.UploadFile(handle, strzipfilename$)
            If (success <> True) Then
                ' Console.WriteLine(sftp.LastErrorText)
                Err.Raise(9, "uploadftp", sftp.LastErrorText)
                Exit Sub
            End If

            '  Close the file.
10:         success = sftp.CloseHandle(handle)
            If (success <> True) Then
                ' Console.WriteLine(sftp.LastErrorText)
                Exit Sub
            End If

            Dim strnewfilename$ = Replace(strfullfilename$, "uploading-", "")
            '   success = sftp.OpenDir("Participants/")
            success = sftp.RenameFileOrDir("Participants/" & strfullfilename$, "Participants/" & strnewfilename$)
            If (success <> True) Then
                ' Console.WriteLine(sftp.LastErrorText)
                Exit Sub
            End If
            '  success = sftp.RemoveFile("ParticipantCOMPLETE.zip")



            haserrors = 0
            '   Console.WriteLine("Success.")
        Catch ex As Exception
            MsgBox(ex.Message)
            clsdb.insert_errorlog(0, "FTPSend", "SendtoFTP", ex.Message, Erl)
            haserrors = 1
        End Try


        clsdb.insert_activity(0, startsync, Date.Now, "TransferFilestoFTP", 0, strfilename$, haserrors)


    End Sub
    Function zipfile(strfilename$, ByRef strzipfilename$) As Boolean
        Dim zip As New Chilkat.Zip()

        Dim success As Boolean

        '  Any string unlocks the component for the 1st 30-days.
        success = zip.UnlockComponent("RIDGESZIP_pv5MG82PoMsG")
        If (success <> True) Then
            Console.WriteLine(zip.LastErrorText)
            Return success
            Exit Function
        End If

        Dim f As New FileInfo(strfilename$)

        If f.Exists Then
            strzipfilename$ = Replace(strfilename$, f.Extension, ".zip")
        End If
        success = zip.NewZip(strzipfilename$)
        If (success <> True) Then
            Console.WriteLine(zip.LastErrorText)
            Return success
            Exit Function
        End If


        '  In this example, the file we wish to zip is /temp/abc123/HelloWorld123.txt

        '  Add a reference to a single file by calling AppendOneFileOrDir
        '  Note: You may use either forward or backward slashes.
        Dim saveExtraPath As Boolean = False
        success = zip.AppendOneFileOrDir(strfilename$, saveExtraPath)
        If (success <> True) Then
            Console.WriteLine(zip.LastErrorText)
            Return success
            Exit Function
        End If


        success = zip.WriteZipAndClose()
        If (success <> True) Then
            Console.WriteLine(zip.LastErrorText)
            Return success
            Exit Function
        End If

        Return success
    End Function
    Function zipfiles(strdirectory$, ByRef strzipfilename$) As Boolean
        Dim zip As New Chilkat.Zip()

        Dim success As Boolean
        Try

       
        '  Any string unlocks the component for the 1st 30-days.
        success = zip.UnlockComponent("RIDGESZIP_pv5MG82PoMsG")
        If (success <> True) Then
                Console.WriteLine(zip.LastErrorText)
                Return success
            Exit Function
        End If
        Dim d As New DirectoryInfo(strdirectory$)
        Dim files() As FileInfo

        strzipfilename$ = strdirectory$ & My.Settings.ijoinclientid & ".zip"

        success = zip.NewZip(strzipfilename$)
        If (success <> True) Then
                Console.WriteLine(zip.LastErrorText)
                Return success
            Exit Function
        End If
            files = d.GetFiles("*.csv")
        For Each f As FileInfo In files




            '  In this example, the file we wish to zip is /temp/abc123/HelloWorld123.txt

            '  Add a reference to a single file by calling AppendOneFileOrDir
            '  Note: You may use either forward or backward slashes.
            Dim saveExtraPath As Boolean = False
            success = zip.AppendOneFileOrDir(f.FullName, saveExtraPath)
            If (success <> True) Then

                    Return success
                    Exit Function

            End If
        Next

        success = zip.WriteZipAndClose()
        If (success <> True) Then

                Return success
            Exit Function
        End If
        Catch ex As Exception
            success = False
            MsgBox("Zipfiles - " & ex.Message)
        End Try
        Return success
    End Function
    Function encryptpassword(strpassword$) As String


        Dim crypt As New Chilkat.Crypt2()

        Dim success As Boolean = crypt.UnlockComponent("RIDGESCrypt_O8t28LY7NFHB")
        If (success <> True) Then
            Return success
            Exit Function
        End If


        '  AES is also known as Rijndael.
        crypt.CryptAlgorithm = "aes"

        '  CipherMode may be "ecb" or "cbc"
        crypt.CipherMode = "cbc"

        '  KeyLength may be 128, 192, 256
        crypt.KeyLength = 256

        '  The padding scheme determines the contents of the bytes
        '  that are added to pad the result to a multiple of the
        '  encryption algorithm's block size.  AES has a block
        '  size of 16 bytes, so encrypted output is always
        '  a multiple of 16.
        crypt.PaddingScheme = 0

        '  EncodingMode specifies the encoding of the output for
        '  encryption, and the input for decryption.
        '  It may be "hex", "url", "base64", or "quoted-printable".
        crypt.EncodingMode = "hex"

        '  An initialization vector is required if using CBC mode.
        '  ECB mode does not use an IV.
        '  The length of the IV is equal to the algorithm's block size.
        '  It is NOT equal to the length of the key.
        Dim ivHex As String = "000102030405060708090A0B0C0D0E0F"
        crypt.SetEncodedIV(ivHex, "hex")

        '  The secret key must equal the size of the key.  For
        '  256-bit encryption, the binary secret key is 32 bytes.
        '  For 128-bit encryption, the binary secret key is 16 bytes.
        Dim keyHex As String = "000102030405060708090A0B0C0D0E0F101112131415161718191A1B1C1D1E1F"
        crypt.SetEncodedKey(keyHex, "hex")

        '  Encrypt a string...
        '  The input string is 44 ANSI characters (i.e. 44 bytes), so
        '  the output should be 48 bytes (a multiple of 16).
        '  Because the output is a hex string, it should
        '  be 96 characters long (2 chars per byte).
        ' Dim encStr As String = crypt.EncryptStringENC("The quick brown fox jumps over the lazy dog.")
        Dim encStr As String = crypt.EncryptStringENC(strpassword$)
        ' Console.WriteLine(encStr)
        Return encStr
        '  Now decrypt:
        ' Dim decStr As String = crypt.DecryptStringENC(encStr)
        '  Console.WriteLine(decStr)
    End Function


    Function decryptpassword(strpassword$) As String


        Dim crypt As New Chilkat.Crypt2()

        Dim success As Boolean = crypt.UnlockComponent("RIDGESCrypt_O8t28LY7NFHB")
        If (success <> True) Then
            Return success
            Exit Function
        End If


        '  AES is also known as Rijndael.
        crypt.CryptAlgorithm = "aes"

        '  CipherMode may be "ecb" or "cbc"
        crypt.CipherMode = "cbc"

        '  KeyLength may be 128, 192, 256
        crypt.KeyLength = 256

        '  The padding scheme determines the contents of the bytes
        '  that are added to pad the result to a multiple of the
        '  encryption algorithm's block size.  AES has a block
        '  size of 16 bytes, so encrypted output is always
        '  a multiple of 16.
        crypt.PaddingScheme = 0

        '  EncodingMode specifies the encoding of the output for
        '  encryption, and the input for decryption.
        '  It may be "hex", "url", "base64", or "quoted-printable".
        crypt.EncodingMode = "hex"

        '  An initialization vector is required if using CBC mode.
        '  ECB mode does not use an IV.
        '  The length of the IV is equal to the algorithm's block size.
        '  It is NOT equal to the length of the key.
        Dim ivHex As String = "000102030405060708090A0B0C0D0E0F"
        crypt.SetEncodedIV(ivHex, "hex")

        '  The secret key must equal the size of the key.  For
        '  256-bit encryption, the binary secret key is 32 bytes.
        '  For 128-bit encryption, the binary secret key is 16 bytes.
        Dim keyHex As String = "000102030405060708090A0B0C0D0E0F101112131415161718191A1B1C1D1E1F"
        crypt.SetEncodedKey(keyHex, "hex")

        '  Encrypt a string...
        '  The input string is 44 ANSI characters (i.e. 44 bytes), so
        '  the output should be 48 bytes (a multiple of 16).
        '  Because the output is a hex string, it should
        '  be 96 characters long (2 chars per byte).
        ' Dim encStr As String = crypt.EncryptStringENC("The quick brown fox jumps over the lazy dog.")

        '  Now decrypt:
        Dim decStr As String = crypt.DecryptStringENC(strpassword)
        Return decStr
        '  Console.WriteLine(decStr)
    End Function
    Sub main()

        Dim frm As New Form1
        Application.Run(frm)
    End Sub
End Module
