Imports System.IO

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        txtclient.Text = My.Settings.ijoinclientid
        txtuser.Text = My.Settings.ijoinuser
        Dim strencrypt$ = decryptpassword(My.Settings.ijoinpassword)
        txtpassword.Text = strencrypt$
        txtdirectory.Text = My.Settings.syncdirectory
        txtsyncexport.Text = My.Settings.syncexport
        txtreliusimport.Text = My.Settings.reliusimportdirectory

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Settings.ijoinclientid = txtclient.Text
        My.Settings.ijoinuser = txtuser.Text
        Dim strencrypt$ = encryptpassword(txtpassword.Text)
        My.Settings.ijoinpassword = strencrypt$
        My.Settings.syncdirectory = txtdirectory.Text
        My.Settings.syncexport = txtsyncexport.Text
        My.Settings.syncimport = txtsyncimport.Text
        My.Settings.reliusimportdirectory = txtreliusimport.Text
        My.Settings.Save()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        sendtoftptest(0)
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(sender As Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        Select Case e.Tab.Key.ToString.ToLower
            Case "activity"
                refreshactivity()
                refresherrors()
        End Select
    End Sub
    Sub refreshactivity()


        Dim dt As DataTable = clsdb.getdatatable("select syncid,  substr(startsync,0) as StartSync,substr(endsync,0) as 'EndSync', SyncMachine, SyncType, HashTotal, FileName, HasErrors  from activity")

        Dim dtclone As DataTable = dt.Clone
        dtclone.Columns("startsync").DataType = GetType(System.DateTime)
        dtclone.Columns("endsync").DataType = GetType(System.DateTime)
        For Each dr As DataRow In dt.Rows
            dtclone.ImportRow(dr)
        Next


        grdactivity.DataSource = dtclone
    End Sub
    Sub refresherrors()

        Dim dt As DataTable = clsdb.getdatatable("select syncid, ErrorType, ErrorModule, ErrorDescription,ErrorNumber,ErrorMachine,substr(errordate,0) as 'ErrorDate' from errorlog")
        Dim dtclone As DataTable = dt.Clone
        dtclone.Columns("ErrorDate").DataType = GetType(System.DateTime)

        For Each dr As DataRow In dt.Rows
            dtclone.ImportRow(dr)
        Next
        grderrors.DataSource = dtclone
    End Sub

    Private Sub grdactivity_InitializeLayout(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdactivity.InitializeLayout
        e.Layout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.None
        e.Layout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.None
        e.Layout.Override.BorderStyleHeader = Infragistics.Win.UIElementBorderStyle.None
        e.Layout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False
        e.Layout.Bands(0).Columns("startsync").Format = "MM/dd/yyyy hh:mm:ss tt"
        e.Layout.Bands(0).Columns("endsync").Format = "MM/dd/yyyy hh:mm:ss tt"
        e.Layout.Bands(0).Columns("HasErrors").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        e.Layout.Bands(0).Columns("filename").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.URL

        Dim dt As DataTable = grdactivity.DataSource
        lblactivityrows.Text = "Rows: " & dt.Rows.Count
    End Sub

    Private Sub UltraButton1_Click(sender As Object, e As EventArgs) Handles UltraButton1.Click
        Dim dt As DataTable = clsdb.getdatatable(" select * from " & cboTables.Value)
        grddata.DataSource = dt
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        clearstatus("")
        updatestatus("Ready to Process", "EXPORT")
        Me.Cursor = Cursors.WaitCursor
        Dim startsync As DateTime = Date.Now
        Try
            Dim dtplangroup As DataTable = clsdb.getdatatable("select * from plangroups where active = 1")
            Dim bInsert As Boolean = False
            If chkParticipant.Checked Then
                bInsert = (chkInsertPart.Checked)


                For Each drplan As DataRow In dtplangroup.Rows
                    Dim intmoved% = movefiletoprocess("participant_" & drplan("planid"))
                    If intmoved% = 1 Then
                        clsdb.impParticipant(My.Settings.syncexport & "process\participant_" & drplan("planid") & ".csv", "", drplan("planid"), bInsert)
                    Else
                        Exit Try
                    End If
                Next

                Dim strfilemoved$ = Replace(My.Settings.syncexport & "process\participant" & My.Settings.ijoinclientid & ".csv", "process\", "process\to_send\")
                '  File.Move(My.Settings.syncexport & "process\participant" & My.Settings.ijoinclientid & ".csv", strfilemoved$)
                Dim strfilename$ = ""
                clsdb.writefile(startsync, "participant", strfilemoved$)
               
            End If

            If chkFundAllocations.Checked Then
                bInsert = (chkInsertFundAllocations.Checked)
                For Each drplan As DataRow In dtplangroup.Rows
                    Dim intmoved% = movefiletoprocess("fundallocations_" & drplan("planid"))
                    If intmoved% = 1 Then
                        clsdb.impfundallocations(My.Settings.syncexport & "process\fundallocations_" & drplan("planid") & ".csv", "", drplan("planid"), bInsert)
                    Else
                        Exit Try

                    End If

                Next

                Dim strfilemoved$ = Replace(My.Settings.syncexport & "process\fundallocations" & My.Settings.ijoinclientid & ".csv", "process\", "process\to_send\")
                '  File.Move(My.Settings.syncexport & "process\fundallocations" & My.Settings.ijoinclientid & ".csv", strfilemoved$)

                clsdb.writefile(startsync, "fundallocations", strfilemoved$)

             
            End If

            If chkBene.Checked Then
                bInsert = (chkInsertBene.Checked)
                For Each drplan As DataRow In dtplangroup.Rows
                    Dim intmoved% = movefiletoprocess("beneficiary_" & drplan("planid"))
                    If intmoved% = 1 Then
                        clsdb.impBeneficiary(My.Settings.syncexport & "process\beneficiary_" & drplan("planid") & ".csv", "", drplan("planid"), bInsert)
                    End If

                Next

                Dim strfilemoved$ = Replace(My.Settings.syncexport & "process\beneficiary" & My.Settings.ijoinclientid & ".csv", "process\", "process\to_send\")
                '  File.Move(My.Settings.syncexport & "process\fundallocations" & My.Settings.ijoinclientid & ".csv", strfilemoved$)

                clsdb.writefile(startsync, "beneficiary", strfilemoved$)

               
            End If




            If chkFTPFiles.Checked Then
                sendtoftp(0, My.Settings.syncexport & "process\to_send\participant" & My.Settings.ijoinclientid & ".csv")
            End If

            If chkplaninfo.Checked Then
                bInsert = (chkInsertPlanInfo.Checked)

                For Each FileFound As String In Directory.GetFiles(My.Settings.syncexport & "process\", "planinfo*.csv")
                    File.Delete(FileFound)
                Next


                For Each drplan As DataRow In dtplangroup.Rows
                    Dim intmoved% = movefiletoprocess("planinfo_" & drplan("planid"))
                    If intmoved% = 1 Then
                        clsdb.impplaninfo(My.Settings.syncexport & "process\planinfo_" & drplan("planid") & ".csv", "", drplan("planid"), bInsert)
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        updatestatus("Done Processing", "EXPORT")
        Cursor = Cursors.Default
    End Sub
    Public Sub clearstatus(strstatustext$)
        txtstatus.Text = ""
        txtimportstatus.Text = ""
        Application.DoEvents()
    End Sub
    Public Sub updatestatus(strstatustext$, strtextbox$)

        If strtextbox$ = "EXPORT" Then
            txtstatus.Text += Format(Date.Now, "MM/dd/yyyy hh:mm:ss") & vbTab & strstatustext$ & vbCrLf
        Else
            txtimportstatus.Text += Format(Date.Now, "MM/dd/yyyy hh:mm:ss") & vbTab & strstatustext$ & vbCrLf
        End If
        Application.DoEvents()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        clearstatus("")
        updatestatus("Ready to Process", "IMP")
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim bInsert As Boolean = (chkImportenrollments.Checked)
            Dim bProcess As Boolean = (chkProcess.Checked)
            If chkdownload.Checked Then
                '  bInsert = (chkInsertPart.Checked)
                ' clsdb.impParticipant(txtimportdir.Text, "", bInsert)
                '  If chkFTPpart.Checked Then
                getfromftp(0, txtpartfilename.Text, bInsert, bProcess)
                'End If
            End If

            Dim dtplangroup As DataTable = clsdb.getdatatable("select distinct plancontractnumber as planid from enrollments where exported = 0")
            Dim strplan$ = ""
            If bProcess Then
                For Each dr As DataRow In dtplangroup.Rows
                    strplan$ = dr("planid")
                    clsdb.buildEnrollmentFile(strplan$)
                    generate_DER_xmlfile(My.Settings.syncimport & "import\" & txtpartfilename.Text, strplan$)

                    clsdb.buildEnrollmentAllocationFile(strplan)
                    generate_alloc_xmlfile(My.Settings.syncimport & "import\" & txtpartfilename.Text, strplan$)
                Next
            End If



            If chkImportenrollments.Checked Then
                '  clsdb.buildEnrollmentAllocationFile(strplan)
                '   generate_alloc_xmlfile(My.Settings.syncimport & "import\" & txtpartfilename.Text, strplan$)
                ' clsdb.impfundallocations(txtfundallocationfilename.Text, "", bInsert)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        updatestatus("Done Processing", "IMP")
        Cursor = Cursors.Default
    End Sub

    Private Sub cboTables_ValueChanged(sender As Object, e As EventArgs) Handles cboTables.ValueChanged
        Dim dt As DataTable = clsdb.getdatatable(" select * from " & cboTables.Value)
        grddata.DataSource = dt
    End Sub

    Private Sub UltraButton2_Click(sender As Object, e As EventArgs) Handles UltraButton2.Click
        Dim retval = MsgBox("Clear Activity Log?  This will Delete all Records.", MsgBoxStyle.YesNo)
        If retval = vbYes Then
            Dim sql$ = " delete from activity"
            clsdb.doQuery(sql$)
        End If
    End Sub

    Private Sub UltraButton3_Click(sender As Object, e As EventArgs) Handles UltraButton3.Click
        Dim retval = MsgBox("Clear Error Log?  This will Delete all Records.", MsgBoxStyle.YesNo)
        If retval = vbYes Then
            Dim sql$ = " delete from errorlog"
            clsdb.doQuery(sql$)
        End If
    End Sub

    Private Sub grderrors_InitializeLayout(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grderrors.InitializeLayout
        e.Layout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.None
        e.Layout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.None
        e.Layout.Override.BorderStyleHeader = Infragistics.Win.UIElementBorderStyle.None
        e.Layout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False

        Dim dt As DataTable = grderrors.DataSource
        lblerrorrows.Text = "Rows: " & dt.Rows.Count
    End Sub

    Private Sub grddata_InitializeLayout(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grddata.InitializeLayout
        Dim dt As DataTable = grddata.DataSource
        lbldatarows.Text = "Rows: " & dt.Rows.Count
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '  generate_alloc_xmlfile("")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        generate_DER_xmlfile("", "")
    End Sub

    Private Sub UltraButton4_Click(sender As Object, e As EventArgs) Handles UltraButton4.Click
        Dim dt As DataTable = clsdb.getdatatable(" select * from plangroups ")
        grdplangroups.DataSource = dt
    End Sub

    Private Sub grdplangroups_InitializeLayout(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdplangroups.InitializeLayout
        e.Layout.Bands(0).Columns("active").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
    End Sub

    Private Sub UltraButton5_Click(sender As Object, e As EventArgs) Handles UltraButton5.Click
        Dim dt As DataTable = grdplangroups.DataSource
        For Each dr As DataRow In dt.Rows
            Dim sql$ = " update plangroups set active = " & dr("active") & " where planid = '" & dr("planid") & "' and plangrpname = '" & dr("plangrpname") & "'"

            clsdb.doQuery(sql$)
        Next
    End Sub

  

End Class
