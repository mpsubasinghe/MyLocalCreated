Imports System.Data.SqlClient

Public Class Form1
    Dim tables As DataTableCollection
    Dim source1 As New BindingSource
    Dim ds As DataSet
    Dim ds3 As DataSet
    Dim ds6 As New DataSet
    Dim Sql As String
    Dim connection As SqlConnection
    Dim dataadapter As SqlDataAdapter
    Dim dataadapterH As SqlDataAdapter
    Dim dataadapterS As SqlDataAdapter
    Dim SqlEx As SqlCommand
    Dim cmdBuilder As SqlCommandBuilder
    Dim dschange As DataSet
    Dim DetDs As DataSet
    Dim EditDs As DataSet
    Dim dsTO As DataSet


    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        connection_SMARTER.Open()

        UCompanyCode = "1100"
        UUsername = "dbuser"

        Me.Visible = True
        Timer1.Enabled = True
    End Sub

    Private Sub Form1_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            'NotifyIcon1.Visible = True
            'NotifyIcon1.Icon = SystemIcons.Application
            'NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
            'NotifyIcon1.BalloonTipTitle = "Verificador corriendo"
            'NotifyIcon1.BalloonTipText = "Verificador corriendo"
            'NotifyIcon1.ShowBalloonTip(50000)
            ''Me.Hide()
            'ShowInTaskbar = False
        End If
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()
        ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
        NotifyIcon1.Visible = False
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        NotifyIcon1.Visible = True
        NotifyIcon1.Icon = SystemIcons.Application
        NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
        NotifyIcon1.BalloonTipTitle = "Dispatch Upload"
        NotifyIcon1.BalloonTipText = "Running Uploading Process"
        NotifyIcon1.ShowBalloonTip(50000)
        Me.Hide()
        ShowInTaskbar = False
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Button2.Enabled = False

        RadioButton1.Checked = True
        DispatchUpload_View()

        DispatchUpload()

        RadioButton2.Checked = True
        DispatchUpload_View()

        Button2.Enabled = True
    End Sub

    Private Sub DispatchUpload_View()
        DataGridView1.Rows.Clear()

        Dim CusQry As String
        Using connection As New SqlConnection(connectionString)
            Dim commands As New SqlCommand("Select CustomerCode from DispatchCustomers", connection)
            connection.Open()
            Dim readers As SqlDataReader = commands.ExecuteReader
            CusQry = ""
            If readers.HasRows Then
                Do While readers.Read
                    If CusQry = "" Then
                        CusQry = " InvoiceType = '" & readers.GetValue(0) & "'"
                    Else
                        CusQry = CusQry & " or InvoiceType = '" & readers.GetValue(0) & "'"
                    End If
                Loop
            End If

        End Using


        If RadioButton1.Checked = True Then
            Sql = "Select CompanyCode,DistributorCode,Date,DivCode,DGDInvoiceID from Dispatch where CompanyCode = '" & UCompanyCode & "' and (" & CusQry & ")  and SUPdated = 'False' Group by CompanyCode,DistributorCode,Date,DivCode,DGDInvoiceID Order By DGDInvoiceID "
        End If

        If RadioButton2.Checked = True Then
            Sql = "Select CompanyCode,DistributorCode,Date,DivCode,DGDInvoiceID from Dispatch where Date='" & DateTime.Now.ToString("yyyy-MM-dd") & "' AND  CompanyCode = '" & UCompanyCode & "' and (" & CusQry & ")  and SUPdated = 'True' Group by CompanyCode,DistributorCode,Date,DivCode,DGDInvoiceID Order By DGDInvoiceID "
        End If

        connection = New SqlConnection(connectionString)
        dataadapter = New SqlDataAdapter(Sql, connection)
        ds3 = New DataSet
        connection.Open()
        dataadapter.Fill(ds3, "Dispatch")
        connection.Close()


        If ds3.Tables(0).Rows.Count > 0 Then
            For a = 0 To ds3.Tables(0).Rows.Count - 1
                Application.DoEvents()
                DataGridView1.Rows.Add(New String() {ds3.Tables(0).Rows(a).Item("DGDInvoiceID").ToString, ds3.Tables(0).Rows(a).Item("DistributorCode").ToString, ds3.Tables(0).Rows(a).Item("Date").ToString, DateTime.Now, ds3.Tables(0).Rows(a).Item("DivCode").ToString, ds3.Tables(0).Rows(a).Item("CompanyCode").ToString})
                'pNo = (ds.Tables(0).Rows(a).Item("ProductID"))
            Next
        End If

    End Sub

    Private Sub DispatchUpload()
      

        Dim connection1 As New SqlConnection(connectionString)
        Dim command1 As New SqlCommand("Select SUser from Session where CompanyCode = '" & UCompanyCode & "' and SType = 'DispatchUpload'", connection1)
        connection1.Open()
        Dim reader1 As SqlDataReader = command1.ExecuteReader
        Dim Suser As String

        If reader1.HasRows Then
            reader1.Read()
            Suser = reader1.GetString(0)

            '      MsgBox("Freeissue upload Process is currently running by User : " & Suser & ". Please try again later")
            Exit Sub

        Else

            connection = New SqlConnection(connectionString)
            Sql = "Insert into Session (CompanyCode,SType,SUser) values ('" & UCompanyCode & "','DispatchUpload','" & UUsername & "')"
            SqlEx = New SqlCommand(Sql, connection)
            If connection.State = ConnectionState.Open Then
                connection.Close()
            End If
            connection.Open()
            SqlEx.ExecuteNonQuery()
            connection.Close()

        End If
        connection1.Close()















        ' '' '' '' ''Dim Nav As New Panel
        ' '' '' '' ''Nav = DispatchPnl

        ' '' '' '' ''Nav.Dock = DockStyle.Top
        ' '' '' '' ''XtraScrollableControl1.Controls.Add(Nav)
        ' '' '' '' ''Nav.Show()



        'connection = New SqlConnection(connectionStringSmarter)
        'Sql = "Delete from FreeIssueChild where CompanyCode = '" & UCompanyCode & "'"
        'SqlEx = New SqlCommand(Sql, connection)
        'connection.Open()
        'SqlEx.ExecuteNonQuery()
        'connection.Close()



        Dim CusQry As String
        Using connection As New SqlConnection(connectionString)
            Dim commands As New SqlCommand("Select CustomerCode from DispatchCustomers", connection)
            connection.Open()
            Dim readers As SqlDataReader = commands.ExecuteReader
            CusQry = ""
            If readers.HasRows Then
                Do While readers.Read
                    If CusQry = "" Then
                        CusQry = " InvoiceType = '" & readers.GetValue(0) & "'"
                    Else
                        CusQry = CusQry & " or InvoiceType = '" & readers.GetValue(0) & "'"
                    End If
                Loop
            End If

        End Using

        Dim row As DataRow


        Sql = "Select * from Dispatch where CompanyCode = '" & UCompanyCode & "' and (" & CusQry & ")  and SUPdated = 'False' order by DivCode,DistributorCode,DGDInvoiceID,LineNumber"
        connection = New SqlConnection(connectionString)
        dataadapter = New SqlDataAdapter(Sql, connection)
        ds = New DataSet
        connection.Open()
        dataadapter.Fill(ds, "Dispatch")
        connection.Close()



        Sql = "Select * from Dispatch where 1=0"
        connection = New SqlConnection(connectionString)
        dataadapter = New SqlDataAdapter(Sql, connection)
        dsTO = New DataSet
        connection.Open()
        dataadapter.Fill(dsTO, "DispatchTO")
        connection.Close()


        DispPro.Value = 0
        DispPro.Maximum = ds.Tables(0).Rows.Count
        Dim Prc As Double
        Dim TemPrc As String = ""

        Dim Pass, Ava, Block As Boolean

        For i = 0 To ds.Tables(0).Rows.Count - 1
            DispPro.Value = DispPro.Value + 1
            DispPro.Refresh()
            row = dsTO.Tables(0).NewRow
            'MsgBox(ds.Tables(0).Rows(i).Item("DistributorCode"))
            'MsgBox(ds.Tables(0).Rows(i).Item("DGDInvoiceID"))
            'MsgBox(ds.Tables(0).Rows(i).Item("UnitRate").ToString.Trim)


            Using connection As New SqlConnection(connectionString)
                Dim commands As New SqlCommand("Select PriceNonVAT from Product where CompanyCode = '" & UCompanyCode & "' and  ProductID = '" & ds.Tables(0).Rows(i).Item("ProductID") & "' and DivCode = '" & ds.Tables(0).Rows(i).Item("DivCode") & "' order by EffectiveDate desc", connection)
                connection.Open()
                Dim readers As SqlDataReader = commands.ExecuteReader
                CusQry = ""
                If readers.HasRows Then
                    readers.Read()
                    TemPrc = readers.GetValue(0)
                Else
                    TemPrc = ""
                End If
            End Using


            Pass = True
            Ava = True
            Block = True

            ' '' ''  Using connection As New SqlConnection(connectionStringSmarter)
            '' ''Dim commands1 As New SqlCommand("Select * from Dispatch where CompanyCode = '" & UCompanyCode & "' and  DGDInvoiceID = '" & ds.Tables(0).Rows(i).Item("DGDInvoiceID") & "'", connection_SMARTER)
            ' '' ''  connection.Open()
            '' ''Dim readers1 As SqlDataReader = commands1.ExecuteReader
            '' ''If readers1.HasRows Then
            '' ''    Ava = True
            '' ''Else
            '' ''    Ava = False
            '' ''End If
            ' '' ''  End Using


            ds6 = GetDataSQL("", "p2", "p3", "p4", "p5", "@dt", "pn2", "pn3", "pn4", "pn5", "Select * from Dispatch where CompanyCode = '" & UCompanyCode & "' and  DGDInvoiceID = '" & ds.Tables(0).Rows(i).Item("DGDInvoiceID") & "'")
            '   MsgBox(ds5.Tables(0).Rows.Count)
            If ds6.Tables(0).Rows.Count > 0 Then
                Ava = True
            Else
                Ava = False
            End If

            Using connection As New SqlConnection(connectionString)
                Dim commands As New SqlCommand("Select * from DispatchRejectProduct  where ProductID = '" & ds.Tables(0).Rows(i).Item("ProductID") & "'", connection)
                connection.Open()
                Dim readers As SqlDataReader = commands.ExecuteReader
                If readers.HasRows Then
                    Block = True
                Else
                    Block = False
                End If
            End Using

            If ds.Tables(0).Rows(i).Item("UnitRate").ToString.Trim = "" Then
                If TemPrc = "" Then
                    Pass = False
                End If
            End If


            If Block = True Then
                Pass = False
            End If

            If Ava = True Then
                Pass = False
            End If


            If Pass = True Then


                If ds.Tables(0).Rows(i).Item("UnitRate").ToString.Trim = "" Then
                    'If TemPrc = "" Then
                    '    MsgBox("Blank UnitRate found in dispatch, the product code not found product master. (" & ds.Tables(0).Rows(i).Item("ProductID") & ")")
                    '    GoTo last
                    'End If
                    row("UnitRate") = TemPrc
                    Prc = TemPrc
                Else
                    row("UnitRate") = ds.Tables(0).Rows(i).Item("UnitRate")
                    Prc = ds.Tables(0).Rows(i).Item("UnitRate")
                End If
                row("CompanyCode") = ds.Tables(0).Rows(i).Item("CompanyCode")
                row("DistributorCode") = ds.Tables(0).Rows(i).Item("DistributorCode")
                row("Date") = ds.Tables(0).Rows(i).Item("Date")
                row("LineNumber") = ds.Tables(0).Rows(i).Item("LineNumber")
                row("InvoiceType") = ds.Tables(0).Rows(i).Item("InvoiceType")
                row("ProductID") = ds.Tables(0).Rows(i).Item("ProductID").ToString.Trim & Prc.ToString
                row("POID") = Strings.Replace(ds.Tables(0).Rows(i).Item("POID"), "'", "")
                'MsgBox(ds.Tables(0).Rows(i).Item("DistributorCode"))
                'MsgBox(ds.Tables(0).Rows(i).Item("DGDInvoiceID"))
                row("DGDInvoiceID") = ds.Tables(0).Rows(i).Item("DGDInvoiceID")
                row("Quantity") = ds.Tables(0).Rows(i).Item("Quantity")
                row("GRSVAL") = ds.Tables(0).Rows(i).Item("GRSVAL")

                If ds.Tables(0).Rows(i).Item("GRSVAL") > 0 Then
                    row("Discount") = Strings.Replace(ds.Tables(0).Rows(i).Item("Discount"), "-", "")
                ElseIf ds.Tables(0).Rows(i).Item("GRSVAL") < 0 Then
                    row("Discount") = ds.Tables(0).Rows(i).Item("Discount")
                End If

                row("Vat") = ds.Tables(0).Rows(i).Item("Vat")
                row("BatchNo") = ds.Tables(0).Rows(i).Item("BatchNo")
                row("BatchNo2") = ds.Tables(0).Rows(i).Item("BatchNo2")
                row("Bank") = ds.Tables(0).Rows(i).Item("Bank")
                row("ChequeNo") = ds.Tables(0).Rows(i).Item("ChequeNo")
                row("DivCode") = ds.Tables(0).Rows(i).Item("DivCode")
                row("SectorCode2") = ds.Tables(0).Rows(i).Item("SectorCode2")
                row("UnitPerCase") = ds.Tables(0).Rows(i).Item("UnitPerCase")
                row("UnitofMeasure") = ds.Tables(0).Rows(i).Item("UnitofMeasure")
                row("ExpiryDate") = ds.Tables(0).Rows(i).Item("ExpiryDate")
                row("SUPdated") = ds.Tables(0).Rows(i).Item("SUPdated")
                dsTO.Tables(0).Rows.Add(row)


            End If
            'row(0) = ds.Tables(0).Rows(i).Item("CompanyCode")
            'row(1) = ds.Tables(0).Rows(i).Item("DistributorCode")
            'row(2) = ds.Tables(0).Rows(i).Item("Date")
            'row(3) = ds.Tables(0).Rows(i).Item("LineNumber")
            'row(4) = ds.Tables(0).Rows(i).Item("InvoiceType")
            'row(5) = ds.Tables(0).Rows(i).Item("ProductID").ToString.Trim & ds.Tables(0).Rows(i).Item("UnitRate").ToString.Trim
            'row(6) = ds.Tables(0).Rows(i).Item("POID")
            'row(7) = ds.Tables(0).Rows(i).Item("DGDInvoiceID")
            'row(8) = ds.Tables(0).Rows(i).Item("Quantity")
            'row(9) = ds.Tables(0).Rows(i).Item("UnitRate")
            'row(10) = ds.Tables(0).Rows(i).Item("Discount")
            'row(11) = ds.Tables(0).Rows(i).Item("Vat")
            'row(12) = ds.Tables(0).Rows(i).Item("GRSVAL")
            'row(13) = ds.Tables(0).Rows(i).Item("BatchNo")
            'row(14) = ds.Tables(0).Rows(i).Item("BatchNo2")
            'row(15) = ds.Tables(0).Rows(i).Item("Bank")
            'row(16) = ds.Tables(0).Rows(i).Item("ChequeNo")
            'row(17) = ds.Tables(0).Rows(i).Item("DivCode")
            'row(18) = ds.Tables(0).Rows(i).Item("SectorCode2")
            'row(19) = ds.Tables(0).Rows(i).Item("UnitPerCase")
            'row(20) = ds.Tables(0).Rows(i).Item("UnitofMeasure")
            'row(21) = ds.Tables(0).Rows(i).Item("ExpiryDate")
            'row(22) = ds.Tables(0).Rows(i).Item("SUPdated")

            'TemPrc = ds.Tables(0).Rows(i).Item("UnitRate")

            FIDStslbl.Text = "Updating Dispatch to Smater.. " & i
            FIDDatlbl.Text = Now
            Application.DoEvents()
        Next i







        Try
            '' '' Using connection As New SqlConnection(connectionStringSmarter)
            ' ''Using sqlBulkCopy As New SqlBulkCopy(connection_SMARTER.ConnectionString, SqlBulkCopyOptions.Default)
            ' ''    'Set the database table name
            ' ''    sqlBulkCopy.DestinationTableName = "Dispatch"
            ' ''    ' connection.Open()
            ' ''    sqlBulkCopy.WriteToServer(dsTO.Tables(0))
            ' ''    ' connection.Close()

            ' ''End Using
            '' ''  End Using


            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(connection_SMARTER)
                bulkCopy.DestinationTableName = "Dispatch"
                Try
                    ' Write from the source to the destination.
                    bulkCopy.WriteToServer(dsTO.Tables(0))

                Catch ex As Exception
                    MsgBox(ex.Message)
                Finally
                    ' Close the SqlDataReader. The SqlBulkCopy
                    ' object is automatically closed at the end
                    ' of the Using block.

                End Try
            End Using


            '    Dim tt(ds.Tables(0).Columns.Count - 1)

            For x = 0 To ds.Tables(0).Rows.Count - 1

                connection = New SqlConnection(connectionString)
                Sql = "Update Dispatch set SUPdated = 'True'  where CompanyCode = '" & ds.Tables(0).Rows(x).Item("CompanyCode") & "' and  DistributorCode = '" & ds.Tables(0).Rows(x).Item("DistributorCode") & "' and Date ='" & ds.Tables(0).Rows(x).Item("Date") & "' and LineNumber ='" & ds.Tables(0).Rows(x).Item("LineNumber") & "' and InvoiceType ='" & ds.Tables(0).Rows(x).Item("InvoiceType") & "' and ProductID ='" & ds.Tables(0).Rows(x).Item("ProductID") & "' and POID ='" & ds.Tables(0).Rows(x).Item("POID") & "' and DGDInvoiceID ='" & ds.Tables(0).Rows(x).Item("DGDInvoiceID") & "' and Quantity ='" & ds.Tables(0).Rows(x).Item("Quantity") & "' and UnitRate ='" & ds.Tables(0).Rows(x).Item("UnitRate") & "'  and DivCode ='" & ds.Tables(0).Rows(x).Item("DivCode") & "' and UnitPerCase ='" & ds.Tables(0).Rows(x).Item("UnitPerCase") & "' and UnitofMeasure ='" & ds.Tables(0).Rows(x).Item("UnitofMeasure") & "'"
                SqlEx = New SqlCommand(Sql, connection)
                connection.Open()
                SqlEx.ExecuteNonQuery()
                connection.Close()

            Next



        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


Last:

        connection = New SqlConnection(connectionString)
        Sql = "Delete from Session  where CompanyCode = '" & UCompanyCode & "' and  SType ='DispatchUpload' and SUser = '" & UUsername & "'"
        SqlEx = New SqlCommand(Sql, connection)
        connection.Open()
        SqlEx.ExecuteNonQuery()
        connection.Close()

        DispPro.Value = 0

        'Nav.Visible = False
        'Nav.Dock = DockStyle.None









    End Sub

   
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        RadioButton1.Checked = True
        DispatchUpload_View()

        DispatchUpload()

        RadioButton2.Checked = True
        DispatchUpload_View()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton2.CheckedChanged

    End Sub

    Private Sub RadioButton1_Click(sender As Object, e As System.EventArgs) Handles RadioButton1.Click
        DispatchUpload_View()
    End Sub

    Private Sub RadioButton2_Click(sender As Object, e As System.EventArgs) Handles RadioButton2.Click
        DispatchUpload_View()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
