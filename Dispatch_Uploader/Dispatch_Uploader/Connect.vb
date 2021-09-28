Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms

Module Connect
    'Public connectionString As String = "Data Source=10.3.2.26,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficeDBC;User ID=admin;Password=a;"
    'Public connectionString As String = "Data Source=10.1.6.31\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficeDBC;User ID=sa;Password=admin@Sfa99;"
    'Public connectionString As String = "Data Source=10.1.6.31\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficePPL;User ID=sa;Password=admin@Sfa99;"
    'Public connectionString As String = "Data Source=10.1.6.31\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficeCWM;User ID=sa;Password=admin@Sfa99;"

    Public connectionStringSmarter As String = "Data Source=SQL5016.Smarterasp.net;Initial Catalog=DB_9AAA2F_SFAHeadoffice;User Id=DB_9AAA2F_SFAHeadoffice_admin;Password=dbcl1234;Connect Timeout=360"
    '  Public connectionStringSmarter As String = "Data Source=SQL5016.Smarterasp.net;Initial Catalog=DB_9AAA2F_SFAHeadoffice;User Id=DB_9AAA2F_SFAHeadoffice_admin;Password=dbcl1234;"
    Public UCompanyCode, UUsername, UUserType, UPassword As String

    Public connection_SMARTER As New SqlConnection(connectionStringSmarter)
    '...........    ................................................................
    '
    '
    Public connectionString As String = "Data Source= SFADB.CFLB.NET,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficeDBC;User ID=sa;Password=admin@Sfa99;"
    'Public connectionString As String = "Data Source=SFADB.CFLB.NET,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficeDBC;User ID=sa;Password=admin@Sfa99;"
    'Public connectionString As String = "Data Source=SFADB.CFLB.NET,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficePPL;User ID=sa;Password=admin@Sfa99;"
    ' Public connectionString As String = "Data Source=SFADB.CFLB.NET,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficeCWM;User ID=sa;Password=admin@Sfa99;"

    '========================================
    'For Clone Machine

    'Public connectionString As String = "Data Source= sfadb-dr.cflb.net,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficeDBC;User ID=sa;Password=admin@Sfa99;"
    'Public connectionString As String = "Data Source=10.1.6.36,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficePPL;User ID=sa;Password=admin@Sfa99;"
    'Public connectionString As String = "Data Source=10.1.6.36,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficeCWM;User ID=sa;Password=admin@Sfa99;"
    '============================================

    '.........................................

    Dim tables As DataTableCollection
    Dim source1 As New BindingSource
    Dim ds As DataSet
    Dim Sql As String
    Dim connection As SqlConnection
    Dim dataadapter As SqlDataAdapter
    Dim dataadapterH As SqlDataAdapter
    Dim dataadapterS As SqlDataAdapter
    Dim SqlEx As SqlCommand
    Dim cmdBuilder As SqlCommandBuilder
    Dim dschange As DataSet
    Function GetDataSQL(ByVal p1 As String, ByVal p2 As String, ByVal p3 As String, ByVal p4 As String, ByVal p5 As String, ByVal pn1 As String, ByVal pn2 As String, ByVal pn3 As String, ByVal pn4 As String, ByVal pn5 As String, ByVal sql As String) As System.Data.DataSet
        Try

            ' con.Open()
            Dim queryString As String = sql
            Dim dbCommand As System.Data.IDbCommand = New System.Data.SqlClient.SqlCommand
            dbCommand.CommandText = queryString
            dbCommand.Connection = connection_SMARTER

            Dim dbParam_p1 As System.Data.IDataParameter = New System.Data.SqlClient.SqlParameter
            dbParam_p1.ParameterName = pn1
            dbParam_p1.Value = p1
            dbParam_p1.DbType = System.Data.DbType.String
            dbCommand.Parameters.Add(dbParam_p1)

            Dim dbParam_p2 As System.Data.IDataParameter = New System.Data.SqlClient.SqlParameter
            dbParam_p2.ParameterName = pn2
            dbParam_p2.Value = p2
            dbParam_p2.DbType = System.Data.DbType.String
            dbCommand.Parameters.Add(dbParam_p2)

            Dim dbParam_p3 As System.Data.IDataParameter = New System.Data.SqlClient.SqlParameter
            dbParam_p3.ParameterName = pn3
            dbParam_p3.Value = p3
            dbParam_p3.DbType = System.Data.DbType.String
            dbCommand.Parameters.Add(dbParam_p3)

            Dim dataAdapter As System.Data.IDbDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            dataAdapter.SelectCommand = dbCommand
            Dim dataSet As System.Data.DataSet = New System.Data.DataSet
            dataAdapter.Fill(dataSet)

            'con.Close()
            Return dataSet

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            '  con.Close()
        End Try

    End Function
    Public Sub WSDLRes(ByVal WSDL As String, ByVal WResponse As String)
        Dim connection2 = New SqlConnection(connectionString)
        Sql = "Insert into WSDLLog (WSDLName,WResponse,WUser,WTimeStamp) Values ('" & WSDL & "','" & Strings.Replace(WResponse, "'", "") & "','" & UUsername & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "')"
        SqlEx = New SqlCommand(Sql, connection2)
        connection2.Open()
        SqlEx.ExecuteNonQuery()
        connection2.Close()

    End Sub

    Public Sub WSDLResLog(ByVal ResTable As DataTable)
        Try
            Using connection As New SqlConnection(connectionString)
                Using sqlBulkCopy As New SqlBulkCopy(connection.ConnectionString, SqlBulkCopyOptions.Default)
                    'Set the database table name
                    sqlBulkCopy.DestinationTableName = "WSDLLog"
                    connection.Open()
                    sqlBulkCopy.WriteToServer(ResTable)
                    connection.Close()

                End Using
            End Using

        Catch ex As Exception
            '         MsgBox(ex.ToString)
        End Try

    End Sub
End Module
