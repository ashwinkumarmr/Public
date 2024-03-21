Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataReader
Module Module1
    Public sConnectionString As String
    Public sServer As String
    Public sUserID As String
    Public sPassword As String
    Public sDatabase As String
    Public sProfessionals As String
    Public sBasePath As String = ""
    Public dtb As New DataTable
    Public sEventClass As String
    Public sSMTPServer As String
    Public sSMTPPort As String
    Public sSMTPSecurity As Boolean
    Public sSMTPADDRESS As String
    Public sFailureTOADDRESS As String
    Public sFailureCCADDRESS As String
    Public sFailureBCCADDRESS As String
    Public sSMTPPWD As String
    Public bDebug As Boolean
    Public bError As Boolean = True
    Public Function GetImageDirectory() As String
        Dim sSQL As String = "Select top 1 Value from Prolawini where ident = 'ImagesDir' and Section ='StmnLedgerPrefForm' "
        GetImageDirectory = ""
        Dim oCom As SqlCommand
        Dim oConn As New SqlConnection
        Dim oDR As SqlDataReader = Nothing
        Try
            oConn = New System.Data.SqlClient.SqlConnection(MakeConnection())
            oConn.Open()
            oCom = New System.Data.SqlClient.SqlCommand
            oCom.CommandTimeout = 0
            oCom.Connection = oConn
            oCom.CommandText = sSQL
            oDR = oCom.ExecuteReader
            If oDR.HasRows Then
                While oDR.Read()
                    GetImageDirectory = oDR(0)
                End While
            End If
        Catch ex As Exception
            Form1.ListBox1.Items.Add("GetImageDirectory error " & ex.Message)
        Finally
            oDR.Close()
            oDR = Nothing
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try
        'Form1.ListBox1.Items.Add("GetIMageDirectory value " & GetImageDirectory)
        Return GetImageDirectory
    End Function
    Public Function MakeConnection() As String
        MakeConnection = ""
        Try

            MakeConnection = ConfigurationManager.ConnectionStrings.Item("ARMailerDB").ConnectionString
            Return MakeConnection

            sServer = ConfigurationManager.AppSettings("Server")
            sDatabase = ConfigurationManager.AppSettings("Database")
            sPassword = ConfigurationManager.AppSettings("Password")
            sUserID = ConfigurationManager.AppSettings("UserID")

            If ConfigurationManager.AppSettings("UseTrustedConnection") = "N" Then
                MakeConnection = "Network Library=DBMSSOCN; Initial Catalog=" & sDatabase & ";Data Source= " & sServer & ";uid=" & sUserID & ";pwd=" & sPassword
            Else
                MakeConnection = "Network Library=DBMSSOCN; Initial Catalog=" & sDatabase & ";Data Source= " & sServer & ";Trusted_Connection=yes"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        'If bDebug Then
        'MakeConnection = "Network Library=DBMSSOCN; Initial Catalog=FPOHC001;Data Source=jsql17;uid=jSQLLLC;pwd=jSQLLLC"
        'End If

        Return MakeConnection
    End Function
End Module
