Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataReader
Imports System.IO
Imports System
Imports System.Text
Imports System.Threading
Imports System.Net.Mail
Imports BCL.easyPDF8.Interop.EasyPDFPrinter
'Imports BCL.easyPDF.Printer


Imports MailKit.Net.Smtp
Imports MimeKit
Imports VB = Microsoft.VisualBasic
Imports SmtpClient = System.Net.Mail.SmtpClient
Imports MailKit.Net
Imports MailKit.Security





Public Class Form1
  
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        sUserID = ConfigurationManager.AppSettings("UserID")
        sPassword = ConfigurationManager.AppSettings("Password")
        sProfessionals = "datacon"
        sSMTPServer = ConfigurationManager.AppSettings("SMTPServer")
        sSMTPPort = ConfigurationManager.AppSettings("SMTPPort")
        sSMTPSecurity = ConfigurationManager.AppSettings("SMTPSecurity")
        sSMTPADDRESS = ConfigurationManager.AppSettings("SMTPADDRESS")
        sFailureTOADDRESS = ConfigurationManager.AppSettings("FailureTOADDRESS")
        sFailureCCADDRESS = ConfigurationManager.AppSettings("FailureCCADDRESS")
        sFailureBCCADDRESS = ConfigurationManager.AppSettings("FailureBCCADDRESS")
        sSMTPPWD = ConfigurationManager.AppSettings("SMTPPWD")
        sEventClass = ConfigurationManager.AppSettings("EventClass")
        bDebug = ConfigurationManager.AppSettings("Debug")
        BuildDataTable()
        GetGeneratingOffice()

    End Sub
    Function GetProfessionals(ByRef sMatters As String) As String
        Dim sSQL As String
        Dim oCom As SqlCommand
        Dim oConn As New SqlConnection
        Dim oDR As SqlDataReader = Nothing
        GetProfessionals = ""
        sSQL = "SELECT top 1 Professionals.Professionals from Professionals, MattersProfessionals MP where MP.Professionals = Professionals.Professionals and MP.AssignedType in ('Responsible','Responsible Attorney') and MP.Matters = '" & sMatters & "'"
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
                    GetProfessionals = oDR(0)
                End While
            End If
        Catch ex As Exception
            ListBox1.Items.Add("GetProfessionals error " & ex.Message)
        Finally
            oDR.Close()
            oDR = Nothing
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try
        If GetProfessionals.Length < 2 Then
            GetProfessionals = "Datacon"
        End If
        Return GetProfessionals
    End Function
    Function GetsBasePath() As String
        GetsBasePath = ""
        Dim sSQL As String
        Dim oCom As SqlCommand
        Dim oConn As New SqlConnection
        Dim oDR As SqlDataReader = Nothing

        sSQL = "SELECT value from Prolawini where Ident = 'docdir'"
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
                    GetsBasePath = oDR(0) & "\"
                End While
            End If
        Catch ex As Exception
            ListBox1.Items.Add("GetsBasePath error " & ex.Message)
        Finally
            oDR.Close()
            oDR = Nothing
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try
        Return GetsBasePath


    End Function
    Private Sub GetGeneratingOffice()

        Dim sSQL As String
        Dim oCom As SqlCommand
        Dim oConn As New SqlConnection
        Dim oDR As SqlDataReader = Nothing
        sSQL = "Select Distinct QGENERATINGOFFI from Matters where isnull(QGENERATINGOFFI  ,'')<>'' order by QGENERATINGOFFI"
        Try
            oConn = New System.Data.SqlClient.SqlConnection(MakeConnection())
            oConn.Open()
            oCom = New System.Data.SqlClient.SqlCommand
            oCom.CommandTimeout = 0
            oCom.Connection = oConn
            oCom.CommandText = sSQL
            oDR = oCom.ExecuteReader()
            If Not oDR.HasRows Then
                ListBox1.Items.Add("Errors in GetGeneratingOffice routine.")
            End If
            While oDR.Read()
                Try
                    CheckedListBox2.Items.Add(oDR(0).ToString)

                Catch ex As Exception
                    ListBox1.Items.Add("Error in GetGeneratingOffice " & ex.Message)
                End Try
            End While
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine()
        Finally
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try

    End Sub





    Function GetDocDir(ByRef sMatters As String, ByRef sEventTypes As String) As String
        Dim sSQL As String
        Dim oCom As SqlCommand
        Dim oConn As New SqlConnection
        Dim oDR As SqlDataReader = Nothing
        sProfessionals = GetProfessionals(sMatters)
        GetDocDir = ""
        sSQL = "SELECT dbo.FN_CreateDocPath('" & sMatters & "','" & sEventTypes & "','" & sProfessionals & "')"
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
                    GetDocDir = oDR(0)
                End While
            End If
        Catch ex As Exception
            ListBox1.Items.Add("GetDocDir error " & ex.Message)
        Finally
            oDR.Close()
            oDR = Nothing
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try
        Return GetDocDir
    End Function
    Private Function GetIncrements() As String
        GetIncrements = ""
        Dim sSQL As String
        Dim oCom As SqlCommand
        Dim oConn As New SqlConnection
        Dim oDR As SqlDataReader = Nothing
        sSQL = "Select Increment + 1 from Increments where V7Tables='Documents' and Site = @@ServerName"
        Try
            oConn = New System.Data.SqlClient.SqlConnection(MakeConnection())
            oConn.Open()
            oCom = New System.Data.SqlClient.SqlCommand
            oCom.CommandTimeout = 0
            oCom.Connection = oConn
            oCom.CommandText = sSQL
            oDR = oCom.ExecuteReader()
            If Not oDR.HasRows Then
                ListBox1.Items.Add("Errors.")
            End If
            While oDR.Read()
                Try
                    GetIncrements = oDR(0)
                Catch ex As Exception
                    ListBox1.Items.Add("Error in GetIncrements " & ex.Message)
                End Try
            End While
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine()
        Finally
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try
        Return GetIncrements
    End Function
    Sub SetIncrements()

        Dim sSQL As String
        Dim oCom As SqlCommand
        Dim oConn As New SqlConnection
        sSQL = "Update Increments Set Increment = Increment + 1 from Increments where V7Tables='Documents' and Site = @@ServerName"
        Try
            oConn = New System.Data.SqlClient.SqlConnection(MakeConnection())
            oConn.Open()
            oCom = New System.Data.SqlClient.SqlCommand
            oCom.CommandTimeout = 0
            oCom.Connection = oConn
            oCom.CommandText = sSQL
            oCom.ExecuteNonQuery()
            Console.WriteLine("Complete SetIncrement")
            Console.WriteLine()
        Catch ex As Exception

            Console.WriteLine(ex.Message)
            Console.WriteLine()
        Finally
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try

    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub RefreshContactEmailAddresses()
        'update jSQLARDocuments SET emailaddress = dbo.GetEmailAddress(Contacts)  where CONTACTS is not null and LEN(IsNull(emailaddress,'')) < 3 and dbo.GetEmailAddress(Contacts) <> '~' and Processed = 'N'
        Dim sSQL As String = "update jSQLARDocuments SET emailaddress = dbo.GetEmailAddress(Contacts)  where CONTACTS is not null and LEN(IsNull(emailaddress,'')) < 3 and dbo.GetEmailAddress(Contacts) <> '~' and Processed = 'N'"
        Dim oCom As SqlCommand = Nothing
        Dim oConn As SqlConnection = Nothing
        Try
            oConn = New System.Data.SqlClient.SqlConnection(MakeConnection())
            oConn.Open()
            oCom = New System.Data.SqlClient.SqlCommand
            oCom.CommandTimeout = 0
            oCom.Connection = oConn
            oCom.CommandText = sSQL
            oCom.ExecuteNonQuery()
            ListBox1.Items.Add("Completed Refreshing Contacts Email Addresses")
        Catch ex As Exception
            ListBox1.Items.Add("Failed Refreshing Contacts Email Addresses " & ex.Message)
        Finally
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try
    End Sub
    Private Sub BuildFunctions()
        Dim sSQL As String = ""
        Dim oCom As New SqlCommand
        Dim oConn As New SqlConnection

        sSQL = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[GetClaimsExaminerEmail]') AND type in (N'TF'))"
        sSQL = sSQL & " BEGIN execute dbo.sp_executesql @statement = N'"
        sSQL = sSQL & "CREATE FUNCTION dbo.GetClaimsExaminerEmail(@Matters varchar(36)) " &
                     " RETURNS varchar(70) " &
                    " As  BEGIN " &
                    " DECLARE @EmailAddress varchar(70) " &
                    " SET @EmailAddress = (Select Top 1 P.phoneNo from MattersQClaims MQC, PHone P WHERE MQC.Matters = @Matters and P.Contacts = MQC.QEXAMINER and MQC.QEXAMINER is not null and P.PhoneType in (''email'',''e-mail''))" &
                    " RETURN IsNull(@EmailAddress,'''')" &
                    " END '"
        sSQL = sSQL & " END "

        Try
            oConn = New System.Data.SqlClient.SqlConnection(MakeConnection())
            oConn.Open()
            oCom = New System.Data.SqlClient.SqlCommand
            oCom.CommandTimeout = 0
            oCom.Connection = oConn
            oCom.CommandText = sSQL
            oCom.ExecuteNonQuery()
        Catch ex As Exception
            ListBox1.Items.Add("Error Creating function GetClaimsExaminerEmail." & ex.Message)
        Finally
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try

    End Sub
    Private Sub Populate_Selections_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim sSQL As String = ""
        Dim sACE As String = ""
        'Make database connection
        Dim oDR As SqlDataReader = Nothing
        Dim oCom As SqlCommand = Nothing
        Dim oConn As SqlConnection = Nothing
        dtb.Rows.Clear()
        CheckedListBox1.Items.Clear()
        RefreshContactEmailAddresses()
        If CheckBox2.Checked Then
            sACE = ""
        Else
            sACE = " and Matters.QREFERRALSOURCE <> 'ACE' "
        End If

        oConn = New System.Data.SqlClient.SqlConnection(MakeConnection())
        oConn.Open()
        oCom = New System.Data.SqlClient.SqlCommand
        oCom.CommandTimeout = 300
        oCom.Connection = oConn

        ' Get the selected item's check state.  
        '  Added three lines of code to each select per Mary Kay wanting things filtered out by the document creatio process and not the billing process.  
        For i = 0 To CheckedListBox2.CheckedItems.Count - 1

            If CheckedListBox3.GetItemCheckState(0) = CheckState.Checked And sSQL.Length = 0 Then
                sSQL = "Select jSQLARDocuments.Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from Events, jSQLARDocuments, MattersQClaims, Matters " &
            " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType = '60 day letter' and Matters.MatterId = jSQLARDocuments.MatterID " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
            " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
            " and Events.Events = jSQLARDocuments.Events " &
          " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
            " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
            " and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " &
            sACE
            ElseIf CheckedListBox3.GetItemCheckState(0) = CheckState.Checked And sSQL.Length > 0 Then
                sSQL = sSQL & " Union Select jSQLARDocuments.Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from  Events,jSQLARDocuments, MattersQClaims, Matters " &
            " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType = '60 day letter'  and Matters.MatterId = jSQLARDocuments.MatterID " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
            " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
            " and Events.Events = jSQLARDocuments.Events " &
          " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
            " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
            " and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " & sACE
            End If
            If CheckedListBox3.GetItemCheckState(1) = CheckState.Checked And sSQL.Length = 0 Then
                sSQL = "Select jSQLARDocuments.Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from  Events,jSQLARDocuments, MattersQClaims, Matters " &
            " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType = '90 day letter' and Matters.MatterId = jSQLARDocuments.MatterID " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
            " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
                      " and Events.Events = jSQLARDocuments.Events " &
          " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
            " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
            " and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " & sACE
                '" and Matters.MatterId not like 'ACE%' and Matters.MatterId not like '%ACE%' and Matters.MatterId not like '%ACE' " & _
            ElseIf CheckedListBox3.GetItemCheckState(1) = CheckState.Checked And sSQL.Length > 0 Then
                sSQL = sSQL & " Union Select jSQLARDocuments.Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from  Events,jSQLARDocuments, MattersQClaims, Matters " &
            " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType = '90 day letter' and Matters.MatterId = jSQLARDocuments.MatterID " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
            " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
           " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
            " and Events.Events = jSQLARDocuments.Events " &
            " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
            " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
            " and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " & sACE
            End If
            If CheckedListBox3.GetItemCheckState(2) = CheckState.Checked And sSQL.Length = 0 Then
                sSQL = "Select jSQLARDocuments.Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from  Events,jSQLARDocuments, MattersQClaims, Matters " &
            " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType = 'Agenct Cap Letter' and Matters.MatterId = jSQLARDocuments.MatterID " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
            " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
            " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
            " and Events.Events = jSQLARDocuments.Events " &
                        " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
            " and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " & sACE
                '" and Matters.MatterId not like 'ACE%' and Matters.MatterId not like '%ACE%' and Matters.MatterId not like '%ACE' " & _
            ElseIf CheckedListBox3.GetItemCheckState(2) = CheckState.Checked And sSQL.Length > 0 Then
                sSQL = sSQL & "UNION Select jSQLARDocuments.Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from  Events,jSQLARDocuments, MattersQClaims, Matters " &
            " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType = 'Agenct Cap Letter' and Matters.MatterId = jSQLARDocuments.MatterID " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
            " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
            " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
          " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
            " and Events.Events = jSQLARDocuments.Events " &
                        " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
            " and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
            " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " & sACE
            End If
            'SIR Initial Letter
            If CheckedListBox3.GetItemCheckState(3) = CheckState.Checked And sSQL.Length = 0 Then

                sSQL = "Select jSQLARDocuments.Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from  Events,jSQLARDocuments, MattersQClaims, Matters " &
                " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType = 'Initial Letter' and Matters.MatterId = jSQLARDocuments.MatterID " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
                " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
               " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
                " and Events.Events = jSQLARDocuments.Events " &
                               " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
                " and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " & sACE

            ElseIf CheckedListBox3.GetItemCheckState(3) = CheckState.Checked And sSQL.Length > 0 Then
                sSQL = sSQL & " Union Select jSQLARDocuments.Events Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from  Events, jSQLARDocuments, MattersQClaims, Matters " &
                " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType = 'Initial Letter' and Matters.MatterId = jSQLARDocuments.MatterID " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
                " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
              " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
              " and Events.Events = jSQLARDocuments.Events " &
                              " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
                " and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " & sACE
            End If
            If CheckedListBox3.GetItemCheckState(4) = CheckState.Checked And sSQL.Length = 0 Then
                sSQL = "Select jSQLARDocuments.Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from  Events,jSQLARDocuments, MattersQClaims, Matters " &
                " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType = 'Get ALL AR Data' and Matters.MatterId = jSQLARDocuments.MatterID " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
                " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
                              " and Events.Events = jSQLARDocuments.Events " &
              " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
                " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
                " and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " & sACE
            ElseIf CheckedListBox3.GetItemCheckState(4) = CheckState.Checked And sSQL.Length > 0 Then
                sSQL = sSQL & " Union Select jSQLARDocuments.Events Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from  Events,jSQLARDocuments, MattersQClaims, Matters " &
                " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType = 'Get ALL AR Data' and Matters.MatterId = jSQLARDocuments.MatterID " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
                " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
                " and Events.Events = jSQLARDocuments.Events " &
               " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
                " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
                " and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " & sACE
            End If

            If CheckedListBox3.GetItemCheckState(5) = CheckState.Checked And sSQL.Length = 0 Then
                sSQL = "Select jSQLARDocuments.Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END email, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from  Events,jSQLARDocuments, MattersQClaims, Matters " &
                " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType like 'Dunning%' and Matters.MatterId = jSQLARDocuments.MatterID " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
                " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
                " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
                " and Events.Events = jSQLARDocuments.Events " &
                " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " & sACE
                '" and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
            ElseIf CheckedListBox3.GetItemCheckState(5) = CheckState.Checked And sSQL.Length > 0 Then
                sSQL = sSQL & " Union Select jSQLARDocuments.Events, CASE WHEN emailaddress = '~' or emailaddress is null then dbo.GetEmailAddress(Matters.Matters) ELSE emailaddress END email, Matters.ShortDesc,Matters.MatterId, stmnDate, Matters.AreaofLaw, Ltrim(Rtrim(FullName)), ltrim(rtrim(Company)), Events.DocDir, StmnNo, LetterType, EventType, REPLACE(SubjectLine,'@',''), REPLACE(Template,char(39), char(34)), IsNull(Contacts,''), dbo.GetClaimsExaminerEmail(Matters.Matters) from  Events,jSQLARDocuments, MattersQClaims, Matters " &
                " where Matters.Matters = MattersQClaims.Matters and Processed = 'N' and LetterType like 'Dunning%' and Matters.MatterId = jSQLARDocuments.MatterID " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QDEDUCTIBLESATI,'N') = 'Y') " &
                " and Matters.Matters not in (Select Matters from MattersQClaims where IsNull(QHOLDARINVOICES,'N') = 'Y') " &
                " and jSQLARDocuments.DocumentDate >= getdate() - 45 " &
                " and IsNull(Matters.QREFERRALSOURCE,'') not in ('GB','ESIS','AGRI')" &
                " and Events.Events = jSQLARDocuments.Events " &
                " and Matters.QGENERATINGOFFI = '" & CheckedListBox2.CheckedItems(i).ToString() & "'" &
                " and DBO.GetMinClaimsNumber(dbo.GetCRN(Matters.Matters)) < '2017-07-17' " &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'For%') " &
                " and Matters.Matters not in (Select Matters from MatterCategories where Category like 'Bank%') " & sACE
            End If

            'Run this before query so we get the latest/most complete result set with valid email addresses. Do not limit updates to the area in question

            oCom.CommandText = sSQL
            oDR = oCom.ExecuteReader()

            If Not oDR.HasRows Then
                ListBox1.Items.Add("No AR letters to process.")
            End If
            While oDR.Read()
                Try
                    If oDR(1).ToString.Length > 0 And oDR(1).ToString.LastIndexOf("@") > 0 Then
                        CheckedListBox1.Items.Add(oDR(2) & " (MatterID: " & oDR(3) & ")" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "PK:" & oDR(0) & vbTab & "SITE:" & sServer & vbTab & "COMPLETE")
                    Else
                        If Not CheckBox1.Checked Then
                            CheckedListBox1.Items.Add("No valid E-mail address for Biling Contact. " & oDR(2) & " (MatterID: " & oDR(3) & ")" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "PK:" & oDR(0) & vbTab & "SITE:" & sServer & vbTab & "COMPLETE")
                        End If
                        ListBox1.Items.Add(ConfigurationManager.AppSettings("Server").ToString & " - No valid E-mail address for Biling Contact.  MatterID: " & oDR(3))
                    End If
                    dtb.Rows.Add(ConfigurationManager.AppSettings("Server").ToString, oDR(0), oDR(6), oDR(7), oDR(3), oDR(2), oDR(4), oDR(9), oDR(8), oDR(1), oDR(10), oDR(11), oDR(12), oDR(13), oDR(14), oDR(15).ToString)
                Catch ex As Exception
                    ListBox1.Items.Add("Error in Inner Populate_Selections_Click " & ex.Message)
                End Try
            End While
        Next
        Try
            oDR.Close()
            oDR = Nothing
        oCom = Nothing
        oConn.Close()
        oConn = Nothing

        Catch ex As Exception

        End Try

        Button1.Enabled = True
            Button2.Enabled = True
        Button4.Enabled = True

    End Sub
    Private Sub BuildDataTable()
        dtb.Columns.Add("Site", GetType(String))
        dtb.Columns.Add("Events", GetType(String))
        dtb.Columns.Add("FullName", GetType(String))
        dtb.Columns.Add("Company", GetType(String))
        dtb.Columns.Add("MatterID", GetType(String))
        dtb.Columns.Add("Description", GetType(String))
        dtb.Columns.Add("StmnDate", GetType(Date))
        dtb.Columns.Add("StmnNo", GetType(Integer))
        dtb.Columns.Add("DocDir", GetType(String))
        dtb.Columns.Add("emailaddress", GetType(String))
        dtb.Columns.Add("LetterTypes", GetType(String))
        dtb.Columns.Add("EventTypes", GetType(String))
        dtb.Columns.Add("SubjectLine", GetType(String))
        dtb.Columns.Add("Template", GetType(String))
        dtb.Columns.Add("Contacts", GetType(String))
        dtb.Columns.Add("ClaimsExaminerEmail", GetType(String))
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        For idx As Integer = 0 To CheckedListBox2.Items.Count - 1
            CheckedListBox2.SetItemCheckState(idx, CheckState.Checked)
        Next
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        For idx As Integer = 0 To CheckedListBox2.Items.Count - 1
            CheckedListBox2.SetItemCheckState(idx, CheckState.Unchecked)
        Next
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For idx As Integer = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemCheckState(idx, CheckState.Unchecked)
        Next
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        For idx As Integer = 0 To CheckedListBox3.Items.Count - 1
            CheckedListBox3.SetItemCheckState(idx, CheckState.Unchecked)
        Next
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        For idx As Integer = 0 To CheckedListBox3.Items.Count - 1
            CheckedListBox3.SetItemCheckState(idx, CheckState.Checked)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For idx As Integer = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemCheckState(idx, CheckState.Checked)
        Next
    End Sub
    
   
    Private Function ParseEventsPK(ByVal sResultString As String) As String
        ParseEventsPK = ""
        Dim iStartIndex As Integer = sResultString.LastIndexOf("PK:")
        Dim iEndIndex As Integer = sResultString.LastIndexOf("SITE:")
        Try
            ParseEventsPK = sResultString.Substring(iStartIndex, iEndIndex - iStartIndex)
            ParseEventsPK = Replace(ParseEventsPK, "PK:", "")
            ParseEventsPK = Replace(ParseEventsPK, vbTab, "")
        Catch ex As Exception
            ListBox1.Items.Add("Error in ParseEventsPK " & ex.Message)
        End Try

        Return ParseEventsPK
    End Function
    Private Function ParseServer(ByVal sResultString As String) As String
        ParseServer = ""
        Dim iStartIndex As Integer = sResultString.LastIndexOf("SITE:")
        Dim iEndIndex As Integer = sResultString.LastIndexOf("COMPLETE")
        Try
            ParseServer = sResultString.Substring(iStartIndex, iEndIndex - iStartIndex)
            ParseServer = Replace(ParseServer, "SITE:", "")
            ParseServer = Replace(ParseServer, "COMPLETE:", "")
            ParseServer = Replace(ParseServer, vbTab, "")
            ParseServer = Trim(ParseServer)

        Catch ex As Exception
            ListBox1.Items.Add("Error in ParseServer " & ex.Message)
        End Try

        Return ParseServer
    End Function
    ' GetEmailAddress(sServerName, LocalContact)
    Function GetEmailAddress(ByVal sLocalContact As String, ByVal sBadEmailAddress As String) As String
        GetEmailAddress = ""
        Dim sSQL As String
        Dim oCom As SqlCommand
        Dim oConn As New SqlConnection
        Dim oDR As SqlDataReader = Nothing
        sSQL = "SELECT top 1 phoneno  from Phone where phoneType in ('email','e-mail') and Phoneno <> '" & sBadEmailAddress & "' and Contacts = '" & sLocalContact & "'"
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
                    GetEmailAddress = oDR(0)
                End While
            End If
        Catch ex As Exception
            ListBox1.Items.Add("GetEmailAddress error " & ex.Message)
        Finally
            oDR.Close()
            oDR = Nothing
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try
        Return GetEmailAddress
    End Function
    Function GetContact(ByVal sStmnNo As String) As String
        GetContact = ""
        Dim sSQL As String
        Dim oCom As SqlCommand
        Dim oConn As New SqlConnection
        Dim oDR As SqlDataReader = Nothing
        sSQL = "SELECT top 1 Contacts from StmnLedger where LedgerType = 'S' and STmnNo = '" & sStmnNo & "'"
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
                    GetContact = oDR(0)
                End While
            End If
        Catch ex As Exception
            ListBox1.Items.Add("GetContact error " & ex.Message)
        Finally
            oDR.Close()
            oDR = Nothing
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try
        Return GetContact
    End Function
    Function GetMattersPK(ByVal sMatterId As String) As String
        GetMattersPK = ""
        Dim sSQL As String
        Dim oCom As SqlCommand
        Dim oConn As New SqlConnection
        Dim oDR As SqlDataReader = Nothing
        sSQL = "SELECT Matters from MAtters where MatterId = '" & sMatterId & "'"
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
                    GetMattersPK = oDR(0)
                End While
            End If
        Catch ex As Exception
            ListBox1.Items.Add("GetMattersPK error " & ex.Message)
        Finally
            oDR.Close()
            oDR = Nothing
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try
        Return GetMattersPK
    End Function

    Function GetPK() As String
        GetPK = ""
        Dim sSQL As String
        Dim oCom As SqlCommand
        Dim oConn As New SqlConnection
        Dim oDR As SqlDataReader = Nothing
        sSQL = "SELECT Cast(NewID() as varchar(36))"
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
                    GetPK = oDR(0)

                End While
            End If
        Catch ex As Exception
            ListBox1.Items.Add("GetPK error " & ex.Message)
        Finally
            oDR.Close()
            oDR = Nothing
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try
        Return GetPK
    End Function
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Read the selected item list and compair to the table created, pull up a document, convert to PDF and formulate the email.
        Dim itemChecked As Object
        Dim itemIndex As Integer = -1
        Dim sEventsPK As String = ""
        Dim sNewEventsPK = GetPK()
        Dim sServerName As String = ""
        Dim sLocalDocDir As String = ""
        Dim sLocalsMatterID As String = ""
        Dim sLocalDescriptions As String = ""
        Dim sLocalEmailAddresss As String = ""
        Dim sLocalMattersPK As String = ""
        Dim sLocalPDFDocDir As String = ""
        Dim sLocalEventTypes As String = ""
        Dim LocalSubjectLine As String = ""
        Dim LocalTemplate As String = ""
        Dim LocalContact As String = ""
        Dim sLocalClaimsExaminerEmail As String = ""
        For Each itemChecked In CheckedListBox1.CheckedItems
            Try
                bError = False
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                itemIndex = -1
                sEventsPK = ""
                sServerName = ""
                sLocalDocDir = ""
                sLocalsMatterID = ""
                sLocalDescriptions = ""
                sLocalEmailAddresss = ""
                sLocalMattersPK = ""
                sLocalPDFDocDir = ""
                sLocalEventTypes = ""
                sLocalClaimsExaminerEmail = ""
                'itemIndex = CheckedListBox1.Items.IndexOf(itemChecked)
                'ListBox1.Items.Add(CheckedListBox1.Items.IndexOf(itemChecked).ToString())
                'ListBox1.Items.Add(itemIndex.ToString())
                'ListBox1.Items.Add(itemChecked.ToString())

                sServerName = ConfigurationManager.AppSettings("Server")
                sEventsPK = ParseEventsPK(itemChecked.ToString())
                'ListBox1.Items.Add(sEventsPK)

                'ListBox1.Items.Add(sServerName)
                sLocalClaimsExaminerEmail = RetrieveClaimsExaminerEmail(sEventsPK)
                sLocalDocDir = RetrieveDocDir(sEventsPK)
                sLocalsMatterID = RetrieveMatters(sEventsPK)
                sLocalDescriptions = RetrievDesc(sEventsPK)
                sLocalEmailAddresss = RetrievEmail(sEventsPK)

                'sLocalEmailAddresss = RetrievEmail(sServerName, sEventsPK)
                sLocalEventTypes = RetrievEventType(sEventsPK)
                sLocalMattersPK = GetMattersPK(sLocalsMatterID)
                sLocalPDFDocDir = ConverttoPDF(sLocalDocDir, sLocalMattersPK, sLocalEventTypes)
                LocalTemplate = RetrieveTemplate(sEventsPK)
                LocalSubjectLine = RetrieveSubjectLine(sEventsPK)
                SetIncrements()
                LocalContact = RetrieveContacts(sEventsPK)
                Try
                    If LocalContact.Length = 0 Or LocalContact = "Financial Lines" Then
                        ListBox1.Items.Add("Searching for New Statement Contacts PK")
                        LocalContact = GetContact(RetrieveSTMNNo(sEventsPK))
                    End If

                Catch ex As Exception
                    ListBox1.Items.Add("Error with evaluating localContact " & LocalContact)
                End Try

                Try
                    If CountCharacter(sLocalEmailAddresss, "@") <> 1 Then
                        sLocalEmailAddresss = GetEmailAddress(LocalContact, sLocalEmailAddresss)
                        If CountCharacter(sLocalEmailAddresss, "@") <> 1 Then
                            ListBox1.Items.Add("Bad Email Address: " & sLocalEmailAddresss)
                            bError = True
                        End If
                    End If
                Catch ex As Exception
                    ListBox1.Items.Add("Error Seeking alt email address " & sLocalEmailAddresss)
                End Try

                'build email
                If Not bError Then
                    Try
                        BuildAREmail(LocalTemplate, sLocalPDFDocDir, sLocalEmailAddresss, sLocalsMatterID, LocalSubjectLine, RetrieveSTMNNo(sEventsPK), GetImageDirectory() & "\", LocalContact, sLocalClaimsExaminerEmail)

                    Catch ex As Exception
                        bError = True
                    End Try

                Else
                    Try
                        BuildFailureEmail(LocalTemplate, sLocalPDFDocDir, sLocalEmailAddresss, sLocalsMatterID, LocalSubjectLine, RetrieveSTMNNo(sEventsPK), GetImageDirectory() & "\", LocalContact)
                        ListBox1.Items.Add("Failure Email Sent to  " & sFailureTOADDRESS.ToString)

                    Catch ex As Exception
                        bError = True
                    End Try
                End If

                'bulild child event
                If Not bDebug And Not bError Then
                    Try
                        BuildEvent(sLocalMattersPK, sLocalPDFDocDir, "PDFDefault", "Converted to PDF", sLocalEmailAddresss, sLocalDescriptions, sLocalsMatterID, sEventsPK)
                    Catch ex As Exception
                        ListBox1.Items.Add("Error building Event " & ex.Message)
                    End Try
                Else
                    If Not bDebug Then
                        ListBox1.Items.Add("Error condition hit, file not generated, email not sent. ")
                        ListBox1.Items.Add(sLocalDocDir)
                    End If
                    ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                End If

            Catch ex As Exception
                ListBox1.Items.Add("Not working(Button4_Click) " & ex.Message)
            End Try
        Next
        Me.Cursor = System.Windows.Forms.Cursors.Default
        ListBox1.Items.Add("Completed emailing selected A/R documents.")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
    End Sub
    Sub BuildEvent(ByVal sMatters As String, ByVal sDocDir As String, ByVal sEventType As String, ByVal sComment As String, ByRef sEmailAddress As String, ByRef sShortDesc As String, ByRef sMatterID As String, ByVal sParentEvents As String)
        Dim oCom As New SqlCommand
        Dim oConn As SqlConnection = Nothing
        Dim sEvents As String = GetPK()
        Dim sDocNo As String
        sDocNo = Replace(Replace(Mid(sDocDir, sDocDir.LastIndexOf("\") + 2, 20), ".PDF", ""), ".pdf", "")
        sProfessionals = GetProfessionals(sMatters)
        Try
            If Not bDebug And Not bError Then
                oConn = New SqlConnection(MakeConnection())
                oConn.Open()
                oCom.Connection = oConn
                Try
                    oCom.CommandText = "INSERT INTO EVENTS(Events, EventKind, ShortNote, EventTypes, EventDate, RTF, ProfSet, AddingProfessionals, AddingDateTime, SearchDoneDate, DocDir, EventClass, Notes, IsReverseView, UseMatterPortalSecurity, IsNew, IsRecurring, ILSKey, eventsparent)" &
                                       " SELECT '" & sEvents & "', 'O', '" & Replace(Replace(sComment, Chr(39), ""), Chr(34), "") & "','" & sEventType & "', CONVERT(VARCHAR(10), GETDATE(), 101), '" & sDocNo & "', '" & sEvents & "','" & sProfessionals & "',CONVERT(VARCHAR(10), GETDATE(), 101), '1899-12-30 00:00:00.000','" & sDocDir & "',Case when '" & sEventClass & "' = '' then NULL Else '" & sEventClass & "' END,'" & Replace(Replace(sComment, Chr(39), ""), Chr(34), "") & "','N','N','N','N','AllView','" & sParentEvents & "'"
                    oCom.ExecuteNonQuery()

                Catch ex As Exception
                    ListBox1.Items.Add("Error Creating Events Record.  " & ex.Message)
                    bError = True
                End Try
                Try
                    oCom.CommandText = "INSERT INTO EVENTMATTERS(EventMatters, Events, Matters)" &
                                       " SELECT '" & sEvents & "', '" & sEvents & "', '" & sMatters & "'"
                    oCom.ExecuteNonQuery()

                Catch ex As Exception
                    ListBox1.Items.Add("Error Creating EventMatters Record.  " & ex.Message)
                    bError = True
                End Try
                Try
                    oCom.CommandText = "INSERT INTO EVENTPROFS(EventPROFS, PROFSET, PROFESSIONALS)" &
                                       " SELECT '" & sEvents & "', '" & sEvents & "', '" & sProfessionals & "'"
                    oCom.ExecuteNonQuery()

                Catch ex As Exception
                    ListBox1.Items.Add("Error Creating EventProfs Record.  " & ex.Message)
                    bError = True
                End Try
                Try
                    oCom.CommandText = "INSERT INTO EVENTTracking(EVENTTracking, Events, PROFESSIONALS, TrackingType, TrackingDate)" &
                                       " SELECT '" & sEvents & "', '" & sEvents & "', '" & sProfessionals & "', 'Converted to PDF', getdate()"
                    oCom.ExecuteNonQuery()

                Catch ex As Exception
                    ListBox1.Items.Add("Error Creating EventTracking Record.  " & ex.Message)
                    bError = True
                End Try
                Try
                    oCom.CommandText = "INSERT INTO EVENTRemind(EVENTREMIND, Events, RemindType, SearchDate, RemindDate)" &
                                       " SELECT '" & sEvents & "', '" & sEvents & "', 'D',CONVERT(VARCHAR(10), GETDATE(), 101),CONVERT(VARCHAR(10), GETDATE(), 101)"
                    oCom.ExecuteNonQuery()
                Catch ex As Exception
                    ListBox1.Items.Add("Error Creating EventRemind Record.  " & ex.Message)
                    bError = True
                End Try


                '  Build EventClassFolder And EVentClassFolderParent table entries

                Dim sEventClassPK As String = GetEventClassPK(sEventClass)
                Dim sEventClassFolder As String = GetEventClassFolder(sMatters, sEventClassPK)
                If sEventClassFolder.Length < 3 Then
                    sEventClassFolder = GetPK()
                End If


                Try

                    '  Only inserts if the record doesn't exist.
                    oCom.CommandText = "If (Select COUNT(*) from sysobjects where name ='EventClassFolder' and XTYPE = 'U') > 0 " &
                                   " BEGIN Insert into EventClassFolder(EventClassFolder, EventClass, MattersContacts, IsMatters, AddingDateTime) " &
                                   " Select '" & sEventClassFolder & "', '" & sEventClassPK & "','" & sMatters & "', 'Y', GETDATE() where '" & sEventClassPK & "' not in (Select EventClass from EventClassFolder where MattersContacts = '" & sMatters & "' and EventClass = '" & sEventClassPK & "') END "

                    oCom.ExecuteNonQuery()

                Catch ex As System.Exception

                    bError = True
                End Try
                '  Insert EventClassFolderParent
                Try
                    oCom.CommandText = "If (Select COUNT(*) from sysobjects where name ='EventClassFolderParent' and XTYPE = 'U') > 0 " &
                                   " BEGIN Insert into EventClassFolderParent(EventClassFolderParent, Events, EventClassFolder)" &
                        " Select newid(), '" & sEvents & "',  '" & sEventClassFolder & "'  WHERE '" & sEvents & "' in (Select Events from Events)  END"
                    oCom.ExecuteNonQuery()
                Catch ex As System.Exception
                    bError = True
                End Try






                Try
                    If Not bError Then
                        If Not bDebug Then
                            oCom.CommandText = "Update jSQLARDocuments SET  Processed  = 'Y' where Events = '" & sParentEvents & "'"
                        Else
                            oCom.CommandText = "Update jSQLARDocuments SET  Processed  = 'T' where Events = '" & sParentEvents & "'"
                        End If
                        oCom.ExecuteNonQuery()
                    End If
                Catch ex As Exception
                    ListBox1.Items.Add("Error Updating existing jSQLARDcoument Record.  " & ex.Message)
                End Try
                Try

                    If Not bError Then
                        If Not bDebug Then
                            oCom.CommandText = "INSERT INTO jSQLARDocuments(Matters,Events, EventType, DocumentDate, Comment, DocDir, LetterType, EmailAddress, Processed,  ShortDesc, MatterID )" &
                                                " SELECT '" & sMatters & "','" & sEvents & "', '" & sEventType & "', getdate(),'" & Replace(Replace(sComment, Chr(39), ""), Chr(34), "") & "','" & sDocDir & "','" & Replace(Replace(sComment, Chr(39), ""), Chr(34), "") & "','" & sEmailAddress & "','N','" & Replace(Replace(sShortDesc, Chr(39), ""), Chr(34), "") & "','" & sMatterID & "'"
                            oCom.ExecuteNonQuery()
                        End If
                    End If

                Catch ex As Exception
                    ListBox1.Items.Add("Error Inserting new jSQLARDcoument Record for PDF.  " & ex.Message)
                End Try
                If Not bError Then
                    If Not bDebug Then
                        oCom.CommandText = "Update Events Set Notes = REPLACE(REPLACE(cast(Notes as varchar(max)),char(39),''), char(34),'') + char(13) + char(10) + 'Document emailed to " & sEmailAddress & "' where Events = '" & sParentEvents & "'"
                        Try
                            oCom.ExecuteNonQuery()
                        Catch ex As Exception
                            ListBox1.Items.Add("Error updating Event to reflect it was emailed.  " & ex.Message)
                            ListBox1.Items.Add(oCom.CommandText)
                        End Try
                    End If
                End If

            Else
                ListBox1.Items.Add("DEBUG MODE, skiping creating event.")
            End If

        Catch Ex As Exception
            ListBox1.Items.Add("BuildEvent error " & Ex.Message)
        Finally
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try
    End Sub

    Private Function GetEventClassFolder(ByVal sMatters As String, ByVal sEventClassPK As String) As String
        GetEventClassFolder = ""
        Dim sSQL As String

        sSQL = "select top 1 EventClassFolder.EventClassFolder from EventClassfolder where " &
            " MattersContacts = '" & sMatters & "' and EventClass = '" & sEventClassPK & "'"
        'Make database connection
        Dim oDR As SqlDataReader = Nothing
        Dim oCom As SqlCommand
        Dim oConn As SqlConnection = Nothing
        Try
            oConn = New System.Data.SqlClient.SqlConnection(MakeConnection())
            oConn.Open()
            oCom = New System.Data.SqlClient.SqlCommand
            oCom.Connection = oConn
            oCom.CommandText = sSQL
            oDR = oCom.ExecuteReader()

            While oDR.Read()
                GetEventClassFolder = oDR(0)
            End While
        Catch ex As System.Exception

        Finally
            oDR.Close()
            oDR = Nothing
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try

        Return GetEventClassFolder
    End Function

    Private Function GetEventClassPK(ByVal sEventClassDesc As String) As String
        Dim sSQL As String
        GetEventClassPK = ""
        sSQL = "Select Top 1 Eventclass from EventClass where EventClassDesc =  '" & sEventClassDesc & "' and EventClassDesc <> 'NONE'"
        'Make database connection
        Dim oDR As SqlDataReader = Nothing
        Dim oCom As SqlCommand
        Dim oConn As SqlConnection = Nothing
        Try
            oConn = New System.Data.SqlClient.SqlConnection(MakeConnection())
            oConn.Open()
            oCom = New System.Data.SqlClient.SqlCommand
            oCom.Connection = oConn
            oCom.CommandText = sSQL
            oDR = oCom.ExecuteReader()

            While oDR.Read()
                Try
                    GetEventClassPK = oDR(0)
                Catch ex As System.Exception

                End Try
            End While
        Catch ex As System.Exception
        Finally
            oDR.Close()
            oDR = Nothing
            oCom = Nothing
            oConn.Close()
            oConn = Nothing
        End Try

        Return GetEventClassPK

    End Function



    Private Sub BuildFailureEmail(ByVal sBody As String, ByVal sAttachment As String, ByVal semailaddress As String, ByVal sMatterID As String, ByRef sSubjectLline As String, ByRef sStmnNo As String, ByVal sImageDirectory As String, ByVal sContacts As String)
        'ListBox1.Items.Add("Entering BuildAREmail")
        Dim emailAttachment As New Mail
        Try
            Dim sLocalSubjectline As String = " Email Failure for MatterID: " & sMatterID & " " & sSubjectLline
            If semailaddress.Length > 0 And semailaddress.LastIndexOf("@") > 0 Then
                emailAttachment.SendFailureAttachments(sBody, semailaddress.ToString, sMatterID, sLocalSubjectline, sAttachment, sLocalSubjectline, sStmnNo, sImageDirectory, sContacts)
            End If
        Catch ex As System.Exception
            ListBox1.Items.Add("Error Sending Failure Email " & ex.Message)
        End Try

    End Sub
    Private Sub BuildAREmail(ByVal sBody As String, ByVal sAttachment As String, ByVal semailaddress As String, ByVal sMatterID As String, ByRef sSubjectLline As String, ByRef sStmnNo As String, ByVal sImageDirectory As String, ByVal sContacts As String, ByVal sClaimsExaminerEmail As String)
        'ListBox1.Items.Add("Entering BuildAREmail")
        Dim emailAttachment As New Mail
        Try
            Dim sLocalSubjectline As String = Replace(sSubjectLline, "@", "")
            If semailaddress.Length > 0 And semailaddress.LastIndexOf("@") > 0 Then
                emailAttachment.SendAttachments(sBody, semailaddress.ToString, sMatterID, sLocalSubjectline, sAttachment, sLocalSubjectline, sStmnNo, sImageDirectory, sContacts, sClaimsExaminerEmail)
            End If
        Catch ex As System.Exception
            ListBox1.Items.Add("Error Sending Email " & ex.Message)
        End Try

    End Sub
    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        Dim cnt As Integer = 0
        For Each c As Char In value
            If c = ch Then
                cnt += 1
            End If
        Next
        Return cnt
    End Function
    Private Function RetrieveMatters(ByVal sEventsPK As String) As String
        RetrieveMatters = ""
        For Each row As DataRow In dtb.Rows
            'strDetail = row.Item("Detail")
            If row.Item("Events").ToString = sEventsPK Then
                RetrieveMatters = row.Item("MatterID")
            End If
        Next row
        'ListBox1.Items.Add("RetrieveMatters " & RetrieveMatters)
        Return RetrieveMatters
    End Function
    Private Function RetrievEventType(ByVal sEventsPK As String) As String
        RetrievEventType = ""
        For Each row As DataRow In dtb.Rows
            'strDetail = row.Item("Detail")
            If row.Item("Events").ToString = sEventsPK Then
                RetrievEventType = row.Item("EventTypes")
            End If
        Next row
        'ListBox1.Items.Add("RetrievEventType " & RetrievEventType)
        Return RetrievEventType
    End Function
    Private Function RetrieveSTMNNo(ByVal sEventsPK As String) As String
        RetrieveSTMNNo = ""
        For Each row As DataRow In dtb.Rows
            'strDetail = row.Item("Detail")
            If row.Item("Events").ToString = sEventsPK Then
                RetrieveSTMNNo = row.Item("StmnNo")
            End If
        Next row
        'ListBox1.Items.Add("RetrieveSTMNNo " & RetrieveSTMNNo)
        Return RetrieveSTMNNo
    End Function
    Private Function RetrieveContacts(ByVal sEventsPK As String) As String
        RetrieveContacts = ""
        For Each row As DataRow In dtb.Rows
            'strDetail = row.Item("Detail")
            If row.Item("Events").ToString = sEventsPK Then
                RetrieveContacts = row.Item("Contacts")
            End If
        Next row
        'ListBox1.Items.Add("RetrieveContacts " & RetrieveContacts)
        Return RetrieveContacts
    End Function
    Private Function RetrieveClaimsExaminerEmail(ByVal sEventsPK As String) As String
        RetrieveClaimsExaminerEmail = ""
        For Each row As DataRow In dtb.Rows
            'strDetail = row.Item("Detail")
            If row.Item("Events").ToString = sEventsPK Then
                RetrieveClaimsExaminerEmail = row.Item("ClaimsExaminerEmail")
            End If
        Next row
        'ListBox1.Items.Add("RetrievEmail " & RetrievEmail)
        If RetrieveClaimsExaminerEmail.Length < 1 Then
            RetrieveClaimsExaminerEmail = ""
        End If
        Return RetrieveClaimsExaminerEmail
    End Function


    Private Function RetrievEmail(ByVal sEventsPK As String) As String
        RetrievEmail = ""
        For Each row As DataRow In dtb.Rows
            'strDetail = row.Item("Detail")
            If row.Item("Events").ToString = sEventsPK Then
                RetrievEmail = row.Item("emailaddress")
            End If
        Next row
        'ListBox1.Items.Add("RetrievEmail " & RetrievEmail)
        Return RetrievEmail
    End Function
    Private Function RetrieveTemplate(ByVal sEventsPK As String) As String
        RetrieveTemplate = ""
        For Each row As DataRow In dtb.Rows
            'strDetail = row.Item("Detail")
            If row.Item("Events").ToString = sEventsPK Then
                RetrieveTemplate = row.Item("Template")
            End If
        Next row
        ' ListBox1.Items.Add("RetrieveTemplate " & RetrieveTemplate)
        Return RetrieveTemplate
    End Function
    Private Function RetrieveSubjectLine(ByVal sEventsPK As String) As String
        RetrieveSubjectLine = ""
        For Each row As DataRow In dtb.Rows
            'strDetail = row.Item("Detail")
            If row.Item("Events").ToString = sEventsPK Then
                RetrieveSubjectLine = Replace(row.Item("SubjectLine"), Chr(34), "")
            End If
        Next row
        ' ListBox1.Items.Add("RetrieveSubjectLine " & RetrieveSubjectLine)
        Return RetrieveSubjectLine
    End Function

    Private Function RetrievDesc(ByVal sEventsPK As String) As String
        RetrievDesc = ""
        For Each row As DataRow In dtb.Rows
            'strDetail = row.Item("Detail")
            If row.Item("Events").ToString = sEventsPK Then
                RetrievDesc = row.Item("Description")
            End If
        Next row
        'ListBox1.Items.Add("RetrievDesc " & RetrievDesc)
        Return RetrievDesc
    End Function
    Private Function RetrieveDocDir(ByVal sEventsPK As String) As String
        RetrieveDocDir = ""
        For Each row As DataRow In dtb.Rows
            'strDetail = row.Item("Detail")
            If row.Item("Events").ToString = sEventsPK Then
                RetrieveDocDir = row.Item("DocDir")
            End If
        Next row
        'ListBox1.Items.Add("Retrieved DocDir value from site " & sServerName)
        'ListBox1.Items.Add("is " & RetrieveDocDir)
        Return RetrieveDocDir
    End Function
    Private Function ConverttoPDF(ByVal sDocDir As String, ByVal sMatters As String, ByVal sEventTypes As String) As String
        ' REturns the URL for the new PDF document.  


        Dim Printer As Printer = New Printer()

        Printer.LicenseKey = "472D-F1FD-5F87-C68B-4770-D638"

        ConverttoPDF = ""
        Try

        Catch ex As Exception
            ListBox1.Items.Add("Error initializing easyPDF.Printer.8 " & ex.Message)
        End Try
        Try
            If File.Exists(sDocDir) Then
                ' Sfile is the file with the pdf extension
                ConverttoPDF = GetDocDir(sMatters, sEventTypes) & ".PDF"
                If Not File.Exists(ConverttoPDF) Then
                    Try
                        Dim printjob As PrintJob = Printer.PrintJob
                        printjob.PrintOut(sDocDir, ConverttoPDF)

                    Catch ex As Exception
                        'nO COMMENT ON PURPOSE - TRY SECONDARY SOLUTION FIRST - FILE PATH TOO LONG ON ORIGINAL.     
                        bError = True
                    End Try
                    If bError Then
                        Try
                            ConverttoPDF = GetsBasePath() & GetIncrements() & ".PDF"
                            Dim printjob As PrintJob = Printer.PrintJob
                            'ListBox1.Items.Add("Trying path:" & ConverttoPDF)
                            'oPrintJob.PrintOut("C:\data\testing.docx", "C:\data\testing.pdf")
                            printjob.PrintOut(sDocDir, ConverttoPDF)
                            If File.Exists(ConverttoPDF) Then
                                bError = False
                            Else
                                ListBox1.Items.Add("Error creating PDF File " & ConverttoPDF)
                                bError = True
                                ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                            End If
                        Catch ex As Exception
                            bError = True
                        End Try
                    End If
                End If
            Else
                ListBox1.Items.Add("Source File Missing " & sDocDir.ToString)
                ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                bError = True
            End If

        Catch ex As Exception
            ListBox1.Items.Add("ERROR IN FILE EXTENSION (TIFF): " & ex.Message)
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        End Try

        Printer = Nothing

        Return ConverttoPDF
    End Function
    Declare Function SetProcessDPIAware Lib "user32.dll" () As Boolean

    Public Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Button11.Enabled = True
        Dim FS As FileStream
        FS = New FileStream("CHUBB_EMail_Log.txt", FileMode.Append, FileAccess.Write)
        Dim SW As StreamWriter
        SW = New StreamWriter(FS, Encoding.Default)
        SW.Write("")

        For Each o As Object In ListBox1.Items
            Try
                SW.WriteLine(o)

            Catch ex As System.Exception
                'No feedback required
                Button10.Enabled = False
            End Try
        Next
        Try
            SW.Close()
        Catch ex As System.Exception
            'No feedback required
        End Try
    End Sub
    Private Sub Button62_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Try
            If System.IO.File.Exists("CHUBB_EMail_Log.txt") = True Then
                Process.Start("CHUBB_EMail_Log.txt")
            Else
                Button11.Enabled = False
            End If

        Catch ex As System.Exception
            'No feedback required
        End Try
    End Sub

End Class

Public NotInheritable Class Mail
    Public Sub send(ByRef sMessage As String)
        Try
            Dim smtpServer As New SmtpClient()
            Dim mail As New MailMessage()
            smtpServer.UseDefaultCredentials = False
            smtpServer.Credentials = New Net.NetworkCredential(sSMTPADDRESS, sSMTPPWD)
            smtpServer.Port = sSMTPPort
            smtpServer.EnableSsl = sSMTPSecurity
            smtpServer.Host = sSMTPServer

            mail = New MailMessage()
            mail.From = New MailAddress("Jeffrey.Sweeney@jsqlllc.com")
            mail.To.Add("Jeffrey.Sweeney@jsqlllc.com")
            mail.Subject = "License Expiration Utilization - Document Import Utility"
            mail.Body = "The firm " & ConfigurationManager.AppSettings("Firm") & sMessage
            smtpServer.Send(mail)
        Catch ex As System.Exception

        End Try

    End Sub
    Public Sub SendFailureAttachments(ByRef sMessage As String, ByRef sRecipients As String, ByRef sMatterID As String, ByRef sShortDesc As String, ByVal sDocDir As String, ByVal sLocalSubjectline As String, ByRef sStmnNo As String, ByVal sImageDirectory As String, ByVal scontacts As String)
        If Not bError Then
            Try
                Dim sSQL As String
                Dim bSkipAttachments As Boolean = True
                Dim oCom As SqlCommand
                Dim oConn As New SqlConnection
                Dim oDR As SqlDataReader = Nothing
                Dim smtpServer As New SmtpClient()
                Dim mail As New MailMessage()
                Dim data As System.Net.Mail.Attachment = Nothing
                Dim insidedata As System.Net.Mail.Attachment = Nothing
                smtpServer.UseDefaultCredentials = False
                smtpServer.Credentials = New Net.NetworkCredential(sSMTPADDRESS, sSMTPPWD)
                smtpServer.Port = sSMTPPort
                smtpServer.EnableSsl = sSMTPSecurity
                smtpServer.Host = sSMTPServer
                mail = New MailMessage()
                mail.From = New MailAddress(sSMTPADDRESS, "Chubb HC Finance")

                mail.To.Add(sFailureTOADDRESS)
                mail.CC.Add(sFailureCCADDRESS)
                mail.Bcc.Add(sFailureBCCADDRESS)
                mail.Subject = Replace(Replace(Replace(Replace(sLocalSubjectline, ":", ""), "@", ""), Chr(34), ""), Chr(39), "")


                mail.Body = Replace(Replace(Replace(Replace(sMessage, ":", ""), "@", ""), Chr(34), ""), Chr(39), "")
                data = New System.Net.Mail.Attachment(sDocDir)
                mail.Attachments.Add(data)
                '  Attachments
                '  Good for initial letter.

                sSQL = " Select StmnImageFileName from stmnledger s, stmnledgerfiles SF " &
                         " where SF.StmnImage = S.StmnImage and S.LedgerType= 'S' " &
                         " and S.StmnNo =  '" & sStmnNo & "'" &
                         "  union " &
                        " Select distinct StmnImageFileName from stmnledger S, stmnledgerfiles SF, Matters M, jSQLARDocuments J" &
                        " where SF.StmnImage = S.StmnImage and S.LedgerType= 'S' and S.Contacts = '" & scontacts & "' " &
                        " and s.Matters = m.Matters and M.MatterID = j.MatterId and J.StmnNo = '" & sStmnNo & "'" &
                        " and S.StmnNo <>  '" & sStmnNo & "' and Comment in ('60 day letter','90 day letter','120 day letter', 'Initial letter','Agency Cap letter') " &
                        " and StmnString is not null and StmnString <>''" &
                        " and S.StmnDate >= '2016-11-30' and S.total <> S.totalpaid"

                If bSkipAttachments Then
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
                                If File.Exists(sImageDirectory & oDR(0).ToString) Then
                                    Try
                                        insidedata = New System.Net.Mail.Attachment(sImageDirectory & oDR(0).ToString)
                                        mail.Attachments.Add(insidedata)
                                    Catch ex As Exception
                                        Form1.ListBox1.Items.Add("Error sending Invoice Attachments " & ex.Message)
                                    End Try
                                Else
                                    Form1.ListBox1.Items.Add("File doesn't exist " & sImageDirectory & oDR(0).ToString)
                                End If
                            End While
                        Else
                            Form1.ListBox1.Items.Add("No records to send ")
                        End If
                    Catch ex As Exception
                        Form1.ListBox1.Items.Add("SendAttachments error " & ex.Message)
                    Finally
                        oDR.Close()
                        oDR = Nothing
                        oCom = Nothing
                        oConn.Close()
                        oConn = Nothing
                    End Try
                End If

                smtpServer.Send(mail)
                Form1.ListBox1.Items.Add("Email Successfully sent for MatterID. " & sMatterID)
            Catch ex As System.Exception
                Form1.ListBox1.Items.Add("Exception in SendAttachments " & ex.Message)
                bError = True
            End Try
            Form1.ListBox1.SelectedIndex = Form1.ListBox1.Items.Count - 1
        End If
    End Sub
    Public Sub SendAttachments(ByRef sMessage As String, ByRef sRecipients As String, ByRef sMatterID As String, ByRef sShortDesc As String, ByVal sDocDir As String, ByVal sLocalSubjectline As String, ByRef sStmnNo As String, ByVal sImageDirectory As String, ByVal scontacts As String, ByVal sClaimsExaminerEmail As String)
        If Not bError Then
            Try

                If StrReverse(sImageDirectory).Substring(0, 2) = "\\" Then
                    sImageDirectory = sImageDirectory.Substring(0, sImageDirectory.Length - 1)

                End If
                Dim sSQL As String
                Dim bSkipAttachments As Boolean = True
                Dim oCom As SqlCommand
                Dim oConn As New SqlConnection
                Dim oDR As SqlDataReader = Nothing
                'NEW VARIABLES
                Dim sAttachment As New MimePart
                Dim smtpServer As New MailKit.Net.Smtp.SmtpClient()
                Dim mail As New MimeMessage()
                Dim builder = New BodyBuilder
                Dim Multipart = New Multipart("mixed")
                Dim sServer As String = ConfigurationManager.AppSettings("SMTPServer")
                Dim sSMTPPort As String = ConfigurationManager.AppSettings("SMTPPort")
                Dim sUserID As String = ConfigurationManager.AppSettings("SMTPADDRESS")
                Dim sPassword As String = ConfigurationManager.AppSettings("SMTPPWD")
                Dim sSMTPDelay As String = ConfigurationManager.AppSettings("SMTPDelay")
                Dim UserSSL As Boolean = ConfigurationManager.AppSettings("SMTPSecurity").ToString



                'Old Variables
                'Dim smtpServer As New SmtpClient()
                'Dim mail As New MailMessage()
                Dim data As System.Net.Mail.Attachment = Nothing
                Dim insidedata As String
                'smtpServer.UseDefaultCredentials = False
                'smtpServer.Credentials = New Net.NetworkCredential(sSMTPADDRESS, sSMTPPWD)
                'smtpServer.Port = sSMTPPort
                'smtpServer.EnableSsl = sSMTPSecurity
                'smtpServer.Host = sSMTPServer
                ' mail = New MailMessage()
                'mail.From = New MailAddress(sSMTPADDRESS, "Chubb HC Finance")

                If Not bDebug Then
                    mail.From.Add(MailboxAddress.Parse(sUserID))
                    mail.To.Add(MailboxAddress.Parse(sRecipients))
                    'mail.To.Add(MailboxAddress.Parse("emoser@chubb.com"))
                    mail.Bcc.Add(MailboxAddress.Parse("Jeffrey.Sweeney@jsqlllc.com"))
                    mail.Bcc.Add(MailboxAddress.Parse("AccountsReceivableClaimsHC@CHUBB.com"))
                    mail.Subject = Replace(Replace(Replace(sLocalSubjectline, ":", ""), Chr(34), ""), Chr(39), "")

                    Try
                        ' if selected and email address exists for the examiner then send them a copy - BCC
                        If Form1.CheckBox3.Checked And sClaimsExaminerEmail.Length > 0 Then
                            mail.Bcc.Add(MailboxAddress.Parse(sClaimsExaminerEmail))
                        End If
                    Catch ex As Exception
                    End Try
                Else
                    mail.To.Add(MailboxAddress.Parse("Jeffrey.Sweeney@jsqlllc.com"))
                    mail.Subject = "TEST " & Replace(Replace(Replace(sLocalSubjectline, ":", ""), Chr(34), ""), Chr(39), "")
                    Try
                        If Form1.CheckBox3.Checked And sClaimsExaminerEmail.Length > 0 Then
                            mail.Bcc.Add(MailboxAddress.Parse("Jeffrey.Sweeney@jsqlllc.com"))
                        End If
                    Catch ex As Exception
                    End Try
                End If
                builder.TextBody = Replace(Replace(sMessage, Chr(34), ""), Chr(39), "")

                'Main Letter in PDF Format

                'sAttachment.FileName = sDocDir
                'sAttachment.Content = New MimeContent(File.OpenRead(sDocDir), ContentEncoding.Default)
                'sAttachment.ContentDisposition = New ContentDisposition(ContentDisposition.Attachment)
                'sAttachment.FileName = Path.GetFileName(sDocDir)
                'sAttachment.ContentTransferEncoding = ContentEncoding.Base64

                builder.Attachments.Add(sDocDir)



                'Multipart.Add(builder.ToMessageBody())
                'Multipart.Add(sAttachment)


                'Updated statement to capture additional attachments.  
                sSQL = " Select distinct StmnImageFileName from stmnledger s, stmnledgerfiles SF " &
                        " where SF.StmnImage = S.StmnImage and S.LedgerType= 'S' " &
                        " and S.StmnNo =  '" & sStmnNo & "'" &
                        " union " &
                        " Select distinct StmnImageFileName from stmnledger S, stmnledgerfiles SF, Matters M, jSQLARDocuments J" &
                        " where SF.StmnImage = S.StmnImage and S.LedgerType= 'S' " &
                        " and s.Matters = m.Matters and M.MatterID = j.MatterId and J.StmnNo = '" & sStmnNo & "'" &
                        " and S.StmnNo <>  '" & sStmnNo & "' and Comment in ('120 day letter', 'Initial letter','Agency cap letter','Get ALL AR Data','Dunning letter') " &
                        " and StmnString is not null and StmnString <>''" &
                        " and S.StmnDate >= '2016-11-30' and S.total <> S.totalpaid" &
                 " and s.StmnDate < getdate() - 120 "    'Needs to be removed
                'and S.Contacts = '" & scontacts & "'  removed.
                '" and S.Contacts = '" & scontacts & "'  " &
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
                            If File.Exists(sImageDirectory & oDR(0).ToString) Then
                                Try
                                    insidedata = sImageDirectory & oDR(0).ToString
                                    '  Add Attachments
                                    'Main Letter in PDF Format
                                    'sAttachment.FileName = insidedata
                                    'sAttachment.Content = New MimeContent(File.OpenRead(insidedata), ContentEncoding.Default)
                                    'sAttachment.ContentDisposition = New ContentDisposition(ContentDisposition.Attachment)
                                    'sAttachment.FileName = Path.GetFileName(insidedata)
                                    'sAttachment.ContentTransferEncoding = ContentEncoding.Base64



                                    'Multipart.Add(builder.ToMessageBody())
                                    'Multipart.Add(sAttachment)

                                    builder.Attachments.Add(insidedata)

                                Catch ex As Exception
                                    Form1.ListBox1.Items.Add("Error sending Invoice Attachments " & ex.Message)
                                End Try
                            Else
                                Form1.ListBox1.Items.Add("File doesn't exist " & sImageDirectory & oDR(0).ToString)
                            End If
                        End While
                    Else
                        Form1.ListBox1.Items.Add("No records to send ")
                    End If
                Catch ex As Exception
                    Form1.ListBox1.Items.Add("SendAttachments error " & ex.Message)
                Finally
                    oDR.Close()
                    oDR = Nothing
                    oCom = Nothing
                    oConn.Close()
                    oConn = Nothing
                End Try



                'mail.Body = Multipart

                mail.Body = builder.ToMessageBody()

                Try
                    Dim smtp As New MailKit.Net.Smtp.SmtpClient()
                    smtp.Connect(sServer, 587, MailKit.Security.SecureSocketOptions.StartTls)
                    Try
                        smtp.Authenticate(sUserID, sPassword)
                    Catch ex As Exception

                    End Try

                    Try
                        smtp.Send(mail)
                        Console.WriteLine("Sending New Arise Report")
                        smtp.Disconnect(True)
                    Catch ex As Exception

                    End Try



                Catch ex As System.Exception
                    bError = True
                Finally
                    mail = Nothing

                End Try

                Form1.ListBox1.Items.Add("Email Successfully sent for MatterID. " & sMatterID & " to " & sRecipients)
                Thread.Sleep(1000)
            Catch ex As System.Exception
                Form1.ListBox1.Items.Add("Exception in SendAttachments " & ex.Message)
                bError = True
            End Try
            Form1.ListBox1.SelectedIndex = Form1.ListBox1.Items.Count - 1
        End If
    End Sub
    Public Sub SendAttachments_Orig(ByRef sMessage As String, ByRef sRecipients As String, ByRef sMatterID As String, ByRef sShortDesc As String, ByVal sDocDir As String, ByVal sLocalSubjectline As String, ByRef sStmnNo As String, ByVal sImageDirectory As String, ByVal scontacts As String, ByVal sClaimsExaminerEmail As String)
        If Not bError Then
            Try

                If StrReverse(sImageDirectory).Substring(0, 2) = "\\" Then
                    sImageDirectory = sImageDirectory.Substring(0, sImageDirectory.Length - 1)

                End If
                Dim sSQL As String
                Dim bSkipAttachments As Boolean = True
                Dim oCom As SqlCommand
                Dim oConn As New SqlConnection
                Dim oDR As SqlDataReader = Nothing
                Dim smtpServer As New SmtpClient()
                Dim mail As New MailMessage()
                Dim data As System.Net.Mail.Attachment = Nothing
                Dim insidedata As System.Net.Mail.Attachment = Nothing
                smtpServer.UseDefaultCredentials = False
                smtpServer.Credentials = New Net.NetworkCredential(sSMTPADDRESS, sSMTPPWD)
                smtpServer.Port = sSMTPPort
                smtpServer.EnableSsl = sSMTPSecurity
                smtpServer.Host = sSMTPServer
                mail = New MailMessage()
                mail.From = New MailAddress(sSMTPADDRESS, "Chubb HC Finance")

                If Not bDebug Then
                    mail.To.Add(sRecipients)
                    'mail.CC.Add("HouseCounselFinanceDept@chubb.com")
                    mail.Bcc.Add("Jeffrey.sweeney@comcast.net")
                    mail.Bcc.Add("AccountsReceivableClaimsHC@CHUBB.com")
                    mail.Subject = Replace(Replace(Replace(sLocalSubjectline, ":", ""), Chr(34), ""), Chr(39), "")
                    Try
                        ' if selected and email address exists for the examiner then send them a copy - BCC
                        If Form1.CheckBox3.Checked And sClaimsExaminerEmail.Length > 0 Then
                            mail.Bcc.Add(sClaimsExaminerEmail)
                        End If
                    Catch ex As Exception
                    End Try
                Else
                    mail.To.Add("Jeffrey.sweeney@comcast.net")
                    ' mail.To.Add("HouseCounselFinanceDept@chubb.com")
                    'mail.CC.Add("mconeill@chubb.com")
                    mail.Subject = "TEST " & Replace(Replace(Replace(sLocalSubjectline, ":", ""), Chr(34), ""), Chr(39), "")
                    Try
                        If Form1.CheckBox3.Checked And sClaimsExaminerEmail.Length > 0 Then
                            mail.Bcc.Add("Jeffrey.sweeney@jsqlllc.com")
                        End If
                    Catch ex As Exception
                    End Try
                End If

                mail.Body = Replace(Replace(sMessage, Chr(34), ""), Chr(39), "")
                data = New System.Net.Mail.Attachment(sDocDir)
                mail.Attachments.Add(data)
                '  Attachments
                '  Good for initial letter can have more than one attachment and 120 day letter.

                'sSQL = "Select StmnImageFileName from stmnledger s, stmnledgerfiles SF" & _
                '        " where SF.StmnImage = S.StmnImage and S.LedgerType= 'S' " & _
                '        " and S.StmnNo = '" & sStmnNo & "'"
                'Updated statement to capture additional attachments.  
                sSQL = " Select StmnImageFileName from stmnledger s, stmnledgerfiles SF " &
                        " where SF.StmnImage = S.StmnImage and S.LedgerType= 'S' " &
                        " and S.StmnNo =  '" & sStmnNo & "'" &
                        " union " &
                        " Select distinct StmnImageFileName from stmnledger S, stmnledgerfiles SF, Matters M, jSQLARDocuments J" &
                        " where SF.StmnImage = S.StmnImage and S.LedgerType= 'S' " &
                        " and s.Matters = m.Matters and M.MatterID = j.MatterId and J.StmnNo = '" & sStmnNo & "'" &
                        " and S.StmnNo <>  '" & sStmnNo & "' and Comment in ('120 day letter', 'Initial letter','Agency cap letter','Get ALL AR Data','Dunning letter') " &
                        " and StmnString is not null and StmnString <>''" &
                        " and S.StmnDate >= '2016-11-30' and S.total <> S.totalpaid" &
                 " and s.StmnDate < getdate() - 120 "    'Needs to be removed
                'and S.Contacts = '" & scontacts & "'  removed.
                '" and S.Contacts = '" & scontacts & "'  " &
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
                            If File.Exists(sImageDirectory & oDR(0).ToString) Then
                                Try
                                    insidedata = New System.Net.Mail.Attachment(sImageDirectory & oDR(0).ToString)
                                    mail.Attachments.Add(insidedata)
                                Catch ex As Exception
                                    Form1.ListBox1.Items.Add("Error sending Invoice Attachments " & ex.Message)
                                End Try
                            Else
                                Form1.ListBox1.Items.Add("File doesn't exist " & sImageDirectory & oDR(0).ToString)
                            End If
                        End While
                    Else
                        Form1.ListBox1.Items.Add("No records to send ")
                    End If
                Catch ex As Exception
                    Form1.ListBox1.Items.Add("SendAttachments error " & ex.Message)
                Finally
                    oDR.Close()
                    oDR = Nothing
                    oCom = Nothing
                    oConn.Close()
                    oConn = Nothing
                End Try


                smtpServer.Send(mail)

                Form1.ListBox1.Items.Add("Email Successfully sent for MatterID. " & sMatterID & " to " & sRecipients)
                Thread.Sleep(10000)
            Catch ex As System.Exception
                Form1.ListBox1.Items.Add("Exception in SendAttachments " & ex.Message)
                bError = True
            End Try
            Form1.ListBox1.SelectedIndex = Form1.ListBox1.Items.Count - 1
        End If
    End Sub


End Class
