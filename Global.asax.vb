Imports System
Imports System.Web
Imports System.Web.Security
Imports System.Web.SessionState
Imports System.Threading
Imports System.Globalization
Imports System.Configuration
Imports HsinYi.EIP.ProjectManagement.BusinessLogicLayer

Namespace HsinYi.EIP.ProjectManagement.Web

    Public Class Global
        Inherits System.Web.HttpApplication

        Public Const CfgKeyPMConnString As String = "PMConnectionString"
        Public Const CfgKeyEIPConnString As String = "EIPConnectionString"
        Public Const CfgKeySMTPServer As String = "SMTPServer"

        Public Const CfgKeyAttachmentWebPath As String = "AttachmentWebPath"
        Public Const CfgKeyAttachmentSrvPath As String = "AttachmentSrvPath"


        Public Const CfgKeyUserAcctSource As String = "UserAccountSource"
        Public Const CfgKeyDefaultRole As String = "DefaultRoleForNewUser"
        Public Const CfgKeyFirstDayOfWeek As String = "FirstDayOfWeek"

        Public UserRoles As String = "userroles"
        Public UserRolesMobile As String = "userrolesmobile"

#Region " ����]�p�u�㲣�ͪ��{���X "

        Public Sub New()
            MyBase.New()

            '��������]�p�u��һݪ��I�s�C
            InitializeComponent()

            '�b InitializeComponent() �I�s����[�J�Ҧ�����l�]�w

        End Sub

        '������]�p�u�㪺���n��
        Private components As System.ComponentModel.IContainer

        '�`�N: �H�U������]�p�u��һݪ��{��
        '�z�i�H�ϥΤ���]�p�u��i��ק�A
        '�ФŨϥε{���X�s�边�i��ק�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            components = New System.ComponentModel.Container
        End Sub

#End Region

         

        Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
            ' �Ұ����ε{���ɤ޵o
        End Sub

        Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
            ' �Ұʤu�@���q�ɤ޵o
        End Sub

        Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
            ' ��C�@�ӭn�D�}�l�ɤ޵o
        End Sub

        Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
            ' �������Ҩϥήɤ޵o
          
            Dim userInformation As String = [String].Empty

            If Request.IsAuthenticated = True Then
                ' Create the roles cookie if it doesn't exist yet for this session.
                If (Request.Cookies(UserRoles) Is Nothing Or Request.Cookies(UserRoles) Is "") Then
                    ' Retrieve the user's role and ID information and add it to
                    ' the cookie
                    Dim strUserLoginID As String = User.Identity.Name
                    Dim IdentityName As String
                    IdentityName = Mid(strUserLoginID, InStrRev(strUserLoginID, "\") + 1)
                    'strUserLoginID.Substring(strUserLoginID.LastIndexOf("\") + 1, strUserLoginID.Length - strUserLoginID.LastIndexOf("\") - 1)


                    Dim usr As New PMUser(IdentityName)
                    If Not usr.Load() Then
                        ' The user was not found in the Time Tracker database so add them using
                        ' the default role.  Specifying a UserID of 0 will result in the user being 
                        ' inserted into the database.

                        '�p�G���OProjectMember�ɪ��@�k
                        'Dim newUser As New PMUser(0, Context.User.Identity.Name, [String].Empty, ConfigurationSettings.AppSettings(CfgKeyDefaultRole))
                        'newUser.Save()
                        'usr = newUser
                        Response.Write("�n�J�̨������~�A�L�k�s���M�׸�T�C")

                    End If
                   
                    ' Create a string to persist the role and user id
                    userInformation = usr.UserID & ";" & usr.Role & ";" & usr.Name & ";" & usr.DisplayName

                    ' Create a cookie authentication ticket.
                    Dim ticket As New FormsAuthenticationTicket(1, User.Identity.Name, DateTime.Now, DateTime.Now.AddHours(1), False, userInformation)
                    ' version
                    ' user name
                    ' issue time
                    ' expires every hour
                    ' don't persist cookie

                    ' Encrypt the ticket
                    Dim cookieStr As [String] = FormsAuthentication.Encrypt(ticket)

                    ' Send the cookie to the client
                    Response.Cookies(UserRoles).Value = cookieStr
                    Response.Cookies(UserRoles).Path = "/"
                    Response.Cookies(UserRoles).Expires = DateTime.Now.AddMinutes(1)

                    ' Add our own custom principal to the request containing the user's identity, the user id, and
                    ' the user's role 
                    Context.User = New CustomPrincipal(User.Identity, usr.UserID, usr.Role, usr.Name, usr.DisplayName)
                Else
                    ' Get roles from roles cookie
                    Dim ticket As FormsAuthenticationTicket = FormsAuthentication.Decrypt(Context.Request.Cookies(UserRoles).Value)
                    userInformation = ticket.UserData

                    ' Add our own custom principal to the request containing the user's identity, the user id, and
                    ' the user's role from the auth ticket
                    Dim info As String() = userInformation.Split(New Char() {";"c})
                    Context.User = New CustomPrincipal(User.Identity, Convert.ToInt32(info(0).ToString()), info(1).ToString(), info(2).ToString(), info(3).ToString())
                End If
            End If
          

        End Sub

        '*********************************************************************
        '  
        ' GetApplicationPath Method
        '
        ' This method returns the correct relative path when installing
        ' the portal on a root web site instead of virtual directory
        '
        '*********************************************************************

        Public Shared Function GetApplicationPath(ByVal request As HttpRequest) As String
            Dim path As String = String.Empty
            Try
                If request.ApplicationPath <> "/" Then
                    path = request.ApplicationPath
                End If
            Catch e As Exception
                Throw e
            End Try
            Return path
        End Function

        Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
            ' �o�Ϳ��~�ɤ޵o
        End Sub

        Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
            ' ��u�@���q�����ɤ޵o
        End Sub

        Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
            ' �����ε{�������ɤ޵o
        End Sub

    End Class

End Namespace
