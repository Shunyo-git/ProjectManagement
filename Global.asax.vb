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

#Region " 元件設計工具產生的程式碼 "

        Public Sub New()
            MyBase.New()

            '此為元件設計工具所需的呼叫。
            InitializeComponent()

            '在 InitializeComponent() 呼叫之後加入所有的初始設定

        End Sub

        '為元件設計工具的必要項
        Private components As System.ComponentModel.IContainer

        '注意: 以下為元件設計工具所需的程序
        '您可以使用元件設計工具進行修改，
        '請勿使用程式碼編輯器進行修改。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            components = New System.ComponentModel.Container
        End Sub

#End Region

         

        Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
            ' 啟動應用程式時引發
        End Sub

        Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
            ' 啟動工作階段時引發
        End Sub

        Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
            ' 於每一個要求開始時引發
        End Sub

        Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
            ' 嘗試驗證使用時引發
          
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

                        '如果不是ProjectMember時的作法
                        'Dim newUser As New PMUser(0, Context.User.Identity.Name, [String].Empty, ConfigurationSettings.AppSettings(CfgKeyDefaultRole))
                        'newUser.Save()
                        'usr = newUser
                        Response.Write("登入者身分錯誤，無法存取專案資訊。")

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
            ' 發生錯誤時引發
        End Sub

        Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
            ' 於工作階段結束時引發
        End Sub

        Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
            ' 於應用程式結束時引發
        End Sub

    End Class

End Namespace
