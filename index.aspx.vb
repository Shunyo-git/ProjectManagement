Imports System
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class index
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm


        '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
        '請勿刪除或移動它。
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
            '請勿使用程式碼編輯器進行修改。
            InitializeComponent()
        End Sub

#End Region

        Private m_CurrentUser As String
        Private m_CurrentUserID As Integer

        Private m_CurrentUserName As String
        Private m_CurrentUserIsManager As Boolean
        Private m_CurrentUserIsSuperManager As Boolean


        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


            'If Session("CurrentUser") Is Nothing Or Strings.Trim(Session("CurrentUser")) = "" Then
            '    getCurrentUserInfo()

            'End If
            
        End Sub
       
       
        
        Private Function getCurrentUserInfo() As Boolean
            '  取得使用者身份

            Dim strUserLoginID As String = User.Identity.Name

            Dim identification As String

            identification = strUserLoginID.Substring(strUserLoginID.LastIndexOf("\") + 1, strUserLoginID.Length - strUserLoginID.LastIndexOf("\") - 1)



            Try

                '取得目前帳號
                If strUserLoginID.IndexOf("HSINYI\") < 0 Then
                    Page.RegisterStartupScript("alert", ServerScriptHelper.ShowAlertMsg("帳號錯誤! 請重新登入!  "))
                    Return False
                Else
                    'strUserLoginID.Replace("HSINYI\", "")
                    ' strUserLoginID = Mid(strUserLoginID, strUserLoginID.IndexOf("\") + 2)
                    strUserLoginID = identification
                End If

                Dim objUser As EIPSysSecurity.clsUser = New EIPSysSecurity.clsUser

                '取得目前帳號的EIP姓名

                m_CurrentUserID = (objUser.getUserID(strUserLoginID))
                m_CurrentUser = strUserLoginID

                Dim sys As EIPSysSecurity.clsSystem = New EIPSysSecurity.clsSystem(m_CurrentUser)
                m_CurrentUserName = sys.CurrentUser()

                '取得EIP UserID
                If m_CurrentUserID <= 0 Or IsDBNull(m_CurrentUserID) Then
                    Page.RegisterStartupScript("alert", ServerScriptHelper.ShowAlertMsg("系統內無您的帳號資料(" & strUserLoginID & ")! 請重新登入!  "))
                    Return False
                End If

                Session("CurrentUserID") = m_CurrentUserID
                Session("CurrentUser") = m_CurrentUser
                Session("CurrentUserName") = m_CurrentUserName

                '取得資料夾權限

                Session("IsSuper") = sys.IsSuper
            Catch ex As Exception
                Response.Write("error 01:" & ex.Message)
                Response.End()
            End Try




            Return True

        End Function

        Public Sub println(ByVal str As String)
            Response.Write(str & "<BR>")
        End Sub

    End Class

End Namespace


