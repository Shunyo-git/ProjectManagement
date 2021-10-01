
'Imports System.Data.SqlClient



Public Class chkLogin
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

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




        If Session("CurrentUser") Is Nothing Or Strings.Trim(Session("CurrentUser")) = "" Then
            getCurrentUserInfo()

        End If
        If Request("fromUrl") Is Nothing Then
            Response.Redirect("index.aspx")
        Else
            Response.Redirect(Request("fromUrl"))
        End If

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


        'If Session("IsSuper") Then
        '    Session("isForderRead") = True
        '    Session("isForderEdit") = True
        '    Session("isForderDel") = True
        '    Session("isForderNew") = True
        '    Session("isForderOwner") = True
        'Else
        '    Try
        '        Dim objFolder As clsFolder
        '        objFolder = New clsFolder(m_folderID, m_CurrentUserID)
        '        objFolder.Load()
        '        Session("isForderRead") = objFolder.Security(KM_Read) 'km_Edit,km_New,KM_Del,km_read
        '        Session("isForderEdit") = objFolder.Security(KM_Edit)
        '        Session("isForderDel") = objFolder.Security(KM_Del)
        '        Session("isForderNew") = objFolder.Security(KM_New)
        '        Session("isForderOwner") = objFolder.Security(KM_Owner)

        '        If Not Session("isForderRead") Then
        '            Page.RegisterStartupScript("alert", GlobalFunctions.ShowAlertMsg("您沒有進入資料夾的權限!"))
        '            Return False
        '        End If
        '    Catch ex As Exception
        '        println("error 02 :" & ex.Message)
        '        println("m_CurrentUserID :" & m_CurrentUserID)
        '        Response.End()
        '    End Try


        'End If

        Return True

    End Function

    Public Sub println(ByVal str As String)
        Response.Write(str & "<BR>")
    End Sub

End Class
