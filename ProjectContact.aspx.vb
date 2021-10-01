Imports System
Imports System.Data
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class ProjectContact
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents Users As System.Web.UI.WebControls.DataGrid

        '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
        '請勿刪除或移動它。
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
            '請勿使用程式碼編輯器進行修改。
            InitializeComponent()
        End Sub

#End Region
        Private _projectID As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            ' Obtain project id from the QueryString.
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            ' Ensure that the visiting user has access to view the current page
            If Not Roles.isEnabledViewProject(_projectID) Then
                Session("m_ErrorMessage") = "系統拒絕存取，您沒有權限存取此專案資訊。" '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If

            If Not Page.IsPostBack Then
                If _projectID <> 0 Then
                    LoadMembers()
                End If

            End If
        End Sub

        '*******************************************************
        '
        ' LoadUsers method populate the DataGrid with a list of Time Tracker users.
        '
        '*******************************************************

        Private Sub LoadMembers()
            Dim usrs As UsersCollection = PMUser.GetMembers(_projectID)

            ' Check for sorting when binding grid
            ' SortGridData(usrs, SortField, SortAscending)
            Users.DataSource = usrs
            Users.DataKeyField = "UserID"
            Users.DataBind()
        End Sub 'LoadMembers
    End Class

End Namespace

