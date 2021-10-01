Imports System
Imports System.Data
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web
    Public Class GroupDetail
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents lblDescription As System.Web.UI.WebControls.Label
        Protected WithEvents Label4 As System.Web.UI.WebControls.Label
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents lblName As System.Web.UI.WebControls.Label
        Protected WithEvents userNamelbl As System.Web.UI.WebControls.Label
        Protected WithEvents lbtnDelete As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnEdit As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents lblMembers As System.Web.UI.WebControls.Label

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
        Private _groupID As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            ' Obtain project id from the QueryString.
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            If Request.QueryString("GroupID") Is Nothing Then
                _groupID = 0
            Else
                _groupID = Convert.ToInt32(Request.QueryString("GroupID"))
            End If

            If Not Roles.isEnabledViewProject(_projectID) Then
                Session("m_ErrorMessage") = "系統拒絕存取，您沒有權限存取此專案資訊。" '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If


            If Roles.isEnabledModifyProject(_projectID) Then
                PanelModify.Visible = True

            Else
                PanelModify.Visible = False
            End If

            If Not IsPostBack Then

                If PMGroup.GetProjectGroups(_projectID).Count <= 1 Then
                    lbtnDelete.Attributes.Add("OnClick", "alert('此組別為專案唯一的組別，不可刪除。');return false;")
                Else
                    lbtnDelete.Attributes.Add("OnClick", "return chkConfirmDelete();")
                End If
                ' Load project with _projID when project id exists in the QueryString
                If _groupID <> 0 And _projectID <> 0 Then
                    BindInfo()
                Else
                    Session("m_ErrorMessage") = "系統錯誤，無法存取群組資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If
            End If
        End Sub
        '*********************************************************************
        '
        ' The BindInfo method retrieves the list of roles 
        ' and Load User information if user exists
        '
        '*********************************************************************

        Private Sub BindInfo()
            Dim group As New PMGroup(_groupID)
            group.Load()

            _projectID = group.ProjectID
            lblName.Text = group.Name
            lblDescription.Text = group.Description

            Dim user As PMUser
            For Each user In group.Members
                If Not Strings.Trim(lblMembers.Text) Is String.Empty Then
                    lblMembers.Text += "、"
                End If
                lblMembers.Text += user.DisplayName
            Next user

        End Sub 'BindInfo

        Private Sub lbtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnDelete.Click
            PMGroup.RemoveGroup(_groupID)
            Response.Redirect([String].Format("ProjectGroup.aspx?projectID={0}&index=-1&projectIndex={1}", _projectID, TabItem.ProjectTabIndex.ProjectGroup), True)
        End Sub

        Private Sub lbtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnEdit.Click
            Response.Redirect([String].Format("GroupModify.aspx?projectID={0}&index=-1&projectIndex={1}&GroupID={2}", _projectID, TabItem.ProjectTabIndex.ProjectGroup, _groupID), False)
        End Sub
    End Class
End Namespace

