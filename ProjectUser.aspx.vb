Imports System
Imports System.Data
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web


    Public Class ProjectUser
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents Users As System.Web.UI.WebControls.DataGrid
        Protected WithEvents lbtnNew As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents lblRecordCount As System.Web.UI.WebControls.Label
        Protected WithEvents lbNextPage As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbPrevPage As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lblTotalPage As System.Web.UI.WebControls.Label
        Protected WithEvents lblNowPage As System.Web.UI.WebControls.Label
        Protected WithEvents lbLastPage As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbFirstPage As System.Web.UI.WebControls.LinkButton

        '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
        '請勿刪除或移動它。
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
            '請勿使用程式碼編輯器進行修改。
            InitializeComponent()
        End Sub

#End Region
        ' Use User component to capture user's input during edit mode
        Protected _userInput As New PMUser(0, String.Empty, String.Empty, String.Empty)
        ' Contains the id of current project
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
            If Roles.isEnabledModifyProject(_projectID) Then
                PanelModify.Visible = True
            Else
                PanelModify.Visible = False
            End If
 
            If Not Page.IsPostBack Then
                If _projectID <> 0 Then
                    LoadMembers()
                End If

            End If
        End Sub 'Page_Load
        '*******************************************************
        '
        ' GetRoles methods returns a list of User Roles from Components.Roles
        '
        '*******************************************************

        Protected Function GetRoles() As RolesCollection
            Return Roles.GetRoles()
        End Function 'GetRoles
        '*******************************************************
        '
        ' LoadUsers method populate the DataGrid with a list of Time Tracker users.
        '
        '*******************************************************

        Private Sub LoadMembers()
            Dim usrs As UsersCollection = PMUser.GetMembers(_projectID)

            ' Check for sorting when binding grid
            SortGridData(usrs, SortField, SortAscending)
            lblRecordCount.Text = usrs.Count
            Users.DataSource = usrs
            Users.DataKeyField = "UserID"
            Users.DataBind()
            showPage()

        End Sub 'LoadMembers
        '*******************************************************
        '
        ' SortGridData methods sorts the Users Grid based on which
        ' sort field is being selected.  Also does reverse sorting based on the boolean.
        '
        '*******************************************************

        Private Sub SortGridData(ByVal list As UsersCollection, ByVal sortField As String, ByVal asc As Boolean)
            Dim sortCol As UsersCollection.UserFields = UsersCollection.UserFields.InitValue

            Select Case sortField
                Case "Name"
                    sortCol = UsersCollection.UserFields.Name
                Case "RoleName"
                    sortCol = UsersCollection.UserFields.RoleName
                Case "Department"
                    sortCol = UsersCollection.UserFields.SourceUserDepartment
                Case "Title"
                    sortCol = UsersCollection.UserFields.SourceUSerTitle
                Case "Description"
                    sortCol = UsersCollection.UserFields.Description

            End Select

            list.Sort(sortCol, asc)
        End Sub 'SortGridData

        '*******************************************************
        '
        ' Users_Sort server event handler on this page is used to
        ' sorts the DataGrid according to the sort field
        ' being selected.
        '
        '*******************************************************

        Private Sub Users_Sort(ByVal [source] As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles Users.SortCommand
            SortField = e.SortExpression
            LoadMembers()
        End Sub 'Users_Sort

        '*******************************************************
        '
        ' Users_PageIndexChanged server event handler on this page is used
        ' for changing page index
        '
        '*******************************************************

        Private Sub Users_PageIndexChanged(ByVal [source] As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles Users.PageIndexChanged
            Users.EditItemIndex = -1
            Users.CurrentPageIndex = e.NewPageIndex
            LoadMembers()
        End Sub 'Users_PageIndexChanged

        Private Sub NewUserButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim mainIndex As Integer = IIf(Request("index") Is Nothing, 0, Convert.ToInt32(Request("index")))
            Dim adminIndex As Integer = IIf(Request("adminindex") Is Nothing, 0, Convert.ToInt32(Request("adminindex")))

            Response.Redirect([String].Format("UserDetail.aspx?index={0}&adminindex={1}", mainIndex, adminIndex), False)
        End Sub 'NewUserButton_Click

        Property SortField() As String
            Get
                Dim o As Object = ViewState("SortField")
                If o Is Nothing Then
                    Return String.Empty
                End If
                Return CStr(o)
            End Get

            Set(ByVal Value As String)
                ' Same as current sort file, toggle sort direction
                If Value = SortField Then
                    SortAscending = Not SortAscending
                End If
                ViewState("SortField") = Value
            End Set
        End Property

        '*********************************************************************
        '
        ' SortAscending property is tracked in ViewState
        '
        '*********************************************************************

        Property SortAscending() As Boolean
            Get
                Dim o As Object = ViewState("SortAscending")
                If o Is Nothing Then
                    Return True
                End If
                Return CBool(o)
            End Get

            Set(ByVal Value As Boolean)
                ViewState("SortAscending") = Value
            End Set
        End Property

        Private Sub lbtnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnNew.Click
            Response.Redirect([String].Format("UserModify.aspx?projectID={0}&index=-1&projectIndex=2&memberID=0", _projectID), False)
        End Sub

        Sub PagerButtonClick(ByVal sender As Object, ByVal e As EventArgs)
            'used by external paging UI
            Dim arg As String = sender.CommandArgument

            Select Case arg
                Case "最前頁"
                    Users.CurrentPageIndex = 0

                Case "上一頁"
                    If (Users.CurrentPageIndex > 0) Then
                        Users.CurrentPageIndex -= 1
                    End If
                Case "下一頁"
                    If (Users.CurrentPageIndex < (Users.PageCount - 1)) Then
                        Users.CurrentPageIndex += 1
                    End If
                Case "最末頁"
                    Users.CurrentPageIndex = Users.PageCount - 1
                Case Else
                    'page number
                    'dgGroups.CurrentPageIndex = Convert.ToInt32(arg)
            End Select

            LoadMembers()
        End Sub



        Private Sub showPage()

            lbFirstPage.Enabled = Users.CurrentPageIndex > 0
            lbPrevPage.Enabled = (Users.CurrentPageIndex > 0)
            lbNextPage.Enabled = (Users.CurrentPageIndex + 1 < Users.PageCount)
            lbLastPage.Enabled = (Users.CurrentPageIndex + 1 < Users.PageCount)

            lblNowPage.Text = Users.CurrentPageIndex + 1
            lblTotalPage.Text = Users.PageCount

        End Sub

    End Class 'ProjectUser
End Namespace 'Hsinyi.EIP.ProjectManagement.Web