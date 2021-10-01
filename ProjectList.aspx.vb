Imports System
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity

Namespace Hsinyi.EIP.ProjectManagement.Web
 
Public Class ProjectList
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
    Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents ProjectsGrid As System.Web.UI.WebControls.DataGrid
        Protected WithEvents NewProjectButton As System.Web.UI.WebControls.Button
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

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            '在這裡放置使用者程式碼以初始化網頁
            ' Ensure that the visiting user has access to view the current page
            'If PMSecurity.IsInRole(PMUser.UserRoleAdminPMgr) = False Then
            '    Response.Redirect("AccessDenied.aspx?Index=-1", True)
            'End If

            If Not IsPostBack Then
                BindProjects()
            End If

        End Sub

        Private Sub BindProjects()

            Dim projectList As ProjectsCollection = Project.GetProjects(PMSecurity.GetUserID(), PMSecurity.GetUserRole())

            ' Call method to sort the data before databinding
            SortGridData(projectList, SortField, SortAscending)
            lblRecordCount.Text = projectList.Count
            ProjectsGrid.DataSource = projectList
            ProjectsGrid.DataBind()
            showPage()
        End Sub 'BindProjects

        '*********************************************************************
        '
        ' The SortGridData method is a helper method called when databinding the Projects grid 
        ' to sort the columns of the grid.
        '
        '*********************************************************************

        Private Sub SortGridData(ByVal list As ProjectsCollection, ByVal sortField As String, ByVal asc As Boolean)
            Dim sortCol As ProjectsCollection.ProjectFields = ProjectsCollection.ProjectFields.InitValue

            Select Case sortField
                Case "ProjectName"
                    sortCol = ProjectsCollection.ProjectFields.ProjectName
                Case "Managere"
                    sortCol = ProjectsCollection.ProjectFields.Manager
                Case "ProjectChief"
                    sortCol = ProjectsCollection.ProjectFields.ProjectChief
                Case "StartDate"
                    sortCol = ProjectsCollection.ProjectFields.StartDate
                Case "EndDate"
                    sortCol = ProjectsCollection.ProjectFields.EndDate
                Case "State"
                    sortCol = ProjectsCollection.ProjectFields.State
                Case Else
            End Select

            list.Sort(sortCol, asc)
        End Sub 'SortGridData

        '*********************************************************************
        '
        ' The NewProjectButton_Click event handler redirects to the ProjectDetails.aspx 
        ' page so the user can create a new project.
        '
        '*********************************************************************
        Private Sub NewProjectButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewProjectButton.Click
            Dim mainIndex As Integer = IIf(Request("index") Is Nothing, 0, Convert.ToInt32(Request("index")))
            Dim projectIndex As Integer = IIf(Request("projectIndex") Is Nothing, 0, Convert.ToInt32(Request("projectIndex")))

            Response.Redirect([String].Format("ProjectDetails.aspx?index={0}&projectIndex={1}", mainIndex, projectIndex), True)
        End Sub

        Protected Sub ProjectsGrid_Page(ByVal sender As [Object], ByVal e As DataGridPageChangedEventArgs)
            ProjectsGrid.CurrentPageIndex = e.NewPageIndex

            BindProjects()
        End Sub 'ProjectsGrid_Page

        '*********************************************************************
        '
        ' The ProjectGrid_Sort event handler changes the sortfield for the Projects grid 
        ' and re-binds it.
        '
        '*********************************************************************

        Private Sub ProjectsGrid_Sort(ByVal [source] As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles ProjectsGrid.SortCommand
            SortField = e.SortExpression

            BindProjects()
        End Sub 'ProjectsGrid_Sort

        Property SortField() As String
            Get
                Dim o As Object = ViewState("SortField")
                If o Is Nothing Then
                    Return [String].Empty
                End If
                Return CStr(o)
            End Get

            Set(ByVal Value As String)
                If Value = SortField Then
                    ' same as current sort file, toggle sort direction
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

        Public Function getStateColorName(ByVal intStateID As Integer, ByVal strStateName As String) As String

            Return "<FONT COLOR=""" & ProjectState.ShowStateColor(intStateID).Name & """>" & strStateName & "</FONT>"

        End Function


        Sub PagerButtonClick(ByVal sender As Object, ByVal e As EventArgs)
            'used by external paging UI
            Dim arg As String = sender.CommandArgument

            Select Case arg
                Case "最前頁"
                    ProjectsGrid.CurrentPageIndex = 0

                Case "上一頁"
                    If (ProjectsGrid.CurrentPageIndex > 0) Then
                        ProjectsGrid.CurrentPageIndex -= 1
                    End If
                Case "下一頁"
                    If (ProjectsGrid.CurrentPageIndex < (ProjectsGrid.PageCount - 1)) Then
                        ProjectsGrid.CurrentPageIndex += 1
                    End If
                Case "最末頁"
                    ProjectsGrid.CurrentPageIndex = ProjectsGrid.PageCount - 1
                Case Else
                    'page number
                    'dgGroups.CurrentPageIndex = Convert.ToInt32(arg)
            End Select

            BindProjects()
        End Sub



        Private Sub showPage()

            lbFirstPage.Enabled = ProjectsGrid.CurrentPageIndex > 0
            lbPrevPage.Enabled = (ProjectsGrid.CurrentPageIndex > 0)
            lbNextPage.Enabled = (ProjectsGrid.CurrentPageIndex + 1 < ProjectsGrid.PageCount)
            lbLastPage.Enabled = (ProjectsGrid.CurrentPageIndex + 1 < ProjectsGrid.PageCount)

            lblNowPage.Text = ProjectsGrid.CurrentPageIndex + 1
            lblTotalPage.Text = ProjectsGrid.PageCount

        End Sub

    End Class


End Namespace 'Hsinyi.EIP.ProjectManagement.Web
