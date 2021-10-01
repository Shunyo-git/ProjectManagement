Imports System
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity

Namespace Hsinyi.EIP.ProjectManagement.Web
    Public Class myProjectList
        Inherits System.Web.UI.UserControl

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectsGrid As System.Web.UI.WebControls.DataGrid

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
            If Not Page.IsPostBack Then
                BindProjects()
            End If
        End Sub

        Private Sub BindProjects()

            Dim projectList As ProjectsCollection = Project.GetProjects(PMSecurity.GetUserID(), PMSecurity.GetUserRole())

            ' Call method to sort the data before databinding
            '  SortGridData(projectList, SortField, SortAscending)

            ProjectsGrid.DataSource = projectList
            ProjectsGrid.DataBind()
        End Sub 'BindProjects

        Public Function getStateColorName(ByVal intStateID As Integer, ByVal strStateName As String) As String

            Return "<FONT COLOR=""" & ProjectState.ShowStateColor(intStateID).Name & """>" & strStateName & "</FONT>"

        End Function

        Private Sub ProjectsGrid_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles ProjectsGrid.PageIndexChanged
            ProjectsGrid.EditItemIndex = -1
            ProjectsGrid.CurrentPageIndex = e.NewPageIndex
            BindProjects()
        End Sub
    End Class

End Namespace

