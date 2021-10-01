Imports System
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity

Namespace Hsinyi.EIP.ProjectManagement.Web
    Public Class myProjectList
        Inherits System.Web.UI.UserControl

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

        '���� Web Form �]�p�u��һݪ��I�s�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectsGrid As System.Web.UI.WebControls.DataGrid

        '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
        '�ФŧR���β��ʥ��C
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
            '�ФŨϥε{���X�s�边�i��ק�C
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

