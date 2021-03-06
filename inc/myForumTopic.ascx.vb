Imports System
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class myForumTopic
        Inherits System.Web.UI.UserControl

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents dgList As System.Web.UI.WebControls.DataGrid

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

                BindTopics()

            End If
        End Sub

        Private Sub BindTopics()


            Dim TopicList As PMForumTopicCollection = PMForumTopic.GetTopics(PMSecurity.GetUserID(), PMSecurity.GetUserRole())

            ' Call method to sort the data before databinding
            ' SortGridData(TopicList, SortField, SortAscending)

            dgList.DataSource = TopicList
            dgList.DataBind()
        End Sub 'BindProjects

        Public Function getProjectName(ByVal ProjectID As Integer) As String

            Dim pj As New Project(ProjectID)
            pj.Load()
            Return pj.ProjectName
        End Function

        Private Sub dgList_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles dgList.PageIndexChanged
            dgList.EditItemIndex = -1
            dgList.CurrentPageIndex = e.NewPageIndex
            BindTopics()
        End Sub
    End Class

End Namespace
