Imports System
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class ProjectEvent
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents lbtnNew As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents dgEvents As System.Web.UI.WebControls.DataGrid
        Protected WithEvents EventList As System.Web.UI.HtmlControls.HtmlForm
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

            If Roles.isEnabledModifyProject(_projectID) Or Roles.isProjectMember(_projectID) Then
                PanelModify.Visible = True
            Else
                PanelModify.Visible = False
            End If



            If Not Page.IsPostBack Then
                If _projectID <> 0 Then
                    BindEvents()
                End If

            End If


        End Sub

        Private Sub BindEvents()


            Dim EventList As PMEventsCollection = PMEvent.GetEvents(_projectID)

            ' Call method to sort the data before databinding
            SortGridData(EventList, SortField, SortAscending)
            lblRecordCount.Text = EventList.Count
            dgEvents.DataSource = EventList
            dgEvents.DataBind()
            showPage()
        End Sub 'BindProjects

        Private Sub dgEvents_Sort(ByVal [source] As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles dgEvents.SortCommand
            SortField = e.SortExpression

            BindEvents()
        End Sub 'dgEvents_Sort

        '*********************************************************************
        '
        ' The SortGridData method is a helper method called when databinding the Projects grid 
        ' to sort the columns of the grid.
        '
        '*********************************************************************

        Private Sub SortGridData(ByVal list As PMEventsCollection, ByVal sortField As String, ByVal asc As Boolean)
            Dim sortCol As PMEventsCollection.EventFields = PMEventsCollection.EventFields.InitValue

            Select Case sortField
                Case "EventDate"
                    sortCol = PMEventsCollection.EventFields.EventDate
                Case "Topic"
                    sortCol = PMEventsCollection.EventFields.Topic
                Case "CreatedUserID"
                    sortCol = PMEventsCollection.EventFields.CreatedUserID
                Case "CreatedDate"
                    sortCol = PMEventsCollection.EventFields.CreatedDate

                Case Else
            End Select

            list.Sort(sortCol, asc)
        End Sub 'SortGridData
        '*******************************************************
        '
        ' Users_PageIndexChanged server event handler on this page is used
        ' for changing page index
        '
        '*******************************************************

        Private Sub Groups_PageIndexChanged(ByVal [source] As System.Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles dgEvents.PageIndexChanged
            dgEvents.EditItemIndex = -1
            dgEvents.CurrentPageIndex = e.NewPageIndex
            BindEvents()
        End Sub 'Groups_PageIndexChanged

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
            Response.Redirect([String].Format("EventModify.aspx?projectID={0}&index=-1&projectIndex={1}&EventID=0", _projectID, TabItem.ProjectTabIndex.ProjectEvent), False)
        End Sub

        Sub PagerButtonClick(ByVal sender As Object, ByVal e As EventArgs)
            'used by external paging UI
            Dim arg As String = sender.CommandArgument

            Select Case arg
                Case "最前頁"
                    dgEvents.CurrentPageIndex = 0

                Case "上一頁"
                    If (dgEvents.CurrentPageIndex > 0) Then
                        dgEvents.CurrentPageIndex -= 1
                    End If
                Case "下一頁"
                    If (dgEvents.CurrentPageIndex < (dgEvents.PageCount - 1)) Then
                        dgEvents.CurrentPageIndex += 1
                    End If
                Case "最末頁"
                    dgEvents.CurrentPageIndex = dgEvents.PageCount - 1
                Case Else
                    'page number
                    'dgGroups.CurrentPageIndex = Convert.ToInt32(arg)
            End Select

            BindEvents()
        End Sub



        Private Sub showPage()

            lbFirstPage.Enabled = dgEvents.CurrentPageIndex > 0
            lbPrevPage.Enabled = (dgEvents.CurrentPageIndex > 0)
            lbNextPage.Enabled = (dgEvents.CurrentPageIndex + 1 < dgEvents.PageCount)
            lbLastPage.Enabled = (dgEvents.CurrentPageIndex + 1 < dgEvents.PageCount)

            lblNowPage.Text = dgEvents.CurrentPageIndex + 1
            lblTotalPage.Text = dgEvents.PageCount

        End Sub
    End Class 'ProjectEvent

End Namespace 'Hsinyi.EIP.ProjectManagement.Web
 