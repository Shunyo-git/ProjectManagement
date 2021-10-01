Imports System
Imports System.Data
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer


Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class ProjectGroup
        Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

        '���� Web Form �]�p�u��һݪ��I�s�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents dgGroups As System.Web.UI.WebControls.DataGrid
        Protected WithEvents lbtnNew As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents GroupList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents lbFirstPage As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbNextPage As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbLastPage As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lblNowPage As System.Web.UI.WebControls.Label
        Protected WithEvents lblTotalPage As System.Web.UI.WebControls.Label
        Protected WithEvents lbPrevPage As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lblRecordCount As System.Web.UI.WebControls.Label

        '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
        '�ФŧR���β��ʥ��C
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
            '�ФŨϥε{���X�s�边�i��ק�C
            InitializeComponent()
        End Sub

#End Region
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
                Session("m_ErrorMessage") = "�t�Ωڵ��s���A�z�S���v���s�����M�׸�T�C" '"�z�D���M�צ����A�]���L�v���s�����M�׸�T�C"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If

            If Roles.isEnabledModifyProject(_projectID) Then
                PanelModify.Visible = True
            Else
                PanelModify.Visible = False
            End If



            If Not Page.IsPostBack Then
                If _projectID <> 0 Then
                    LoadGroups()
                End If
            End If

            

        End Sub 'Page_Load


        '*******************************************************
        '
        ' LoadGroups method populate the DataGrid with a list of Time Tracker users.
        '
        '*******************************************************

        Private Sub LoadGroups()
            Dim Groups As GroupsCollection = PMGroup.GetProjectGroups(_projectID)

            ' Check for sorting when binding grid
            SortGridData(Groups, SortField, SortAscending)
            dgGroups.DataSource = Groups
            dgGroups.DataKeyField = "GroupID"
            dgGroups.DataBind()
            lblRecordCount.Text = Groups.Count
            showPage()
        End Sub 'LoadGroups

        '*******************************************************
        '
        ' SortGridData methods sorts the Users Grid based on which
        ' sort field is being selected.  Also does reverse sorting based on the boolean.
        '
        '*******************************************************

        Private Sub SortGridData(ByVal list As GroupsCollection, ByVal sortField As String, ByVal asc As Boolean)
            Dim sortCol As GroupsCollection.GroupsFields = GroupsCollection.GroupsFields.InitValue

            Select Case sortField
                Case "Name"
                    sortCol = GroupsCollection.GroupsFields.Name

                Case "Description"
                    sortCol = GroupsCollection.GroupsFields.Description

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

        Private Sub Groups_Sort(ByVal [source] As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles dgGroups.SortCommand
            SortField = e.SortExpression
            LoadGroups()
        End Sub 'Groups_Sort

       


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
            Response.Redirect([String].Format("GroupModify.aspx?projectID={0}&index=-1&projectIndex={1}&GroupID=0", _projectID, TabItem.ProjectTabIndex.ProjectGroup), False)
        End Sub
        Public Function getGroupMembers(ByVal groupID As Integer) As String
            Dim group As New PMGroup(groupID)
            group.Load()

            Dim strMembers As String
            Dim user As PMUser
            For Each user In group.Members
                If Not Strings.Trim(strMembers) Is String.Empty Then
                    strMembers += "�B"
                End If
                strMembers += user.DisplayName
            Next user
            Return strMembers
        End Function
        '*******************************************************
        '
        ' Users_PageIndexChanged server event handler on this page is used
        ' for changing page index
        '
        '*******************************************************
 

        Sub PagerButtonClick(ByVal sender As Object, ByVal e As EventArgs)
            'used by external paging UI
            Dim arg As String = sender.CommandArgument

            Select Case arg
                Case "�̫e��"
                    dgGroups.CurrentPageIndex = 0
                   
                Case "�W�@��"
                    If (dgGroups.CurrentPageIndex > 0) Then
                        dgGroups.CurrentPageIndex -= 1
                    End If
                Case "�U�@��"
                    If (dgGroups.CurrentPageIndex < (dgGroups.PageCount - 1)) Then
                        dgGroups.CurrentPageIndex += 1
                    End If
                Case "�̥���"
                    dgGroups.CurrentPageIndex = dgGroups.PageCount - 1
                Case Else
                    'page number
                    'dgGroups.CurrentPageIndex = Convert.ToInt32(arg)
            End Select

            LoadGroups()
        End Sub

 

        Private Sub showPage()

            lbFirstPage.Enabled = dgGroups.CurrentPageIndex > 0
            lbPrevPage.Enabled = (dgGroups.CurrentPageIndex > 0)
            lbNextPage.Enabled = (dgGroups.CurrentPageIndex + 1 < dgGroups.PageCount)
            lbLastPage.Enabled = (dgGroups.CurrentPageIndex + 1 < dgGroups.PageCount)

            lblNowPage.Text = dgGroups.CurrentPageIndex + 1
            lblTotalPage.Text = dgGroups.PageCount

        End Sub
    End Class
End Namespace


