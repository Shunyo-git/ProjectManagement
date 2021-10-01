Imports System

Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity

Namespace Hsinyi.EIP.ProjectManagement.Web
 
Public Class ProjectDocument
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
    Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents lbtnNew As System.Web.UI.WebControls.LinkButton
    Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents lblRecordCount As System.Web.UI.WebControls.Label
        Protected WithEvents lbNextPage As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbPrevPage As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lblTotalPage As System.Web.UI.WebControls.Label
        Protected WithEvents lblNowPage As System.Web.UI.WebControls.Label
        Protected WithEvents lbLastPage As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbFirstPage As System.Web.UI.WebControls.LinkButton
        Protected WithEvents dgDocuments As System.Web.UI.WebControls.DataGrid

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
        Private categoryIndex As Integer

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Obtain project id from the QueryString.
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            If Not Request("categoryIndex") Is Nothing Then
                categoryIndex = Convert.ToInt32(Request("categoryIndex"))
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

            Page.RegisterStartupScript("toShowFile", ClientScript.ShowScript("function toShowFile(e) { var strLoc='getAttachment.aspx?ProjectID=" & _projectID.ToString & "&DocumentID=' + e.recordID + '&fileName=' + e.file   ;	 document.nosee.location.href=strLoc; }"))


            If Not Page.IsPostBack Then
                If _projectID <> 0 Then
                    BindDocuments()
                End If

            End If
    End Sub

        Private Sub BindDocuments()
            Dim docs As PMDocumentsCollection


            If Not Request("categoryIndex") Is Nothing Then
                Dim categoryID As Integer = PMDocCategories.GetCategoryBySort(categoryIndex + 1).CategoryID
                docs = PMDocuments.GetDocuments(_projectID, categoryID)
            Else
                docs = PMDocuments.GetDocuments(_projectID)
            End If


            ' Call method to sort the data before databinding
            SortGridData(docs, SortField, SortAscending)
            lblRecordCount.Text = docs.Count
            dgDocuments.DataSource = docs
            dgDocuments.DataBind()
            showPage()

        End Sub 'BindProjects

        Private Sub dgDocuments_Sort(ByVal [source] As System.Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles dgDocuments.SortCommand
            SortField = e.SortExpression

            BindDocuments()
        End Sub 'dgEvents_Sort

        '*********************************************************************
        '
        ' The SortGridData method is a helper method called when databinding the Projects grid 
        ' to sort the columns of the grid.
        '
        '*********************************************************************

        Private Sub SortGridData(ByVal list As PMDocumentsCollection, ByVal sortField As String, ByVal asc As Boolean)
            Dim sortCol As PMDocumentsCollection.DocumentFields = PMDocumentsCollection.DocumentFields.InitValue

            Select Case sortField
                Case "Category"
                    sortCol = PMDocumentsCollection.DocumentFields.CategoryID
                Case "Title"
                    sortCol = PMDocumentsCollection.DocumentFields.Title
                Case "CreatedUserID"
                    sortCol = PMDocumentsCollection.DocumentFields.CreatedUserID
                Case "CreatedDate"
                    sortCol = PMDocumentsCollection.DocumentFields.CreatedDate

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

        Private Sub Groups_PageIndexChanged(ByVal [source] As System.Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles dgDocuments.PageIndexChanged
            dgDocuments.EditItemIndex = -1
            dgDocuments.CurrentPageIndex = e.NewPageIndex
            BindDocuments()
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
            Response.Redirect([String].Format("DocumentModify.aspx?projectID={0}&index=-1&projectIndex={1}&DocumentID=0", _projectID, TabItem.ProjectTabIndex.ProjectDocument), False)
        End Sub

        Public Function showAttachment(ByVal documentID As Integer) As String
            Dim strHtml As String = String.Empty
            Dim strWebPath As String = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Documents/" & _projectID & "/" & documentID & "/"

            Dim doc As New PMDocuments(documentID)
            doc.Load()
            Dim Attachments As PMAttachmentCollection = doc.Attachments

            Dim Attachment As PMAttachment
            For Each Attachment In Attachments
                strHtml += (PMAttachment.showFile(Attachment.AttachmentName, 0, strWebPath, False, documentID))
            Next Attachment

            Return strHtml
        End Function 'BindAttachment

        Sub PagerButtonClick(ByVal sender As Object, ByVal e As EventArgs)
            'used by external paging UI
            Dim arg As String = sender.CommandArgument

            Select Case arg
                Case "最前頁"
                    dgDocuments.CurrentPageIndex = 0

                Case "上一頁"
                    If (dgDocuments.CurrentPageIndex > 0) Then
                        dgDocuments.CurrentPageIndex -= 1
                    End If
                Case "下一頁"
                    If (dgDocuments.CurrentPageIndex < (dgDocuments.PageCount - 1)) Then
                        dgDocuments.CurrentPageIndex += 1
                    End If
                Case "最末頁"
                    dgDocuments.CurrentPageIndex = dgDocuments.PageCount - 1
                Case Else
                    'page number
                    'dgGroups.CurrentPageIndex = Convert.ToInt32(arg)
            End Select

            BindDocuments()
        End Sub



        Private Sub showPage()

            lbFirstPage.Enabled = dgDocuments.CurrentPageIndex > 0
            lbPrevPage.Enabled = (dgDocuments.CurrentPageIndex > 0)
            lbNextPage.Enabled = (dgDocuments.CurrentPageIndex + 1 < dgDocuments.PageCount)
            lbLastPage.Enabled = (dgDocuments.CurrentPageIndex + 1 < dgDocuments.PageCount)

            lblNowPage.Text = dgDocuments.CurrentPageIndex + 1
            lblTotalPage.Text = dgDocuments.PageCount

        End Sub
End Class
End Namespace