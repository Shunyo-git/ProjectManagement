Imports System
Imports System.Data
Imports System.Web.UI.WebControls
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web



Public Class ForumTopicDetail
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents dgReplys As System.Web.UI.WebControls.DataGrid
        Protected WithEvents lbtnNew As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnReply As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents lblSubject As System.Web.UI.WebControls.Label
        Protected WithEvents lblAuthor As System.Web.UI.WebControls.Label
        Protected WithEvents lblModifiedDate As System.Web.UI.WebControls.Label
        Protected WithEvents lblContent As System.Web.UI.WebControls.Label
        Protected WithEvents lbtnEdit As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnDelete As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lblAttachment As System.Web.UI.WebControls.Label

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
        Private _topicID As Integer
        Private _documentID As Integer
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

            ' Obtain project id from the QueryString.
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            If Request.QueryString("TopicID") Is Nothing Then
                _topicID = 0
            Else
                _topicID = Convert.ToInt32(Request.QueryString("TopicID"))
            End If

            If Not Roles.isEnabledViewProject(_projectID) Then
                Session("m_ErrorMessage") = "系統拒絕存取，您沒有權限存取此專案資訊。" '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If


            If Roles.isEnabledModifyProject(_projectID) Or Roles.isProjectMember(_projectID) Then
                PanelModify.Visible = True
            Else
                PanelModify.Visible = False
            End If


            Page.RegisterStartupScript("toShowFile", ClientScript.ShowScript("function toShowFile(e) { var strLoc='getAttachment.aspx?ProjectID=" & _projectID.ToString & "&DocumentID=' + e.recordID + '&fileName=' + e.file   ;	   nosee.location.href=strLoc; }"))

            '刪除回覆
            If Not Request("DelReplyID") Is Nothing Then
                Dim DelReplyID As Integer = Convert.ToInt32(Request.QueryString("DelReplyID"))
                PMForumTopic.Remove(DelReplyID)
            End If

            If Not IsPostBack Then
                lbtnDelete.Attributes.Add("OnClick", "return chkConfirmDelete();")
                ' Load project with _projID when project id exists in the QueryString
                If _topicID <> 0 And _projectID <> 0 Then
                    BindInfo()

                Else
                    Session("m_ErrorMessage") = "系統錯誤，無法存取專案資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If
            End If

        End Sub

        Private Sub BindInfo()
            Dim topic As New PMForumTopic(_topicID)
            topic.Load()

            lblSubject.Text = topic.Subject

            lblModifiedDate.Text = topic.ModifiedDate
            lblAuthor.Text = EIPSysSecurity.clsUser.getUserName(topic.CreatedUserID)

            lblContent.Text = IIf(topic.Content Is String.Empty, "", topic.Content.Replace(vbCrLf, "<BR>"))

            _documentID = topic.DocumentID
             
            lblAttachment.Text = showAttachmentTable(_documentID)
            
            BindReplys()
        End Sub 'BindInfo
        Private Sub BindReplys()
            Dim replys As PMForumTopicCollection = PMForumTopic.GetReplys(_topicID)

            ' Call method to sort the data before databinding
            'SortGridData(EventList, SortField, SortAscending)

            dgReplys.DataSource = replys
            dgReplys.DataBind()
        End Sub
       

        Private Sub lbtnReply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnReply.Click
            Response.Redirect([String].Format("ForumTopicModify.aspx?projectID={0}&index=-1&projectIndex={1}&TopicID=0&ReplyTo={2}", _projectID, TabItem.ProjectTabIndex.ProjectForum, _topicID), True)
        End Sub

        Private Sub lbtnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnNew.Click
            Response.Redirect([String].Format("ForumTopicModify.aspx?projectID={0}&index=-1&projectIndex={1}&TopicID=0&ReplyTo=0", _projectID, TabItem.ProjectTabIndex.ProjectForum), True)
        End Sub


        Private Sub lbtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnDelete.Click
            PMForumTopic.Remove(_topicID)
            Response.Redirect([String].Format("ProjectForum.aspx?projectID={0}&index=-1&projectIndex={1}", _projectID, TabItem.ProjectTabIndex.ProjectForum), True)
        End Sub

        Private Sub lbtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnEdit.Click
            Response.Redirect([String].Format("ForumTopicModify.aspx?projectID={0}&index=-1&projectIndex=6&TopicID={1}&ReplyTo=0", _projectID, _topicID), True)
        End Sub

        Public Function showAttachmentTable(ByVal documentID As Integer) As String

            If documentID <= 0 Then
                Return String.Empty
            End If

            Dim strHtml As String = String.Empty
            Dim strWebPath As String = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Documents/" & _projectID & "/" & documentID & "/"

            Dim doc As New PMDocuments(documentID)
            doc.Load()



            If Not doc.Attachments Is Nothing Then
                Dim Attachments As PMAttachmentCollection = doc.Attachments
                Dim Attachment As PMAttachment
                For Each Attachment In Attachments
                    strHtml += (PMAttachment.showFile(Attachment.AttachmentName, 0, strWebPath, False, documentID))
                Next Attachment
            End If


            Return strHtml
        End Function 'showAttachment

    End Class

End Namespace