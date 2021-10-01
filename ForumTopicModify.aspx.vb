Imports System
Imports System.IO
Imports System.Data
Imports System.Web.UI.WebControls
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class ForumTopicModify
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents lbtnCancel As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnSave As System.Web.UI.WebControls.LinkButton
        Protected WithEvents ErrorMessage As System.Web.UI.WebControls.Label
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents txtFilename As System.Web.UI.WebControls.TextBox
        Protected WithEvents lblAttachment As System.Web.UI.WebControls.Label
        Protected WithEvents hlAttachment As System.Web.UI.WebControls.HyperLink
        Protected WithEvents Label4 As System.Web.UI.WebControls.Label
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents ProjectNameRequiredfieldvalidator As System.Web.UI.WebControls.RequiredFieldValidator
        Protected WithEvents itnTopic As System.Web.UI.WebControls.Label
        Protected WithEvents txtSubject As System.Web.UI.WebControls.TextBox
        Protected WithEvents txtContent As System.Web.UI.WebControls.TextBox
        Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm

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
        Private _replyTo As Integer
        Private _documentID As Integer = 0
        Private _strWebPath As String
        Private _strTempFilePath As String

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            If Request.QueryString("TopicID") Is Nothing Then
                _TopicID = 0
            Else
                _TopicID = Convert.ToInt32(Request.QueryString("TopicID"))
            End If

            If Request.QueryString("ReplyTo") Is Nothing Then
                _ReplyTo = 0
            Else
                _ReplyTo = Convert.ToInt32(Request.QueryString("ReplyTo"))
            End If

            doUpdateFilePath()

            If Not IsPostBack Then

                ' Load project with _projID when project id exists in the QueryString
                If _projectID = 0 Then

                    Session("m_ErrorMessage") = "系統錯誤，無法存取群組資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If

                If _topicID <> 0 Or _replyTo <> 0 Then
                    BindInfo()
                End If

            End If
        End Sub

        Private Sub doUpdateFilePath()
            _strWebPath = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Documents/" & _projectID & "/" & _documentID & "/"
            _strTempFilePath = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentSrvPath) & "Temp\" & PMSecurity.GetUserID & "\"

        End Sub
        Private Sub BindInfo()

            Dim topic As New PMForumTopic


            If _topicID = 0 And _replyTo <> 0 Then ' New Reply
                topic.TopicID = _replyTo
                topic.Load()
                txtSubject.ReadOnly = True
                txtSubject.Text = "RE:" & topic.Subject
            End If

            If _topicID <> 0 And _replyTo = 0 Then ' Modify Topic
                topic.TopicID = _topicID
                topic.Load()

                txtSubject.Text = topic.Subject
                txtContent.Text = topic.Content
                _documentID = topic.DocumentID
                chkAttachment(topic.Attachments)
            End If

            If _topicID <> 0 And _replyTo <> 0 Then ' Modify Reply
                topic.TopicID = _topicID
                topic.Load()
                txtSubject.ReadOnly = True
                txtSubject.Text = topic.Subject
                txtContent.Text = topic.Content
                _documentID = topic.DocumentID
                chkAttachment(topic.Attachments)
            End If

            hlAttachment.NavigateUrl = "javascript:getAttachment(" & _projectID.ToString & "," & _documentID.ToString & ",1,2);"

            Page.RegisterStartupScript("toShowFile", ClientScript.ShowScript("function toShowFile(e) { var strLoc='getAttachment.aspx?ProjectID=" & _projectID.ToString & "&DocumentID=" & _documentID.ToString & "&fileName=' + e.file   ;	 document.nosee.location.href=strLoc; }"))




        End Sub 'BindInfo
        Private Sub chkAttachment(ByVal Attachments As PMAttachmentCollection)
            txtFilename.Text = String.Empty
            Dim Attachment As PMAttachment
            For Each Attachment In Attachments
                lblAttachment.Text += showAttachment(Attachment.AttachmentName)

                If txtFilename.Text.Length > 0 Then txtFilename.Text += ","
                txtFilename.Text += Attachment.AttachmentName
            Next Attachment
        End Sub

        Private Function showAttachment(ByVal filename As String) As String

            Return PMAttachment.showFile(filename, 0, _strWebPath, True, 0)

        End Function 'BindAttachment

        Private Sub lbtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnCancel.Click
            BackToUserPage()
        End Sub

        Private Sub BackToUserPage()
            Dim topicID
            If _replyTo <> 0 Then
                topicID = _replyTo
            Else
                topicID = _topicID
            End If
            Response.Redirect([String].Format("ForumTopicDetail.aspx?ProjectID={0}&TopicID={1}&index=-1&projectIndex={2}", _projectID, topicID, TabItem.ProjectTabIndex.ProjectForum), False)
        End Sub

        Private Sub lbtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnSave.Click
            SaveTopic()

        End Sub

        Private Sub SaveTopic()
            Dim topic As New PMForumTopic(_topicID)
            If _topicID > 0 Then
                topic.Load()
            End If
            topic.Subject = txtSubject.Text
            topic.Content = txtContent.Text
            topic.ReplyTo = _replyTo
            topic.ModifiedUserID = PMSecurity.GetUserID
            topic.ProjectID = _projectID

            If InStr(txtFilename.Text, ".") > 0 Then
                Dim attachmentFileCol As New PMAttachmentCollection
                Dim aryAttacment As Array = Strings.Split(txtFilename.Text, ",")
                Dim filename As String
                For Each filename In aryAttacment
                    If Strings.Trim(filename) <> "" Then
                        Dim attachmentFile As New PMAttachment
                        attachmentFile.AttachmentName = filename
                        attachmentFileCol.Add(attachmentFile)
                        'Response.Write(filename)
                    End If
                Next
                topic.Attachments = attachmentFileCol
            End If

            If topic.Save() Then
                If _replyTo <> 0 Then
                    _topicID = _replyTo
                Else
                    _topicID = topic.TopicID
                End If
                saveAttachmentFiles()
                BackToUserPage()
            Else
                Session("m_ErrorMessage") = "系統錯誤，無法儲存群組資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If
        End Sub

        Private Sub saveAttachmentFiles()
            doUpdateFilePath()
            If Not Directory.Exists(_strWebPath) Then
                Directory.CreateDirectory(_strWebPath)
            End If

            Dim strUplosdFileName As String
            strUplosdFileName = txtFilename.Text
            If Strings.Trim(strUplosdFileName) <> "" Then
                Dim aryFile As Array = Split(strUplosdFileName, ",")
                Dim iFile As String
                For Each iFile In aryFile
                    If Strings.Trim(iFile) <> "" Then
                        If File.Exists(_strTempFilePath & iFile) Then
                            If File.Exists(_strWebPath & iFile) Then
                                File.Delete(_strWebPath & iFile)
                            End If
                            File.Move(_strTempFilePath & iFile, _strWebPath & iFile)

                        Else
                            'ErrorMessage.Text += "儲存檔案失敗，找不到 " & iFile & " 檔案。"
                        End If
                    End If

                Next
            End If
        End Sub
    End Class 'ForumTopicModify

End Namespace


