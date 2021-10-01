Imports System
Imports System.Data
Imports System.Web.UI.WebControls
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web
    Public Class EventDetail
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents Label4 As System.Web.UI.WebControls.Label
        Protected WithEvents lblTel As System.Web.UI.WebControls.Label
        Protected WithEvents passwordlbl As System.Web.UI.WebControls.Label
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents lbtnDelete As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnEdit As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents itnTopic As System.Web.UI.WebControls.Label
        Protected WithEvents lblTopic As System.Web.UI.WebControls.Label
        Protected WithEvents itnDate As System.Web.UI.WebControls.Label
        Protected WithEvents lblEventDate As System.Web.UI.WebControls.Label
        Protected WithEvents lblReason As System.Web.UI.WebControls.Label
        Protected WithEvents lblAnswer As System.Web.UI.WebControls.Label
        Protected WithEvents lblEffect As System.Web.UI.WebControls.Label
        Protected WithEvents lblAttachment As System.Web.UI.WebControls.Label
        Protected WithEvents Label2 As System.Web.UI.WebControls.Label
        Protected WithEvents Label6 As System.Web.UI.WebControls.Label
        Protected WithEvents lblCreatedUserID As System.Web.UI.WebControls.Label
        Protected WithEvents lblModifiedUserID As System.Web.UI.WebControls.Label
        Protected WithEvents lblCreatedDate As System.Web.UI.WebControls.Label
        Protected WithEvents lblModifiedDate As System.Web.UI.WebControls.Label

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
        Private _eventID As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

            ' Obtain project id from the QueryString.
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            If Request.QueryString("EventID") Is Nothing Then
                _eventID = 0
            Else
                _eventID = Convert.ToInt32(Request.QueryString("EventID"))
            End If

            If Not Roles.isEnabledViewProject(_projectID) Then
                Session("m_ErrorMessage") = "系統拒絕存取，您沒有權限存取此專案資訊。" '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If


            If Roles.isEnabledModifyProject(_projectID) Then
                PanelModify.Visible = True
            Else
                PanelModify.Visible = False
            End If


            Page.RegisterStartupScript("toShowFile", ClientScript.ShowScript("function toShowFile(e) { var strLoc='getAttachment.aspx?ProjectID=" & _projectID.ToString & "&EventID=" & _eventID.ToString & "&fileName=' + e.file   ;	document.nosee.location.href=(strLoc ; }"))


            If Not IsPostBack Then
                lbtnDelete.Attributes.Add("OnClick", "return chkConfirmDelete();")
                ' Load project with _projID when project id exists in the QueryString
                If _eventID <> 0 And _projectID <> 0 Then
                    BindInfo()
                Else
                    Session("m_ErrorMessage") = "系統錯誤，無法存取專案資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If
            End If
        End Sub

        Private Sub BindInfo()
            Dim pjevent As New PMEvent(_eventID)
            pjevent.Load()
            _projectID = pjevent.ProjectID
            lblTopic.Text = pjevent.Topic

            lblEventDate.Text = pjevent.EventDate
            lblReason.Text = IIf(pjevent.Reason Is String.Empty, "", pjevent.Reason.Replace(vbCrLf, "<BR>"))
            'lblBackground.Text = IIf(pjevent.Background Is String.Empty, "", pjevent.Background.Replace(vbCrLf, "<BR>"))
            lblAnswer.Text = IIf(pjevent.Answer Is String.Empty, "", pjevent.Answer.Replace(vbCrLf, "<BR>"))
            lblEffect.Text = IIf(pjevent.Effect Is String.Empty, "", pjevent.Effect.Replace(vbCrLf, "<BR>"))
            lblCreatedUserID.Text = EIPSysSecurity.clsUser.getUserName(pjevent.CreatedUserID)
            lblCreatedDate.Text = pjevent.CreatedDate
            lblModifiedUserID.Text = EIPSysSecurity.clsUser.getUserName(pjevent.ModifiedUserID)
            lblModifiedDate.Text = pjevent.ModifiedDate

            If PMSecurity.GetUserID = pjevent.ModifiedUserID Then
                PanelModify.Visible = True
            End If

            Dim Attachment As PMAttachment
            For Each Attachment In pjevent.Attachments
                lblAttachment.Text += showAttachment(Attachment.AttachmentName)
            Next Attachment

        End Sub 'BindInfo
        Private Function showAttachment(ByVal filename As String) As String

            Dim strWebPath As String = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Events/" & _projectID & "/" & _eventID & "/"
            Return PMAttachment.showFile(filename, 0, strWebPath, False, 0)
        End Function 'BindAttachment

        Private Sub lbtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnEdit.Click
            Response.Redirect([String].Format("EventModify.aspx?projectID={0}&index=-1&projectIndex={1}&EventID={2}", _projectID, TabItem.ProjectTabIndex.ProjectEvent, _eventID), False)
        End Sub 'lbtnEdit_Click

        Private Sub lbtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnDelete.Click
            PMEvent.Remove(_eventID)
            BackToProjectList()
        End Sub

        Private Sub BackToProjectList()
            Response.Redirect([String].Format("ProjectEvent.aspx?projectID={0}&index=-1&projectIndex={1}", _projectID, TabItem.ProjectTabIndex.ProjectEvent), False)
        End Sub


    End Class

End Namespace

