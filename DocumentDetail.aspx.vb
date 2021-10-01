Imports System
Imports System.IO
Imports System.Data
Imports System.Web.UI.WebControls
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class DocumentDetail
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents ErrorMessage As System.Web.UI.WebControls.Label
        Protected WithEvents lbtnSave As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnCancel As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents txtFilename As System.Web.UI.WebControls.TextBox
        Protected WithEvents lblAttachment As System.Web.UI.WebControls.Label
        Protected WithEvents hlAttachment As System.Web.UI.WebControls.HyperLink
        Protected WithEvents Label4 As System.Web.UI.WebControls.Label
        Protected WithEvents txtDescription As System.Web.UI.WebControls.TextBox
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents ProjectNameRequiredfieldvalidator As System.Web.UI.WebControls.RequiredFieldValidator
        Protected WithEvents txtTitle As System.Web.UI.WebControls.TextBox
        Protected WithEvents itnTopic As System.Web.UI.WebControls.Label
        Protected WithEvents Label2 As System.Web.UI.WebControls.Label
        Protected WithEvents ddlCategory As System.Web.UI.WebControls.DropDownList
        Protected WithEvents lbtnEdit As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnDelete As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lblTitle As System.Web.UI.WebControls.Label
        Protected WithEvents lblCategory As System.Web.UI.WebControls.Label
        Protected WithEvents lblDescription As System.Web.UI.WebControls.Label
        Protected WithEvents Label3 As System.Web.UI.WebControls.Label
        Protected WithEvents lblCreatedUserID As System.Web.UI.WebControls.Label
        Protected WithEvents lblCreatedDate As System.Web.UI.WebControls.Label
        Protected WithEvents Label6 As System.Web.UI.WebControls.Label
        Protected WithEvents lblModifiedUserID As System.Web.UI.WebControls.Label
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
        Private _documentID As Integer
       
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            If Request.QueryString("DocumentID") Is Nothing Then
                _documentID = 0
            Else
                _documentID = Convert.ToInt32(Request.QueryString("DocumentID"))
            End If

            If Roles.isEnabledModifyProject(_projectID) Then
                PanelModify.Visible = True
            Else
                PanelModify.Visible = False
            End If

            Page.RegisterStartupScript("toShowFile", ClientScript.ShowScript("function toShowFile(e) { var strLoc='getAttachment.aspx?ProjectID=" & _projectID.ToString & "&DocumentID=" & _documentID.ToString & "&fileName=' + e.file   ;	 document.nosee.location.href=strLoc; }"))

            If Not IsPostBack Then
                lbtnDelete.Attributes.Add("OnClick", "return chkConfirmDelete();")
                ' Load project with _projID when project id exists in the QueryString
                If _documentID <> 0 And _projectID <> 0 Then
                    BindInfo()
                Else
                    Session("m_ErrorMessage") = "系統錯誤，無法存取專案資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If
            End If
        End Sub
        
       
        Private Sub BindInfo()

            Dim doc As New PMDocuments(_documentID)
            doc.Load()
            _projectID = doc.ProjectID
            lblTitle.Text = doc.Title
            lblDescription.Text = doc.Description

            lblCategory.Text = doc.CategoryName
            lblCreatedUserID.Text = EIPSysSecurity.clsUser.getUserName(doc.CreatedUserID)
            lblCreatedDate.Text = doc.CreatedDate
            lblModifiedUserID.Text = EIPSysSecurity.clsUser.getUserName(doc.ModifiedUserID)
            lblModifiedDate.Text = doc.ModifiedDate


            'If PMSecurity.GetUserID = doc.ModifiedUserID Then
            '    PanelModify.Visible = True
            'Else

            '    If Not Roles.isEnabledModifyProject(_projectID) Then
            '        Session("m_ErrorMessage") = "系統拒絕存取，您沒有權限存修改專案資訊。" '"您非本專案成員，因此無權限存取此專案資訊。"
            '        Response.Redirect("AlertMessage.aspx?Index=-1", True)
            '        PanelModify.Visible = False
            '    Else
            '        PanelModify.Visible = True
            '    End If
            'End If

            If PMSecurity.GetUserID = doc.ModifiedUserID Then
                PanelModify.Visible = True
            End If

            Dim Attachment As PMAttachment
            For Each Attachment In doc.Attachments
                lblAttachment.Text += showAttachment(Attachment.AttachmentName)
            Next Attachment

        End Sub 'BindInfo



        Private Function showAttachment(ByVal filename As String) As String

            Dim strWebPath As String = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Documents/" & _projectID & "/" & _documentID & "/"
            Return PMAttachment.showFile(filename, 0, strWebPath, False, 0)
        End Function 'BindAttachment

        


        Private Sub BackToProjectList()
            Response.Redirect([String].Format("ProjectDocument.aspx?ProjectID={0}&index=-1&projectIndex={1}", _projectID, TabItem.ProjectTabIndex.ProjectDocument), False)
        End Sub

        Private Sub lbtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnEdit.Click
            Response.Redirect([String].Format("DocumentModify.aspx?projectID={0}&index=-1&projectIndex={1}&DocumentID={2}", _projectID, TabItem.ProjectTabIndex.ProjectDocument, _documentID), False)
        End Sub 'lbtnEdit_Click

        Private Sub lbtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnDelete.Click
            PMDocuments.Remove(_documentID)
            BackToProjectList()
        End Sub


    End Class

End Namespace


