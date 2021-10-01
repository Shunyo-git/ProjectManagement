Imports System
Imports System.IO
Imports System.Data
Imports System.Web.UI.WebControls
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class DocumentModify
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            Me.ID = "DocumentDetail"

        End Sub
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents lblAttachment As System.Web.UI.WebControls.Label
        Protected WithEvents Label4 As System.Web.UI.WebControls.Label
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents itnTopic As System.Web.UI.WebControls.Label
        Protected WithEvents Label2 As System.Web.UI.WebControls.Label
        Protected WithEvents lblModifiedDate As System.Web.UI.WebControls.Label
        Protected WithEvents lblModifiedUserID As System.Web.UI.WebControls.Label
        Protected WithEvents Label6 As System.Web.UI.WebControls.Label
        Protected WithEvents lblCreatedDate As System.Web.UI.WebControls.Label
        Protected WithEvents lblCreatedUserID As System.Web.UI.WebControls.Label
        Protected WithEvents Label3 As System.Web.UI.WebControls.Label
        Protected WithEvents lbtnDelete As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnEdit As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents lblTitle As System.Web.UI.WebControls.Label
        Protected WithEvents lblCategory As System.Web.UI.WebControls.Label
        Protected WithEvents lblDescription As System.Web.UI.WebControls.Label
        Protected WithEvents ErrorMessage As System.Web.UI.WebControls.Label
        Protected WithEvents lbtnSave As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnCancel As System.Web.UI.WebControls.LinkButton
        Protected WithEvents txtTitle As System.Web.UI.WebControls.TextBox
        Protected WithEvents ProjectNameRequiredfieldvalidator As System.Web.UI.WebControls.RequiredFieldValidator
        Protected WithEvents ddlCategory As System.Web.UI.WebControls.DropDownList
        Protected WithEvents txtDescription As System.Web.UI.WebControls.TextBox
        Protected WithEvents hlAttachment As System.Web.UI.WebControls.HyperLink
        Protected WithEvents txtFilename As System.Web.UI.WebControls.TextBox
        Protected WithEvents DocumentModify As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents DocumentDetail As System.Web.UI.HtmlControls.HtmlForm

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
        Private _strWebPath As String
        Private _strTempFilePath As String
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

            doUpdateFilePath()

            hlAttachment.NavigateUrl = "javascript:getAttachment(" & _projectID.ToString & "," & _documentID.ToString & ",1,2);"
            Page.RegisterStartupScript("toShowFile", ClientScript.ShowScript("function toShowFile(e) { var strLoc='getAttachment.aspx?ProjectID=" & _projectID.ToString & "&DocumentID=" & _documentID.ToString & "&fileName=' + e.file   ;	 document.nosee.location.href=strLoc; }"))

            If Not IsPostBack Then



                BindCategories()

                ' Load project with _projID when project id exists in the QueryString
                If _projectID = 0 Then

                    Session("m_ErrorMessage") = "系統錯誤，無法存取群組資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If

                If _documentID <> 0 Then
                    BindInfo()
                End If

            End If
        End Sub
        Private Sub doUpdateFilePath()
            _strWebPath = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Documents/" & _projectID & "/" & _documentID & "/"
            _strTempFilePath = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentSrvPath) & "Temp\" & PMSecurity.GetUserID & "\"

        End Sub
        Private Sub BindCategories()
            ddlCategory.DataSource = PMDocCategories.GetCategorys
            ddlCategory.DataTextField = "Name"
            ddlCategory.DataValueField = "CategoryID"
            ddlCategory.DataBind()
        End Sub

        Private Sub BindInfo()

            Dim doc As New PMDocuments(_documentID)
            doc.Load()
            _projectID = doc.ProjectID
            txtTitle.Text = doc.Title
            txtDescription.Text = doc.Description

            ddlCategory.SelectedValue = doc.CategoryID


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

            txtFilename.Text = String.Empty
            Dim Attachment As PMAttachment
            For Each Attachment In doc.Attachments
                lblAttachment.Text += showAttachment(Attachment.AttachmentName)

                If txtFilename.Text.Length > 0 Then txtFilename.Text += ","
                txtFilename.Text += Attachment.AttachmentName
            Next Attachment

        End Sub 'BindInfo



        Private Function showAttachment(ByVal filename As String) As String

            Return PMAttachment.showFile(filename, 0, _strWebPath, True, 0)

        End Function 'BindAttachment

        Private Sub SaveDocument()
            Dim doc As New PMDocuments(_documentID)
            doc.Title = txtTitle.Text
            doc.ProjectID = _projectID
            doc.Description = txtDescription.Text
            doc.CategoryID = ddlCategory.SelectedValue
            doc.ModifiedUserID = PMSecurity.GetUserID


            Dim attachmentFileCol As New PMAttachmentCollection
            Dim aryAttacment As Array = Strings.Split(txtFilename.Text, ",")
            Dim filename As String
            For Each filename In aryAttacment
                If Strings.Trim(filename) <> "" Then
                    Dim attachmentFile As New PMAttachment
                    attachmentFile.AttachmentName = filename

                    attachmentFileCol.Add(attachmentFile)
                End If
            Next

            doc.Attachments = attachmentFileCol


            If doc.Save() Then
                _documentID = doc.DocumentID
                doUpdateFilePath()
                saveAttachmentFiles()
                BackToUserPage()
            Else
                Session("m_ErrorMessage") = "系統錯誤，無法儲存群組資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If

        End Sub

        Private Sub saveAttachmentFiles()

            Dim TargetPath As String = Server.MapPath(_strWebPath)
            If Not Directory.Exists(TargetPath) Then
                Directory.CreateDirectory(TargetPath)
            End If

            Dim strUplosdFileName As String
            strUplosdFileName = txtFilename.Text
            If Strings.Trim(strUplosdFileName) <> "" Then
                Dim aryFile As Array = Split(strUplosdFileName, ",")
                Dim iFile As String
                For Each iFile In aryFile
                    If Strings.Trim(iFile) <> "" Then
                        If File.Exists(_strTempFilePath & iFile) Then
                            'Response.Write(_strTempFilePath & iFile)
                            'Response.Write(_strWebPath & iFile)
                            File.Copy(_strTempFilePath & iFile, TargetPath & iFile, True)
                        Else
                            'ErrorMessage.Text += "儲存檔案失敗，找不到 " & iFile & " 檔案。"
                        End If
                    End If

                Next
            End If
        End Sub


        Private Sub BackToUserPage()
            Response.Redirect([String].Format("DocumentDetail.aspx?ProjectID={0}&DocumentID={1}&index=-1&projectIndex={2}", _projectID, _documentID, TabItem.ProjectTabIndex.ProjectDocument), False)
        End Sub

        Private Sub lbtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnCancel.Click
            BackToUserPage()
        End Sub

        Private Sub lbtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnSave.Click
            SaveDocument()
        End Sub


    End Class

End Namespace


