Imports System
Imports System.IO
Imports System.Data
Imports System.Web.UI.WebControls
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class EventModify
        Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

        '���� Web Form �]�p�u��һݪ��I�s�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectNameRequiredfieldvalidator As System.Web.UI.WebControls.RequiredFieldValidator
        Protected WithEvents lbtnCancel As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnSave As System.Web.UI.WebControls.LinkButton
        Protected WithEvents ErrorMessage As System.Web.UI.WebControls.Label
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents Label4 As System.Web.UI.WebControls.Label
        Protected WithEvents lblTel As System.Web.UI.WebControls.Label
        Protected WithEvents passwordlbl As System.Web.UI.WebControls.Label
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents itnDate As System.Web.UI.WebControls.Label
        Protected WithEvents itnTopic As System.Web.UI.WebControls.Label
        Protected WithEvents txtReason As System.Web.UI.WebControls.TextBox
        Protected WithEvents txtAnswer As System.Web.UI.WebControls.TextBox
        Protected WithEvents txtTopic As System.Web.UI.WebControls.TextBox
        Protected WithEvents txtEventDate As DatePicker.iX.Controls.DatePicker
        Protected WithEvents txtEffect As System.Web.UI.WebControls.TextBox
        Protected WithEvents lblAttachment As System.Web.UI.WebControls.Label
        Protected WithEvents hlAttachment As System.Web.UI.WebControls.HyperLink
        Protected WithEvents txtFilename As System.Web.UI.WebControls.TextBox

        '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
        '�ФŧR���β��ʥ��C
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
            '�ФŨϥε{���X�s�边�i��ק�C
            InitializeComponent()
        End Sub

#End Region
        Private _projectID As Integer
        Private _eventID As Integer
        Private _strWebPath As String
        Private _strTempFilePath As String

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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

            doUpdateFilePath()
            hlAttachment.NavigateUrl = "javascript:getAttachment(" & _projectID.ToString & "," & _eventID.ToString & ",1,1);"
            Page.RegisterStartupScript("toShowFile", ClientScript.ShowScript("function toShowFile(e) { var strLoc='getAttachment.aspx?ProjectID=" & _projectID.ToString & "&EventID=" & _eventID.ToString & "&fileName=' + e.file   ;	 document.nosee.location.href=strLoc; }"))

            If Not IsPostBack Then

                ' Load project with _projID when project id exists in the QueryString
                If _projectID = 0 Then

                    Session("m_ErrorMessage") = "�t�ο��~�A�L�k�s���s�ո�T�C"  '"�z�D���M�צ����A�]���L�v���s�����M�׸�T�C"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If

                If _eventID <> 0 Then
                    BindInfo()

                End If

            End If
        End Sub
        Private Sub doUpdateFilePath()
            _strWebPath = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Events/" & _projectID & "/" & _eventID & "/"
            _strTempFilePath = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentSrvPath) & "Temp\" & PMSecurity.GetUserID & "\"

        End Sub
        Private Sub BindInfo()
            Dim pjevent As New PMEvent(_eventID)
            pjevent.Load()
            _projectID = pjevent.ProjectID
            txtTopic.Text = pjevent.Topic

            txtEventDate.Text = pjevent.EventDate
            txtReason.Text = pjevent.Reason
            '  txtBackground.Text = pjevent.Background
            txtAnswer.Text = pjevent.Answer
            txtEffect.Text = pjevent.Effect

            'If PMSecurity.GetUserID = pjevent.ModifiedUserID Then
            '    PanelModify.Visible = True
            'Else
            '    '�b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���
            '    If Not Roles.isEnabledModifyProject(_projectID) Then
            '        Session("m_ErrorMessage") = "�t�Ωڵ��s���A�z�S���v���s�ק�M�׸�T�C" '"�z�D���M�צ����A�]���L�v���s�����M�׸�T�C"
            '        Response.Redirect("AlertMessage.aspx?Index=-1", True)
            '        PanelModify.Visible = False
            '    Else
            '        PanelModify.Visible = True
            '    End If
            'End If

            txtFilename.Text = String.Empty
            Dim Attachment As PMAttachment
            For Each Attachment In pjevent.Attachments
                lblAttachment.Text += showAttachment(Attachment.AttachmentName)

                If txtFilename.Text.Length > 0 Then txtFilename.Text += ","
                txtFilename.Text += Attachment.AttachmentName
            Next Attachment

        End Sub 'BindInfo

        Private Function showAttachment(ByVal filename As String) As String

            Return PMAttachment.showFile(filename, 0, _strWebPath, True, 0)

        End Function 'BindAttachment

        Private Sub SaveEvent()
            Dim pjevent As New PMEvent(_eventID)
            pjevent.Topic = txtTopic.Text
            pjevent.EventDate = txtEventDate.Text
            pjevent.Reason = txtReason.Text
            '   pjevent.Background = txtBackground.Text
            pjevent.Answer = txtAnswer.Text
            pjevent.Effect = txtEffect.Text
            pjevent.ModifiedUserID = PMSecurity.GetUserID
            pjevent.ProjectID = _projectID

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
            pjevent.Attachments = attachmentFileCol

            If pjevent.Save() Then
                _eventID = pjevent.EventID
                saveAttachmentFiles()
                BackToDeatilPage()
            Else
                Session("m_ErrorMessage") = "�t�ο��~�A�L�k�x�s�s�ո�T�C"  '"�z�D���M�צ����A�]���L�v���s�����M�׸�T�C"
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
                            File.Move(_strTempFilePath & iFile, _strWebPath & iFile)
                        Else
                            'ErrorMessage.Text += "�x�s�ɮץ��ѡA�䤣�� " & iFile & " �ɮסC"
                        End If
                    End If

                Next
            End If
        End Sub


        Private Sub BackToDeatilPage()
            Response.Redirect([String].Format("EventDetail.aspx?ProjectID={0}&EventID={1}&index=-1&projectIndex={2}", _projectID, _eventID, TabItem.ProjectTabIndex.ProjectEvent), True)
        End Sub

        Private Sub lbtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnCancel.Click
            BackToDeatilPage()
        End Sub

        Private Sub lbtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnSave.Click
            SaveEvent()
        End Sub
    End Class

End Namespace
