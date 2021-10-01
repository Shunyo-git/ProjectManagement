Imports System
Imports System.IO
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace HsinYi.EIP.ProjectManagement.Web
    Public Class Attachment
        Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

        '���� Web Form �]�p�u��һݪ��I�s�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents File1 As System.Web.UI.HtmlControls.HtmlInputFile
        Protected WithEvents btnUpload As System.Web.UI.WebControls.Button
        Protected WithEvents txtFilename As System.Web.UI.WebControls.TextBox
        Protected WithEvents lblFileNameList As System.Web.UI.WebControls.Label
        Protected WithEvents txtFileSize As System.Web.UI.WebControls.TextBox
        Protected WithEvents lblHidden As System.Web.UI.WebControls.Label
        Protected WithEvents PanelReadonly As System.Web.UI.WebControls.Panel
        Protected WithEvents PanelWrite As System.Web.UI.WebControls.Panel
        Protected WithEvents PanelUpload As System.Web.UI.WebControls.Panel
        Protected WithEvents btnSave As System.Web.UI.WebControls.Button
        Protected WithEvents PanelView As System.Web.UI.WebControls.Panel

        '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
        '�ФŧR���β��ʥ��C
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
            '�ФŨϥε{���X�s�边�i��ק�C
            InitializeComponent()
        End Sub

#End Region
        Private _userID As Integer = 0
        Private _strWebPath As String 'Web Path
        Private _strTempFilePath As String 'Server Path
        Private _intRecordID As Integer
        Private _strTableName As String
        Private _strColnumName As String
        Private _strFilePath As String 'Server Path
        Private _blnIsWrite As Boolean
        Private _attachmentType As PMAttachment.AttachmentTypes
        Private _projectID As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

            '�b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���
            _userID = PMSecurity.GetUserID
            _blnIsWrite = (Request("mode") = 1)

            If Request.QueryString("AttachmentType") Is Nothing Or Request.QueryString("ProjectID") Is Nothing Then
                Session("m_ErrorMessage") = "�t�ο��~�A�L�k�s���ɮ׸�T�C"
                Response.Redirect("AlertMessage.aspx?Index=-1", False)
            Else
                _attachmentType = Request.QueryString("AttachmentType")
                _projectID = Integer.Parse(Request.QueryString("ProjectID"))
            End If

            If Not Request("AttachmentFiles") Is Nothing Then
                txtFilename.Text = Request("AttachmentFiles")
            End If

            If Request("RecordID") Is Nothing Then
                _intRecordID = 0
            Else
                _intRecordID = Integer.Parse(Request("RecordID"))
            End If

            Page.RegisterStartupScript("toShowFile", ClientScript.ShowScript("function toShowFile(e) { var strLoc='" & ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Temp/" & _userID & "/'+e.file ;	window.open(strLoc);	}"))

            If Not Roles.isEnabledModifyProject(_projectID) And Not Roles.isProjectMember(_projectID) Then
                Session("m_ErrorMessage") = "�t�Ωڵ��s���A�z�S���v���s�ק�M�׸�T�C"
                Response.Redirect("AlertMessage.aspx?Index=-1", False)
            End If



            If _blnIsWrite Then
                ' PanelReadonly.Visible = False
                ' PanelWrite.Visible = True
                PanelUpload.Visible = True
            Else
                ' PanelReadonly.Visible = True
                'PanelWrite.Visible = False
                PanelUpload.Visible = False
            End If

            
            Select Case _attachmentType
                Case PMAttachment.AttachmentTypes.EventAttachment
                    _strWebPath = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Events/" & _projectID & "/" & _intRecordID & "/"

                Case PMAttachment.AttachmentTypes.DocumentAttachment
                    _strWebPath = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Documents/" & _projectID & "/" & _intRecordID & "/"

                Case Else
                    Session("m_ErrorMessage") = "�t�ο��~�A�L�k�s���ɮ׸�T�C"
                    Response.Redirect("AlertMessage.aspx?Index=-1", False)
            End Select

            _strFilePath = Server.MapPath(_strWebPath)
            _strTempFilePath = ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentSrvPath) & "Temp\" & _userID & "\"



            lblHidden.Text = "<input type=hidden name=RecordID value=" & Request("RecordID") & ">"
            lblHidden.Text += "<input type=hidden name=AttachmentTypes value=" & Request("AttachmentTypes") & ">"
            lblHidden.Text += "<input type=hidden name=mode value=" & Request("mode") & ">"

            btnSave.Attributes.Add("onclick", "SaveAttachment();")
            If Not Page.IsPostBack Then
                'If _intRecordID > 0 Then

                getSourceFolderFiles()
                'End If

            End If


            'If Strings.Trim(lblFileNameList.Text) = "" Then
            '    PanelView.Visible = False
            'Else
            '    PanelView.Visible = True
            'End If

        End Sub


#Region " �B�z�W���ɮ� "
        '�ƻs�O���ɮר�Ȧs��Ƨ�
        Private Sub getSourceFolderFiles()

            If Directory.Exists(_strTempFilePath) Then
                If Strings.Trim(txtFilename.Text) = "" Then
                    'Response.Write(_strTempFilePath)
                    PMAttachment.delDirectoryFiles(_strTempFilePath)
                Else
                    If Directory.Exists(_strFilePath) Then
                        Dim di As New DirectoryInfo(_strFilePath)
                        Dim fiArr As FileInfo() = di.GetFiles()
                        Dim f As FileInfo

                        For Each f In fiArr
                            f.CopyTo(_strTempFilePath & f.Name, True)

                        Next f
                    End If
                End If
            Else
                Directory.CreateDirectory(_strTempFilePath)
            End If

 
            chkUploadFileShow(_strTempFilePath)
        End Sub
        '��ܼȦs��Ƨ��ɮ�
        Private Sub chkUploadFileShow(ByVal strPatn As String)
            Dim iFile As FileInfo
            Dim strFile As String
            Dim intFileSize As Integer

            Dim di As New DirectoryInfo(strPatn)
            Dim fiArr As FileInfo() = di.GetFiles()
            Dim f As FileInfo

            For Each f In fiArr
                strFile = f.Name
                intFileSize = f.Length
                If Strings.Trim(txtFilename.Text) <> "" Then
                    txtFilename.Text += ","
                End If
                If Strings.InStr(txtFilename.Text, f.Name) < 0 Then
                    txtFilename.Text += f.Name
                End If

                lblFileNameList.Text += PMAttachment.showFile(Trim(strFile), intFileSize, _strWebPath, _blnIsWrite, 0)
            Next

            
        End Sub


        Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click

            If Strings.Trim(File1.PostedFile.FileName) = "" Then
                Page.RegisterStartupScript("alert", ClientScript.ShowAlertMsg("�п���ɮ�!!"))
                Exit Sub
            End If

            UploadFiles()

        End Sub

        Private Sub UploadFiles()
            Dim strFileName As String
            Dim strFileSize As String
            Dim intFileSize As String

            '���o�W���ɦW
            strFileName = Path.GetFileName(File1.PostedFile.FileName)


            '���o�ɮפj�p
            intFileSize = File1.PostedFile.ContentLength
            'If intFileSize < 1024 Then
            '    strFileSize = intFileSize & "�줸��"
            'Else
            '    strFileSize = Int(intFileSize / 1024).ToString & " KB"
            'End If


            '�}�ү��s�ؿ�
            If Not Directory.Exists(_strTempFilePath) Then
                Directory.CreateDirectory(_strTempFilePath)
            End If


            '�x�s�Ȧs�ɮ�
            File1.PostedFile.SaveAs(_strTempFilePath & strFileName)

            '��s�Ȧs�ɮײM��
            If Strings.Trim(txtFilename.Text) <> "" Then
                txtFilename.Text += ","
            End If

            txtFilename.Text += strFileName

            '��s�Ȧs�ɮפj�p�M��
            If Strings.Trim(txtFileSize.Text) <> "" Then
                txtFileSize.Text += ","
            End If

            txtFileSize.Text += strFileName & "_" & intFileSize.ToString

            '����ɮ�
            lblFileNameList.Text += PMAttachment.showFile(strFileName, intFileSize, _strWebPath, _blnIsWrite, 0)
            'chkUploadFileShow(m_strTempFilePath)
        End Sub






        'Public Function chkRecordFileShow(ByVal intRecordID As Integer, ByVal strFileName As String) As String

        '    Dim strReplyFilePath As String
        '    Dim strUplosdFileName As String
        '    Dim strHtml As String
        '    strReplyFilePath = "/pub/pub/u00021/T_00917/" & intRecordID & "/"

        '    If Not IsDBNull(strFileName) Then
        '        strUplosdFileName = strFileName
        '    Else
        '        strUplosdFileName = ""
        '    End If


        '    If Strings.Trim(strUplosdFileName) <> "" Then
        '        Dim aryFile As Array = Split(strUplosdFileName, ",")
        '        Dim iFile As String
        '        Dim aryFileSize As Array = Split(txtFileSize.Text, ",")
        '        Dim i As Int16 = 0

        '        For Each iFile In aryFile
        '            strHtml += getFile(Trim(iFile), aryFileSize(i), strReplyFilePath, False)
        '        Next

        '    End If

        '    Return strHtml
        'End Function

        Private Sub saveAttachmentFiles()


            If Not Directory.Exists(_strFilePath) Then
                Directory.CreateDirectory(_strFilePath)
            End If

            Dim strUplosdFileName As String
            strUplosdFileName = txtFilename.Text
            If Strings.Trim(strUplosdFileName) <> "" Then
                Dim aryFile As Array = Split(strUplosdFileName, ",")
                Dim iFile As String
                For Each iFile In aryFile
                    If File.Exists(_strTempFilePath & iFile) Then
                        File.Move(_strTempFilePath & iFile, _strFilePath & iFile)
                    End If

                    ' clsStudyGroup.addStudyFile(m_intRecordID, iFile)
                Next
            End If
        End Sub

#End Region

        Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

            'Select Case _strTableName
            '    Case "u00021"

            '        '  clsStudyGroup.delStudyFiles(m_intRecordID)
            'End Select
            If _intRecordID > 0 Then
                PMAttachment.delDirectoryFiles(_strFilePath)
                saveAttachmentFiles()
            End If


            Response.Write(ClientScript.ShowScript("window.close();"))
        End Sub

        Private Sub PrintLn(ByVal Value As String)
            Response.Write(Value & "<BR>")
        End Sub
    End Class
End Namespace

