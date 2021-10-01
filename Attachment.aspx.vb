Imports System
Imports System.IO
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace HsinYi.EIP.ProjectManagement.Web
    Public Class Attachment
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
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

        '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
        '請勿刪除或移動它。
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
            '請勿使用程式碼編輯器進行修改。
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

            '在這裡放置使用者程式碼以初始化網頁
            _userID = PMSecurity.GetUserID
            _blnIsWrite = (Request("mode") = 1)

            If Request.QueryString("AttachmentType") Is Nothing Or Request.QueryString("ProjectID") Is Nothing Then
                Session("m_ErrorMessage") = "系統錯誤，無法存取檔案資訊。"
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
                Session("m_ErrorMessage") = "系統拒絕存取，您沒有權限存修改專案資訊。"
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
                    Session("m_ErrorMessage") = "系統錯誤，無法存取檔案資訊。"
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


#Region " 處理上傳檔案 "
        '複製記錄檔案到暫存資料夾
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
        '顯示暫存資料夾檔案
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
                Page.RegisterStartupScript("alert", ClientScript.ShowAlertMsg("請選擇檔案!!"))
                Exit Sub
            End If

            UploadFiles()

        End Sub

        Private Sub UploadFiles()
            Dim strFileName As String
            Dim strFileSize As String
            Dim intFileSize As String

            '取得上傳檔名
            strFileName = Path.GetFileName(File1.PostedFile.FileName)


            '取得檔案大小
            intFileSize = File1.PostedFile.ContentLength
            'If intFileSize < 1024 Then
            '    strFileSize = intFileSize & "位元組"
            'Else
            '    strFileSize = Int(intFileSize / 1024).ToString & " KB"
            'End If


            '開啟站存目錄
            If Not Directory.Exists(_strTempFilePath) Then
                Directory.CreateDirectory(_strTempFilePath)
            End If


            '儲存暫存檔案
            File1.PostedFile.SaveAs(_strTempFilePath & strFileName)

            '更新暫存檔案清單
            If Strings.Trim(txtFilename.Text) <> "" Then
                txtFilename.Text += ","
            End If

            txtFilename.Text += strFileName

            '更新暫存檔案大小清單
            If Strings.Trim(txtFileSize.Text) <> "" Then
                txtFileSize.Text += ","
            End If

            txtFileSize.Text += strFileName & "_" & intFileSize.ToString

            '顯示檔案
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

