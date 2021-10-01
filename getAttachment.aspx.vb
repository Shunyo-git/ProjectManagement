Imports System.IO
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class getAttachment
        Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

        '���� Web Form �]�p�u��һݪ��I�s�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub

        '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
        '�ФŧR���β��ʥ��C
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
            '�ФŨϥε{���X�s�边�i��ק�C
            InitializeComponent()
        End Sub

#End Region

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            '�b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���



            Dim fileName As String
            Dim EventID As Integer
            Dim DocumentID As Integer
            Dim _projectID As Integer
            Dim fileSpec As String

            If Request("fileName") Is Nothing Or Request("ProjectID") Is Nothing Or (Request("DocumentID") Is Nothing And Request("EventID") Is Nothing) Then
                Page.RegisterStartupScript("CloseWin", "self.close();")
                Exit Sub
            End If

            fileName = Request("fileName")
            EventID = IIf(Request("EventID") Is Nothing, 0, Convert.ToInt32(Request("EventID")))
            DocumentID = IIf(Request("DocumentID") Is Nothing, 0, Convert.ToInt32(Request("DocumentID")))
            _projectID = IIf(Request("ProjectID") Is Nothing, 0, Convert.ToInt32(Request("ProjectID")))

            If Not Roles.isEnabledViewProject(_projectID) Then
                Session("m_ErrorMessage") = "�t�Ωڵ��s���A�z�S���v���s�����M�׸�T�C" '"�z�D���M�צ����A�]���L�v���s�����M�׸�T�C"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If



            If DocumentID > 0 Then
                fileSpec = Server.MapPath(ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Documents/" & _projectID & "/" & DocumentID & "/" & fileName)
            ElseIf EventID > 0 Then
                fileSpec = Server.MapPath(ConfigurationSettings.AppSettings(Web.Global.CfgKeyAttachmentWebPath) & "Events/" & _projectID & "/" & EventID & "/" & fileName)
            Else
                Page.RegisterStartupScript("CloseWin", "self.close();")
                Exit Sub
            End If


            Dim MyFileStream As FileStream
            Dim FileSize As Long

            MyFileStream = New FileStream(fileSpec, FileMode.Open)
            FileSize = MyFileStream.Length

            Dim Buffer(CInt(FileSize)) As Byte
            MyFileStream.Read(Buffer, 0, CInt(FileSize))
            MyFileStream.Close()



            Response.Clear()
            Response.AddHeader("Content-Disposition", "attachment;  filename=" + HttpUtility.UrlEncode(fileName))
            Response.ContentType = "application/octet-stream"
            Response.BinaryWrite(Buffer)
            'Response.Charset = "big5"
 
        End Sub


    End Class

End Namespace

