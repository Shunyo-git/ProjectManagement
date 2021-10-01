Imports System.IO
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class getAttachment
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub

        '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
        '請勿刪除或移動它。
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
            '請勿使用程式碼編輯器進行修改。
            InitializeComponent()
        End Sub

#End Region

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            '在這裡放置使用者程式碼以初始化網頁



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
                Session("m_ErrorMessage") = "系統拒絕存取，您沒有權限存取此專案資訊。" '"您非本專案成員，因此無權限存取此專案資訊。"
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

