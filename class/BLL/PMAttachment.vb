Imports System
Imports System.Data
Imports System.Configuration
Imports System.Text
Imports System.IO
Imports EIPSysSecurity

Imports Hsinyi.EIP.ProjectManagement.DataAccessLayer

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
    Public Class PMAttachment
        Private _attachmentID As Integer
        Private _eventID As Integer
        Private _projectID As Integer
        Private _attachmentName As String
        Private _createdUserID As Integer
        Private _createdDate As System.DateTime
        Private _modifiedUserID As Integer
        Private _modifiedDate As System.DateTime
        Private _isDel As Boolean

        Private _strWebPath As String 'Web Path
        Private _strTempFilePath As String 'Server Path
        Private _strFilePath As String 'Server Path
        Private _intRecordID As Integer = 0

        Private _attachmentType As AttachmentTypes


#Region "  專案屬性存取 "
        Public Property RecordID() As Integer
            Get
                Return _intRecordID
            End Get
            Set(ByVal Value As Integer)
                _intRecordID = Value
            End Set
        End Property

        Public Property AttachmentID() As Integer
            Get
                Return _attachmentID
            End Get
            Set(ByVal Value As Integer)
                _attachmentID = Value
            End Set
        End Property

        Public Property AttachmentName() As String
            Get
                Return _attachmentName
            End Get
            Set(ByVal Value As String)
                _attachmentName = Value
            End Set
        End Property

        Public Property EventID() As Integer
            Get
                Return _eventID
            End Get
            Set(ByVal Value As Integer)
                _eventID = Value
            End Set
        End Property

        Public Property ProjectID() As Integer
            Get
                Return _projectID
            End Get
            Set(ByVal Value As Integer)
                _projectID = Value
            End Set
        End Property

        Public Property CreatedUserID() As Integer
            Get
                Return _createdUserID
            End Get
            Set(ByVal Value As Integer)
                _createdUserID = Value
            End Set
        End Property

        Public Property ModifiedUserID() As Integer
            Get
                Return _modifiedUserID
            End Get
            Set(ByVal Value As Integer)
                _modifiedUserID = Value
            End Set
        End Property



        Public Property CreatedDate() As DateTime
            Get
                Return _createdDate
            End Get

            Set(ByVal Value As DateTime)
                _createdDate = Value
            End Set
        End Property

        Public Property ModifiedDate() As DateTime
            Get
                Return _modifiedDate
            End Get
            Set(ByVal Value As DateTime)
                _modifiedDate = Value
            End Set
        End Property

        Public Property isDel() As Boolean
            Get
                Return _isDel
            End Get
            Set(ByVal Value As Boolean)
                _isDel = Value
            End Set
        End Property

#End Region

        Public Enum AttachmentTypes
            InitValue
            EventAttachment
            DocumentAttachment
        End Enum 'AttachmentTypes


        Public Sub New()
        End Sub 'New

        Public Sub New(ByVal AttachmentID As Integer)
            _attachmentID = AttachmentID
        End Sub 'New
        Public Sub load()

        End Sub
         
        Public Shared Function GetAttachmentName(ByVal AttachmentID As Integer) As String
            Return Convert.ToString(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetAttachmentName", AttachmentID))
        End Function
        Public Shared Function getImage(ByVal strFilename As String) As String
            Dim i
            Dim strExtension
            Dim strImg
            Dim objStrings As New clsStrings
            strExtension = objStrings.subString(strFilename, ".", 1, 1, 1)
            Select Case LCase(strExtension)
                Case "bmp"
                    strImg = "<img src='/images/bmp.gif'>"
                Case "gif"
                    strImg = "<img src='/images/gif.gif'>"
                Case "jpg"
                    strImg = "<img src='/images/jpg.gif'>"
                Case "ppt"
                    strImg = "<img src='/images/icon-doc-ppt.gif'>"
                Case "doc"
                    strImg = "<img src='/images/icon-doc-word.gif'>"
                Case "xls"
                    strImg = "<img src='/images/icon-doc-excel.gif'>"
                Case "asp"
                    strImg = "<img src='/images/icon-doc-asp.gif'>"
                Case "zip"
                    strImg = "<img src='/images/zip.gif'>"
                Case "rar"
                    strImg = "<img src='/images/rar.gif'>"
                Case Else
                    strImg = "<img src='/images/icon-doc.gif'>"
            End Select
            getImage = strImg
        End Function
        Public Shared Function showFile(ByVal strFilename As String, ByVal intFileSize As Integer, ByVal strPath As String, ByVal blnIsEnabledDelBtn As Boolean, ByVal intRecoedID As Integer) As String
            Dim strType
            Dim strImage
            Dim strTemp = String.Empty
            Dim objFile

            If InStr(strFilename, ".") > 0 Then
                Dim strFileSize As String

                If intFileSize <= 0 Then
                    strFileSize = ""
                ElseIf intFileSize < 1024 Then
                    strFileSize = " (" & intFileSize.ToString & "位元組) "
                Else
                    strFileSize = " (" & Int(intFileSize / 1024).ToString & " KB) "
                End If

                strImage = PMAttachment.getImage(strFilename)
                If blnIsEnabledDelBtn Then
                    strTemp = strTemp & "<table class=fieldset cellpadding=0 cellspacing=3 border=0 width=90% id='f_" & strFilename & "'" & _
                                         "   >" & _
                                         "	<tr>" & _
                                         "		<td> " & strImage & "<a href=""#"" onclick='toShowFile(this)' recordID='" & intRecoedID.ToString & "' file='" & (strFilename) & "'  >" & strFilename & "</a>  " & strFileSize & "  </td>" & _
                                         "		<td align=right>" & _
                                         "		<input type=button value='刪除' class=button recordID='" & intRecoedID.ToString & "'  file='" & strFilename & "' onclick='toClear(this)' style='WIDTH:70px'>" & _
                                         "		</td>" & _
                                         "	</tr>" & _
                                         "</table>"
                Else
                    strTemp = strTemp & "<table class=fieldset cellpadding=0 cellspacing=3 border=0 width=90% id='f_" & strFilename & "'" & _
                                         "  file='" & strFilename & "' >" & _
                                         "	<tr>" & _
                                         "		<td> " & strImage & "<a href=""#"" onclick='toShowFile(this)' recordID='" & intRecoedID.ToString & "' file='" & (strFilename) & "'  >" & strFilename & "</a>  " & strFileSize & "  </td>" & _
                                          "	</tr>" & _
                                         "</table>"
                End If
            End If


            Return strTemp
        End Function

        '清除資料夾檔案
        Public Shared Function delDirectoryFiles(ByVal strPath As String) As String

            Dim iFile As String
            Dim strFile As String
            If Directory.Exists(strPath) Then
                For Each iFile In Directory.GetFiles(strPath)
                    strFile += " file : " & iFile & "<BR>"
                    'If File.Exists(iFile) Then
                    strFile += " Delete : " & iFile & "<BR>"
                    File.Delete(iFile)
                    'End If
                Next
                'Directory.Delete(strPath)
            End If

        End Function

        Public Shared Function getAttachments(ByVal _Attachments As PMAttachmentCollection) As String



            Dim Attachments As New StringBuilder(_Attachments.Count)
            If _Attachments.Count <= 0 Then
                Return String.Empty
            End If

            Dim index As Integer = 1
            Dim file As PMAttachment
            For Each file In _Attachments
                Attachments.Append(file.AttachmentName)

                ' Don't append separator if this is the last item
                If index <> _Attachments.Count Then
                    Attachments.Append(",")
                End If
                index += 1
            Next file
            Return Attachments.ToString()

        End Function 'getAttachments

    End Class 'PMAttachment
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

