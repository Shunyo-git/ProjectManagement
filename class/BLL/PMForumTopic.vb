Imports System
Imports System.Data
Imports System.Text
Imports System.Configuration
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.DataAccessLayer

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
    '****************************************************************************
    '
    ' Event Class
    '
    ' The Event class represents a list of Project Events.
    '
    '****************************************************************************


    Public Class PMForumTopic
        Private _topicID As Integer
        Private _projectID As Integer
        Private _subject As String
        Private _content As String
        Private _replyTo As Integer
        Private _documentID As Integer

        Private _createdUserID As Integer
        Private _createdDate As Date
        Private _modifiedUserID As Integer
        Private _modifiedDate As Date
        Private _isDel As Boolean
        Private _LastPostUserID As Integer
        Private _LastPostDate As Date
        Private _Replies As Integer
        Private _Attachments As PMAttachmentCollection

#Region "  專案屬性存取 "
        Public Property LastPostDate() As Date
            Get
                Return _LastPostDate
            End Get
            Set(ByVal Value As Date)
                _LastPostDate = Value
            End Set
        End Property

        Public Property LastPostUserID() As Integer
            Get
                Return _LastPostUserID
            End Get
            Set(ByVal Value As Integer)
                _LastPostUserID = Value
            End Set
        End Property
         

        Public Property Replies() As Integer
            Get
                Return _Replies
            End Get
            Set(ByVal Value As Integer)
                _Replies = Value
            End Set
        End Property
        Public Property TopicID() As Integer
            Get
                Return _topicID
            End Get
            Set(ByVal Value As Integer)
                _topicID = Value
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
        Public Property Subject() As String
            Get
                Return _subject

            End Get
            Set(ByVal Value As String)
                _subject = Value
            End Set
        End Property
        Public Property Content() As String
            Get
                Return _content

            End Get
            Set(ByVal Value As String)
                _content = Value
            End Set
        End Property
        Public Property ReplyTo() As Integer
            Get
                Return _replyTo

            End Get
            Set(ByVal Value As Integer)
                _replyTo = Value
            End Set
        End Property
        Public Property DocumentID() As Integer
            Get
                Return _documentID

            End Get
            Set(ByVal Value As Integer)
                _documentID = Value
            End Set
        End Property
        Public Property Attachments() As PMAttachmentCollection
            Get
                Return _Attachments
            End Get
            Set(ByVal Value As PMAttachmentCollection)
                _Attachments = Value
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

        Public ReadOnly Property CreatedUserName() As String
            Get
                Return EIPSysSecurity.clsUser.getUserName(_createdUserID)
            End Get

        End Property
        Public ReadOnly Property ModifiedUserName() As String
            Get
                Return EIPSysSecurity.clsUser.getUserName(_modifiedUserID)
            End Get

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

        Public Sub New()
        End Sub 'New

        Public Sub New(ByVal TopicID As Integer)
            _topicID = TopicID
        End Sub 'New

        Public Sub New(ByVal projectID As Integer, ByVal subject As String, ByVal content As String, ByVal replyTo As Integer, ByVal documentID As Integer, ByVal modifiedUserID As Integer)
            _projectID = projectID
            _subject = subject
            _content = content
            _replyTo = replyTo
            _documentID = documentID
            _modifiedUserID = modifiedUserID

        End Sub 'New

        '*********************************************************************
        '
        ' Retrieves the list of all the Topics
        '
        '*********************************************************************
        Public Overloads Shared Function GetTopics() As PMForumTopicCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListAllTopics")

            ' Separate Data into a Collection of Topics
            Dim topics As New PMForumTopicCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim topic As New PMForumTopic
                    topic.TopicID = Convert.ToInt32(r("TopicID"))
                    topic.ProjectID = Convert.ToInt32(r("ProjectID"))
                    topic.Subject = r("Subject").ToString()
                    topic.Content = r("Content").ToString()
                    topic.ReplyTo = Convert.ToInt32(r("ReplyTo"))
                    topic.DocumentID = Convert.ToInt32(r("DocumentID"))
                    topic.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    topic.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    topic.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    topic.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))
                    If IsDBNull(r("LastPostDate")) Then
                        topic.LastPostDate = Convert.ToDateTime(r("CreatedDate"))
                        topic.LastPostUserID = Convert.ToInt32(r("LastPostUserID"))
                    Else
                        topic.LastPostDate = Convert.ToDateTime(r("LastPostDate"))
                        topic.LastPostUserID = Convert.ToInt32(r("CreatedUserID"))
                    End If
                    topic.Replies = Convert.ToInt32(r("Replies"))
                    topics.Add(topic)
                Next r

            End If
            Return topics
        End Function 'GetTopics

        Public Overloads Shared Function GetTopics(ByVal projectID As Integer) As PMForumTopicCollection
            'Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListEvents", projectID)
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListTopics", projectID)

            ' Separate Data into a Collection of projects
            Dim topics As New PMForumTopicCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim topic As New PMForumTopic
                    topic.TopicID = Convert.ToInt32(r("TopicID"))
                    topic.ProjectID = Convert.ToInt32(r("ProjectID"))
                    topic.Subject = r("Subject").ToString()
                    topic.Content = r("Content").ToString()
                    topic.ReplyTo = Convert.ToInt32(r("ReplyTo"))
                    topic.DocumentID = Convert.ToInt32(r("DocumentID"))
                    topic.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    topic.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    topic.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    topic.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))
                    If IsDBNull(r("LastPostDate")) Then
                        topic.LastPostDate = Convert.ToDateTime(r("CreatedDate"))
                        topic.LastPostUserID = Convert.ToInt32(r("CreatedUserID"))
                    Else
                        topic.LastPostDate = Convert.ToDateTime(r("LastPostDate"))
                        topic.LastPostUserID = Convert.ToInt32(r("LastPostUserID"))
                    End If
                    topic.Replies = Convert.ToInt32(r("Replies"))

                    topics.Add(topic)
                Next r

            End If
            Return topics
        End Function 'GetTopics

        Public Overloads Shared Function GetTopics(ByVal UserID As Integer, ByVal RoleID As Integer) As PMForumTopicCollection
            'Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListEvents", projectID)
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListMyTopics", UserID, RoleID)

            ' Separate Data into a Collection of projects
            Dim topics As New PMForumTopicCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim topic As New PMForumTopic
                    topic.TopicID = Convert.ToInt32(r("TopicID"))
                    topic.ProjectID = Convert.ToInt32(r("ProjectID"))
                    topic.Subject = r("Subject").ToString()
                    topic.Content = r("Content").ToString()
                    topic.ReplyTo = Convert.ToInt32(r("ReplyTo"))
                    topic.DocumentID = Convert.ToInt32(r("DocumentID"))
                    topic.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    topic.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    topic.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    topic.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))

                    If IsDBNull(r("LastPostDate")) Then
                        topic.LastPostDate = Convert.ToDateTime(r("CreatedDate"))
                        topic.LastPostUserID = Convert.ToInt32(r("CreatedUserID"))

                    Else
                        topic.LastPostDate = Convert.ToDateTime(r("LastPostDate"))
                        topic.LastPostUserID = Convert.ToInt32(r("LastPostUserID"))
                    End If
                    topic.Replies = Convert.ToInt32(r("Replies"))

                    topics.Add(topic)
                Next r

            End If
            Return topics
        End Function 'GetTopics
        Public Shared Function GetReplys(ByVal TopicID As Integer) As PMForumTopicCollection
            'Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListEvents", projectID)
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListReplies", TopicID)

            ' Separate Data into a Collection of projects
            Dim topics As New PMForumTopicCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim topic As New PMForumTopic
                    topic.TopicID = Convert.ToInt32(r("TopicID"))
                    topic.ProjectID = Convert.ToInt32(r("ProjectID"))
                    topic.Subject = r("Subject").ToString()
                    topic.Content = r("Content").ToString()
                    topic.ReplyTo = Convert.ToInt32(r("ReplyTo"))
                    topic.DocumentID = Convert.ToInt32(r("DocumentID"))
                    topic.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    topic.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    topic.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    topic.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))
                    'topic.LastPostDate = Convert.ToDateTime(r("LastPostDate"))
                    'topic.Replies = Convert.ToInt32(r("Replies"))
                    'topic.LastPostUserID = Convert.ToInt32(r("LastPostUserID"))

                    topics.Add(topic)
                Next r

            End If
            Return topics
        End Function 'GetReplys

        '*********************************************************************
        '
        ' Populates the object with the Topic info from the database.
        '
        '*********************************************************************

        Public Function Load() As Boolean
            ' The Get Projects stored procedure returns a dataset with 3 tables: Projects, Members, and Categories.
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetTopic", _topicID)

            If ds.Tables(0).Rows.Count < 1 Then
                Return False
            End If

            _documentID = Convert.ToInt32(ds.Tables(0).Rows(0)("DocumentID"))
            _projectID = Convert.ToInt32(ds.Tables(0).Rows(0)("ProjectID"))
            _replyTo = Convert.ToInt32(ds.Tables(0).Rows(0)("replyTo"))
            _subject = ds.Tables(0).Rows(0)("Subject").ToString()
            _content = ds.Tables(0).Rows(0)("Content").ToString()

            _createdUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("CreatedUserID"))
            _createdDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("CreatedDate"))
            _modifiedUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("ModifiedUserID"))
            _modifiedDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("ModifiedDate"))

            If _replyTo = 0 Then

                If IsDBNull(ds.Tables(0).Rows(0)("LastPostDate")) Then
                    _LastPostDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("CreatedDate"))
                    _LastPostUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("CreatedUserID"))
                Else
                    _LastPostDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("LastPostDate"))
                    _LastPostUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("LastPostUserID"))
                End If

                _Replies = Convert.ToInt32(ds.Tables(0).Rows(0)("Replies"))


            End If

            _Attachments = New PMAttachmentCollection
            Dim row As DataRow
            For Each row In ds.Tables(1).Rows
                Dim Attachment As New PMAttachment

                Attachment.AttachmentID = Convert.ToInt32(row("AttachmentID"))
                Attachment.AttachmentName = Convert.ToString(row("AttachmentName"))
                Attachment.ModifiedUserID = Convert.ToInt32(row("ModifiedUserID"))
                Attachment.ModifiedDate = Convert.ToDateTime(row("ModifiedDate"))
                Attachment.CreatedUserID = Convert.ToInt32(row("CreatedUserID"))
                Attachment.CreatedDate = Convert.ToDateTime(row("CreatedDate"))
                _Attachments.Add(Attachment)
            Next row


            Return True
        End Function 'Load

        '*********************************************************************
        '
        ' Remove static method
        '
        ' Removes a project from database.
        ' This method relis on cascading delete to delete the associated categories,
        ' projectmembers, and time entries
        '
        '*********************************************************************

        Public Shared Sub Remove(ByVal topicID As Integer)
            SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_DeleteTopic", topicID)
        End Sub 'Remove
        '*********************************************************************
        '
        ' Calls Insert or Save based on the _projectID
        '
        '*********************************************************************
        Public Function showValue() As String
            Dim strTmp As String

            strTmp = _topicID.ToString & "," & _projectID.ToString & "," & _replyTo & "," & _documentID & "," & _subject & "," & _content & "," & PMAttachment.getAttachments(_Attachments) & "," & _modifiedUserID
            Return strTmp
        End Function
        Public Function Save() As Boolean
            If _topicID = 0 Then
                Return Insert()
            Else
                If _topicID > 0 Then
                    Return Update()
                Else
                    _topicID = 0
                    Return False
                End If
            End If
        End Function 'Save
        '*********************************************************************
        '
        ' Inserts a project along with its members and categories into the databse
        '
        '*********************************************************************

        Private Function Insert() As Boolean

            Dim blnResult As Boolean

            Try
                _topicID = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_AddTopic", _projectID, _replyTo, _subject, _content, PMAttachment.getAttachments(_Attachments), _modifiedUserID))

                blnResult = True

            Catch ex As Exception
                Throw ex
                blnResult = False
            End Try


            Return blnResult

        End Function 'Insert
        '*********************************************************************
        '
        ' Updates a project along with its members and categories.
        '
        '*********************************************************************

        Private Function Update() As Boolean
            Dim blnResult As Boolean

            Try
                SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_UpdateTopic", _topicID, _projectID, _replyTo, _documentID, _subject, _content, PMAttachment.getAttachments(_Attachments), _modifiedUserID)
                blnResult = True
            Catch ex As Exception
                Throw ex
                blnResult = False
            End Try

            Return blnResult
        End Function 'Update


    End Class

End Namespace

