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

    Public Class PMEvent
        Private _eventID As Integer
        Private _projectID As Integer
        Private _topic As String
        Private _eventDate As Date
        Private _reason As String
        Private _background As String
        Private _answer As String
        Private _effect As String
        Private _createdUserID As Integer
        Private _createdDate As Date
        Private _modifiedUserID As Integer
        Private _modifiedDate As Date
        Private _isDel As Boolean
        Private _Attachments As PMAttachmentCollection

#Region "  專案屬性存取 "
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

        Public Property Topic() As String
            Get
                Return _topic

            End Get
            Set(ByVal Value As String)
                _topic = Value
            End Set
        End Property
        Public Property Reason() As String
            Get
                Return _reason

            End Get
            Set(ByVal Value As String)
                _reason = Value
            End Set
        End Property
        Public Property Background() As String
            Get
                Return _background

            End Get
            Set(ByVal Value As String)
                _background = Value
            End Set
        End Property
        Public Property Answer() As String
            Get
                Return _answer

            End Get
            Set(ByVal Value As String)
                _answer = Value
            End Set
        End Property
        Public Property Effect() As String
            Get
                Return _effect

            End Get
            Set(ByVal Value As String)
                _effect = Value
            End Set
        End Property
        Public Property EventDate() As DateTime
            Get
                Return _eventDate
            End Get

            Set(ByVal Value As DateTime)
                _eventDate = Value
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

        Public Sub New(ByVal EventID As Integer)
            _eventID = EventID
        End Sub 'New

        Public Sub New(ByVal projectID As Integer, ByVal topic As String, ByVal eventDate As Date, ByVal reason As String, ByVal background As String, ByVal answer As String, ByVal effect As String, ByVal modifiedUserID As Integer)
            _projectID = projectID
            _topic = topic
            _eventDate = eventDate
            _reason = reason
            _background = background
            _answer = answer
            _effect = effect
            _modifiedUserID = modifiedUserID

        End Sub 'New

        '*********************************************************************
        '
        ' Retrieves the list of all the projects
        '
        '*********************************************************************
        Public Overloads Shared Function GetEvents() As PMEventsCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListAllEvents")

            ' Separate Data into a Collection of projects
            Dim events As New PMEventsCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim Pjevent As New PMEvent
                    Pjevent.EventID = Convert.ToInt32(r("EventID"))
                    Pjevent.ProjectID = Convert.ToInt32(r("ProjectID"))
                    Pjevent.EventDate = Convert.ToDateTime(r("EventDate"))
                    Pjevent.Topic = r("Topic").ToString()
                    Pjevent.Reason = r("Reason").ToString()
                    Pjevent.Background = r("Background").ToString()
                    Pjevent.Answer = r("Answer").ToString()
                    Pjevent.Effect = r("Effect").ToString()
                    Pjevent.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    Pjevent.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    Pjevent.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    Pjevent.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))


                    events.Add(Pjevent)
                Next r

            End If
            Return events
        End Function 'GetEvents
        Public Overloads Shared Function GetEvents(ByVal projectID As Integer) As PMEventsCollection
            'Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListEvents", projectID)
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListEvents", projectID)

            ' Separate Data into a Collection of projects
            Dim events As New PMEventsCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim Pjevent As New PMEvent
                    Pjevent.EventID = Convert.ToInt32(r("EventID"))
                    Pjevent.ProjectID = Convert.ToInt32(r("ProjectID"))
                    Pjevent.EventDate = Convert.ToDateTime(r("EventDate"))
                    Pjevent.Topic = r("Topic").ToString()
                    Pjevent.Reason = r("Reason").ToString()
                    Pjevent.Background = r("Background").ToString()
                    Pjevent.Answer = r("Answer").ToString()
                    Pjevent.Effect = r("Effect").ToString()
                    Pjevent.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    Pjevent.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    Pjevent.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    Pjevent.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))


                    events.Add(Pjevent)
                Next r

            End If
            Return events
        End Function 'GetEvents

        '*********************************************************************
        '
        ' Populates the object with the Event info from the database.
        '
        '*********************************************************************

        Public Function Load() As Boolean
            ' The Get Projects stored procedure returns a dataset with 3 tables: Projects, Members, and Categories.
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetEvent", _eventID)

            If ds.Tables(0).Rows.Count < 1 Then
                Return False
            End If


            _eventID = Convert.ToInt32(ds.Tables(0).Rows(0)("EventID"))
            _projectID = Convert.ToInt32(ds.Tables(0).Rows(0)("ProjectID"))
            _eventDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("EventDate"))
            _topic = ds.Tables(0).Rows(0)("Topic").ToString()
            _reason = ds.Tables(0).Rows(0)("Reason").ToString()
            _background = ds.Tables(0).Rows(0)("Background").ToString()
            _answer = ds.Tables(0).Rows(0)("Answer").ToString()
            _effect = ds.Tables(0).Rows(0)("Effect").ToString()
            _createdUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("CreatedUserID"))
            _createdDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("CreatedDate"))
            _modifiedUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("ModifiedUserID"))
            _modifiedDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("ModifiedDate"))

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

        Public Shared Sub Remove(ByVal eventID As Integer)
            SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_DeleteEvent", eventID)
        End Sub 'Remove
        '*********************************************************************
        '
        ' Calls Insert or Save based on the _projectID
        '
        '*********************************************************************

        Public Function Save() As Boolean
            If _eventID = 0 Then
                Return Insert()
            Else
                If _eventID > 0 Then
                    Return Update()
                Else
                    _eventID = 0
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
                _eventID = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_AddEvent", _projectID, _eventDate, _topic, _reason, _background, _answer, _effect, PMAttachment.getAttachments(_Attachments), _modifiedUserID))

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
                SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_UpdateEvent", _eventID, _projectID, _eventDate, _topic, _reason, _background, _answer, _effect, PMAttachment.getAttachments(_Attachments), _modifiedUserID)
                blnResult = True
            Catch ex As Exception
                Throw ex
                blnResult = False
            End Try

            Return blnResult
        End Function 'Update


    End Class 'PMEvent
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

