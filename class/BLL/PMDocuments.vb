Imports System
Imports System.Data
Imports System.Text
Imports System.Configuration
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.DataAccessLayer

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    '****************************************************************************
    '
    ' PMDocuments Class
    '
    ' The PMDocuments class represents a list of Project  Documents.
    '
    '****************************************************************************

    Public Class PMDocuments
        Private _documentID As Integer
        Private _projectID As Integer
        Private _categoryID As Integer

        Private _title As String
        Private _description As String

        Private _createdUserID As Integer
        Private _createdDate As Date
        Private _modifiedUserID As Integer
        Private _modifiedDate As Date
        Private _isDel As Boolean

        Private _Attachments As PMAttachmentCollection

#Region "  專案屬性存取 "
        Public Property DocumentID() As Integer
            Get
                Return _documentID

            End Get
            Set(ByVal Value As Integer)
                _documentID = Value

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
        Public Property CategoryID() As Integer
            Get
                Return _categoryID
            End Get
            Set(ByVal Value As Integer)
                _categoryID = Value
            End Set
        End Property
        Public ReadOnly Property CategoryName() As String
            Get
                Return PMDocCategories.GetCategoryName(_categoryID)

            End Get
            
        End Property
        Public Property Title() As String
            Get
                Return _title

            End Get
            Set(ByVal Value As String)
                _title = Value
            End Set
        End Property
        Public Property Description() As String
            Get
                Return _description

            End Get
            Set(ByVal Value As String)
                _description = Value
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

        Public Sub New(ByVal DocumentID As Integer)
            _documentID = DocumentID
        End Sub 'New

        Public Sub New(ByVal projectID As Integer, ByVal categoryID As Integer, ByVal title As String, ByVal description As String, ByVal modifiedUserID As Integer)
            _projectID = projectID
            _categoryID = categoryID
            _title = title
            _description = description
            _modifiedUserID = modifiedUserID

        End Sub 'New
        '*********************************************************************
        '
        ' Retrieves the list of all the Documents
        '
        '*********************************************************************
        Public Overloads Shared Function GetDocuments() As PMDocumentsCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListAllDocuments")

            ' Separate Data into a Collection of projects
            Dim docs As New PMDocumentsCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim doc As New PMDocuments
                    doc.DocumentID = Convert.ToInt32(r("DocumentID"))
                    doc.ProjectID = Convert.ToInt32(r("ProjectID"))
                    doc.CategoryID = Convert.ToInt32(r("CategoryID"))
                    doc.Title = r("Title").ToString()
                    doc.Description = r("Description").ToString()
                    
                    doc.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    doc.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    doc.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    doc.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))


                    docs.Add(doc)
                Next r

            End If
            Return docs
        End Function 'GetDocuments

        Public Overloads Shared Function GetDocuments(ByVal projectID As Integer) As PMDocumentsCollection
            'Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListEvents", projectID)
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListPMDocuments", projectID)

            ' Separate Data into a Collection of projects
            Dim docs As New PMDocumentsCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim doc As New PMDocuments
                    doc.DocumentID = Convert.ToInt32(r("DocumentID"))
                    doc.ProjectID = Convert.ToInt32(r("ProjectID"))
                    doc.CategoryID = Convert.ToInt32(r("CategoryID"))

                    doc.Title = r("Title").ToString()
                    doc.Description = r("Description").ToString()

                    doc.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    doc.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    doc.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    doc.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))

                    docs.Add(doc)
                Next r

            End If
            Return docs
        End Function 'GetDocuments

        Public Overloads Shared Function GetDocuments(ByVal projectID As Integer, ByVal categoryID As Integer) As PMDocumentsCollection
            'Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListEvents", projectID)
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListPMDocumentsByCategory", projectID, categoryID)

            ' Separate Data into a Collection of projects
            Dim docs As New PMDocumentsCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim doc As New PMDocuments
                    doc.DocumentID = Convert.ToInt32(r("DocumentID"))
                    doc.ProjectID = Convert.ToInt32(r("ProjectID"))
                    doc.CategoryID = Convert.ToInt32(r("CategoryID"))
                    doc.Title = r("Title").ToString()
                    doc.Description = r("Description").ToString()

                    doc.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    doc.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    doc.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    doc.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))

                    docs.Add(doc)
                Next r

            End If
            Return docs
        End Function 'GetDocuments

        Public Overloads Shared Function GetMyDocuments(ByVal UserID As Integer, ByVal RoleID As Integer) As PMDocumentsCollection
            'Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListEvents", projectID)
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListMyDocuments", UserID, RoleID)

            ' Separate Data into a Collection of projects
            Dim docs As New PMDocumentsCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim doc As New PMDocuments
                    doc.DocumentID = Convert.ToInt32(r("DocumentID"))
                    doc.ProjectID = Convert.ToInt32(r("ProjectID"))
                    doc.CategoryID = Convert.ToInt32(r("CategoryID"))
                    doc.Title = r("Title").ToString()
                    doc.Description = r("Description").ToString()

                    doc.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    doc.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    doc.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    doc.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))

                    docs.Add(doc)
                Next r

            End If
            Return docs
        End Function 'GetDocuments
        '*********************************************************************
        '
        ' Populates the object with the project info from the database.
        '
        '*********************************************************************

        Public Function Load() As Boolean
            ' The Get Projects stored procedure returns a dataset with 3 tables: Projects, Members, and Categories.
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetDocument", _documentID)

            If ds.Tables(0).Rows.Count < 1 Then
                Return False
            End If


            Me.DocumentID = Convert.ToInt32(ds.Tables(0).Rows(0)("DocumentID"))
            Me.ProjectID = Convert.ToInt32(ds.Tables(0).Rows(0)("ProjectID"))
            Me.CategoryID = Convert.ToInt32(ds.Tables(0).Rows(0)("CategoryID"))
            Me.Title = ds.Tables(0).Rows(0)("Title").ToString()
            Me.Description = ds.Tables(0).Rows(0)("Description").ToString()

            Me.CreatedUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("CreatedUserID"))
            Me.CreatedDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("CreatedDate"))
            Me.ModifiedUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("ModifiedUserID"))
            Me.ModifiedDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("ModifiedDate"))

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

        Public Shared Sub Remove(ByVal DocumentID As Integer)
            SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_DeleteDocument", DocumentID)
        End Sub 'Remove
        '*********************************************************************
        '
        ' Calls Insert or Save based on the _projectID
        '
        '*********************************************************************

        Public Function Save() As Boolean
            If _documentID = 0 Then
                Return Insert()
            Else
                If _documentID > 0 Then
                    Return Update()
                Else
                    _documentID = 0
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
                _documentID = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_AddDocument", _projectID, _categoryID, _title, _description, PMAttachment.getAttachments(_Attachments), _modifiedUserID))
 
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
                SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_UpdateDocument", _documentID, _projectID, _categoryID, _title, _description, PMAttachment.getAttachments(_Attachments), _modifiedUserID)
                blnResult = True
            Catch ex As Exception
                Throw ex
                blnResult = False
            End Try

            Return blnResult
        End Function 'Update


    End Class 'PMEvent
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

