Imports System
Imports System.Data
Imports System.Configuration
Imports System.Text
Imports HsinYi.EIP.ProjectManagement.DataAccessLayer
Imports EIPSysSecurity


Namespace HsinYi.EIP.ProjectManagement.BusinessLogicLayer

    Public Class Project
        Private _ProjectID As Integer
        Private _ProjectName As String
        Private _ProjectDescription As String
        Private _ManagerUserID As Integer
        Private _ManagerUserName As String
        Private _ChiefUserID As Integer
        Private _ChiefUserName As String
        Private _StartDate As DateTime
        Private _EndDate As DateTime
        Private _StateID As Integer
        Private _StateName As String
        Private _TypeID As Integer
        Private _TypeName As String
        Private _SupplementaryUnits As String
        Private _SupplementarySubsidy As String
        Private _AssistUnits As String
        Private _CreatedUserID As Integer
        Private _CreatedDate As DateTime
        Private _ModifiedUserID As Integer
        Private _ModifiedDate As DateTime
        Private _isDel As Boolean

        Private _members As UsersCollection

#Region "  專案屬性存取 "

        Public Property projectID() As Integer
            Get
                Return _ProjectID
            End Get
            Set(ByVal Value As Integer)
                _ProjectID = Value
            End Set
        End Property

        Public Property ProjectName() As String
            Get
                Return _ProjectName
            End Get
            Set(ByVal Value As String)
                _ProjectName = Value
            End Set
        End Property


        Public Property ProjectDescription() As String
            Get
                Return _ProjectDescription
            End Get
            Set(ByVal Value As String)
                _ProjectDescription = Value
            End Set
        End Property

        Public Property ManagerUserID() As Integer
            Get
                Return _ManagerUserID
            End Get
            Set(ByVal Value As Integer)
                _ManagerUserID = Value
            End Set
        End Property

        Public Property ChiefUserID() As Integer
            Get
                Return _ChiefUserID
            End Get
            Set(ByVal Value As Integer)
                _ChiefUserID = Value
            End Set
        End Property

        Public Property ChiefName() As String
            Get
                Return _ChiefUserName
            End Get
            Set(ByVal Value As String)
                _ChiefUserName = Value
            End Set
        End Property

        Public Property ManagerUserName() As String
            Get
                Return _ManagerUserName
            End Get
            Set(ByVal Value As String)
                _ManagerUserName = Value
            End Set
        End Property

        Public Property StartDate() As Date
            Get
                Return _StartDate
            End Get
            Set(ByVal Value As Date)
                _StartDate = Value
            End Set
        End Property

        Public Property EndDate() As Date
            Get
                Return _EndDate
            End Get
            Set(ByVal Value As Date)
                _EndDate = Value
            End Set
        End Property

        Public Property StateID() As Integer
            Get
                Return _StateID
            End Get
            Set(ByVal Value As Integer)
                _StateID = Value
            End Set
        End Property

        Public Property StateName() As String
            Get
                Return _StateName
            End Get
            Set(ByVal Value As String)
                _StateName = Value
            End Set
        End Property

        Public Property TypeID() As Integer
            Get
                Return _TypeID
            End Get
            Set(ByVal Value As Integer)
                _TypeID = Value
            End Set
        End Property

        Public Property TypeName() As String
            Get
                Return _TypeName
            End Get
            Set(ByVal Value As String)
                _TypeName = Value
            End Set
        End Property

        Public Property SupplementaryUnits() As String
            Get
                Return _SupplementaryUnits
            End Get
            Set(ByVal Value As String)
                _SupplementaryUnits = Value
            End Set
        End Property

        Public Property SupplementarySubsidy() As String
            Get
                Return _SupplementarySubsidy
            End Get
            Set(ByVal Value As String)
                _SupplementarySubsidy = Value
            End Set
        End Property

        Public Property AssistUnits() As String
            Get
                Return _AssistUnits
            End Get
            Set(ByVal Value As String)
                _AssistUnits = Value
            End Set
        End Property
        Public Property CreatedUserID() As Integer
            Get
                Return _CreatedUserID
            End Get
            Set(ByVal Value As Integer)
                _CreatedUserID = Value
            End Set
        End Property

        Public Property ModifiedUserID() As Integer
            Get
                Return _ModifiedUserID
            End Get
            Set(ByVal Value As Integer)
                _ModifiedUserID = Value
            End Set
        End Property

        Public ReadOnly Property getCreatedDate() As DateTime
            Get
                Return _CreatedDate
            End Get

        End Property

        Public ReadOnly Property getModifiedDate() As DateTime
            Get
                Return _ModifiedDate
            End Get

        End Property

        Private Property CreatedDate() As DateTime
            Get
                Return _CreatedDate
            End Get

            Set(ByVal Value As DateTime)
                _CreatedDate = Value
            End Set
        End Property

        Private Property ModifiedDate() As DateTime
            Get
                Return _ModifiedDate
            End Get
            Set(ByVal Value As DateTime)
                _ModifiedDate = Value
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
        Public Property Members() As UsersCollection
            Get
                Return _members
            End Get
            Set(ByVal Value As UsersCollection)
                _members = Value
            End Set
        End Property
#End Region


        Public Sub New()
        End Sub 'New

        Public Sub New(ByVal projectID As Integer)
            _ProjectID = projectID
        End Sub 'New

        Public Sub New(ByVal projectID As Integer, ByVal name As String, ByVal description As String, ByVal managerUserID As Integer, ByVal chiefUserID As Integer, ByVal stateID As Integer, ByVal typeID As Integer, ByVal StartDate As Date, ByVal EndDate As Date)
            _ProjectID = projectID
            _ProjectName = name
            _ProjectDescription = description
            _ManagerUserID = managerUserID
            _StateID = stateID
            _TypeID = typeID
            _ChiefUserID = chiefUserID
            _StartDate = StartDate
            _EndDate = EndDate
        End Sub 'New

     
        '*********************************************************************
        '
        ' Retrieves the list of all the projects
        '
        '*********************************************************************
        Public Overloads Shared Function GetProjects() As ProjectsCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListAllProjects")

            ' Separate Data into a Collection of projects
            Dim projects As New ProjectsCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim pj As New Project
                    pj.projectID = Convert.ToInt32(r("ProjectID"))
                    pj.ProjectName = r("ProjectName").ToString()
                    pj.ProjectDescription = r("ProjectDescription").ToString()
                    pj.ManagerUserID = Convert.ToInt32(r("ManagerUserID"))
                    pj.ManagerUserName = EIPSysSecurity.clsUser.getUserName(pj.ManagerUserID)
                    pj.ChiefUserID = Convert.ToInt32(r("ChiefUserID"))
                    pj.ChiefName = EIPSysSecurity.clsUser.getUserName(pj.ChiefUserID)
                    pj.StartDate = Convert.ToDateTime(r("StartDate"))
                    pj.EndDate = Convert.ToDateTime(r("EndDate"))
                    pj.StateID = Convert.ToInt32(r("StateID"))
                    pj.StateName = r("StateName").ToString()
                    pj.TypeID = Convert.ToInt32(r("TypeID"))
                    pj.TypeName = r("TypeName").ToString()
                    pj.SupplementaryUnits = r("SupplementaryUnits").ToString()
                    pj.SupplementarySubsidy = r("SupplementarySubsidy").ToString()
                    pj.AssistUnits = r("AssistUnits").ToString()
                    pj.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    pj.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    pj.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    pj.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))



                    projects.Add(pj)
                Next r

            End If
            Return projects
        End Function 'GetProjects

        '*********************************************************************
        '
        ' Retrieves a list of projects based on the user's role
        '
        '*********************************************************************

        Public Overloads Shared Function GetProjects(ByVal userID As Integer, ByVal role As String) As ProjectsCollection
            Dim firstName As String = String.Empty
            Dim lastName As String = String.Empty

            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListProjects", userID, Convert.ToInt32(role))

            Dim projects As New ProjectsCollection
            Dim r As DataRow
            If ds.Tables.Count > 0 Then


                For Each r In ds.Tables(0).Rows
                    Dim pj As New Project
                    pj.projectID = Convert.ToInt32(r("ProjectID"))
                    pj.ProjectName = r("ProjectName").ToString()
                    pj.ProjectDescription = r("ProjectDescription").ToString()
                    pj.ManagerUserID = Convert.ToInt32(r("ManagerUserID"))
                    pj.ManagerUserName = EIPSysSecurity.clsUser.getUserName(pj.ManagerUserID) 'r("ManagerUserName").ToString()


                    pj.ChiefUserID = Convert.ToInt32(r("ChiefUserID"))
                    pj.ChiefName = EIPSysSecurity.clsUser.getUserName(pj.ChiefUserID)
                    pj.StartDate = Convert.ToDateTime(r("StartDate"))
                    pj.EndDate = Convert.ToDateTime(r("EndDate"))
                    pj.StateID = Convert.ToInt32(r("StateID"))
                    pj.StateName = r("StateName").ToString()
                    pj.TypeID = Convert.ToInt32(r("TypeID"))
                    pj.TypeName = r("TypeName").ToString()

                    pj.SupplementaryUnits = r("SupplementaryUnits").ToString()
                    pj.SupplementarySubsidy = r("SupplementarySubsidy").ToString()
                    pj.AssistUnits = r("AssistUnits").ToString()
                    pj.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    pj.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    pj.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    pj.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))


                    projects.Add(pj)
                Next r

            End If
            Return projects
        End Function 'GetProjects

        '*********************************************************************
        '
        ' Retrieves a list of projects based on project membership
        '
        '*********************************************************************

        Public Overloads Shared Function GetProjects(ByVal queryUserID As Integer, ByVal userID As Integer) As ProjectsCollection
            Dim projects As New ProjectsCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListProjectsWithMembership", queryUserID, userID)

            If ds.Tables.Count > 0 Then

                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim pj As New Project
                    pj.projectID = Convert.ToInt32(r("ProjectID"))
                    pj.ProjectName = r("ProjectName").ToString()
                    pj.ProjectDescription = r("ProjectDescription").ToString()
                    pj.ManagerUserID = Convert.ToInt32(r("ManagerUserID"))
                    pj.ManagerUserName = EIPSysSecurity.clsUser.getUserName(pj.ManagerUserID) ' r("ManagerUserName").ToString()

                    pj.ChiefUserID = Convert.ToInt32(r("ChiefUserID"))
                    pj.ChiefName = EIPSysSecurity.clsUser.getUserName(pj.ChiefUserID)
                    pj.StartDate = Convert.ToDateTime(r("StartDate"))
                    pj.EndDate = Convert.ToDateTime(r("EndDate"))
                    pj.StateID = Convert.ToInt32(r("StateID"))
                    pj.StateName = r("StateName").ToString()
                    pj.TypeID = Convert.ToInt32(r("TypeID"))
                    pj.TypeName = r("TypeName").ToString()
                    pj.SupplementaryUnits = r("SupplementaryUnits").ToString()
                    pj.SupplementarySubsidy = r("SupplementarySubsidy").ToString()
                    pj.AssistUnits = r("AssistUnits").ToString()
                    pj.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    pj.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    pj.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    pj.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))


                    projects.Add(pj)
                Next r

            End If
            Return projects
        End Function 'GetProjects

        '*********************************************************************
        '
        ' Populates the object with the project info from the database.
        '
        '*********************************************************************

        Public Function Load() As Boolean
            ' The Get Projects stored procedure returns a dataset with 3 tables: Projects, Members, and Categories.
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetProject", _ProjectID)

            If ds.Tables(0).Rows.Count < 1 Then
                Return False
            End If

            _ProjectName = Convert.ToString(ds.Tables(0).Rows(0)("ProjectName"))
            _ProjectDescription = Convert.ToString(ds.Tables(0).Rows(0)("ProjectDescription"))
            _ManagerUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("ManagerUserID"))
            _ManagerUserName = EIPSysSecurity.clsUser.getUserName(_ManagerUserID)
            _ChiefUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("ChiefUserID"))
            _ChiefUserName = EIPSysSecurity.clsUser.getUserName(_ChiefUserID)
            _StartDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("StartDate"))
            _EndDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("EndDate"))
            _StateID = Convert.ToInt32(ds.Tables(0).Rows(0)("StateID"))
            _StateName = ProjectState.GetStateName(_StateID)
            _TypeID = Convert.ToInt32(ds.Tables(0).Rows(0)("TypeID"))
            _TypeName = ProjectType.GetTypeName(_TypeID)
            _SupplementaryUnits = Convert.ToString(ds.Tables(0).Rows(0)("SupplementaryUnits"))
            _SupplementarySubsidy = Convert.ToString(ds.Tables(0).Rows(0)("SupplementarySubsidy"))
            _AssistUnits = Convert.ToString(ds.Tables(0).Rows(0)("AssistUnits"))
            _CreatedUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("CreatedUserID"))
            _CreatedDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("CreatedDate"))
            _ModifiedUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("ModifiedUserID"))
            _ModifiedDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("ModifiedDate"))
            _isDel = Convert.ToBoolean(ds.Tables(0).Rows(0)("isDel"))



            _members = New UsersCollection
            Dim row As DataRow
            For Each row In ds.Tables(1).Rows
                Dim user As New PMUser
                user.UserID = Convert.ToInt32(row("UserID"))
                user.UserName = EIPSysSecurity.clsUser.getUserLoginID(user.UserID)
                user.DisplayName = EIPSysSecurity.clsUser.getUserName(user.UserID)
                user.Role = Convert.ToInt32(row("RoleID"))
                user.RoleName = Convert.ToString(row("RoleName"))
                user.Description = Convert.ToString(row("Description"))
                user.Connects = Convert.ToString(row("Connects"))
                user.SourceUserTitle = Convert.ToString(row("SourceUserTitle"))
                user.SourceUserDepartment = Convert.ToString(row("SourceUserDepartment"))
                user.ModifiedUserID = PMSecurity.GetUserID
                user.Name = EIPSysSecurity.clsUser.getUserLoginID(user.UserID)

                _members.Add(user)
            Next row

            '_categories = New CategoriesCollection
            'For Each row In ds.Tables(2).Rows
            '    Dim cat As New Category
            '    cat.ProjectID = _ProjectID
            '    cat.CategoryID = Convert.ToInt32(row("CategoryID"))
            '    cat.Name = Convert.ToString(row("Name"))
            '    cat.Abbreviation = Convert.ToString(row("CategoryShortName"))
            '    cat.EstDuration = Convert.ToDecimal(row("EstDuration"))
            '    _categories.Add(cat)
            'Next row

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

        Public Shared Sub Remove(ByVal projectID As Integer)
            SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_DeleteProject", projectID)
        End Sub 'Remove
        '*********************************************************************
        '
        ' Calls Insert or Save based on the _projectID
        '
        '*********************************************************************

        Public Function Save() As Boolean
            If _ProjectID = 0 Then
                Return Insert()
            Else
                If _ProjectID > 0 Then
                    Return Update()
                Else
                    _ProjectID = 0
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
                _ProjectID = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_AddProject", _ProjectName, _ProjectDescription, _ManagerUserID, _ChiefUserID, _StartDate, _EndDate, _StateID, _TypeID, _SupplementaryUnits, _SupplementarySubsidy, _AssistUnits, _ModifiedUserID))

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
                SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_UpdateProject", _ProjectID, _ProjectName, _ProjectDescription, _ManagerUserID, _ChiefUserID, _StartDate, _EndDate, _StateID, _TypeID, _SupplementaryUnits, _SupplementarySubsidy, _AssistUnits, _ModifiedUserID)
                blnResult = True
            Catch ex As Exception
                Throw ex
                blnResult = False
            End Try

            Return blnResult
        End Function 'Update

    End Class 'Project

End Namespace 'EIP.ProjectManagement.BusinessLogicLayer

