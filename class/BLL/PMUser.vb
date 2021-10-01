Imports System
Imports System.Data
Imports System.Configuration
Imports Hsinyi.EIP.ProjectManagement.DataAccessLayer
Imports EIPSysSecurity

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    Public Class PMUser

        'Public Const UserRoleNone As String = "0"
        'Public Const UserRoleAdministrator As String = "1"
        'Public Const FlowManagr As String = "2"
        'Public Const HRMaster As String = "3"
        'Public Const UserRoleProjectManager As String = "4"
        'Public Const UserRoleProjectChief As String = "5"
        'Public Const UserRoleProjectMember As String = "6"
        'Public Const UserRoleUser As String = "7" ' 

        'Public Const EipSecurityUser As Int16 = 1
        'Public Const ProjectManagementUser As Int16 = 2

        'Public Const UserRoleAdminPMgr As String = UserRoleAdministrator + "," + UserRoleProjectManager
        'Public Const UserRolePMgrConsultant As String = UserRoleProjectManager + "," + UserRoleConsultant

        Private _memberID As Integer = 0
        Private _userID As Integer

        Private _userName As String
        Private _Name As String
        'Private _identityName As String

        Private _projectID As Integer
        Private _role As String = Roles.UserRoleNone
        Private _roleName As String   
        Private _Connects As String
        Private _sourceUserTitle As String
        Private _sourceUserDepartment As String
        Private _description As String
        Private _CurrentUserID As Integer

        Private _groupID As Integer
        Private _groupName As String
        Private _displayName As String = String.Empty
        Private _firstName As String = String.Empty
        Private _lastName As String = String.Empty
        Private _floor As String
        Private _extension As String

        Public Sub New()
        End Sub 'New

        Public Sub New(ByVal MemberID As Integer)
            _memberID = MemberID

        End Sub 'New

        Public Sub New(ByVal UserName As String)
            Me.UserName = UserName

        End Sub 'New
        Public Sub New(ByVal UserName As String, ByVal projectID As Integer)
            Me.UserName = UserName
            _projectID = projectID
        End Sub 'New

        Public Sub New(ByVal UserID As Integer, ByVal projectID As Integer)
            Me.UserID = UserID
            _projectID = projectID
        End Sub 'New
        Public Sub New(ByVal UserID As Integer, ByVal UserName As String, ByVal Name As String, ByVal Role As String)
            _userID = UserID
            _userName = UserName
            _displayName = Name
            _role = Role

        End Sub 'New
        Public Property Group() As Integer
            Get
                Return _groupID
            End Get
            Set(ByVal Value As Integer)
                _groupID = Value

            End Set
        End Property
        Public Property GroupName() As String
            Get
                Return _groupName
            End Get
            Set(ByVal Value As String)
                _groupName = Value
            End Set
        End Property
        Public Property DisplayName() As String
            Get
                If Not Strings.Trim(_displayName) = "" Then
                    Return _displayName
                Else
                    Return GetDisplayName(_userName, _firstName, _lastName) 'EIPSysSecurity.clsUser.getUserName(_userID)
                End If

            End Get
            Set(ByVal Value As String)
                _displayName = Value
            End Set
        End Property

        Public Property FirstName() As String
            Get
                Return _firstName
            End Get
            Set(ByVal Value As String)
                _firstName = Value
            End Set
        End Property

        Public Property LastName() As String
            Get
                Return _lastName
            End Get
            Set(ByVal Value As String)
                _lastName = Value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal Value As String)
                _Name = Value
            End Set
        End Property

        'Public Property IdentityName() As String
        '    Get
        '        Return _identityName
        '    End Get
        '    Set(ByVal Value As String)
        '        _identityName = Value
        '    End Set
        'End Property

        Public Property ProjectID() As String
            Get
                Return _projectID
            End Get
            Set(ByVal Value As String)
                _projectID = Value
            End Set
        End Property

        Public Property Role() As String
            Get
                Return _role
            End Get
            Set(ByVal Value As String)
                _role = Value
            End Set
        End Property

        Public Property RoleName() As String
            Get
                Return _roleName
            End Get
            Set(ByVal Value As String)
                _roleName = Value
            End Set
        End Property
        Public Property MemberID() As Integer
            Get
                Return _memberID
            End Get
            Set(ByVal Value As Integer)
                _memberID = Value

            End Set
        End Property
        Public Property UserID() As Integer
            Get
                Return _userID
            End Get
            Set(ByVal Value As Integer)
                _userID = Value
                _userName = EIPSysSecurity.clsUser.getUserLoginID(_userID)
            End Set
        End Property

        Public Property UserName() As String
            Get
                Return _userName
            End Get
            Set(ByVal Value As String)
                _userName = Value
                _userID = EIPSysSecurity.clsUser.getUserID(_userName)
            End Set
        End Property
        Public Property Connects() As String
            Get
                Return _Connects
            End Get
            Set(ByVal Value As String)
                _Connects = Value
            End Set
        End Property
        Public Property SourceUserTitle() As String
            Get
                Return _sourceUserTitle
            End Get
            Set(ByVal Value As String)
                _sourceUserTitle = Value
            End Set
        End Property
        Public Property SourceUserDepartment() As String
            Get
                Return _sourceUserDepartment
            End Get
            Set(ByVal Value As String)
                _sourceUserDepartment = Value
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
        Public Property ModifiedUserID() As Integer
            Get
                Return _CurrentUserID
            End Get
            Set(ByVal Value As Integer)
                _CurrentUserID = Value
            End Set
        End Property

        Public Property Floor() As String
            Get
                Return _floor
            End Get
            Set(ByVal Value As String)
                _floor = Value
            End Set
        End Property

        Public Property Extension() As String
            Get
                Return _extension
            End Get
            Set(ByVal Value As String)
                _extension = Value
            End Set
        End Property
        '*********************************************************************
        '
        ' GetAllUsers Static Method
        ' Retrieves a list of all users.
        '
        '*********************************************************************

        Public Shared Function GetAllUsers(ByVal userID As Integer) As UsersCollection
            Return GetUsers(userID, Roles.UserRoleAdministrator)
        End Function 'GetAllUsers

        '*********************************************************************
        '
        ' GetUsers Static Method
        ' Retrieves a list of users based on the specified userID and role.
        ' The list returned is restricted by role.  For instance, users with
        ' the role of Administrator can see all users, while users with the
        ' role of Consultant can only see themselves.
        '
        '*********************************************************************

        Public Shared Function GetUsers(ByVal userID As Integer, ByVal role As String) As UsersCollection
            Dim firstName As String = String.Empty
            Dim lastName As String = String.Empty

            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListUsers", userID, Convert.ToInt32(role))
            Dim users As New UsersCollection

            ' Separate Data into a collection of Users.
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows
                Dim usr As New PMUser
                usr.MemberID = Convert.ToInt32(r("MemberID"))
                usr.ProjectID = Convert.ToInt32(r("projectID"))
                usr.UserID = Convert.ToInt32(r("UserID"))
                'usr.IdentityName = EIPSysSecurity.clsUser.getUserLoginID(usr.UserID)
                usr.UserName = EIPSysSecurity.clsUser.getUserLoginID(usr.UserID)
                usr.DisplayName = GetDisplayName(usr.UserName, usr.FirstName, usr.LastName)
                usr.Role = r("RoleID").ToString()
                usr.RoleName = r("RoleName").ToString()
                usr.Name = GetDisplayName(usr.UserName, usr.FirstName, usr.LastName)
                usr.Connects = r("Connects").ToString()
                usr.SourceUserTitle = r("SourceUserTitle").ToString()
                usr.SourceUserDepartment = r("SourceUserDepartment").ToString()
                usr.Description = r("Description").ToString()
                usr.FirstName = firstName
                usr.LastName = lastName
                usr.Floor = EIPSysSecurity.clsUser.getFloor(r("UserID"))
                usr.Extension = EIPSysSecurity.clsUser.getExtension(r("UserID"))
                users.Add(usr)
            Next r
            Return users
        End Function 'GetUsers

        Public Shared Function IsProjectMember(ByVal projectID As Integer, ByVal UserID As Integer) As Boolean
            Dim blnResult As Boolean
            Dim MemberID As Integer = 0
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetMemberByUserID", UserID, projectID)
            ' Separate Data into a collection of Users.
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows
                MemberID = Convert.ToInt32(r("MemberID"))
            Next

            If MemberID > 0 Then
                blnResult = True
            Else
                blnResult = False
            End If

            ds.Clear()
            ds.Dispose()

            Return blnResult
        End Function

        Public Shared Function GetMembers(ByVal projectID As Integer) As UsersCollection
            Dim firstName As String = String.Empty
            Dim lastName As String = String.Empty

            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListMembers", projectID)
            Dim users As New UsersCollection

            ' Separate Data into a collection of Users.
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows
                Dim usr As New PMUser
                usr.MemberID = Convert.ToInt32(r("MemberID"))
                usr.ProjectID = Convert.ToInt32(r("projectID"))
                usr.Group = GetUserGroupID(Convert.ToInt32(r("MemberID")), projectID)
                usr.GroupName = PMGroup.GetGroupName(usr.Group)
                usr.UserID = Convert.ToInt32(r("UserID"))
                'usr.IdentityName = EIPSysSecurity.clsUser.getUserLoginID(usr.UserID)
                usr.UserName = EIPSysSecurity.clsUser.getUserLoginID(usr.UserID)
                usr.DisplayName = GetDisplayName(usr.UserName, usr.FirstName, usr.LastName)

                usr.Role = r("RoleID").ToString()
                usr.RoleName = r("RoleName").ToString()
                usr.Name = GetDisplayName(usr.UserName, usr.FirstName, usr.LastName)
                usr.Connects = r("Connects").ToString()
                usr.SourceUserTitle = r("SourceUserTitle").ToString()
                usr.SourceUserDepartment = r("SourceUserDepartment").ToString()
                usr.Description = r("Description").ToString()
                usr.FirstName = firstName
                usr.LastName = lastName
                usr.Floor = EIPSysSecurity.clsUser.getFloor(r("UserID"))
                usr.Extension = EIPSysSecurity.clsUser.getExtension(r("UserID"))
                users.Add(usr)
            Next r
            Return users
        End Function 'GetMembers
        '*********************************************************************
        '
        ' GetDisplayName static method
        ' Gets the user's first and last name from the specified TTUser account source, which is
        ' set in Web.confg.
        '
        '*********************************************************************

        Public Shared Function GetDisplayName(ByVal userName As String, ByRef firstName As String, ByRef lastName As String) As String
            Dim displayName As String = String.Empty
            Dim dbName As String = String.Empty


            ' The DirectoryHelper class will attempt to get the user's first 
            ' and last name from the specified account source.
            If userName Is Nothing Then
                displayName = "[使用者名稱空白]"
            Else

                DirectoryHelper.FindUser(userName, firstName, lastName)


                'If firstName.Length > 0 Or lastName.Length > 0 Then
                If (Strings.Trim(firstName) Is String.Empty) Or (Strings.Trim(lastName) Is String.Empty) Then
                    ' If the first and last name could not be retrieved, return the EIP UserName.
                    dbName = EIPSysSecurity.clsUser.getUserName(EIPSysSecurity.clsUser.getUserID(userName))

                    If Not dbName Is String.Empty Then
                        displayName = dbName
                    Else
                        displayName = userName
                    End If
                Else
                    displayName = lastName & " " & firstName
                End If
            End If

            Return displayName
        End Function 'GetDisplayName

        'Public Shared Function GetDisplayNameFromDB(ByVal userName As String) As String
        '    Dim displayName As String = String.Empty
        '    displayName = CStr(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetUserDisplayName", userName))
        '    Return displayName
        'End Function

        '*********************************************************************
        '
        ' ListManagers Static Method
        ' Retrieves a list of users with the role of Project Manager.
        '
        '*********************************************************************

        Public Shared Function ListManagers() As UsersCollection
            Dim firstName As String = String.Empty
            Dim lastName As String = String.Empty

            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListManagers")
            Dim managersArray As New UsersCollection

            ' Separate Data into a list of collections.
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows
                Dim usr As New PMUser
                usr.MemberID = Convert.ToInt32(r("MemberID"))

                usr.UserID = Convert.ToInt32(r("UserID"))
                usr.UserName = EIPSysSecurity.clsUser.getUserLoginID(usr.UserID)
                'usr.IdentityName = EIPSysSecurity.clsUser.getUserLoginID(usr.UserID)
                usr.DisplayName = GetDisplayName(usr.UserName, usr.FirstName, usr.LastName)

                usr.Role = r("RoleID").ToString()
                usr.RoleName = r("RoleName").ToString()
                usr.Name = GetDisplayName(usr.UserName, firstName, lastName)
                usr.Connects = r("Connects").ToString()
                usr.SourceUserTitle = r("SourceUserTitle").ToString()
                usr.SourceUserDepartment = r("SourceUserDepartment").ToString()
                usr.Description = r("Description").ToString()
                usr.FirstName = firstName
                usr.LastName = lastName
                usr.Floor = EIPSysSecurity.clsUser.getFloor(r("UserID"))
                usr.Extension = EIPSysSecurity.clsUser.getExtension(r("UserID"))
                managersArray.Add(usr)
            Next r
            Return managersArray
        End Function 'ListManagers


        '*********************************************************************
        '
        ' Remove static method
        ' Removes a user from database
        '
        '*********************************************************************

        Public Shared Sub RemoveMember(ByVal MemberID As Integer)
            SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_DeleteMember", MemberID)
        End Sub 'RemoveMember

        Public Shared Sub RemoveUser(ByVal ID As Integer)
            SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_DeleteUser", ID)
        End Sub 'RemoveUser

        '*********************************************************************
        '
        ' Load method
        ' Retrieve user information from the data access layer
        ' returns True if user information is loaded successfully, false otherwise.
        '
        '*********************************************************************

        Public Function LoadProjectMember() As Boolean
            ' Get the user's information from the database

            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetProjectUser", _userID, _projectID)

            If ds.Tables(0).Rows.Count < 1 Then
                Return False
            End If
            Dim dr As DataRow = ds.Tables(0).Rows(0)

            _memberID = Convert.ToInt32(dr("MemberID"))
            _userName = EIPSysSecurity.clsUser.getUserLoginID(_userID)
            _role = dr("RoleID").ToString()
            _roleName = dr("RoleName").ToString()
            _displayName = GetDisplayName(_userName, _firstName, _lastName)
            _Connects = dr("Connects").ToString()
            _sourceUserTitle = dr("SourceUserTitle").ToString()
            _sourceUserDepartment = dr("SourceUserDepartment").ToString()
            _description = dr("Description").ToString()
            _groupID = GetUserGroupID(_memberID, ProjectID)
            _groupName = PMGroup.GetGroupName(_groupID)
            _floor = EIPSysSecurity.clsUser.getFloor(_userID)
            _extension = EIPSysSecurity.clsUser.getExtension(_userID)
            Return True
        End Function 'Load

        Public Function Load() As Boolean
            ' Get the user's information from the database

            If _memberID <= 0 Then
                _displayName = GetDisplayName(_userName, _firstName, _lastName)

                _role = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetUserRole", _userID))
                _roleName = Roles.GetRolesName(_role)

            Else

                Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetProjectMember", _memberID)

                If ds.Tables(0).Rows.Count < 1 Then
                    Return False
                End If
                Dim dr As DataRow = ds.Tables(0).Rows(0)
                _memberID = Convert.ToInt32(dr("MemberID"))
                _projectID = Convert.ToInt32(dr("ProjectID"))
                _userID = Convert.ToInt32(dr("userID"))
                _userName = EIPSysSecurity.clsUser.getUserLoginID(_userID)
                _role = dr("RoleID").ToString()
                _roleName = dr("RoleName").ToString()
                _displayName = GetDisplayName(_userName, _firstName, _lastName)
                _Connects = dr("Connects").ToString()
                _sourceUserTitle = dr("SourceUserTitle").ToString()
                _sourceUserDepartment = dr("SourceUserDepartment").ToString()
                _description = dr("Description").ToString()
                _groupID = GetUserGroupID(_memberID, _projectID)
                _groupName = PMGroup.GetGroupName(_groupID)
                _floor = EIPSysSecurity.clsUser.getFloor(_userID)
                _extension = EIPSysSecurity.clsUser.getExtension(_userID)
            End If


            Return True
        End Function 'Load

        '*********************************************************************
        '
        ' Save method
        ' Add or update user information in the database depending on the TT_UserID.
        ' Returns True if saved successfully, false otherwise.
        '
        '*********************************************************************

        Public Overloads Function Save() As Boolean
            Dim isUserFound As Boolean = False
            Dim isUserActiveManager As Boolean = True
            Return Save(False, isUserFound, isUserActiveManager)
        End Function 'Save

        '*********************************************************************
        '
        ' Save method
        ' Add or update user information in the database depending on the TTUserID.
        ' Returns True if saved successfully, false otherwise.
        '
        '*********************************************************************

        Public Overloads Function Save(ByVal checkUsername As Boolean, ByRef isUserFound As Boolean, ByRef isUserActiveManager As Boolean) As Boolean
            ' Determines whether object needs update or to be inserted.

            If _memberID = 0 Then
                Return Insert(checkUsername, isUserFound)
            Else
                If _memberID > 0 Then
                    Return Update(isUserActiveManager)
                Else
                    _memberID = 0
                    Return False
                End If
            End If
        End Function 'Save

        Private Function Insert(ByVal checkUsername As Boolean, ByRef isUserFound As Boolean) As Boolean
            Dim firstName As String = String.Empty
            Dim lastName As String = String.Empty
            isUserFound = False

            If ConfigurationSettings.AppSettings(Web.Global.CfgKeyUserAcctSource) <> "None" Then
                ' Check to see if the user is in the NT SAM or Active Directory before inserting them
                ' into the Time Tracker database.  If a first or last name is returned, the user exists and
                ' can be inserted into the Time Tracker database.
                If checkUsername Then
                    PMUser.GetDisplayName(_userName, firstName, lastName)
                    isUserFound = firstName <> String.Empty Or lastName <> String.Empty
                End If
            Else
                checkUsername = False
                isUserFound = True
            End If


            If checkUsername And isUserFound Or Not checkUsername Then

                _memberID = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_AddUser", _userID, Convert.ToInt32(_role), _groupID, _projectID, _description, _Connects, _sourceUserTitle, _sourceUserDepartment, _CurrentUserID))
                isUserFound = True
            End If
            Return _memberID > 0
        End Function 'Insert

        Private Function Update(ByRef isUserActiveManger As Boolean) As Boolean
            ' if new user role is a consultant, check if user is a active manager of one or more project. if so, no update is applied 
            If _role = Roles.UserRoleUser Then
                If Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetManagerProjectCount", _userID)) > 0 Then
                    isUserActiveManger = True
                    Return False
                Else
                    isUserActiveManger = False
                End If
            End If
            Return 0 < Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_UpdateUser", _memberID, _userID, Convert.ToInt32(_role), _groupID, _projectID, _description, _Connects, _sourceUserTitle, _sourceUserDepartment, _CurrentUserID))
        End Function 'Update

        Public Shared Function GetUserGroupID(ByVal MemberID As Integer, ByVal ProjectID As Integer) As Integer
            Return Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetUserGroup", MemberID, ProjectID))
        End Function

        '*********************************************************************
        '
        ' UsersDB.Login() Method 
        '
        ' The Login method validates a email/password pair against credentials
        ' stored in the users database.  If the email/password pair is valid,
        ' the method returns user's name.
        '
        ' Other relevant sources:
        '     + UserLogin Stored Procedure
        '
        '*********************************************************************

        'Public Function Login(ByVal email As String, ByVal password As String) As String

        '    Dim userName As String
        '    userName = CStr(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "TT_UserLogin", email, password))

        '    If Not userName Is Nothing Or userName Is "" Then
        '        Return userName
        '    Else
        '        Return String.Empty
        '    End If

        'End Function

    End Class 'PMUser

End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer