Imports System
Imports System.Data
Imports System.Configuration

Imports System.Text
Imports Hsinyi.EIP.ProjectManagement.DataAccessLayer

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
    Public Class PMGroup
        Private _name As String
        Private _groupID As Integer
        Private _projectID As Integer
        Private _description As String
        Private _createdUserID As Integer
        Private _createdDate As DateTime
        Private _modifiedUserID As Integer
        Private _modifiedDate As DateTime
        Private _isDel As Boolean
        Private _members As UsersCollection
#Region "  專案屬性存取 "
        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(ByVal Value As String)
                _name = Value
            End Set
        End Property

        Public Property GroupID() As Integer
            Get
                Return _groupID
            End Get
            Set(ByVal Value As Integer)
                _groupID = Value
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
        Public Property Description() As String
            Get
                Return _description
            End Get
            Set(ByVal Value As String)
                _description = Value
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

        Public Sub New(ByVal GroupID As Integer)
            _groupID = GroupID
        End Sub 'New


        Public Shared Function GetGroupName(ByVal GroupID As Integer) As String
            Return Convert.ToString(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetGroupName", GroupID))
        End Function

        '*********************************************************************
        '
        ' Populates the object with the project info from the database.
        '
        '*********************************************************************

        Public Function Load() As Boolean
            ' The Get Projects stored procedure returns a dataset with 3 tables: Projects, Members, and Categories.
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetGroup", _groupID)

            If ds.Tables(0).Rows.Count < 1 Then
                Return False
            End If

            _name = Convert.ToString(ds.Tables(0).Rows(0)("GroupName"))
            _description = Convert.ToString(ds.Tables(0).Rows(0)("GroupDescription"))
            _projectID = Convert.ToInt32(ds.Tables(0).Rows(0)("ProjectID"))
            _groupID = Convert.ToInt32(ds.Tables(0).Rows(0)("GroupID"))

            _CreatedUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("CreatedUserID"))
            _CreatedDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("CreatedDate"))
            _ModifiedUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("ModifiedUserID"))
            _ModifiedDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("ModifiedDate"))
            _isDel = Convert.ToBoolean(ds.Tables(0).Rows(0)("isDel"))

            _members = New UsersCollection
            Dim row As DataRow
            For Each row In ds.Tables(1).Rows
                Dim user As New PMUser
                user.MemberID = Convert.ToInt32(row("MemberID"))
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

            Return True
        End Function 'Load

        Public Shared Function GetGroups() As GroupsCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListAllGroups")
            Dim GroupArray As New GroupsCollection

            ' Separate Data into a collection of Roles
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows
                Dim Group As New PMGroup
                Group.GroupID = Convert.ToInt32(r("GroupID"))
                Group.Name = r("GroupName").ToString()
                Group.ProjectID = Convert.ToInt32(r("ProjectID"))
                Group.Description = r("GroupDescription").ToString()
                GroupArray.Add(Group)
            Next r
            Return GroupArray
        End Function 'GetGroups


        Public Shared Function GetProjectGroups(ByVal projectID As Integer) As GroupsCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListGroups", projectID)
            Dim GroupArray As New GroupsCollection

            ' Separate Data into a collection of Roles
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows
                Dim Group As New PMGroup
                Group.GroupID = Convert.ToInt32(r("GroupID"))
                Group.Name = r("GroupName").ToString()
                Group.ProjectID = Convert.ToInt32(r("ProjectID"))
                Group.Description = r("GroupDescription").ToString()
                GroupArray.Add(Group)
            Next r
            Return GroupArray
        End Function 'GetRoles

        Public Shared Sub RemoveGroup(ByVal ID As Integer)
            SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_DeleteGroup", ID)
        End Sub 'RemoveGroup
        '*********************************************************************
        '
        ' Calls Insert or Save based on the _projectID
        '
        '*********************************************************************

        Public Function Save() As Boolean
            If _groupID = 0 Then
                Return Insert()
            Else
                If _groupID > 0 Then
                    Return Update()
                Else
                    _groupID = 0
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
                _groupID = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_AddGroup", _projectID, _name, _description, getSelectedMembers.ToString(), _modifiedUserID))

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
                SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_UpdateGroup", _groupID, _projectID, _name, _description, getSelectedMembers.ToString, _ModifiedUserID)
                blnResult = True
            Catch ex As Exception
                Throw ex
                blnResult = False
            End Try

            Return blnResult
        End Function 'Update

        Private Function getSelectedMembers() As String
            Dim selectedMembers As New StringBuilder(_members.Count)
            Dim index As Integer = 1
            Dim user As PMUser
            For Each user In _members
                selectedMembers.Append(user.MemberID)

                ' Don't append separator if this is the last item
                If index <> _members.Count Then
                    selectedMembers.Append(",")
                End If
                index += 1
            Next user
            Return selectedMembers.ToString()
        End Function
    End Class 'PMGroup 
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
