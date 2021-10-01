Imports System
Imports System.Data
Imports System.Configuration
Imports Hsinyi.EIP.ProjectManagement.DataAccessLayer

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    '****************************************************************************
    '
    ' Roles Class
    '
    ' The Roles class represents a list of custom Time Tracker roles.
    '
    '****************************************************************************
    


    Public Class Roles

        Public Const UserRoleNone As String = "0"
        Public Const UserRoleAdministrator As String = "1"
        Public Const FlowManagr As String = "2"
        Public Const HRMaster As String = "3"
        Public Const UserRoleProjectManager As String = "4"
        Public Const UserRoleProjectChief As String = "5"
        Public Const UserRoleProjectMember As String = "6"
        Public Const UserRoleUser As String = "7"



        Public Const UserEnabledViewAllProject As String = UserRoleAdministrator + "," + FlowManagr + "," + HRMaster '+ "," + UserRoleProjectManager + "," + UserRoleProjectChief + "," + UserRoleProjectMember
        'Public Const UserEnabledModifyAllProject As String =   FlowManagr  
        Public Const UserEnabledModifySelfProject As String = UserRoleProjectManager + "," + UserRoleProjectChief


        Private _name As String
        Private _roleID As Integer

        Public Enum ProjectRoleType
            InitValue
            Read_Project_All
            New_Project_All
            Edit_Project_All
            Delete_Project_All
            Read_Project_Self
            New_Project_Self
            Edit_Project_Self
            Delete_Project_Self
            Read_Doc_All
            New_Doc_All
            Edit_Doc_All
            Delete_Doc_All
            Read_Doc_Self
            New_Doc_Self
            Edit_Doc_Self
            Delete_Doc_Self
            Read_Forum_All
            New_Forum_All
            Edit_Forum_All
            Delete_Forum_All
            Read_Forum_Self
            New_Forum_Self
            Edit_Forum_Self
            Delete_Forum_Self

        End Enum 'ProjectRoleType

        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(ByVal Value As String)
                _name = Value
            End Set
        End Property

        Public Property RoleID() As Integer
            Get
                Return _roleID
            End Get
            Set(ByVal Value As Integer)
                _roleID = Value
            End Set
        End Property

        Public Shared Function GetRolesName(ByVal RoleID As Integer) As String
            Return Convert.ToString(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetRolesName", RoleID))
        End Function

        Public Shared Function GetRoles() As RolesCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListAllRoles")
            Dim roleArray As New RolesCollection

            ' Separate Data into a collection of Roles
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows
                Dim role As New Roles
                role.RoleID = Convert.ToInt32(r("RoleID"))
                role.Name = r("RoleName").ToString()

                roleArray.Add(role)
            Next r
            Return roleArray
        End Function 'GetRoles

        Public Shared Function GetProjectRoles() As RolesCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListProjectRoles")
            Dim roleArray As New RolesCollection

            ' Separate Data into a collection of Roles
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows
                Dim role As New Roles
                role.RoleID = Convert.ToInt32(r("RoleID"))
                role.Name = r("RoleName").ToString()

                roleArray.Add(role)
            Next r
            Return roleArray
        End Function 'GetRoles

        '判斷使用者是否具有修改專案資料的權限
        Public Shared Function isEnabledModifyProject(ByVal _ProjectID As Integer) As Boolean

            If PMSecurity.IsInRole(Roles.FlowManagr) Then
                Return True
            Else
                Return (PMSecurity.IsInRole(Roles.UserEnabledModifySelfProject) And isProjectMember(_ProjectID))
            End If
        End Function


        '判斷使用者是否具有檢視專案的權限
        Public Shared Function isEnabledViewProject(ByVal _ProjectID As Integer) As Boolean

            If PMSecurity.IsInRole(Roles.UserEnabledViewAllProject) Then
                Return True
            Else
                Return isProjectMember(_ProjectID)
            End If

        End Function

        '判斷使用者是否為專案成員
        Public Shared Function isProjectMember(ByVal _ProjectID As Integer) As Boolean
            Return PMUser.IsProjectMember(_ProjectID, PMSecurity.GetUserID)
        End Function
    End Class 'Roles 
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer