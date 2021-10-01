 


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

    Public Class ProjectType
        Private _typeName As String
        Private _typeID As Integer

        Public Property Name() As String
            Get
                Return _typeName
            End Get
            Set(ByVal Value As String)
                _typeName = Value
            End Set
        End Property

        Public Property TypeID() As Integer
            Get
                Return _typeID
            End Get
            Set(ByVal Value As Integer)
                _typeID = Value
            End Set
        End Property

        Public Shared Function GetTypeName(ByVal TypeID As Integer) As String
            Return Convert.ToString(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetTypeName", TypeID))
        End Function
        Public Shared Function GetTypes() As RolesCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListAllTypes")
            Dim StateArray As New RolesCollection

            ' Separate Data into a collection of ProjectStates
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows
                Dim ProjectType As New ProjectType
                ProjectType.TypeID = Convert.ToInt32(r("TypeID"))
                ProjectType.Name = r("TypeName").ToString()

                StateArray.Add(ProjectType)
            Next r
            Return StateArray
        End Function 'GetTypes
    End Class 'ProjectType 
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer