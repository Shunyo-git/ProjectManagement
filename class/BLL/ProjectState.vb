 
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

    Public Class ProjectState
        Private _stateName As String
        Private _stateID As Integer

        Public Property Name() As String
            Get
                Return _stateName
            End Get
            Set(ByVal Value As String)
                _stateName = Value
            End Set
        End Property

        Public Property StateID() As Integer
            Get
                Return _stateID
            End Get
            Set(ByVal Value As Integer)
                _stateID = Value
            End Set
        End Property

        Public Shared Function GetStateName(ByVal StateID As Integer) As String
            Return Convert.ToString(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetStateName", StateID))
        End Function
        Public Shared Function GetStates() As RolesCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListAllStates")
            Dim StateArray As New RolesCollection

            ' Separate Data into a collection of ProjectStates
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows
                Dim state As New ProjectState
                state.StateID = Convert.ToInt32(r("StateID"))
                state.Name = r("StateName").ToString()

                StateArray.Add(state)
            Next r
            Return StateArray
        End Function 'GetStates

        Public Shared Function ShowStateColor(ByVal intStateID As Integer) As Color
            Select Case intStateID
                Case 1
                    Return Color.SlateBlue
                Case 2
                    Return Color.Orange
                Case 3
                    Return Color.RoyalBlue
                Case 4
                    Return Color.Crimson
                Case Else
                    Return Color.Black
            End Select

        End Function

      
    End Class 'Roles 
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer