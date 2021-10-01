Imports System

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    '*********************************************************************
    '
    ' TabItem Class
    '
    ' TabItem class used to represent one tab in the Banner.ascx user control
    ' or one tab in the AdminTabs.ascx user control.  A class is used instead
    ' of a struct so a parameterized constructor could be used.
    '
    '*********************************************************************

    Public Class TabItem
        Private _name As String
        Private _path As String

        Public Enum ProjectTabIndex
            ProjectDetail
            ProjectTimeline
            ProjectGroup
            ProjectUser
            ProjectCalendar
            ProjectEvent
            ProjectDocument
            ProjectForum
            ProjectContact
        End Enum

        Public Enum CategoryTabIndex
            Plan
            Timeline
            SOP
            Record
            Form
            Reference
            CaseResult
            Benefits
        End Enum

        Public Sub New(ByVal newName As String, ByVal newPath As String)
            _name = newName
            _path = newPath
        End Sub 'New

        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(ByVal Value As String)
                _name = Value
            End Set
        End Property

        Public Property Path() As String
            Get
                Return _path
            End Get
            Set(ByVal Value As String)
                _path = Value
            End Set
        End Property
    End Class 'TabItem
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer