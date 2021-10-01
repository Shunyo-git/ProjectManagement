Imports System
Imports System.Collections

Namespace HsinYi.EIP.ProjectManagement.BusinessLogicLayer

    Public Class PMAppointCollection
        Inherits ArrayList
        Public Enum EventFields
            InitValue
            AppointID
            AppointTypeID
            AppointLevelID
            ProjectID
            Subject
            Location
            AppointDate
            StartTime
            EndTime
            CreatedUserID
            CreatedDate
            ModifiedUserID
            ModifiedDate
        End Enum 'EventFields
    End Class

End Namespace

