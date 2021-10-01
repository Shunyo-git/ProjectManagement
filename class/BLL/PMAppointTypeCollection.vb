Imports System
Imports System.Collections

Namespace HsinYi.EIP.ProjectManagement.BusinessLogicLayer

    Public Class PMAppointTypeCollection
        Inherits ArrayList

        Public Enum AppointTypeFields
            InitValue
        End Enum 'AppointTypeFields

        ' Put new enumerations here

        Public Overloads Sub Sort(ByVal sortField As AppointTypeFields, ByVal isAscending As Boolean)
        End Sub 'Sort

    End Class

End Namespace

