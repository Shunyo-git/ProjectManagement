Imports System
Imports System.Collections

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    '*********************************************************************
    '
    ' RolesCollection Class 
    '
    '  This is a skeleton class for creating sortable collection.
    '
    '*********************************************************************

    Public Class RolesCollection
        Inherits ArrayList

        Public Enum RolesFields
            InitValue
        End Enum 'RolesFields

        ' Put new enumerations here

        Public Overloads Sub Sort(ByVal sortField As RolesFields, ByVal isAscending As Boolean)
        End Sub 'Sort
    End Class 'RolesCollection ' Implement Sort method here
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

' Place IComparer implementations here