Imports System
Imports System.Collections

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    '*********************************************************************
    '
    ' UsersCollection Class 
    '
    ' The UsersCollection is a Custom User collection used 
    ' to represet a list of User objects.
    '
    '
    '*********************************************************************

    Public Class UsersCollection
        Inherits ArrayList

        Public Enum UserFields
            InitValue
            Name
            Description
            RoleName
            SourceUserName
            SourceUserDepartment
            SourceUSerTitle

        End Enum 'UserFields

        Public Overloads Sub Sort(ByVal sortField As UserFields, ByVal isAscending As Boolean)
            Select Case sortField
                Case UserFields.Name
                    MyBase.Sort(New NameComparer)
                Case UserFields.RoleName
                    MyBase.Sort(New RoleNameComparer)
            End Select
            If Not isAscending Then
                MyBase.Reverse()
            End If
        End Sub 'Sort

        Private NotInheritable Class NameComparer
            Implements IComparer 'ToDo: Add Implements Clauses for implementation methods of these interface(s)

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
                Dim first As PMUser = CType(x, PMUser)
                Dim second As PMUser = CType(y, PMUser)
                Return first.Name.CompareTo(second.Name)
            End Function 'Compare
        End Class 'NameComparer

        Private NotInheritable Class RoleNameComparer
            Implements IComparer 'ToDo: Add Implements Clauses for implementation methods of these interface(s)

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
                Dim first As PMUser = CType(x, PMUser)
                Dim second As PMUser = CType(y, PMUser)
                Return first.RoleName.CompareTo(second.RoleName)
            End Function 'Compare
        End Class 'RoleNameComparer
    End Class 'UsersCollection
End Namespace 'ASPNET.StarterKit.TimeTracker.BusinessLogicLayer