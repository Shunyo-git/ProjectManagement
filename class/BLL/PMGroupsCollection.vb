 
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

    Public Class GroupsCollection
        Inherits ArrayList

        Public Enum GroupsFields
            InitValue
            GroupID
            Name
            Description
            ProjectID
        End Enum 'GroupsFields

        ' Put new enumerations here

        Public Overloads Sub Sort(ByVal sortField As GroupsFields, ByVal isAscending As Boolean)
            Select Case sortField
                Case GroupsFields.Name
                    MyBase.Sort(New NameComparer)
                Case GroupsFields.Description
                    MyBase.Sort(New DescriptionComparer)
                Case GroupsFields.GroupID
                    MyBase.Sort(New GroupIDComparer)
                Case GroupsFields.ProjectID
                    MyBase.Sort(New ProjectIDComparer)
            End Select

            If Not isAscending Then
                MyBase.Reverse() ' 反向整個 System.Collections.ArrayList 中元素的順序。  
            End If
        End Sub 'Sort

        Private NotInheritable Class ProjectIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMGroup = CType(x, PMGroup)
                Dim second As PMGroup = CType(y, PMGroup)
                Return first.ProjectID.CompareTo(second.ProjectID)
            End Function 'Compare
        End Class 'ProjectIDComparer

        Private NotInheritable Class GroupIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMGroup = CType(x, PMGroup)
                Dim second As PMGroup = CType(y, PMGroup)
                Return first.GroupID.CompareTo(second.GroupID)
            End Function 'Compare
        End Class 'GroupIDComparer

        Private NotInheritable Class NameComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMGroup = CType(x, PMGroup)
                Dim second As PMGroup = CType(y, PMGroup)
                Return first.Name.CompareTo(second.Name)
            End Function 'Compare
        End Class 'NameComparer

        Private NotInheritable Class DescriptionComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMGroup = CType(x, PMGroup)
                Dim second As PMGroup = CType(y, PMGroup)
                Return first.Description.CompareTo(second.Description)
            End Function 'Compare
        End Class 'DescriptionComparer
    End Class 'GroupsCollection ' Implement Sort method here
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

' Place IComparer implementations here
