 

Imports System
Imports System.Collections

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    '*********************************************************************
    '
    ' DocCategoriesCollection Class 
    '
    '  This is a skeleton class for creating sortable collection.
    '
    '*********************************************************************

    Public Class PMDocCategoriesCollection
        Inherits ArrayList

        Public Enum DocCategoriesFields
            InitValue
            CategoryID
            Name
            Sort
        End Enum 'DocCategoriesFields

        ' Put new enumerations here
        Public Overloads Sub Sort(ByVal sortField As DocCategoriesFields, ByVal isAscending As Boolean)
            Select Case sortField
                Case DocCategoriesFields.CategoryID
                    MyBase.Sort(New CategoryIDComparer)
                Case DocCategoriesFields.Name
                    MyBase.Sort(New NameComparer)
                Case DocCategoriesFields.Sort
                    MyBase.Sort(New SortComparer)
               

            End Select

            If Not isAscending Then
                MyBase.Reverse() ' 反向整個 System.Collections.ArrayList 中元素的順序。  
            End If
        End Sub 'Sort

        Private NotInheritable Class CategoryIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMDocCategories = CType(x, PMDocCategories)
                Dim second As PMDocCategories = CType(y, PMDocCategories)
                Return first.CategoryID.CompareTo(second.CategoryID)
            End Function 'Compare
        End Class 'CategoryIDComparer

        Private NotInheritable Class NameComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMDocCategories = CType(x, PMDocCategories)
                Dim second As PMDocCategories = CType(y, PMDocCategories)
                Return first.Name.CompareTo(second.Name)
            End Function 'Compare
        End Class 'NameComparer

        Private NotInheritable Class SortComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMDocCategories = CType(x, PMDocCategories)
                Dim second As PMDocCategories = CType(y, PMDocCategories)
                Return first.Sort.CompareTo(second.Sort)
            End Function 'Compare
        End Class 'SortComparer

    End Class 'DocCategoriesCollection ' Implement Sort method here
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

' Place IComparer implementations here