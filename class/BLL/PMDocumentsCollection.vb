 

Imports System
Imports System.Collections

Namespace HsinYi.EIP.ProjectManagement.BusinessLogicLayer
    Public Class PMDocumentsCollection
        Inherits ArrayList
        Public Enum DocumentFields
            InitValue
            DocumentID
            ProjectID
            CategoryID
            Title
            Description
            CreatedUserID
            CreatedDate
            ModifiedUserID
            ModifiedDate
        End Enum 'DocumentFields

        Public Overloads Sub Sort(ByVal sortField As DocumentFields, ByVal isAscending As Boolean)
            Select Case sortField
                Case DocumentFields.DocumentID
                    MyBase.Sort(New DocumentIDComparer)
                Case DocumentFields.ProjectID
                    MyBase.Sort(New ProjectIDComparer)
                Case DocumentFields.CategoryID
                    MyBase.Sort(New CategoryIDComparer)
                Case DocumentFields.Title
                    MyBase.Sort(New TitleComparer)
                
                Case DocumentFields.CreatedUserID
                    MyBase.Sort(New CreatedUserIDComparer)
                Case DocumentFields.CreatedDate
                    MyBase.Sort(New CreatedDateComparer)
                Case DocumentFields.ModifiedUserID
                    MyBase.Sort(New ModifiedUserIDComparer)
                Case DocumentFields.ModifiedDate
                    MyBase.Sort(New ModifiedDateComparer)

            End Select

            If Not isAscending Then
                MyBase.Reverse() ' 反向整個 System.Collections.ArrayList 中元素的順序。  
            End If
        End Sub 'Sort

        Private NotInheritable Class DocumentIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMDocuments = CType(x, PMDocuments)
                Dim second As PMDocuments = CType(y, PMDocuments)
                Return first.DocumentID.CompareTo(second.DocumentID)
            End Function 'Compare
        End Class 'DocumentIDComparer

        Private NotInheritable Class ProjectIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMDocuments = CType(x, PMDocuments)
                Dim second As PMDocuments = CType(y, PMDocuments)
                Return first.ProjectID.CompareTo(second.ProjectID)
            End Function 'Compare
        End Class 'ProjectIDComparer

        Private NotInheritable Class CategoryIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMDocuments = CType(x, PMDocuments)
                Dim second As PMDocuments = CType(y, PMDocuments)
                Return first.CategoryID.CompareTo(second.CategoryID)
            End Function 'Compare
        End Class 'CategoryIDComparer
        Private NotInheritable Class TitleComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMDocuments = CType(x, PMDocuments)
                Dim second As PMDocuments = CType(y, PMDocuments)
                Return first.Title.CompareTo(second.Title)
            End Function 'Compare
        End Class 'TitleComparer

        Private NotInheritable Class CreatedUserIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMEvent = CType(x, PMEvent)
                Dim second As PMEvent = CType(y, PMEvent)
                Return first.CreatedUserID.CompareTo(second.CreatedUserID)
            End Function 'Compare
        End Class 'CreatedUserIDComparer

        Private NotInheritable Class CreatedDateComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMEvent = CType(x, PMEvent)
                Dim second As PMEvent = CType(y, PMEvent)
                Return first.CreatedDate.CompareTo(second.CreatedDate)
            End Function 'Compare
        End Class 'CreatedDateComparer
        Private NotInheritable Class ModifiedUserIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMEvent = CType(x, PMEvent)
                Dim second As PMEvent = CType(y, PMEvent)
                Return first.ModifiedUserID.CompareTo(second.ModifiedUserID)
            End Function 'Compare
        End Class 'ModifiedUserIDComparer

        Private NotInheritable Class ModifiedDateComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMEvent = CType(x, PMEvent)
                Dim second As PMEvent = CType(y, PMEvent)
                Return first.ModifiedDate.CompareTo(second.ModifiedDate)
            End Function 'Compare
        End Class 'ModifiedDateComparer
    End Class
End Namespace


