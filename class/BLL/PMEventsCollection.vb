Imports System
Imports System.Collections

Namespace HsinYi.EIP.ProjectManagement.BusinessLogicLayer

    '*********************************************************************
    '
    ' EventsCollection Class 
    '
    '  This is a skeleton class for creating sortable collection.
    '
    '*********************************************************************

    Public Class PMEventsCollection
        Inherits ArrayList
        Public Enum EventFields
            InitValue
            EventID
            ProjectID
            Topic
            EventDate
            CreatedUserID
            CreatedDate
            ModifiedUserID
            ModifiedDate
        End Enum 'EventFields

        Public Overloads Sub Sort(ByVal sortField As EventFields, ByVal isAscending As Boolean)
            Select Case sortField
                Case EventFields.EventID
                    MyBase.Sort(New EventIDComparer)
                Case EventFields.ProjectID
                    MyBase.Sort(New ProjectIDComparer)
                Case EventFields.Topic
                    MyBase.Sort(New TopicComparer)
                Case EventFields.EventDate
                    MyBase.Sort(New EventDateComparer)
                Case EventFields.CreatedUserID
                    MyBase.Sort(New CreatedUserIDComparer)
                Case EventFields.CreatedDate
                    MyBase.Sort(New CreatedDateComparer)
                Case EventFields.ModifiedUserID
                    MyBase.Sort(New ModifiedUserIDComparer)
                Case EventFields.ModifiedDate
                    MyBase.Sort(New ModifiedDateComparer)

            End Select

            If Not isAscending Then
                MyBase.Reverse() ' 反向整個 System.Collections.ArrayList 中元素的順序。  
            End If
        End Sub 'Sort

        Private NotInheritable Class EventIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMEvent = CType(x, PMEvent)
                Dim second As PMEvent = CType(y, PMEvent)
                Return first.EventID.CompareTo(second.EventID)
            End Function 'Compare
        End Class 'EventIDComparer

        Private NotInheritable Class ProjectIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMEvent = CType(x, PMEvent)
                Dim second As PMEvent = CType(y, PMEvent)
                Return first.ProjectID.CompareTo(second.ProjectID)
            End Function 'Compare
        End Class 'ProjectIDComparer

        Private NotInheritable Class TopicComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMEvent = CType(x, PMEvent)
                Dim second As PMEvent = CType(y, PMEvent)
                Return first.Topic.CompareTo(second.Topic)
            End Function 'Compare
        End Class 'TopicComparer
        Private NotInheritable Class EventDateComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMEvent = CType(x, PMEvent)
                Dim second As PMEvent = CType(y, PMEvent)
                Return first.EventDate.CompareTo(second.EventDate)
            End Function 'Compare
        End Class 'EventDateComparer

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


