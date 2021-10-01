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

    Public Class PMForumTopicCollection
        Inherits ArrayList
        Public Enum TopicFields
            InitValue
            TopicID
            ProjectID
            Subject
            CreatedUserID
            CreatedDate
            ModifiedUserID
            ModifiedDate
            LastPostDate
            Replies
        End Enum 'TopicFields

        Public Overloads Sub Sort(ByVal sortField As TopicFields, ByVal isAscending As Boolean)
            Select Case sortField
                Case TopicFields.TopicID
                    MyBase.Sort(New TopicIDComparer)
                Case TopicFields.ProjectID
                    MyBase.Sort(New ProjectIDComparer)
                Case TopicFields.Subject
                    MyBase.Sort(New SubjectComparer)
                Case TopicFields.CreatedUserID
                    MyBase.Sort(New CreatedUserIDComparer)
                Case TopicFields.CreatedDate
                    MyBase.Sort(New CreatedDateComparer)
                Case TopicFields.ModifiedUserID
                    MyBase.Sort(New ModifiedUserIDComparer)
                Case TopicFields.ModifiedDate
                    MyBase.Sort(New ModifiedDateComparer)
                Case TopicFields.LastPostDate
                    MyBase.Sort(New LastPostDateComparer)
                Case TopicFields.Replies
                    MyBase.Sort(New RepliesComparer)

            End Select

            If Not isAscending Then
                MyBase.Reverse() ' 反向整個 System.Collections.ArrayList 中元素的順序。  
            End If
        End Sub 'Sort

        Private NotInheritable Class LastPostDateComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMForumTopic = CType(x, PMForumTopic)
                Dim second As PMForumTopic = CType(y, PMForumTopic)
                Return first.LastPostDate.CompareTo(second.LastPostDate)
            End Function 'Compare
        End Class 'LastPostDateComparer

        Private NotInheritable Class RepliesComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMForumTopic = CType(x, PMForumTopic)
                Dim second As PMForumTopic = CType(y, PMForumTopic)
                Return first.Replies.CompareTo(second.Replies)
            End Function 'Compare
        End Class 'RepliesComparer

        Private NotInheritable Class TopicIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMForumTopic = CType(x, PMForumTopic)
                Dim second As PMForumTopic = CType(y, PMForumTopic)
                Return first.TopicID.CompareTo(second.TopicID)
            End Function 'Compare
        End Class 'TopicIDComparer

        Private NotInheritable Class ProjectIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMForumTopic = CType(x, PMForumTopic)
                Dim second As PMForumTopic = CType(y, PMForumTopic)
                Return first.ProjectID.CompareTo(second.ProjectID)
            End Function 'Compare
        End Class 'ProjectIDComparer

        Private NotInheritable Class SubjectComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMForumTopic = CType(x, PMForumTopic)
                Dim second As PMForumTopic = CType(y, PMForumTopic)
                Return first.Subject.CompareTo(second.Subject)
            End Function 'Compare
        End Class 'SubjectComparer


        Private NotInheritable Class CreatedUserIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMForumTopic = CType(x, PMForumTopic)
                Dim second As PMForumTopic = CType(y, PMForumTopic)
                Return first.CreatedUserID.CompareTo(second.CreatedUserID)
            End Function 'Compare
        End Class 'CreatedUserIDComparer

        Private NotInheritable Class CreatedDateComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMForumTopic = CType(x, PMForumTopic)
                Dim second As PMForumTopic = CType(y, PMForumTopic)
                Return first.CreatedDate.CompareTo(second.CreatedDate)
            End Function 'Compare
        End Class 'CreatedDateComparer
        Private NotInheritable Class ModifiedUserIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMForumTopic = CType(x, PMForumTopic)
                Dim second As PMForumTopic = CType(y, PMForumTopic)
                Return first.ModifiedUserID.CompareTo(second.ModifiedUserID)
            End Function 'Compare
        End Class 'ModifiedUserIDComparer

        Private NotInheritable Class ModifiedDateComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMForumTopic = CType(x, PMForumTopic)
                Dim second As PMForumTopic = CType(y, PMForumTopic)
                Return first.ModifiedDate.CompareTo(second.ModifiedDate)
            End Function 'Compare
        End Class 'ModifiedDateComparer
    End Class
End Namespace

