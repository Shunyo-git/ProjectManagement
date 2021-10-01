Imports System
Imports System.Collections

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    Public Class PMAttachmentCollection
        Inherits ArrayList

        Public Enum AttachmentFields
            InitValue
            EventID
            ProjectID
            AttachmentID
            AttachmentName
            CreatedDate
            ModifiedDate
        End Enum 'AttachmentFields

        ' Put new enumerations here

        Public Overloads Sub Sort(ByVal sortField As AttachmentFields, ByVal isAscending As Boolean)
            Select Case sortField
                Case AttachmentFields.EventID
                    MyBase.Sort(New EventIDComparer)
                Case AttachmentFields.ProjectID
                    MyBase.Sort(New ProjectIDComparer)
                Case AttachmentFields.AttachmentID
                    MyBase.Sort(New AttachmentIDComparer)
                Case AttachmentFields.AttachmentName
                    MyBase.Sort(New AttachmentNameComparer)
                Case AttachmentFields.CreatedDate
                    MyBase.Sort(New CreatedDateComparer)
                Case AttachmentFields.ModifiedDate
                    MyBase.Sort(New ModifiedDateComparer)
            End Select

            If Not isAscending Then
                MyBase.Reverse() ' 反向整個 System.Collections.ArrayList 中元素的順序。  
            End If
        End Sub 'Sort

        Private NotInheritable Class EventIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMAttachment = CType(x, PMAttachment)
                Dim second As PMAttachment = CType(y, PMAttachment)
                Return first.EventID.CompareTo(second.EventID)
            End Function 'Compare
        End Class 'EventIDComparer

        Private NotInheritable Class ProjectIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMAttachment = CType(x, PMAttachment)
                Dim second As PMAttachment = CType(y, PMAttachment)
                Return first.ProjectID.CompareTo(second.ProjectID)
            End Function 'Compare
        End Class 'ProjectIDComparer

        Private NotInheritable Class AttachmentIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMAttachment = CType(x, PMAttachment)
                Dim second As PMAttachment = CType(y, PMAttachment)
                Return first.AttachmentID.CompareTo(second.AttachmentID)
            End Function 'Compare
        End Class 'AttachmentIDComparer

        Private NotInheritable Class AttachmentNameComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMAttachment = CType(x, PMAttachment)
                Dim second As PMAttachment = CType(y, PMAttachment)
                Return first.AttachmentName.CompareTo(second.AttachmentName)
            End Function 'Compare
        End Class 'AttachmentNameComparer

        Private NotInheritable Class CreatedDateComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMAttachment = CType(x, PMAttachment)
                Dim second As PMAttachment = CType(y, PMAttachment)
                Return first.CreatedDate.CompareTo(second.CreatedDate)
            End Function 'Compare
        End Class 'CreatedDateComparer

        Private NotInheritable Class ModifiedDateComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As PMAttachment = CType(x, PMAttachment)
                Dim second As PMAttachment = CType(y, PMAttachment)
                Return first.ModifiedDate.CompareTo(second.ModifiedDate)
            End Function 'Compare
        End Class 'ModifiedDateComparer

    End Class 'AttachmentCollection

End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
