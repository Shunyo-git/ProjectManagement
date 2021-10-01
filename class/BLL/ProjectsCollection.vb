Imports System
Imports System.Collections

Namespace HsinYi.EIP.ProjectManagement.BusinessLogicLayer

    Public Class ProjectsCollection
        Inherits ArrayList
        Public Enum ProjectFields
            InitValue
            ProjectName
            Manager
            ProjectChief
            StartDate
            EndDate
            State
        End Enum 'ProjectFields


        Public Overloads Sub Sort(ByVal sortField As ProjectFields, ByVal isAscending As Boolean)
            Select Case sortField
                Case ProjectFields.ProjectName
                    MyBase.Sort(New ProjectNameComparer)
                Case ProjectFields.Manager
                    MyBase.Sort(New ManagerComparer)
                Case ProjectFields.ProjectChief
                    MyBase.Sort(New ProjectChiefComparer)
                Case ProjectFields.StartDate
                    MyBase.Sort(New StartDateComparer)
                Case ProjectFields.EndDate
                    MyBase.Sort(New EndDateComparer)
                Case ProjectFields.State
                    MyBase.Sort(New StateComparer)
            End Select

            If Not isAscending Then
                MyBase.Reverse() ' 反向整個 System.Collections.ArrayList 中元素的順序。  
            End If
        End Sub 'Sort

        Private NotInheritable Class ProjectNameComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As Project = CType(x, Project)
                Dim second As Project = CType(y, Project)
                Return first.ProjectName.CompareTo(second.ProjectName)
            End Function 'Compare
        End Class 'NameComparer

        Private NotInheritable Class ManagerComparer
            Implements IComparer

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As Project = CType(x, Project)
                Dim second As Project = CType(y, Project)
                Return first.ManagerUserName.CompareTo(second.ManagerUserName)
            End Function 'Compare
        End Class 'ManagerComparer

        Private NotInheritable Class ProjectChiefComparer
            Implements IComparer

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As Project = CType(x, Project)
                Dim second As Project = CType(y, Project)
                Return first.ChiefName.CompareTo(second.ChiefName)
            End Function 'Compare
        End Class 'ProjectChiefComparer


        Private NotInheritable Class StartDateComparer
            Implements IComparer

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As Project = CType(x, Project)
                Dim second As Project = CType(y, Project)
                Return first.StartDate.CompareTo(second.StartDate)
            End Function 'Compare
        End Class 'StartDateComparer

        Private NotInheritable Class EndDateComparer
            Implements IComparer

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As Project = CType(x, Project)
                Dim second As Project = CType(y, Project)
                Return first.EndDate.CompareTo(second.EndDate)
            End Function 'Compare
        End Class 'EndDateComparer

        Private NotInheritable Class StateComparer
            Implements IComparer

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As Project = CType(x, Project)
                Dim second As Project = CType(y, Project)
                Return first.StateID.CompareTo(second.StateID)
            End Function 'Compare
        End Class 'ProjectChiefCompar

    End Class 'ProjectsCollection
End Namespace 'EIP.ProjectManagement.BusinessLogicLayer
