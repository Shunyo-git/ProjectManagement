Imports System
Imports System.Configuration
Imports System.DirectoryServices

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    '*********************************************************************
    '
    ' DirectoryHelper Class
    '
    ' Uses the DirectorySearcher .NET Framework class to retrieve user information fro
    ' the Active Directory or the WindowsSAM store.
    '
    '*********************************************************************

    Public Class DirectoryHelper
        Private Const _activeDirectoryPath As String = "GC://"
        Private Const _filter As String = "(&(ObjectClass=Person)(SAMAccountName={0}))"
        Private Const _windowsSAMPath As String = "WinNT://"

        Private Shared _path As String = [String].Empty

        '*********************************************************************
        '
        ' Methods are all static -- don't allow construction.
        '
        '*********************************************************************

        Private Sub New()
        End Sub 'New

        '*********************************************************************
        '
        ' This does the core directory searching using the filters specified above.
        '
        '*********************************************************************

        Public Shared Function FindUser(ByVal identification As String, ByRef FirstName As String, ByRef LastName As String) As Boolean
            Dim result As Boolean = False

            ' Determine which method to use to retrieve user information
            ' WindowsSAM
            If ConfigurationSettings.AppSettings(Web.Global.CfgKeyUserAcctSource) = "WindowsSAM" Then
                ' Extract the machine or domain name and the user name from the 
                ' identification string
                Dim samPath As String() = identification.Split(New Char() {"\"c})
                _path = _windowsSAMPath + samPath(0)

                Try
                    ' Find the user 
                    Dim entryRoot As New DirectoryEntry(_path)
                    Dim userEntry As DirectoryEntry = entryRoot.Children.Find(samPath(1), "user")
                    LastName = userEntry.Properties("FullName").Value.ToString()
                Catch
                End Try
                result = True

                ' Active Directory
            ElseIf ConfigurationSettings.AppSettings(Web.Global.CfgKeyUserAcctSource) = "ActiveDirectory" Then
                _path = _activeDirectoryPath

                ' Setup the filter
                identification = identification.Substring(identification.LastIndexOf("\") + 1, identification.Length - identification.LastIndexOf("\") - 1)
                Dim userNameFilter As String = String.Format(_filter, identification)


                ' Get a Directory Searcher to the LDAPPath
                Dim searcher As New DirectorySearcher(_path)
                If searcher Is Nothing Then
                    Return False
                End If

                ' Add the propierties that need to be retrieved
                searcher.PropertiesToLoad.Add("givenName")
                searcher.PropertiesToLoad.Add("sn")

                ' Set the filter for the search
                searcher.Filter = userNameFilter

                Try
                    ' Execute the search
                    Dim search As SearchResult = searcher.FindOne()

                    If Not (search Is Nothing) Then
                        FirstName = SearchResultProperty(search, "givenName")
                        LastName = SearchResultProperty(search, "sn")
                        result = True
                    Else
                        result = False
                    End If

                Catch
                End Try
            Else
                ' The user has not choosen an UserAccountSource or UserAccountSource as None
                ' Usernames will be displayed as "Domain/Username"
                result = False
            End If

            Return result
        End Function 'FindUser

        '*********************************************************************
        '
        ' Retrieves the specified property from the SearchResult collection.
        '
        '*********************************************************************

        Private Shared Function SearchResultProperty(ByVal sr As SearchResult, ByVal field As String) As [String]
            If Not (sr.Properties(field) Is Nothing) Then
                Return CType(sr.Properties(field)(0), [String])
            End If

            Return Nothing
        End Function 'SearchResultProperty
    End Class 'DirectoryHelper
End Namespace 'ASPNET.StarterKit.TimeTracker.BusinessLogicLayer