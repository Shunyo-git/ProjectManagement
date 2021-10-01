Imports System
Imports System.Security.Principal

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    '*********************************************************************
    '
    ' CustomPrincipal Class
    '
    ' The CustomPrincipal class implements the IPrincipal interface so it 
    ' can be used in place of the GenericPrincipal object.  Requirements for 
    ' implementing the IPrincipal interface include implementing the
    ' IIdentity interface and an implementation for IsInRole.  The custom
    ' principal is attached to the current request in Global.asax in the
    ' Authenticate_Request event handler.  The user's role is stored in the
    ' custom principal object in the Global_AcquireRequestState event handler.
    ' 
    '*********************************************************************

    Public Class CustomPrincipal
        Implements IPrincipal

        Private _userID As Integer
        Private _userRole As String = [String].Empty
        Private _name As String
        'Private _identityName As String
        Private _displayName As String
        ' Required to implement the IPrincipal interface.
        Protected _Identity As IIdentity

        Public Sub New()
        End Sub 'New

        Public Sub New(ByVal identity As IIdentity, ByVal userID As Integer, ByVal userRole As String, ByVal name As String, ByVal displayName As String)
            _Identity = identity
            _userID = userID
            _userRole = userRole
            _name = name
            _displayName = displayName

        End Sub 'New

        ' IIdentity property used to retrieve the Identity object attached to
        ' this principal.

        Public ReadOnly Property Identity() As IIdentity Implements IPrincipal.Identity
            Get
                Return _Identity
            End Get
            'Set(ByVal Value As IIdentity)
            '    _Identity = Value
            'End Set
        End Property
        Public ReadOnly Property DisplayName() As String
            Get
                Return _displayName
            End Get

        End Property

        ' The user's ID, created when the user was inserted into the database

        Public Property UserID() As Integer
            Get
                Return _userID
            End Get
            Set(ByVal Value As Integer)
                _userID = Value
            End Set
        End Property

        ' The user's role, as defined in TTUser.

        Public Property UserRole() As String
            Get
                Return _userRole
            End Get
            Set(ByVal Value As String)
                _userRole = Value
            End Set
        End Property

        ' The user's name (either First and Last name, or their network username)

        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(ByVal Value As String)
                _name = Value
            End Set
        End Property


        'Public Property DisplayName() As String
        '    Get
        '        Return _displayName
        '    End Get
        '    Set(ByVal Value As String)
        '        _displayName = Value
        '    End Set
        'End Property

        '*********************************************************************
        '
        ' Checks to see if the current user is a member of AT LEAST ONE of
        ' the roles in the role string.  Returns true if found, otherwise false.
        ' role is a comma-delimited list of role IDs.
        '
        '*********************************************************************

        Public Function IsInRole(ByVal role As String) As Boolean Implements IPrincipal.IsInRole
            Dim roleArray As String() = role.Split(New Char() {","c})

            Dim r As String
            For Each r In roleArray
                If _userRole = r Then
                    Return True
                End If
            Next r
            Return False
        End Function 'IsInRole
    End Class 'CustomPrincipal
End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer