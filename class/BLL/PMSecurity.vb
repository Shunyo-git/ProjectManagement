Imports System
Imports System.Web
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    '*********************************************************************
    '
    ' TTSecurity Class
    '
    ' The TimeTrackerSecurity class encapsulates two helper methods that enable
    ' developers to easily check the role status of the current browser client.
    '
    '*********************************************************************

    Public Class PMSecurity

        '*********************************************************************
        '
        ' TTSecurity.IsInRole() Method
        '
        ' The IsInRole method enables developers to easily check the role
        ' status of the current browser client.
        '
        '*********************************************************************

        Public Shared Function IsInRole(ByVal role As [String]) As Boolean
            Return HttpContext.Current.User.IsInRole(role)
        End Function 'IsInRole

        '*********************************************************************
        '
        ' TTSecurity.Encrypt() Method
        '
        ' The Encrypt method enables Encryt a clean string into hash
        '
        '*********************************************************************
        Public Shared Function Encrypt(ByVal cleanString As String) As String
            Dim clearBytes As [Byte]()
            clearBytes = New UnicodeEncoding().GetBytes(cleanString)
            Dim hashedBytes As [Byte]() = CType(CryptoConfig.CreateFromName("MD5"), HashAlgorithm).ComputeHash(clearBytes)
            Dim hashedText As String = BitConverter.ToString(hashedBytes)
            Return hashedText
        End Function

        Public Shared Function GetUserID() As Integer
            Return CType(HttpContext.Current.User, CustomPrincipal).UserID
        End Function 'GetUserID


        Public Shared Function GetUserRole() As String
            Return CType(HttpContext.Current.User, CustomPrincipal).UserRole
        End Function 'GetUserRole


        Public Shared Function GetName() As String
            Return CType(HttpContext.Current.User, CustomPrincipal).Name
        End Function 'GetName

        'Public Shared Function GetIdentityName() As String
        '    Return CType(HttpContext.Current.User, CustomPrincipal).identityName
        'End Function 'GetName

        Public Shared Function GetDisplayName() As String
            Return CType(HttpContext.Current.User, CustomPrincipal).DisplayName
        End Function 'GetDisplayName

        '*********************************************************************
        '
        ' <summary>
        ' Validates the input text using a Regular Expression and replaces any input expression
        ' characters with empty string.Removes any characters not in [a-zA-Z0-9_]. 
        ' Also removes the exact words script|javascript|vbscript|jscript
        ' <summary>
        ' <remarks>
        ' For a good reference on Regular Expressions, please see
        '	 - http://regexlib.com
        '	 - http://py-howto.sourceforge.net/regex/regex.html
        ' </remarks>
        ' <param name="inputText">The text to validate.</param>
        ' <returns>Sanitized string</returns>
        '
        '*********************************************************************

        Public Shared Function CleanStringRegex(ByVal inputText As String) As String
            Dim options As RegexOptions = RegexOptions.IgnoreCase
            Return ReplaceRegex(inputText, "[^\\\.!?""',\-\w\s@]", options)
        End Function 'CleanStringRegex

        '*********************************************************************
        '
        ' <summary>
        ' Removes designated characters from an input string input text using a Regular Expression.
        ' </summary>
        ' <remarks>
        ' For a good reference on Regular Expressions, please see
        '	 - http://regexlib.com
        '	 - http://py-howto.sourceforge.net/regex/regex.html
        ' </remarks>
        ' <param name="inputText">The text to clean.</param>
        ' <param name="regularExpression">The regular expression</param>
        ' <returns>Sanitized string.</returns>
        '
        '*********************************************************************

        Private Shared Function ReplaceRegex(ByVal inputText As String, ByVal regularExpression As String, ByVal options As RegexOptions) As String
            Dim regex As New regex(regularExpression, options)
            Return regex.Replace(inputText, "")
        End Function 'ReplaceRegex
    End Class 'TTSecurity
End Namespace 'ASPNET.StarterKit.TimeTracker.BusinessLogicLayer