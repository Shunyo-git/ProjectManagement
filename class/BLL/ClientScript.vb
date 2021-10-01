

Imports System.Web

Namespace HsinYi.EIP.ProjectManagement.Web

    Public Class ClientScript
        Public Overloads Shared Function ShowAlertMsg(ByVal intStatus As Int16, ByVal strMsg As String, ByVal strURL As String) As String
            Dim IntGoPage As String
            Dim strScript As String

            If Not IsNumeric(strURL) Then
                IntGoPage = "-1"
            Else
                IntGoPage = strURL
            End If

            strScript = "<script language='JavaScript'>" & _
                    "alert('" & strMsg & "');"
            If CInt(intStatus) = 1 Then
                strScript = strScript & "	window.history.go(" & IntGoPage & ");"
            ElseIf CInt(intStatus) = 2 Then
                strScript = strScript & "	location.href='" & strURL & "'; "
            ElseIf CInt(intStatus) = 3 Then
                strScript = strScript & " window.close(); "

            End If
            strScript = strScript & "</script>"

            Return strScript


        End Function

        Public Overloads Shared Function ShowAlertMsg(ByVal strMsg As String) As String
            Dim IntGoPage As String
            Dim strScript As String


            strScript = "<script language='JavaScript'>" & _
                        "alert('" & strMsg & "');" & _
                        "</script>"

            Return strScript


        End Function
        Public Shared Function ShowScript(ByVal strMsg As String) As String
            Dim IntGoPage As String
            Dim strScript As String


            strScript = "<script language='JavaScript'>" & _
                        strMsg & _
                        "</script>"

            Return strScript


        End Function

        Public Shared Function FilterInput(ByVal strSQL As String) As String
            strSQL = strSQL.Replace("'", "''")
            strSQL = strSQL.Replace("[", "[[]")
            strSQL = strSQL.Replace("%", "[%]")
            strSQL = strSQL.Replace("-", "[-]")
            strSQL = strSQL.Replace("_", "[_]")
            Return strSQL
        End Function
        Public Shared Function CRtoBR(ByVal strVal As String) As String
            Return Strings.Replace(strVal, vbCrLf, "<BR>")

        End Function
    End Class

End Namespace

