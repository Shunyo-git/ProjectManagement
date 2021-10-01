 
Imports System
Imports System.Web
Imports System.Web.Mail

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
    Public Class PMSendMail
        Private m_strBody As String
        Private m_strFromName As String
        Private m_strFrom As String
        Private m_strRecipt As String
        Private m_intType As String
        Private m_strSubject As String
        Private m_Attachments As String
        Private m_BCC As String
        Private m_CC As String
        Private m_server As String = ConfigurationSettings.AppSettings(Web.Global.CfgKeySMTPServer)
        Dim Message As System.Web.Mail.MailMessage
        '--------------------------------------------------------	
        Public Property Body()
            Get
                Return m_strBody
            End Get

            Set(ByVal Value)
                m_strBody = Value
            End Set
        End Property
        '--------------------------------------------------------	
        Public Property Subject()
            Get
                Return m_strSubject
            End Get
            Set(ByVal Value)
                m_strSubject = Value
            End Set
        End Property
        Public Property From()
            Get
                Return m_strFrom
            End Get
            Set(ByVal Value)
                m_strFrom = Value
            End Set
        End Property
        Public Property FromName()
            Get
                Return m_strFromName
            End Get
            Set(ByVal Value)
                m_strFromName = Value
            End Set
        End Property
        Public Property Recipt()
            Get
                Return m_strRecipt
            End Get
            Set(ByVal Value)
                m_strRecipt = Value
            End Set
        End Property
        Public Property BCC()
            Get
                Return m_BCC
            End Get
            Set(ByVal Value)
                m_BCC = Value
            End Set
        End Property
        Public Property CC()
            Get
                Return m_CC
            End Get
            Set(ByVal Value)
                m_CC = Value
            End Set
        End Property
        Public Property SMTPServer()
            Get
                Return m_server
            End Get
            Set(ByVal Value)
                m_server = Value
            End Set
        End Property
        Public Sub AddAttachments(ByVal Value As String)
            Dim MailAttachment As New System.Web.Mail.MailAttachment(Value)
            Message.Attachments.Add(MailAttachment)
            MailAttachment = Nothing
        End Sub
        Sub New()
            Message = New System.Web.Mail.MailMessage
        End Sub
        Public Sub send()



            Dim objMail As System.Web.Mail.SmtpMail


            Message.To = m_strRecipt

            Message.From = m_strFrom
            Message.Subject = m_strSubject
            Message.Body = m_strBody

            If Strings.Trim(m_CC) <> "" Then
                Message.Cc = m_CC
            End If

            If Strings.Trim(m_BCC) <> "" Then
                Message.Bcc = m_strFrom & ";" & m_BCC
            Else
                Message.Bcc = m_strFrom
            End If


            objMail.SmtpServer = m_server

            objMail.Send(Message)


            'objMail.SmtpServer = "192.168.1.1"

            'objMail.Send( _
            '   m_strFrom, m_strRecipt, m_strSubject, m_strBody)

        End Sub

    End Class
End Namespace


