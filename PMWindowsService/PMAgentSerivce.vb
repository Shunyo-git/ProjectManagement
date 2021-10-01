Imports System
Imports System.Diagnostics
Imports System.ServiceProcess
Imports System.Timers
Imports System.Configuration
Imports System.Globalization
Imports EIPSysSecurity
Imports HsinYi.EIP.ProjectManagement.BusinessLogicLayer
Imports HsinYi.EIP.ProjectManagement.DataAccessLayer


Public Class PMAgentSerivce
    Inherits System.ServiceProcess.ServiceBase

#Region " ����]�p�u�㲣�ͪ��{���X "

    Public Sub New()
        MyBase.New()

        '��������]�p�u��һݪ��I�s�C
        InitializeComponent()

        '�b InitializeComponent() �I�s����[�J�Ҧ�����l�Ƨ@�~
        AgentTimer = New System.Timers.Timer
        AgentTimer.Interval = _TimeInterval

        AddHandler AgentTimer.Elapsed, AddressOf OnTimer
    End Sub

    'UserService �мg Dispose �H�M������M��C
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' �B�z�Ǫ��D�i�J�I
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' �b�P�@�ӳB�z�Ǥ��i�H����h�� NT �A��; �Y�n�b�o�ӳB�z�Ǥ�
        ' �[�J�t�@���A�ȡA�Эק�H�U�U��
        ' �H�إ߲ĤG�ӪA�Ȫ���C�Ҧp�A
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New PMAgentSerivce}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    '������]�p�u�㪺���n��
    Private components As System.ComponentModel.IContainer

    ' �`�N: �H�U������]�p�u��һݪ��{��
    '�z�i�H�ϥΤ���]�p�u��i��ק�A
    '�ФŨϥε{���X�s�边�i��ק�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        '
        'PMAgentSerivce
        '
        Me.ServiceName = "PMAgentSerivce"

    End Sub

#End Region

    Protected AgentTimer As System.Timers.Timer
    Private _LogName As String = "Hsin-Yi Service"
    Private _EventSourceName As String = "�M�׺޲z�t��"
    Private _TimeSecInterval As Integer = 1000
    Private _TimeMinInterval As Integer = 1000 * 60
    Private _TimeHourInterval As Integer = _TimeMinInterval * 60
    Private _TimeInterval As Integer = _TimeHourInterval * 1 '�ɶ��������j
    Private _LastRunTime As Date = Date.Now

    Protected Overrides Sub OnStart(ByVal args() As String)
        '  �ҰʪA�Ȫ��u�@�C
      
 
        WriteEventLog("�H�˱M�׺޲z�t�ΥN�z�A�� �A�Ȥw�Ұ�" & vbCrLf & "�U�������ɶ��G" & Date.Now.AddMilliseconds(_TimeInterval).ToString, EventLogEntryType.Information)
        AgentTimer.Enabled = True

    End Sub

    Protected Overrides Sub OnStop()
        '  ����A�ȩҰ��檺�u�@�C
        WriteEventLog("�H�˱M�׺޲z�t�ΥN�z�A�� �A�Ȥv����", EventLogEntryType.Warning)
        AgentTimer.Enabled = False
    End Sub

    Protected Overrides Sub OnPause()
        '  �Ȱ��A�ȩҰ��檺�u�@�C
        WriteEventLog("�H�˱M�׺޲z�t�ΥN�z�A�� �A�Ȥw�Ȱ�", EventLogEntryType.Warning)
        AgentTimer.Enabled = False
    End Sub

    Protected Overrides Sub OnContinue()
        '  �~�A�ȩҰ��檺�u�@�C
        WriteEventLog("�H�˱M�׺޲z�t�ΥN�z�A�� �A�Ȥw�~��", EventLogEntryType.Information)
        AgentTimer.Enabled = True
    End Sub

    Protected Sub OnTimer(ByVal source As Object, ByVal e As ElapsedEventArgs)
        '  �p�ɾ��޵oElapsed �ƥ�ɪ��u�@�C

        If Day(_LastRunTime) < Day(Date.Now) Then
            LoadProjects()
        Else
            WriteEventLog("�H�˱M�׺޲z�t�ΥN�z�A�� �A�Ȱ��� " & vbCrLf & "�U�������ɶ��G" & Date.Now.AddMilliseconds(_TimeInterval).ToString, EventLogEntryType.Information)
        End If

    End Sub

    Private Sub WriteEventLog(ByVal strMsg As String, ByVal EventLogEntryType As EventLogEntryType)
        '�O���ƥ��t�Ψƥ�ï�C
        Dim MyLog As New EventLog

        If Not MyLog.SourceExists(_EventSourceName) Then
            MyLog.CreateEventSource(_EventSourceName, _LogName)
            ' Create Log
        End If
        MyLog.Source = _EventSourceName
        ' Write to the Log
        MyLog.WriteEntry(_EventSourceName, strMsg, EventLogEntryType)

    End Sub

    Private Sub LoadProjects()
        'Ū���M��
        'WriteEventLog("�s�u�r�� :" & ConfigurationSettings.AppSettings("PMConnectionString"), EventLogEntryType.Information)
        Dim blnResult As Boolean

        Try

            Dim projects As ProjectsCollection = Project.GetProjects
            Dim pj As Project
            Dim strMailTo As String = String.Empty
            Dim strVal As String = String.Empty
            For Each pj In projects


                strVal = WriteAppoint(pj.projectID, pj.ProjectName, Date.Now, PMAppoint.TimeOfAppoint.Day)
                strMailTo = getMemberEmail(pj.projectID)
                If strVal.Length > 0 Then
                    SendMail(pj.ProjectName & " - ����M�רƶ��q��", strMailTo, strVal)
                    WriteEventLog("�H�e����M�רƶ��q�� " & vbCrLf & "�M�צW�� : " & pj.ProjectName & vbCrLf & "�M�צ��� : " & strMailTo, EventLogEntryType.Information)
                End If

                If Date.Now.DayOfWeek = DayOfWeek.Friday Then
                    strVal = WriteAppoint(pj.projectID, pj.ProjectName, Date.Now.AddDays(3), PMAppoint.TimeOfAppoint.Week)
                    If strVal.Length > 0 Then
                        SendMail(pj.ProjectName & " -�U�g�M�רƶ��q��", strMailTo, strVal)
                        WriteEventLog("�H�e�U�g�M�רƶ��q�� " & vbCrLf & "�M�צW�� : " & pj.ProjectName & vbCrLf & "�M�צ��� : " & strMailTo, EventLogEntryType.Information)
                    End If
                End If

            Next

            blnResult = True
        Catch ex As Exception
            WriteEventLog("Ū���M�צ�ƾ�o�Ϳ��~ :" & ex.Message, EventLogEntryType.Error)
            blnResult = False
        End Try

        If blnResult Then
            _LastRunTime = Date.Now
        End If


    End Sub

    Private Function getMemberEmail(ByVal projectID As Integer) As String
        'Ū���M�צ���Email�b��
        Dim strVal As String = String.Empty
        Dim user As PMUser

        For Each user In PMUser.GetMembers(projectID)
            If strVal.Length > 0 Then strVal += ";"
            strVal += user.UserName & "@hsin-yi.org.tw"
        Next
        Return strVal
    End Function

    Public Function WriteAppoint(ByVal ProjectID As Integer, ByVal ProjectName As String, ByVal theDate As Date, ByVal TimeOfAppoint As PMAppoint.TimeOfAppoint) As String
        'Ū���M�צ�ƾ�ƶ�
        Dim strVal As String = String.Empty
        Dim appoints As PMAppointCollection
        Dim appoint As PMAppoint

        appoints = PMAppoint.GetAppointments(ProjectID, theDate, TimeOfAppoint)
        If appoints.Count <= 0 Then
            Return String.Empty
        End If

        strVal += "�M�צW�١G" & ProjectName & vbCrLf
        strVal += "======================================================================================" & vbCrLf
        For Each appoint In appoints

            strVal += "�ƶ��W�١G[" & PMAppointType.GetTypeName(appoint.AppointTypeID) & "] " & appoint.Subject & vbCrLf
            Select Case appoint.AppointLevelID
                Case PMAppoint.AppointLevel.Top
                    strVal += "���n���šG ��" & vbCrLf
                Case PMAppoint.AppointLevel.Middle
                    strVal += "���n���šG ��" & vbCrLf
                Case PMAppoint.AppointLevel.Low
                    strVal += "���n���šG �C" & vbCrLf
            End Select
            strVal += "�ƶ��ɶ��G" & appoint.AppointDate.ToLongDateString & "�@" & PMAppoint.getWeekName(appoint.AppointDate.DayOfWeek) & "�@" & appoint.StartTime.ToString("t", DateTimeFormatInfo.InvariantInfo) & " ~ " & appoint.EndTime.ToString("t", DateTimeFormatInfo.InvariantInfo)
            If appoint.isCycle Then
                strVal += " (�g���ƶ�) "
            End If
            strVal += vbCrLf
            strVal += "�a�I�G" & appoint.Location & vbCrLf
            strVal += "�Ƶ��G" & appoint.Description & vbCrLf
            strVal += "======================================================================================" & vbCrLf
            '    ' strVal += " <a href=" & [String].Format("../CalendarDetail.aspx?ProjectID={0}&AppointID={1}&CycleID={2}&index=-1&projectIndex=3", ProjectID, appoint.AppointID, appoint.CycleID) & " >" & appoint.Subject & "</a>"

        Next


        Return strVal

    End Function

    Private Sub SendMail(ByVal strSubject As String, ByVal strMailTo As String, ByVal strBody As String)
        '�H�e�O���q���H��
        Dim mailInfo As String
        Dim mail As New PMSendMail

        mail.From = "�H�˱M�׺޲z�t��<eip@hsin-yi.org.tw>"
        mail.Recipt = strMailTo
        mail.Subject = strSubject
        mail.BCC = "sean@hsin-yi.org.tw"
        mail.Body = strBody
        mail.SMTPServer = ConfigurationSettings.AppSettings("SMTPServer")

        mailInfo += "Recipt:" & strMailTo & vbCrLf
        mailInfo += "Subject:" & strSubject & vbCrLf
        mailInfo += "Body:" & strBody & vbCrLf
        mailInfo += "SMTPServer:" & ConfigurationSettings.AppSettings("SMTPServer")
        Try
            mail.send()
        Catch ex As Exception
            WriteEventLog("�H�e�l��q���o�Ϳ��~ :" & ex.Message & vbCrLf & mailInfo, EventLogEntryType.Error)
        End Try

    End Sub

End Class
