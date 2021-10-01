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

#Region " 元件設計工具產生的程式碼 "

    Public Sub New()
        MyBase.New()

        '此為元件設計工具所需的呼叫。
        InitializeComponent()

        '在 InitializeComponent() 呼叫之後加入所有的初始化作業
        AgentTimer = New System.Timers.Timer
        AgentTimer.Interval = _TimeInterval

        AddHandler AgentTimer.Elapsed, AddressOf OnTimer
    End Sub

    'UserService 覆寫 Dispose 以清除元件清單。
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' 處理序的主進入點
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' 在同一個處理序中可以執行多個 NT 服務; 若要在這個處理序中
        ' 加入另一項服務，請修改以下各行
        ' 以建立第二個服務物件。例如，
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New PMAgentSerivce}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    '為元件設計工具的必要項
    Private components As System.ComponentModel.IContainer

    ' 注意: 以下為元件設計工具所需的程序
    '您可以使用元件設計工具進行修改，
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        '
        'PMAgentSerivce
        '
        Me.ServiceName = "PMAgentSerivce"

    End Sub

#End Region

    Protected AgentTimer As System.Timers.Timer
    Private _LogName As String = "Hsin-Yi Service"
    Private _EventSourceName As String = "專案管理系統"
    Private _TimeSecInterval As Integer = 1000
    Private _TimeMinInterval As Integer = 1000 * 60
    Private _TimeHourInterval As Integer = _TimeMinInterval * 60
    Private _TimeInterval As Integer = _TimeHourInterval * 1 '時間偵測間隔
    Private _LastRunTime As Date = Date.Now

    Protected Overrides Sub OnStart(ByVal args() As String)
        '  啟動服務的工作。
      
 
        WriteEventLog("信誼專案管理系統代理服務 服務已啟動" & vbCrLf & "下次偵測時間：" & Date.Now.AddMilliseconds(_TimeInterval).ToString, EventLogEntryType.Information)
        AgentTimer.Enabled = True

    End Sub

    Protected Overrides Sub OnStop()
        '  停止服務所執行的工作。
        WriteEventLog("信誼專案管理系統代理服務 服務己停止", EventLogEntryType.Warning)
        AgentTimer.Enabled = False
    End Sub

    Protected Overrides Sub OnPause()
        '  暫停服務所執行的工作。
        WriteEventLog("信誼專案管理系統代理服務 服務已暫停", EventLogEntryType.Warning)
        AgentTimer.Enabled = False
    End Sub

    Protected Overrides Sub OnContinue()
        '  繼服務所執行的工作。
        WriteEventLog("信誼專案管理系統代理服務 服務已繼續", EventLogEntryType.Information)
        AgentTimer.Enabled = True
    End Sub

    Protected Sub OnTimer(ByVal source As Object, ByVal e As ElapsedEventArgs)
        '  計時器引發Elapsed 事件時的工作。

        If Day(_LastRunTime) < Day(Date.Now) Then
            LoadProjects()
        Else
            WriteEventLog("信誼專案管理系統代理服務 服務偵測 " & vbCrLf & "下次偵測時間：" & Date.Now.AddMilliseconds(_TimeInterval).ToString, EventLogEntryType.Information)
        End If

    End Sub

    Private Sub WriteEventLog(ByVal strMsg As String, ByVal EventLogEntryType As EventLogEntryType)
        '記錄事件到系統事件簿。
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
        '讀取專案
        'WriteEventLog("連線字串 :" & ConfigurationSettings.AppSettings("PMConnectionString"), EventLogEntryType.Information)
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
                    SendMail(pj.ProjectName & " - 今日專案事項通知", strMailTo, strVal)
                    WriteEventLog("寄送今日專案事項通知 " & vbCrLf & "專案名稱 : " & pj.ProjectName & vbCrLf & "專案成員 : " & strMailTo, EventLogEntryType.Information)
                End If

                If Date.Now.DayOfWeek = DayOfWeek.Friday Then
                    strVal = WriteAppoint(pj.projectID, pj.ProjectName, Date.Now.AddDays(3), PMAppoint.TimeOfAppoint.Week)
                    If strVal.Length > 0 Then
                        SendMail(pj.ProjectName & " -下週專案事項通知", strMailTo, strVal)
                        WriteEventLog("寄送下週專案事項通知 " & vbCrLf & "專案名稱 : " & pj.ProjectName & vbCrLf & "專案成員 : " & strMailTo, EventLogEntryType.Information)
                    End If
                End If

            Next

            blnResult = True
        Catch ex As Exception
            WriteEventLog("讀取專案行事曆發生錯誤 :" & ex.Message, EventLogEntryType.Error)
            blnResult = False
        End Try

        If blnResult Then
            _LastRunTime = Date.Now
        End If


    End Sub

    Private Function getMemberEmail(ByVal projectID As Integer) As String
        '讀取專案成員Email帳號
        Dim strVal As String = String.Empty
        Dim user As PMUser

        For Each user In PMUser.GetMembers(projectID)
            If strVal.Length > 0 Then strVal += ";"
            strVal += user.UserName & "@hsin-yi.org.tw"
        Next
        Return strVal
    End Function

    Public Function WriteAppoint(ByVal ProjectID As Integer, ByVal ProjectName As String, ByVal theDate As Date, ByVal TimeOfAppoint As PMAppoint.TimeOfAppoint) As String
        '讀取專案行事曆事項
        Dim strVal As String = String.Empty
        Dim appoints As PMAppointCollection
        Dim appoint As PMAppoint

        appoints = PMAppoint.GetAppointments(ProjectID, theDate, TimeOfAppoint)
        If appoints.Count <= 0 Then
            Return String.Empty
        End If

        strVal += "專案名稱：" & ProjectName & vbCrLf
        strVal += "======================================================================================" & vbCrLf
        For Each appoint In appoints

            strVal += "事項名稱：[" & PMAppointType.GetTypeName(appoint.AppointTypeID) & "] " & appoint.Subject & vbCrLf
            Select Case appoint.AppointLevelID
                Case PMAppoint.AppointLevel.Top
                    strVal += "重要等級： 高" & vbCrLf
                Case PMAppoint.AppointLevel.Middle
                    strVal += "重要等級： 中" & vbCrLf
                Case PMAppoint.AppointLevel.Low
                    strVal += "重要等級： 低" & vbCrLf
            End Select
            strVal += "事項時間：" & appoint.AppointDate.ToLongDateString & "　" & PMAppoint.getWeekName(appoint.AppointDate.DayOfWeek) & "　" & appoint.StartTime.ToString("t", DateTimeFormatInfo.InvariantInfo) & " ~ " & appoint.EndTime.ToString("t", DateTimeFormatInfo.InvariantInfo)
            If appoint.isCycle Then
                strVal += " (週期事項) "
            End If
            strVal += vbCrLf
            strVal += "地點：" & appoint.Location & vbCrLf
            strVal += "備註：" & appoint.Description & vbCrLf
            strVal += "======================================================================================" & vbCrLf
            '    ' strVal += " <a href=" & [String].Format("../CalendarDetail.aspx?ProjectID={0}&AppointID={1}&CycleID={2}&index=-1&projectIndex=3", ProjectID, appoint.AppointID, appoint.CycleID) & " >" & appoint.Subject & "</a>"

        Next


        Return strVal

    End Function

    Private Sub SendMail(ByVal strSubject As String, ByVal strMailTo As String, ByVal strBody As String)
        '寄送是項通知信函
        Dim mailInfo As String
        Dim mail As New PMSendMail

        mail.From = "信誼專案管理系統<eip@hsin-yi.org.tw>"
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
            WriteEventLog("寄送郵件通知發生錯誤 :" & ex.Message & vbCrLf & mailInfo, EventLogEntryType.Error)
        End Try

    End Sub

End Class
