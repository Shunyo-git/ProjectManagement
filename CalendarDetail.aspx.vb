Imports System
Imports System.IO
Imports System.Data
Imports System.Web.UI.WebControls
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class CalendarDetail
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents lbtnDelete As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnEdit As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents PanelCycle As System.Web.UI.WebControls.Panel
        Protected WithEvents PanelSingle As System.Web.UI.WebControls.Panel
        Protected WithEvents lblSubject As System.Web.UI.WebControls.Label
        Protected WithEvents lblSingleDateStart As System.Web.UI.WebControls.Label
        Protected WithEvents lblCycleDateStart As System.Web.UI.WebControls.Label
        Protected WithEvents lblCycleTimeStart As System.Web.UI.WebControls.Label
        Protected WithEvents lblSingleDateEnd As System.Web.UI.WebControls.Label
        Protected WithEvents lblCycleDateEnd As System.Web.UI.WebControls.Label
        Protected WithEvents lblCycleTimeEnd As System.Web.UI.WebControls.Label
        Protected WithEvents PanelCycleMonth As System.Web.UI.WebControls.Panel
        Protected WithEvents PanelCycleWeek As System.Web.UI.WebControls.Panel
        Protected WithEvents PanelCycleDay As System.Web.UI.WebControls.Panel
        Protected WithEvents lblCycleDayType1 As System.Web.UI.WebControls.Label
        Protected WithEvents lblCycleDayType2 As System.Web.UI.WebControls.Label
        Protected WithEvents lblCycleWeelInterval As System.Web.UI.WebControls.Label
        Protected WithEvents lblWeekDay1 As System.Web.UI.WebControls.Label
        Protected WithEvents lblWeekDay2 As System.Web.UI.WebControls.Label
        Protected WithEvents lblWeekDay3 As System.Web.UI.WebControls.Label
        Protected WithEvents lblWeekDay4 As System.Web.UI.WebControls.Label
        Protected WithEvents lblWeekDay5 As System.Web.UI.WebControls.Label
        Protected WithEvents lblWeekDay6 As System.Web.UI.WebControls.Label
        Protected WithEvents lblWeekDay7 As System.Web.UI.WebControls.Label
        Protected WithEvents lblCycleMonthDay As System.Web.UI.WebControls.Label
        Protected WithEvents lblDescription As System.Web.UI.WebControls.Label
        Protected WithEvents lblCycleMonthInterval As System.Web.UI.WebControls.Label
        Protected WithEvents lblCyclMonthWeekDay As System.Web.UI.WebControls.Label
        Protected WithEvents lblSingle As System.Web.UI.WebControls.Label
        Protected WithEvents lblCycle As System.Web.UI.WebControls.Label
        Protected WithEvents lblLocation As System.Web.UI.WebControls.Label
        Protected WithEvents lblAppointLevel As System.Web.UI.WebControls.Label
        Protected WithEvents lblAppointType As System.Web.UI.WebControls.Label
        Protected WithEvents lbtnDeltetCycle As System.Web.UI.WebControls.LinkButton

        '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
        '請勿刪除或移動它。
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
            '請勿使用程式碼編輯器進行修改。
            InitializeComponent()
        End Sub

#End Region

        Private _projectID As Integer
        Private _appointID As Integer
        Private _cycleID As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            If Request.QueryString("AppointID") Is Nothing Then
                _appointID = 0
            Else
                _appointID = Convert.ToInt32(Request.QueryString("AppointID"))
            End If

            If Request.QueryString("CycleID") Is Nothing Then
                _cycleID = 0
            Else
                _cycleID = Convert.ToInt32(Request.QueryString("CycleID"))
            End If

            If Roles.isEnabledModifyProject(_projectID) Then
                PanelModify.Visible = True
            Else
                PanelModify.Visible = False
            End If


            If Not IsPostBack Then
                lbtnDelete.Attributes.Add("OnClick", "return chkConfirmDelete();")
                lbtnDeltetCycle.Attributes.Add("OnClick", "return chkConfirmDelete();")
                ' Load project with _projID when project id exists in the QueryString
                If (_appointID <> 0 Or _cycleID <> 0) And _projectID <> 0 Then
                    BindInfo()
                Else
                    Session("m_ErrorMessage") = "系統錯誤，無法存取專案資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    ' Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If
            End If
        End Sub

        Private Sub BindInfo()


            Dim appoint As New PMAppoint(_appointID, _cycleID)

            appoint.Load()

            _projectID = appoint.ProjectID
            lblSubject.Text = appoint.Subject
            lblAppointType.Text = PMAppointType.GetTypeName(appoint.AppointTypeID)

            Select Case appoint.AppointLevelID
                Case PMAppoint.AppointLevel.Top
                    lblAppointLevel.BackColor = Color.Red
                    lblAppointLevel.Text = "　高　"
                Case PMAppoint.AppointLevel.Middle
                    lblAppointLevel.BackColor = Color.Gold
                    lblAppointLevel.Text = "　中　"
                Case PMAppoint.AppointLevel.Low
                    lblAppointLevel.BackColor = Color.FromName("#00c000")
                    lblAppointLevel.Text = "　低　"
            End Select

            lblLocation.Text = appoint.Location



            chkCyclePanelVisible(appoint.CycleID > 0)

            If appoint.isCycle Then
                lblCycleDateStart.Text = appoint.CycleStartDate.ToShortDateString
                lblCycleDateEnd.Text = appoint.CycleEndDate.ToShortDateString
                lblCycleTimeStart.Text = appoint.CycleStartTime.ToShortTimeString
                lblCycleTimeEnd.Text = appoint.CycleEndTime.ToShortTimeString


                Select Case appoint.CycleTypeID
                    Case PMAppoint.CycleType.DayCycle
                        PanelCycleDay.Visible = True
                    Case PMAppoint.CycleType.WeekCycle
                        PanelCycleWeek.Visible = True
                    Case PMAppoint.CycleType.MonthCycle
                        PanelCycleMonth.Visible = True
                End Select
                lblCycleDayType1.Visible = (appoint.CycleDayTypeID = PMAppoint.CycleDayType.Everyday)
                lblCycleDayType2.Visible = (appoint.CycleDayTypeID = PMAppoint.CycleDayType.WorkDay)

                lblCycleWeelInterval.Text = appoint.CycleWeekInterval

                'Response.Write(appoint.CycleWeekDays)
                'Response.Write(InStr(appoint.CycleWeekDays, DayOfWeek.Monday))
                'Response.Write(InStr(appoint.CycleWeekDays, DayOfWeek.Tuesday))
                'Response.Write(InStr(appoint.CycleWeekDays, DayOfWeek.Wednesday))
                'Response.Write(InStr(appoint.CycleWeekDays, DayOfWeek.Thursday))
                'Response.Write(InStr(appoint.CycleWeekDays, DayOfWeek.Friday))
                'Response.Write(InStr(appoint.CycleWeekDays, DayOfWeek.Saturday))
                'Response.Write(InStr(appoint.CycleWeekDays, DayOfWeek.Sunday))

                If InStr(appoint.CycleWeekDays, DayOfWeek.Monday) > 0 Then lblWeekDay1.Visible = True
                If InStr(appoint.CycleWeekDays, DayOfWeek.Tuesday) > 0 Then lblWeekDay2.Visible = True
                If InStr(appoint.CycleWeekDays, DayOfWeek.Wednesday) > 0 Then lblWeekDay3.Visible = True
                If InStr(appoint.CycleWeekDays, DayOfWeek.Thursday) > 0 Then lblWeekDay4.Visible = True
                If InStr(appoint.CycleWeekDays, DayOfWeek.Friday) > 0 Then lblWeekDay5.Visible = True
                If InStr(appoint.CycleWeekDays, DayOfWeek.Saturday) > 0 Then lblWeekDay6.Visible = True
                If InStr(appoint.CycleWeekDays, DayOfWeek.Sunday) > 0 Then lblWeekDay7.Visible = True

                lblCycleMonthDay.Visible = (appoint.CycleMonthTypeID = PMAppoint.CycleMonthType.MonthDay)
                lblCycleMonthInterval.Visible = (appoint.CycleMonthTypeID = PMAppoint.CycleMonthType.MonthWeek)
                lblCyclMonthWeekDay.Visible = (appoint.CycleMonthTypeID = PMAppoint.CycleMonthType.MonthWeek)
                If appoint.CycleMonthTypeID = PMAppoint.CycleMonthType.MonthDay Then
                    lblCycleMonthDay.Text = appoint.CycleMonthDay & "日"
                Else

                    If appoint.CycleMonthInterval = 5 Then
                        lblCycleMonthInterval.Text = "最後一個"
                    Else
                        lblCycleMonthInterval.Text = "第 " & appoint.CycleWeekInterval & " 個"
                    End If

                    lblCyclMonthWeekDay.Text = PMAppoint.getWeekName(appoint.CycleWeekDay)

                End If
            Else
                lblSingleDateStart.Text = appoint.StartTime.ToString
                lblSingleDateEnd.Text = appoint.EndTime.ToString

            End If


            If PMSecurity.GetUserID = appoint.ModifiedUserID Then
                PanelModify.Visible = True
            End If



        End Sub 'BindInfo

        Private Sub chkCyclePanelVisible(ByVal blnVal As Boolean)



            lblSingle.Visible = Not blnVal
            PanelSingle.Visible = Not blnVal
            lblSingle.Visible = Not blnVal

            lblCycle.Visible = blnVal
            PanelCycle.Visible = blnVal
            lblCycle.Visible = blnVal

        End Sub

        Private Sub lbtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnDelete.Click
            PMAppoint.Remove(_appointID)
            BackToProjectList()
        End Sub

        Private Sub lbtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnEdit.Click
            Response.Redirect([String].Format("CalendarModify.aspx?ProjectID={0}&index=-1&projectIndex={1}&AppointID={2}&CycleID={3}", _projectID, TabItem.ProjectTabIndex.ProjectCalendar, _appointID, _cycleID), True)
        End Sub
        Private Sub BackToProjectList()
            Response.Redirect([String].Format("ProjectCalendar.aspx?ProjectID={0}&index=-1&projectIndex={1}", _projectID, TabItem.ProjectTabIndex.ProjectCalendar), True)
        End Sub

        Private Sub lbtnDeltetCycle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnDeltetCycle.Click
            PMAppoint.RemoveCycle(_cycleID)
            BackToProjectList()
        End Sub
    End Class

End Namespace


