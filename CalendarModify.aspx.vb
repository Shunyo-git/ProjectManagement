Imports System
Imports System.IO
Imports System.Data
Imports System.Web.UI.WebControls
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class CalendarModify
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents lbtnCancel As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnSave As System.Web.UI.WebControls.LinkButton
        Protected WithEvents ErrorMessage As System.Web.UI.WebControls.Label
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents itnDate As System.Web.UI.WebControls.Label
        Protected WithEvents itnTopic As System.Web.UI.WebControls.Label
        Protected WithEvents txtSubject As System.Web.UI.WebControls.TextBox
        Protected WithEvents Label2 As System.Web.UI.WebControls.Label
        Protected WithEvents Label3 As System.Web.UI.WebControls.Label
        Protected WithEvents txtDescription As System.Web.UI.WebControls.TextBox
        Protected WithEvents ddlAppointType As System.Web.UI.WebControls.DropDownList
        Protected WithEvents LevelList As System.Web.UI.HtmlControls.HtmlSelect
        Protected WithEvents rbtnSigle As System.Web.UI.WebControls.RadioButton
        Protected WithEvents rbtnCycle As System.Web.UI.WebControls.RadioButton
        Protected WithEvents Label4 As System.Web.UI.WebControls.Label
        Protected WithEvents Label6 As System.Web.UI.WebControls.Label
        Protected WithEvents LEGEND1 As System.Web.UI.HtmlControls.HtmlGenericControl
        Protected WithEvents PanelCycleDay As System.Web.UI.WebControls.Panel
        Protected WithEvents PanelCycleWeek As System.Web.UI.WebControls.Panel
        Protected WithEvents PanelCycleMonth As System.Web.UI.WebControls.Panel
        Protected WithEvents Label5 As System.Web.UI.WebControls.Label
        Protected WithEvents Label7 As System.Web.UI.WebControls.Label
        Protected WithEvents Legend2 As System.Web.UI.HtmlControls.HtmlGenericControl
        Protected WithEvents PanelSingle As System.Web.UI.WebControls.Panel
        Protected WithEvents PanelCycle As System.Web.UI.WebControls.Panel
        Protected WithEvents txtSingleStartDate As DatePicker.iX.Controls.DatePicker
        Protected WithEvents lbxSingleStartTime As System.Web.UI.WebControls.ListBox
        Protected WithEvents txtSingleEndDate As DatePicker.iX.Controls.DatePicker
        Protected WithEvents lbxSingleEndTime As System.Web.UI.WebControls.ListBox
        Protected WithEvents txtCycleStartDate As DatePicker.iX.Controls.DatePicker
        Protected WithEvents lbxCycleStartTime As System.Web.UI.WebControls.ListBox
        Protected WithEvents txtCycleEndDate As DatePicker.iX.Controls.DatePicker
        Protected WithEvents lbxCycleEndTime As System.Web.UI.WebControls.ListBox
        Protected WithEvents cboxWeekDay1 As System.Web.UI.WebControls.CheckBox
        Protected WithEvents cboxWeekDay2 As System.Web.UI.WebControls.CheckBox
        Protected WithEvents cboxWeekDay3 As System.Web.UI.WebControls.CheckBox
        Protected WithEvents cboxWeekDay4 As System.Web.UI.WebControls.CheckBox
        Protected WithEvents cboxWeekDay5 As System.Web.UI.WebControls.CheckBox
        Protected WithEvents cboxWeekDay6 As System.Web.UI.WebControls.CheckBox
        Protected WithEvents cboxWeekDay7 As System.Web.UI.WebControls.CheckBox
        Protected WithEvents lboxMonthDay As System.Web.UI.WebControls.ListBox
        Protected WithEvents lboxWeekInterval As System.Web.UI.WebControls.ListBox
        Protected WithEvents rbtnCycleType1 As System.Web.UI.WebControls.RadioButton
        Protected WithEvents rbtnCycleType2 As System.Web.UI.WebControls.RadioButton
        Protected WithEvents rbtnCycleType3 As System.Web.UI.WebControls.RadioButton
        Protected WithEvents rbtnCycleDayType1 As System.Web.UI.WebControls.RadioButton
        Protected WithEvents rbtnCycleDayType2 As System.Web.UI.WebControls.RadioButton
        Protected WithEvents rbtnCycleMonthType1 As System.Web.UI.WebControls.RadioButton
        Protected WithEvents rbtnCycleMonthType2 As System.Web.UI.WebControls.RadioButton
        Protected WithEvents lboxMonthWeekInterval As System.Web.UI.WebControls.ListBox
        Protected WithEvents lboxMonthWeekDay As System.Web.UI.WebControls.ListBox
        Protected WithEvents txtLocation As System.Web.UI.WebControls.TextBox
        Protected WithEvents SubjectRequiredfieldvalidator As System.Web.UI.WebControls.RequiredFieldValidator

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

            If Not IsPostBack Then

                ' Load project with _projID when project id exists in the QueryString
                If _projectID <> 0 Then
                    BindAppointType()
                    BindDates()
                Else
                    Session("m_ErrorMessage") = "系統錯誤，無法存取群組資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If

                If _appointID <> 0 Or _cycleID <> 0 Then
                    BindInfo()

                End If

            End If

        End Sub
        Private Sub BindAppointType()
            ddlAppointType.DataSource = PMAppointType.GetAppointTypes
            ddlAppointType.DataTextField = "TypeName"
            ddlAppointType.DataValueField = "TypeID"
            ddlAppointType.DataBind()



        End Sub

        Private Sub BindDates()
            '   _dayListTable = BusinessLogicLayer.TimeEntry.GetWeek(Convert.ToDateTime(WeekEnding.Text))
            lboxMonthDay.DataSource = PMAppoint.GetDays
            lboxMonthDay.DataValueField = "Date"
            lboxMonthDay.DataTextField = "Day"
            lboxMonthDay.DataBind()

            lboxMonthWeekDay.DataSource = PMAppoint.GetWeek(Date.Now)
            lboxMonthWeekDay.DataValueField = "Week"
            lboxMonthWeekDay.DataTextField = "Day"
            lboxMonthWeekDay.DataBind()
            'lboxMonthWeekDay.Items.FindByText(Convert.ToDateTime(DateTime.Today).ToString("ddd")).Selected = True
        End Sub 'BindDates

        Private Sub BindInfo()




            Dim appoint As New PMAppoint(_appointID, _cycleID)
            appoint.Load()
            _projectID = appoint.ProjectID
            _cycleID = appoint.CycleID

            ddlAppointType.SelectedValue = appoint.AppointTypeID
            LevelList.SelectedIndex = appoint.AppointLevelID - 1

            If _cycleID > 0 Then
                txtSubject.Text = appoint.CycleSubject
                txtCycleStartDate.Text = appoint.CycleStartDate.ToShortDateString
                txtCycleEndDate.Text = appoint.CycleEndDate.ToShortDateString
                lbxCycleStartTime.SelectedValue = Strings.FormatDateTime(appoint.CycleStartTime, DateFormat.ShortTime)
                lbxCycleEndTime.SelectedValue = Strings.FormatDateTime(appoint.CycleEndTime, DateFormat.ShortTime)

                Select Case appoint.CycleTypeID
                    Case PMAppoint.CycleType.DayCycle
                        rbtnCycleType1.Checked = True

                        Select Case appoint.CycleDayTypeID
                            Case PMAppoint.CycleDayType.Everyday
                                rbtnCycleDayType1.Checked = True
                            Case PMAppoint.CycleDayType.WorkDay
                                rbtnCycleDayType2.Checked = True
                        End Select

                    Case PMAppoint.CycleType.WeekCycle
                        rbtnCycleType2.Checked = True
                        lboxWeekInterval.SelectedValue = appoint.CycleWeekInterval

                        If InStr(appoint.CycleWeekDays, DayOfWeek.Monday) > 0 Then cboxWeekDay1.Checked = True
                        If InStr(appoint.CycleWeekDays, DayOfWeek.Tuesday) > 0 Then cboxWeekDay2.Checked = True
                        If InStr(appoint.CycleWeekDays, DayOfWeek.Wednesday) > 0 Then cboxWeekDay3.Checked = True
                        If InStr(appoint.CycleWeekDays, DayOfWeek.Thursday) > 0 Then cboxWeekDay4.Checked = True
                        If InStr(appoint.CycleWeekDays, DayOfWeek.Friday) > 0 Then cboxWeekDay5.Checked = True
                        If InStr(appoint.CycleWeekDays, DayOfWeek.Saturday) > 0 Then cboxWeekDay6.Checked = True
                        If InStr(appoint.CycleWeekDays, DayOfWeek.Sunday) > 0 Then cboxWeekDay7.Checked = True


                    Case PMAppoint.CycleType.MonthCycle
                        rbtnCycleType3.Checked = True
                        Select Case appoint.CycleMonthTypeID
                            Case PMAppoint.CycleMonthType.MonthDay
                                rbtnCycleMonthType1.Checked = True
                                lboxMonthDay.SelectedValue = appoint.CycleMonthDay
                            Case PMAppoint.CycleMonthType.MonthWeek
                                rbtnCycleMonthType2.Checked = True
                                lboxMonthWeekInterval.SelectedValue = appoint.CycleWeekInterval
                                lboxMonthWeekDay.SelectedValue = appoint.CycleWeekDay
                        End Select
                End Select
 
 
                rbtnSigle.Checked = False
                rbtnCycle.Checked = True
 
                chkCyclePanelVisible()
            Else

                txtSubject.Text = appoint.Subject
                txtSingleStartDate.Text = appoint.StartTime.ToShortDateString
                txtSingleEndDate.Text = appoint.EndTime.ToShortDateString
                lbxSingleStartTime.SelectedValue = Strings.FormatDateTime(appoint.StartTime, DateFormat.ShortTime)
                lbxSingleEndTime.SelectedValue = Strings.FormatDateTime(appoint.EndTime, DateFormat.ShortTime)


            End If

        End Sub 'BindInfo

        Private Sub BackToUserPage()
            Response.Redirect([String].Format("CalendarDetail.aspx?ProjectID={0}&AppointID={1}&CycleID={2}&index=-1&projectIndex=3", _projectID, _appointID, _cycleID), True)
        End Sub

        Private Sub rbtnCycleType1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnCycleType1.CheckedChanged
            chkCyclePanelVisible()
        End Sub

        Private Sub rbtnCycleType2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnCycleType2.CheckedChanged
            chkCyclePanelVisible()
        End Sub

        Private Sub rbtnCycleType3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnCycleType3.CheckedChanged
            chkCyclePanelVisible()
        End Sub

        Private Sub chkCyclePanelVisible()
            PanelSingle.Visible = rbtnSigle.Checked
            PanelCycle.Visible = rbtnCycle.Checked

            PanelCycleDay.Visible = rbtnCycleType1.Checked
            PanelCycleWeek.Visible = rbtnCycleType2.Checked
            PanelCycleMonth.Visible = rbtnCycleType3.Checked

        End Sub

        Private Sub rbtnSigle_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnSigle.CheckedChanged
            chkCyclePanelVisible()
        End Sub

        Private Sub rbtnCycle_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnCycle.CheckedChanged
            chkCyclePanelVisible()
        End Sub

        Private Sub lbtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnSave.Click
            If SaveAppoint() Then

                BackToDeatilPage()
            Else
                Session("m_ErrorMessage") = "系統錯誤，無法儲存群組資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If
        End Sub

        Private Function SaveAppoint() As Boolean
            Dim blnResult As Boolean
            Dim appoint As New PMAppoint(_appointID, _cycleID)
            appoint.Subject = txtSubject.Text
            appoint.AppointTypeID = Convert.ToInt32(ddlAppointType.SelectedValue)
            appoint.AppointLevelID = Convert.ToInt32(LevelList.Value)
            appoint.Location = txtLocation.Text
            appoint.Description = txtDescription.Text
           
            appoint.ProjectID = _projectID

            '單一
            If rbtnSigle.Checked Then
                appoint.CycleTypeID = 0
                If Not IsDate(txtSingleStartDate.Text) Or Not IsDate(txtSingleEndDate.Text) Then
                    Session("m_ErrorMessage") = "系統錯誤，日期不可空白。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If
                appoint.StartTime = Convert.ToDateTime(txtSingleStartDate.Text & " " & lbxSingleStartTime.SelectedValue)
                appoint.EndTime = Convert.ToDateTime(txtSingleEndDate.Text & " " & lbxSingleEndTime.SelectedValue)
                appoint.AppointDate = Convert.ToDateTime(txtSingleStartDate.Text & " " & lbxSingleStartTime.SelectedValue)
            Else
                If Not IsDate(txtCycleStartDate.Text) Or Not IsDate(txtCycleEndDate.Text) Then
                    Session("m_ErrorMessage") = "系統錯誤，日期不可空白。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If
                appoint.CycleSubject = txtSubject.Text
                appoint.CycleStartDate = (txtCycleStartDate.Text & " " & lbxCycleStartTime.SelectedValue)
                appoint.CycleEndDate = (txtCycleEndDate.Text & " " & lbxCycleEndTime.SelectedValue)
                appoint.CycleStartTime = (txtCycleStartDate.Text & " " & lbxCycleStartTime.SelectedValue)
                appoint.CycleEndTime = (txtCycleEndDate.Text & " " & lbxCycleEndTime.SelectedValue)
                appoint.StartTime = appoint.CycleStartDate
                appoint.EndTime = appoint.CycleEndDate
                appoint.AppointDate = appoint.CycleStartDate
                '每日
                If rbtnCycleType1.Checked Then
                    appoint.CycleTypeID = PMAppoint.CycleType.DayCycle
                    If rbtnCycleDayType1.Checked Then
                        appoint.CycleDayTypeID = PMAppoint.CycleDayType.Everyday
                    Else
                        appoint.CycleDayTypeID = PMAppoint.CycleDayType.WorkDay
                    End If

                End If
                '每週
                If rbtnCycleType2.Checked Then
                    appoint.CycleTypeID = PMAppoint.CycleType.WeekCycle
                    appoint.CycleWeekInterval = Convert.ToInt32(lboxWeekInterval.SelectedValue)
                    appoint.CycleWeekDays = String.Empty
                    If cboxWeekDay1.Checked Then appoint.CycleWeekDays += Convert.ToString(DayOfWeek.Monday)
                    If cboxWeekDay2.Checked Then appoint.CycleWeekDays += Convert.ToString(DayOfWeek.Tuesday)
                    If cboxWeekDay3.Checked Then appoint.CycleWeekDays += Convert.ToString(DayOfWeek.Wednesday)
                    If cboxWeekDay4.Checked Then appoint.CycleWeekDays += Convert.ToString(DayOfWeek.Thursday)
                    If cboxWeekDay5.Checked Then appoint.CycleWeekDays += Convert.ToString(DayOfWeek.Friday)
                    If cboxWeekDay6.Checked Then appoint.CycleWeekDays += Convert.ToString(DayOfWeek.Saturday)
                    If cboxWeekDay7.Checked Then appoint.CycleWeekDays += Convert.ToString(DayOfWeek.Sunday)

                End If
                '每月
                If rbtnCycleType3.Checked Then
                    appoint.CycleTypeID = PMAppoint.CycleType.MonthCycle

                    If rbtnCycleMonthType1.Checked Then
                        appoint.CycleMonthTypeID = PMAppoint.CycleMonthType.MonthDay
                        appoint.CycleMonthDay = lboxMonthDay.SelectedValue
                    Else
                        appoint.CycleMonthTypeID = PMAppoint.CycleMonthType.MonthWeek
                        appoint.CycleMonthInterval = 1
                        appoint.CycleWeekInterval = lboxMonthWeekInterval.SelectedValue
                        appoint.CycleWeekDay = lboxMonthWeekDay.SelectedValue
                    End If
                End If
            End If




            appoint.ModifiedUserID = PMSecurity.GetUserID
            'PrintLn(appoint.AppointID)
            'PrintLn(appoint.CycleID)
            'PrintLn(appoint.CycleTypeID)
            'PrintLn(appoint.CycleSubject)
            'PrintLn(appoint.CycleStartDate)
            'PrintLn(appoint.CycleEndDate)
            'PrintLn(appoint.CycleStartTime)
            'PrintLn(appoint.CycleEndTime)
            'PrintLn(appoint.CycleMonthInterval)
            'PrintLn(appoint.CycleMonthDay)
            'PrintLn(appoint.CycleWeekInterval)
            'PrintLn(appoint.CycleWeekDay)
            'PrintLn(appoint.CycleWeekDays)
            'PrintLn(appoint.CycleMonthTypeID)
            'PrintLn(appoint.CycleDayTypeID)
            'PrintLn(appoint.AppointTypeID)
            'PrintLn(appoint.AppointLevelID)
            'PrintLn(appoint.ProjectID)
            'PrintLn(appoint.Subject)
            'PrintLn(appoint.Location)
            'PrintLn(appoint.Description)
            'PrintLn(appoint.AppointDate)
            'PrintLn(appoint.StartTime)
            'PrintLn(appoint.EndTime)
            'PrintLn(appoint.ReminderTime)
            'PrintLn(appoint.FinishedTimes)
            'PrintLn(appoint.ModifiedUserID)
            'Response.End()

            blnResult = appoint.Save

            If blnResult Then
                _appointID = appoint.AppointID
                _cycleID = appoint.CycleID
            End If

            Return blnResult

        End Function

        Private Sub BackToDeatilPage()
            Response.Redirect([String].Format("CalendarDetail.aspx?ProjectID={0}&AppointID={1}&CycleID={2}&index=-1&projectIndex={3}", _projectID, _appointID, _cycleID, TabItem.ProjectTabIndex.ProjectCalendar), True)
        End Sub
        Private Sub PrintLn(ByVal strVal As String)
            Response.Write(strVal & ",<BR>")

        End Sub
    End Class



End Namespace


