
Imports System
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity
Imports System.Globalization


Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class ProjectCalendar
        Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

        '���� Web Form �]�p�u��һݪ��I�s�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents lbtnNew As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents PjCalendar As System.Web.UI.WebControls.Calendar

        '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
        '�ФŧR���β��ʥ��C
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
            '�ФŨϥε{���X�s�边�i��ק�C
            InitializeComponent()
        End Sub

#End Region
        Private _projectID As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            ' Obtain project id from the QueryString.
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            ' Ensure that the visiting user has access to view the current page
            If Not Roles.isEnabledViewProject(_projectID) Then
                Session("m_ErrorMessage") = "�t�Ωڵ��s���A�z�S���v���s�����M�׸�T�C" '"�z�D���M�צ����A�]���L�v���s�����M�׸�T�C"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If

            If Roles.isEnabledModifyProject(_projectID) Or Roles.isProjectMember(_projectID) Then
                PanelModify.Visible = True
            Else
                PanelModify.Visible = False
            End If



            If Not Page.IsPostBack Then
                If _projectID <> 0 Then

                    PjCalendar.VisibleDate = Date.Now
                    PjCalendar.SelectedDate = Date.Now

            End If


            End If
        End Sub

        Private Sub lbtnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnNew.Click
            GoToModifyPage()
        End Sub

        Private Sub PjCalendar_DayRender(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DayRenderEventArgs) Handles PjCalendar.DayRender
            Dim d As CalendarDay
            Dim c As TableCell

            d = e.Day
            c = e.Cell

            If d.IsOtherMonth Then
                c.Controls.Clear()
            Else
                Try

                    Dim Hol As String = getActivityAppoint(d.Date) ' getActivityDept(d.Date) 'holidays(d.Date.Month, d.Date.Day)

                    If Hol <> "" Then
                        c.Controls.Add(New LiteralControl(" <br><SPAN style='font-weight: normal;' onclick=""toShow('" & d.Date.ToShortDateString & "');"">" + Hol + "</SPAN>"))
                    End If
                Catch exc As Exception
                    Response.Write(exc.ToString())
                End Try
            End If
            If e.Day.IsWeekend() Then
                c.ForeColor = Color.Red
                c.Font.Bold = True

            End If
        End Sub

        Private Function getActivityAppoint(ByVal theDate As Date) As String
            Dim appoints As PMAppointCollection = PMAppoint.GetAppointments(_projectID, theDate)
            Dim appoint As PMAppoint
            Dim strVal As String

            For Each appoint In appoints


                If appoint.isCycle Then
                    strVal += "<img src=images/icon_cycle.gif  alt=�g���ƶ�>"
                End If
                strVal += appoint.StartTime.ToString("t", DateTimeFormatInfo.InvariantInfo) & "~" & appoint.EndTime.ToString("t", DateTimeFormatInfo.InvariantInfo)
                strVal += "<BR>"
                Select Case appoint.AppointLevelID
                    Case PMAppoint.AppointLevel.Top
                        strVal += "<FONT COLOR=" & Color.Red.ToString & ">"
                    Case PMAppoint.AppointLevel.Middle
                        strVal += "<FONT COLOR=" & Color.Gold.ToString & ">"
                    Case PMAppoint.AppointLevel.Low
                        strVal += "<FONT COLOR=#00c000>"
                End Select

                strVal += "[" & PMAppointType.GetTypeName(appoint.AppointTypeID) & "]</FONT>"

                strVal += " <a href=" & [String].Format("CalendarDetail.aspx?ProjectID={0}&AppointID={1}&CycleID={2}&index=-1&projectIndex=3", _projectID, appoint.AppointID, appoint.CycleID) & ">" & appoint.Subject & "</a>"
                strVal += "<BR>"
            Next

            Return strVal
        End Function

        Private Sub GoToDeatilPage(ByVal appointID As Integer, ByVal cycleID As Integer)
            Response.Redirect([String].Format("CalendarDetail.aspx?ProjectID={0}&AppointID={1}&CycleID={2}&index=-1&projectIndex={3}", _projectID, appointID, cycleID, TabItem.ProjectTabIndex.ProjectCalendar), True)
        End Sub
        Private Sub GoToModifyPage()
            Response.Redirect([String].Format("CalendarModify.aspx?projectID={0}&index=-1&projectIndex={1}&AppointID=0", _projectID, TabItem.ProjectTabIndex.ProjectCalendar), True)
        End Sub


    End Class

End Namespace
