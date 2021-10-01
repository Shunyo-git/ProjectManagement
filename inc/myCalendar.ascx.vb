Imports System
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity
Imports System.Globalization
Imports System.Threading

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class myCalendar
        Inherits System.Web.UI.UserControl

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

        '���� Web Form �]�p�u��һݪ��I�s�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectsGrid As System.Web.UI.WebControls.DataGrid
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
        Private theDate As Date = Date.Now
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

            Dim dayNames() As String = {"��", "�@", "�G", "�T", "�|", "��", "��"}

            Dim culture As New CultureInfo("zh-TW")
            culture.DateTimeFormat.AbbreviatedDayNames = dayNames

            Thread.CurrentThread.CurrentCulture = culture

            If Not Page.IsPostBack Then
                BindProjects()
                PjCalendar.SelectedDate = Date.Now
                PjCalendar.VisibleDate = Date.Now
            End If

        End Sub

        Private Sub BindProjects()

            Dim projectList As ProjectsCollection = Project.GetProjects(PMSecurity.GetUserID(), PMSecurity.GetUserRole())

            ProjectsGrid.DataSource = projectList
            ProjectsGrid.DataBind()
        End Sub 'BindProjects
        Public Function WriteAppoint(ByVal ProjectID As Integer) As String
            Dim strVal As String

            Dim pj As New Project(ProjectID)
            pj.Load()
            strVal += "<li>" & pj.ProjectName & " </li><BR>"

            Dim appoints As PMAppointCollection = PMAppoint.GetAppointments(ProjectID, theDate)
            Dim appoint As PMAppoint
            For Each appoint In appoints
                strVal += ""


                Select Case appoint.AppointLevelID
                    Case PMAppoint.AppointLevel.Top
                        strVal += "<FONT COLOR=" & Color.Red.ToString & ">"
                    Case PMAppoint.AppointLevel.Middle
                        strVal += "<FONT COLOR=" & Color.Gold.ToString & ">"
                    Case PMAppoint.AppointLevel.Low
                        strVal += "<FONT COLOR=#00c000>"
                End Select

                strVal += "[" & PMAppointType.GetTypeName(appoint.AppointTypeID) & "]</FONT>"
                If appoint.isCycle Then
                    strVal += " <img src=images/icon_cycle.gif  alt=�g���ƶ�>"
                End If

                strVal += appoint.StartTime.ToString("t", DateTimeFormatInfo.InvariantInfo)

                strVal += " <a href=" & [String].Format("CalendarDetail.aspx?ProjectID={0}&AppointID={1}&CycleID={2}&index=-1&projectIndex=3", ProjectID, appoint.AppointID, appoint.CycleID) & " >" & appoint.Subject & "</a>"


                strVal += "<BR>"
            Next
            If appoints.Count > 0 Then
                Return strVal
            Else
                Return String.Empty
            End If

        End Function

        Private Sub PjCalendar_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PjCalendar.SelectionChanged
            theDate = PjCalendar.SelectedDate
            BindProjects()
        End Sub



        Private Sub PjCalendar_DayRender(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DayRenderEventArgs) Handles PjCalendar.DayRender

            'Dim c As TableCell

            'c = e.Cell

            'If InStr(c.Text, "�P") > 0 Then
            '    c.Text = c.Text.Replace("�P��", "")
            'Else
            '    c.Text += "a"
            'End If

        End Sub


    End Class

End Namespace


