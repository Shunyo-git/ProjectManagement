Imports System
Imports System.Collections
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class projectTabs
        Inherits System.Web.UI.UserControl

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

        '���� Web Form �]�p�u��һݪ��I�s�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents tabs As System.Web.UI.WebControls.DataList
        Protected WithEvents dlProjectTabs As System.Web.UI.WebControls.DataList
        Protected WithEvents lblProjectName As System.Web.UI.WebControls.Label

        '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
        '�ФŧR���β��ʥ��C
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
            '�ФŨϥε{���X�s�边�i��ק�C
            InitializeComponent()
        End Sub

#End Region
        Private projectID As Integer
        Private mainIndex As Integer


        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

            mainIndex = IIf(Request("index") Is Nothing, 0, Convert.ToInt32(Request("index")))
            projectID = IIf(Request("projectID") Is Nothing, 0, Convert.ToInt32(Request("projectID")))

            If Request.QueryString("projectID") Is Nothing Then
                projectID = 0
            End If

            'If Not Page.IsPostBack Then
            If projectID <> 0 Then
                AddTabItems()
                Dim pj As New Project(projectID)
                pj.Load()
                lblProjectName.Text = "�ثe�M�סG" & pj.ProjectName
                'lblProjectName.Text += "<BR>" & pj.StartDate.ToShortDateString & "~" & pj.EndDate.ToShortDateString

            End If
            'End If




        End Sub 'Page_Load

        Private Sub AddTabItems()
            Dim tabItems As New ArrayList

            ' Always add the project list tab.  If the user was able to see the Administration tab, they'll
            ' always be able to view this tab.

            tabItems.Add(New TabItem("�M�׭���", "ProjectDetails.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("�M�׶i�ת�", "ProjectDocument.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count & "&categoryIndex=" & Convert.ToInt32(TabItem.CategoryTabIndex.Timeline)))

            tabItems.Add(New TabItem("�M�ײէO", "ProjectGroup.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("�H�O�t�m", "ProjectUser.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("����/�q��", "ProjectCalendar.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("�S��ƥ�O��", "ProjectEvent.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("�M�פ��", "ProjectDocument.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("�Q�װ�", "ProjectForum.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("�s����", "ProjectContact.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))

            '�ȱM�׭t�d�H�Y�u�n����̥i�s�W�H�O
            'If PMSecurity.IsInRole(PMUser.FlowManagr) Then
            '    tabItems.Add(New TabItem("�s�W�M��", "ProjectNew.aspx?index=" & mainIndex & "&adminIndex=" & tabItems.Count))
            '    tabItems.Add(New TabItem("�v���޲z", "ProjectNew.aspx?index=" & mainIndex & "&adminIndex=" & tabItems.Count))
            'End If



            dlProjectTabs.DataSource = tabItems
            dlProjectTabs.SelectedIndex = IIf(Request("projectIndex") Is Nothing, 0, Convert.ToInt16(Request("projectIndex")))
            dlProjectTabs.DataBind()
        End Sub
    End Class
End Namespace

