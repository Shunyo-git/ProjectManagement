Imports System
Imports System.Collections
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer


Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class mainTitleTabs
        Inherits System.Web.UI.UserControl


#Region " Web Form �]�p�u�㲣�ͪ��{���X "

        '���� Web Form �]�p�u��һݪ��I�s�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents tabs As System.Web.UI.WebControls.DataList
        Protected WithEvents lblCurrentUser As System.Web.UI.WebControls.Label

        '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
        '�ФŧR���β��ʥ��C
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
            '�ФŨϥε{���X�s�边�i��ק�C
            InitializeComponent()
        End Sub

#End Region

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


            lblCurrentUser.Text = PMSecurity.GetDisplayName & " ( " & Roles.GetRolesName(Convert.ToInt32(PMSecurity.GetUserRole)) & " )  "

            Dim tabItems As New ArrayList

            ' If Request.IsAuthenticated Then
            ' add the Log tab (all users can see this tab)
            tabItems.Add(New TabItem("�ڪ�����", "index.aspx?index=" & tabItems.Count))
            ' End If


            'If TTSecurity.IsInRole(TTUser.UserRoleAdminPMgr) Then
            tabItems.Add(New TabItem("�M�צC��", "ProjectList.aspx?index=" & tabItems.Count))
            ' tabItems.Add(New TabItem("�̷s���", "NewDocument.aspx?index=" & tabItems.Count))
            ' tabItems.Add(New TabItem("�̷s���D", "NewTopic.aspx?index=" & tabItems.Count))
            ' End If

            '�Ȭy�{�޲z�̥i�s�W�M��
            If PMSecurity.IsInRole(Roles.FlowManagr) Then
                tabItems.Add(New TabItem("�s�W�M��", "ProjectModify.aspx?index=" & tabItems.Count))
                ' tabItems.Add(New TabItem("�v���޲z", "ProjectSecurity.aspx?index=" & tabItems.Count))
            End If
            tabs.DataSource = tabItems
            tabs.SelectedIndex = IIf(Request("index") Is Nothing, 0, Convert.ToInt32(Request("index")))
            tabs.DataBind()

        End Sub 'Page_Load
    End Class 'myProjectTabs
End Namespace 'Hsinyi.EIP.ProjectManagement.Web
