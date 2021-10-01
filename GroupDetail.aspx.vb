Imports System
Imports System.Data
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web
    Public Class GroupDetail
        Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

        '���� Web Form �]�p�u��һݪ��I�s�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents lblDescription As System.Web.UI.WebControls.Label
        Protected WithEvents Label4 As System.Web.UI.WebControls.Label
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents lblName As System.Web.UI.WebControls.Label
        Protected WithEvents userNamelbl As System.Web.UI.WebControls.Label
        Protected WithEvents lbtnDelete As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnEdit As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents lblMembers As System.Web.UI.WebControls.Label

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
        Private _groupID As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            ' Obtain project id from the QueryString.
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            If Request.QueryString("GroupID") Is Nothing Then
                _groupID = 0
            Else
                _groupID = Convert.ToInt32(Request.QueryString("GroupID"))
            End If

            If Not Roles.isEnabledViewProject(_projectID) Then
                Session("m_ErrorMessage") = "�t�Ωڵ��s���A�z�S���v���s�����M�׸�T�C" '"�z�D���M�צ����A�]���L�v���s�����M�׸�T�C"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If


            If Roles.isEnabledModifyProject(_projectID) Then
                PanelModify.Visible = True

            Else
                PanelModify.Visible = False
            End If

            If Not IsPostBack Then

                If PMGroup.GetProjectGroups(_projectID).Count <= 1 Then
                    lbtnDelete.Attributes.Add("OnClick", "alert('���էO���M�װߤ@���էO�A���i�R���C');return false;")
                Else
                    lbtnDelete.Attributes.Add("OnClick", "return chkConfirmDelete();")
                End If
                ' Load project with _projID when project id exists in the QueryString
                If _groupID <> 0 And _projectID <> 0 Then
                    BindInfo()
                Else
                    Session("m_ErrorMessage") = "�t�ο��~�A�L�k�s���s�ո�T�C"  '"�z�D���M�צ����A�]���L�v���s�����M�׸�T�C"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If
            End If
        End Sub
        '*********************************************************************
        '
        ' The BindInfo method retrieves the list of roles 
        ' and Load User information if user exists
        '
        '*********************************************************************

        Private Sub BindInfo()
            Dim group As New PMGroup(_groupID)
            group.Load()

            _projectID = group.ProjectID
            lblName.Text = group.Name
            lblDescription.Text = group.Description

            Dim user As PMUser
            For Each user In group.Members
                If Not Strings.Trim(lblMembers.Text) Is String.Empty Then
                    lblMembers.Text += "�B"
                End If
                lblMembers.Text += user.DisplayName
            Next user

        End Sub 'BindInfo

        Private Sub lbtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnDelete.Click
            PMGroup.RemoveGroup(_groupID)
            Response.Redirect([String].Format("ProjectGroup.aspx?projectID={0}&index=-1&projectIndex={1}", _projectID, TabItem.ProjectTabIndex.ProjectGroup), True)
        End Sub

        Private Sub lbtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnEdit.Click
            Response.Redirect([String].Format("GroupModify.aspx?projectID={0}&index=-1&projectIndex={1}&GroupID={2}", _projectID, TabItem.ProjectTabIndex.ProjectGroup, _groupID), False)
        End Sub
    End Class
End Namespace

