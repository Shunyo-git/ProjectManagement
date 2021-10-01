Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    Public Class ProjectModify
        Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

        '���� Web Form �]�p�u��һݪ��I�s�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents txtAssistUnits As System.Web.UI.WebControls.TextBox
        Protected WithEvents txtSupplementarySubsidy As System.Web.UI.WebControls.TextBox
        Protected WithEvents txtSupplementaryUnits As System.Web.UI.WebControls.TextBox
        Protected WithEvents txtDescription As System.Web.UI.WebControls.TextBox
        Protected WithEvents ddlType As System.Web.UI.WebControls.DropDownList
        Protected WithEvents ddlState As System.Web.UI.WebControls.DropDownList
        Protected WithEvents ProjectNameRequiredfieldvalidator As System.Web.UI.WebControls.RequiredFieldValidator
        Protected WithEvents txtProjectName As System.Web.UI.WebControls.TextBox
        Protected WithEvents PanelEdit As System.Web.UI.WebControls.Panel
        Protected WithEvents lbtnCancel As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnSave As System.Web.UI.WebControls.LinkButton
        Protected WithEvents ErrorMessage As System.Web.UI.WebControls.Label
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents txtStartDate As DatePicker.iX.Controls.DatePicker
        Protected WithEvents txtEndDate As DatePicker.iX.Controls.DatePicker
        Protected WithEvents txtManager As HsinYiMemberPicker.HsinYi.EIP.Security.MemberPicker.MemberPicker
        Protected WithEvents txtChief As HsinYiMemberPicker.HsinYi.EIP.Security.MemberPicker.MemberPicker
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm

        '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
        '�ФŧR���β��ʥ��C
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
            '�ФŨϥε{���X�s�边�i��ק�C
            InitializeComponent()
        End Sub

#End Region
        ' Contains the id of current project
        Private _projectID As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

            ' Obtain project id from the QueryString.
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            '�b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���
            If Not Roles.isEnabledModifyProject(_projectID) Then
                Session("m_ErrorMessage") = "�t�Ωڵ��s���A�z�S���v���s�����M�׸�T�C" '"�z�D���M�צ����A�]���L�v���s�����M�׸�T�C"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If




            If Not IsPostBack Then
                BindDefaultStateType()
                ' Load project with _projID when project id exists in the QueryString
                If _projectID <> 0 Then
                    BindProjectEditForm()

                End If
            End If
        End Sub
        Private Sub BindDefaultStateType()
            ddlState.DataSource = ProjectState.GetStates
            ddlState.DataTextField = "Name"
            ddlState.DataValueField = "StateID"
            ddlState.DataBind()

            ddlType.DataSource = ProjectType.GetTypes
            ddlType.DataTextField = "Name"
            ddlType.DataValueField = "TypeID"
            ddlType.DataBind()
        End Sub
        Private Sub BindProjectEditForm()

            Dim project As New project(_projectID)
            project.Load()

            Dim manager As New PMUser(project.ManagerUserID, _projectID)
            manager.LoadProjectMember()
            Dim Chief As New PMUser(project.ChiefUserID, _projectID)
            Chief.LoadProjectMember()

            ' Populate web controls with the project's info. ( for View )
            txtProjectName.Text = project.ProjectName
            txtDescription.Text = project.ProjectDescription

            txtManager.Text = manager.DisplayName
            txtManager.WhoID = manager.UserID
            txtManager.WhoName = EIPSysSecurity.clsUser.getUserName(manager.UserID)
            txtManager.WhoType = "0"
            txtManager.EnableEmpty = False
            txtManager.ValidatorErrorMessage = "�M�׭t�d�H���o�ť�."


            txtChief.Text = Chief.DisplayName
            txtChief.WhoID = Chief.UserID
            txtChief.WhoName = EIPSysSecurity.clsUser.getUserName(Chief.UserID)
            txtChief.WhoType = "0"
            txtChief.EnableEmpty = False
            txtChief.ValidatorErrorMessage = "�D�n����̤��o�ť�."

            txtStartDate.Text = project.StartDate.ToShortDateString
            txtEndDate.Text = project.EndDate.ToShortDateString


            ddlState.SelectedValue = project.StateID



            ddlType.SelectedValue = project.TypeID



            txtSupplementaryUnits.Text = project.SupplementarySubsidy
            txtSupplementarySubsidy.Text = project.SupplementaryUnits
            txtAssistUnits.Text = project.AssistUnits

            manager = Nothing
            Chief = Nothing


        End Sub 'BindProjectEditForm

        Private Sub lbtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnSave.Click
            If SaveProject(_projectID) Then
                ReturnToProjectDetail()
            End If
        End Sub

        Private Function SaveProject(ByVal projectID As Integer) As Boolean

            'Dim selectedMembers As New UsersCollection
            'Dim manager As New PMUser
            'manager.UserID = Convert.ToInt32(txtManager.WhoID)
            'manager.Role = Roles.UserRoleProjectManager
            'manager.Description = "�M�׭t�d�H"
            'manager.CurrentUserID = PMSecurity.GetUserID
            'selectedMembers.Add(manager)

            'Dim Chief As New PMUser
            'Chief.UserID = Convert.ToInt32(txtChief.WhoID)
            'Chief.Role = Roles.UserRoleProjectChief
            'Chief.Description = "�D�n�����"
            'Chief.CurrentUserID = PMSecurity.GetUserID
            'selectedMembers.Add(Chief)


            Dim prj As New Project(projectID, PMSecurity.CleanStringRegex(txtProjectName.Text), PMSecurity.CleanStringRegex(txtDescription.Text), Convert.ToInt32(txtManager.WhoID), Convert.ToInt32(txtChief.WhoID), Convert.ToInt32(ddlState.SelectedValue), Convert.ToInt32(ddlType.SelectedValue), Convert.ToDateTime(txtStartDate.Text), Convert.ToDateTime(txtEndDate.Text))

            prj.SupplementarySubsidy = IIf(IsDBNull(txtSupplementarySubsidy.Text), String.Empty, txtSupplementarySubsidy.Text)
            prj.SupplementaryUnits = IIf(IsDBNull(txtSupplementaryUnits.Text), String.Empty, txtSupplementaryUnits.Text)
            prj.AssistUnits = IIf(IsDBNull(txtAssistUnits.Text), String.Empty, txtAssistUnits.Text)
            prj.ManagerUserID = Convert.ToInt32(txtManager.WhoID)
            prj.ChiefUserID = Convert.ToInt32(txtChief.WhoID)
            'prj.Members = selectedMembers
            prj.ModifiedUserID = PMSecurity.GetUserID

            If prj.Save() Then
                _projectID = prj.projectID
                ErrorMessage.Text = String.Empty
                Return True
            Else
                ErrorMessage.Text = "�o�Ϳ��~.  �L�k�x�s�M�׸��."
                Return False
            End If
        End Function 'SaveProject

        Private Sub ReturnToProjectDetail()

            Response.Redirect([String].Format("ProjectDetails.aspx?projectID={0}&index=-1&projectIndex=0", _projectID), False)
        End Sub 'ReturnToProjectDetail

        Private Sub lbtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnCancel.Click
            ReturnToProjectDetail()
        End Sub
    End Class
End Namespace

