Imports System
Imports System.Collections
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class DocCategoriesTabs
        Inherits System.Web.UI.UserControl

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents dlProjectTabs As System.Web.UI.WebControls.DataList

        '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
        '請勿刪除或移動它。
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
            '請勿使用程式碼編輯器進行修改。
            InitializeComponent()
        End Sub

#End Region

        Private projectID As Integer
        Private mainIndex As Integer
        Private categoryIndex As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            mainIndex = IIf(Request("index") Is Nothing, 0, Convert.ToInt32(Request("index")))
            projectID = IIf(Request("projectID") Is Nothing, 0, Convert.ToInt32(Request("projectID")))
            categoryIndex = IIf(Request("categoryIndex") Is Nothing, 0, Convert.ToInt32(Request("categoryIndex")))

            If Request.QueryString("projectID") Is Nothing Then
                projectID = 0
            End If

            If Request.QueryString("categoryIndex") Is Nothing Then
                categoryIndex = 1
            End If

            If projectID <> 0 Then
                AddTabItems()
            End If


        End Sub

        Private Sub AddTabItems()
            Dim tabItems As New ArrayList

            ' Always add the project list tab.  If the user was able to see the Administration tab, they'll
            ' always be able to view this tab.

            Dim categories As PMDocCategoriesCollection = PMDocCategories.GetCategorys()
            categories.Sort(PMDocCategoriesCollection.DocCategoriesFields.Sort, True)
            Dim category As PMDocCategories
            For Each category In categories
                tabItems.Add(New TabItem(category.Name, [String].Format("ProjectDocument.aspx?index={0}&projectID={1}&projectIndex={2}&categoryIndex={3}", mainIndex, projectID, Convert.ToInt32(TabItem.ProjectTabIndex.ProjectDocument), tabItems.Count)))
            Next
            dlProjectTabs.DataSource = tabItems
            dlProjectTabs.SelectedIndex = IIf(Request("categoryIndex") Is Nothing, -1, Convert.ToInt16(Request("categoryIndex")))
            dlProjectTabs.DataBind()
        End Sub

    End Class

End Namespace
