Public Class WebForm1
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents ItemsGrid As System.Web.UI.WebControls.DataGrid
    Protected WithEvents CheckBox1 As System.Web.UI.WebControls.CheckBox

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If CheckBox1.Checked Then
            ItemsGrid.PagerStyle.Mode = PagerMode.NumericPages
        Else
            ItemsGrid.PagerStyle.Mode = PagerMode.NextPrev
        End If

        If Not IsPostBack Then
            start_index = 0
            ItemsGrid.VirtualItemCount = 100
        End If

        BindGrid()
    End Sub
    Dim start_index As Integer

    Function CreateDataSource() As ICollection

        Dim dt As New DataTable
        Dim dr As DataRow

        dt.Columns.Add(New DataColumn("IntegerValue", GetType(Int32)))
        dt.Columns.Add(New DataColumn("StringValue", GetType(String)))
        dt.Columns.Add(New DataColumn("CurrencyValue", GetType(Double)))

        Dim i As Integer
        For i = start_index To (start_index + ItemsGrid.PageSize) - 1
            dr = dt.NewRow()

            dr(0) = i
            dr(1) = "Item " + i.ToString()
            dr(2) = 1.23 * (i + 1)

            dt.Rows.Add(dr)
        Next i

        Dim dv As New DataView(dt)
        Return dv

    End Function 'CreateDataSource

   

    Sub Grid_Change(ByVal sender As Object, ByVal e As DataGridPageChangedEventArgs)

        ItemsGrid.CurrentPageIndex = e.NewPageIndex
        start_index = ItemsGrid.CurrentPageIndex * ItemsGrid.PageSize
        BindGrid()

    End Sub 'Grid_Change

    Sub BindGrid()

        ItemsGrid.DataSource = CreateDataSource()
        ItemsGrid.DataBind()

    End Sub 'BindGrid 


End Class
