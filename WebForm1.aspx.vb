Public Class WebForm1
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents ItemsGrid As System.Web.UI.WebControls.DataGrid
    Protected WithEvents CheckBox1 As System.Web.UI.WebControls.CheckBox

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
