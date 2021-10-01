Imports System
Imports System.Data
Imports System.Configuration
Imports Hsinyi.EIP.ProjectManagement.DataAccessLayer

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    '****************************************************************************
    '
    ' DocCategory Class
    '
    ' The DocCategory class represents a list of Document Category.
    '
    '****************************************************************************

    Public Class PMDocCategories


        Private _name As String
        Private _categoryID As Integer
        Private _description As String
        Private _sort As Integer

#Region "ÄÝ©Ê³]©w"



        Public Property CategoryID() As Integer
            Get
                Return _categoryID
            End Get
            Set(ByVal Value As Integer)
                _categoryID = Value
            End Set
        End Property
        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(ByVal Value As String)
                _name = Value
            End Set
        End Property



        Public Property Description() As String
            Get
                Return _description
            End Get
            Set(ByVal Value As String)
                _description = Value
            End Set
        End Property

        Public Property Sort() As Integer
            Get
                Return _sort
            End Get
            Set(ByVal Value As Integer)
                _sort = Value
            End Set
        End Property
#End Region

        Public Shared Function GetCategoryName(ByVal CategoryID As Integer) As String
            Return Convert.ToString(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetDocCategoryName", CategoryID))
        End Function

        Public Shared Function GetCategorys() As PMDocCategoriesCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListAllDocCategories")
            Dim CategoriesArray As New PMDocCategoriesCollection

            ' Separate Data into a collection of Roles
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows
                Dim DocCategory As New PMDocCategories
                DocCategory.CategoryID = Convert.ToInt32(r("CategoryID"))
                DocCategory.Name = r("CategoryName").ToString()
                DocCategory.Description = r("Description").ToString()
                DocCategory.Sort = Convert.ToInt32(r("Sort"))
                CategoriesArray.Add(DocCategory)
            Next r

            Return CategoriesArray
        End Function 'GetCategorys

        Public Shared Function GetCategoryBySort(ByVal sort As Integer) As PMDocCategories
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetDocCategoryBySort", sort)

            Dim DocCategory As New PMDocCategories
            ' Separate Data into a collection of Roles
            Dim r As DataRow
            For Each r In ds.Tables(0).Rows

                DocCategory.CategoryID = Convert.ToInt32(r("CategoryID"))
                DocCategory.Name = r("CategoryName").ToString()
                DocCategory.Description = r("Description").ToString()
                DocCategory.Sort = Convert.ToInt32(r("Sort"))

            Next r

            Return DocCategory
        End Function 'GetCategorys

    End Class 'DocCategories 

End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer