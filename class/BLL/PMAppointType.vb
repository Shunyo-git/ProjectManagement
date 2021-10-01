Imports System
Imports System.Data
Imports System.Text
Imports System.Configuration
Imports Hsinyi.EIP.ProjectManagement.DataAccessLayer

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    Public Class PMAppointType

        Private _appointTypeID As Integer
        Private _TypeName As String
        Private _createdUserID As Integer
        Private _createdDate As Date
        Private _modifiedUserID As Integer
        Private _modifiedDate As Date
        Private _isDel As Boolean

#Region "  專案屬性存取 "
        Public Property TypeID() As Integer
            Get
                Return _appointTypeID
            End Get
            Set(ByVal Value As Integer)
                _appointTypeID = Value
            End Set
        End Property


        Public Property TypeName() As String
            Get
                Return _TypeName

            End Get
            Set(ByVal Value As String)
                _TypeName = Value
            End Set
        End Property

        Public Property CreatedUserID() As Integer
            Get
                Return _createdUserID
            End Get
            Set(ByVal Value As Integer)
                _createdUserID = Value
            End Set
        End Property


        Public Property ModifiedUserID() As Integer
            Get
                Return _modifiedUserID
            End Get
            Set(ByVal Value As Integer)
                _modifiedUserID = Value
            End Set
        End Property



        Public Property CreatedDate() As DateTime
            Get
                Return _createdDate
            End Get

            Set(ByVal Value As DateTime)
                _createdDate = Value
            End Set
        End Property

        Public Property ModifiedDate() As DateTime
            Get
                Return _modifiedDate
            End Get
            Set(ByVal Value As DateTime)
                _modifiedDate = Value
            End Set
        End Property

        Public Property isDel() As Boolean
            Get
                Return _isDel
            End Get
            Set(ByVal Value As Boolean)
                _isDel = Value
            End Set
        End Property
#End Region

        Public Sub New()
        End Sub 'New

        Public Sub New(ByVal TypeID As Integer)
            _appointTypeID = TypeID
        End Sub 'New

        Public Overloads Shared Function GetAppointTypes() As PMAppointTypeCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListAllAppointTypes")

            ' Separate Data into a Collection of projects
            Dim types As New PMAppointTypeCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim atype As New PMAppointType
                    atype.TypeID = Convert.ToInt32(r("AppointTypeID"))
                    atype.TypeName = Convert.ToString(r("TypeName"))
                    atype.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    atype.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    atype.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    atype.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))
                    types.Add(atype)
                Next r

            End If
            Return types
        End Function 'GetAppointTypes

        Public Shared Function GetTypeName(ByVal AppointTypeID As Integer) As String
            Return Convert.ToString(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetAppointTypeName", AppointTypeID))
        End Function 'GetTypeName

    End Class 'PMAppointType

End Namespace 'Hsinyi.EIP.ProjectManagement.BusinessLogicLayer


