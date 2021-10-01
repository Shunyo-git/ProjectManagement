Imports System
Imports System.Data
Imports System.Text
Imports System.Configuration
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.DataAccessLayer

Namespace Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

    Public Class PMAppoint

#Region "變數宣告"
        '時間因素
        Private _appointID As Integer
        Private _appointTypeID As Integer
        Private _appointLevelID As AppointLevel
        Private _cycleID As Integer = 0
        Private _projectID As Integer
        Private _subject As String
        Private _location As String
        Private _description As String
        Private _appointDate As Date
        Private _startTime As Date
        Private _endTime As Date
        Private _reminderTime As Integer
        Private _finishedTimes As Integer
        Private _createdUserID As Integer
        Private _createdDate As Date
        Private _modifiedUserID As Integer
        Private _modifiedDate As Date
        Private _isDel As Boolean

        '周期因素
        Private _cycleTypeID As CycleType '1每日 2每周 3每月
        Private _cycleSubject As String
        Private _cycleStartDate As Date
        Private _cycleEndDate As Date
        Private _cycleStartTime As Date
        Private _cycleEndTime As Date
        Private _cycleMonthInterval As Integer '月間格
        Private _cycleMonthDay As Integer '月日期
        Private _cycleWeekInterval As Integer '週間格
        Private _cycleWeekDay As DayOfWeek ' 周幾
        Private _cycleWeekDays As String = String.Empty '周幾(複)
        Private _cycleMonthTypeID As CycleMonthType '1每月X日 2.每X月第X周 星期X
        Private _cycleDayTypeID As CycleDayType '1每日 2.每工作日
#End Region


#Region " 參數 "
        Public Enum CycleType
            InitValue = 0
            DayCycle = 1
            WeekCycle = 2
            MonthCycle = 3
        End Enum

        Public Enum CycleDayType
            InitValue = 0
            Everyday = 1
            WorkDay = 2
        End Enum

        Public Enum CycleMonthType
            InitValue = 0
            MonthDay = 1
            MonthWeek = 2
        End Enum

        Public Enum AppointLevel
            InitValue
            Top
            Middle
            Low
        End Enum

        Public Enum TimeOfAppoint
            InitValue
            Day
            Week
            Month
        End Enum
#End Region

#Region "   一般屬性存取 "
        Public ReadOnly Property isCycle() As Boolean
            Get
                Return _cycleID > 0
            End Get

        End Property
        Public Property AppointID() As Integer
            Get
                Return _appointID
            End Get
            Set(ByVal Value As Integer)
                _appointID = Value
            End Set
        End Property

        Public ReadOnly Property _appointTypeName() As String
            Get
                Return PMAppointType.GetTypeName(_appointTypeID)
            End Get

        End Property

        Public Property AppointTypeID() As Integer
            Get
                Return _appointTypeID
            End Get
            Set(ByVal Value As Integer)
                _appointTypeID = Value
            End Set
        End Property

        Public Property AppointLevelID() As AppointLevel
            Get
                Return _appointLevelID
            End Get
            Set(ByVal Value As AppointLevel)
                _appointLevelID = Value
            End Set
        End Property

        Public Property CycleID() As Integer
            Get
                Return _cycleID
            End Get
            Set(ByVal Value As Integer)
                _cycleID = Value
            End Set
        End Property

        Public Property ProjectID() As Integer
            Get
                Return _projectID
            End Get
            Set(ByVal Value As Integer)
                _projectID = Value
            End Set
        End Property

        Public Property Subject() As String
            Get
                Return _subject

            End Get
            Set(ByVal Value As String)
                _subject = Value
            End Set
        End Property

        Public Property Location() As String
            Get
                Return _location

            End Get
            Set(ByVal Value As String)
                _location = Value
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

        Public Property AppointDate() As Date
            Get
                Return _appointDate

            End Get
            Set(ByVal Value As Date)
                _appointDate = Value
            End Set
        End Property

        Public Property StartTime() As Date
            Get
                Return _startTime

            End Get
            Set(ByVal Value As Date)
                _startTime = Value
            End Set
        End Property

        Public Property EndTime() As Date
            Get
                Return _endTime
            End Get

            Set(ByVal Value As Date)
                _endTime = Value
            End Set
        End Property

        Public Property ReminderTime() As Integer
            Get
                Return _reminderTime
            End Get
            Set(ByVal Value As Integer)
                _reminderTime = Value
            End Set
        End Property

        Public Property FinishedTimes() As Integer
            Get
                Return _finishedTimes
            End Get
            Set(ByVal Value As Integer)
                _finishedTimes = Value
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

#Region "   週期屬性存取 "
        Public Property CycleTypeID() As CycleType
            Get
                Return _cycleTypeID
            End Get
            Set(ByVal Value As CycleType)
                _cycleTypeID = Value
            End Set
        End Property

        Public Property CycleSubject() As String
            Get
                Return _cycleSubject
            End Get
            Set(ByVal Value As String)
                _cycleSubject = Value
            End Set
        End Property

        Public Property CycleStartDate() As Date
            Get
                Return _cycleStartDate
            End Get
            Set(ByVal Value As Date)
                _cycleStartDate = Value
            End Set
        End Property

        Public Property CycleEndDate() As Date
            Get
                Return _cycleEndDate
            End Get
            Set(ByVal Value As Date)
                _cycleEndDate = Value
            End Set
        End Property

        Public Property CycleStartTime() As Date
            Get
                Return _cycleStartTime
            End Get
            Set(ByVal Value As Date)
                _cycleStartTime = Value
            End Set
        End Property

        Public Property CycleEndTime() As Date
            Get
                Return _cycleEndTime
            End Get
            Set(ByVal Value As Date)
                _cycleEndTime = Value
            End Set
        End Property

        Public Property CycleMonthInterval() As Integer
            Get
                Return _cycleMonthInterval
            End Get
            Set(ByVal Value As Integer)
                _cycleMonthInterval = Value
            End Set
        End Property

        Public Property CycleMonthDay() As Integer
            Get
                Return _cycleMonthDay
            End Get
            Set(ByVal Value As Integer)
                _cycleMonthDay = Value
            End Set
        End Property

        Public Property CycleWeekInterval() As Integer
            Get
                Return _cycleWeekInterval
            End Get
            Set(ByVal Value As Integer)
                _cycleWeekInterval = Value
            End Set
        End Property

        Public Property CycleWeekDay() As DayOfWeek
            Get
                Return _cycleWeekDay
            End Get
            Set(ByVal Value As DayOfWeek)
                _cycleWeekDay = Value
            End Set
        End Property

        Public Property CycleWeekDays() As String
            Get
                Return _cycleWeekDays
            End Get
            Set(ByVal Value As String)
                _cycleWeekDays = Value
            End Set
        End Property

        Public Property CycleMonthTypeID() As CycleMonthType
            Get
                Return _cycleMonthTypeID
            End Get
            Set(ByVal Value As CycleMonthType)
                _cycleMonthTypeID = Value
            End Set
        End Property

        Public Property CycleDayTypeID() As CycleDayType
            Get
                Return _cycleDayTypeID
            End Get
            Set(ByVal Value As CycleDayType)
                _cycleDayTypeID = Value
            End Set
        End Property

#End Region

        Public Sub New()
        End Sub 'New

        Public Sub New(ByVal AppointID As Integer)
            _appointID = AppointID
        End Sub 'New

        Public Sub New(ByVal AppointID As Integer, ByVal CycleID As Integer)
            _appointID = AppointID
            _cycleID = CycleID
        End Sub 'New

        Public Overloads Shared Function GetAppointments() As PMAppointCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListAllAppointments")

            ' Separate Data into a Collection of projects
             Return LoadAppoints(ds)
        End Function 'GetAppointments

        Public Overloads Shared Function GetAppointments(ByVal projectID As Integer) As PMAppointCollection
              Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListAppointments", projectID)

            ' Separate Data into a Collection of projects
             Return LoadAppoints(ds)

        End Function 'GetAppointments

        Public Overloads Shared Function GetAppointments(ByVal projectID As Integer, ByVal theDate As Date) As PMAppointCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListDateAppointments", projectID, theDate, TimeOfAppoint.Day)

            ' Separate Data into a Collection of projects
             Return LoadAppoints(ds)

        End Function 'GetAppointments
        Public Overloads Shared Function GetAppointments(ByVal projectID As Integer, ByVal AppointDate As Date, ByVal TimeOfAppoint As TimeOfAppoint) As PMAppointCollection


            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListDateAppointments", projectID, AppointDate, TimeOfAppoint)

            ' Separate Data into a Collection of projects
            Return LoadAppoints(ds)

        End Function 'GetAppointments

        Private Shared Function LoadAppoints(ByVal ds As DataSet) As PMAppointCollection
            Dim Appointments As New PMAppointCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim Appoint As New PMAppoint
                    Appoint.AppointID = Convert.ToInt32(r("AppointID"))
                    Appoint.ProjectID = Convert.ToInt32(r("ProjectID"))
                    Appoint.AppointTypeID = Convert.ToInt32(r("AppointTypeID"))
                    Appoint.AppointLevelID = Convert.ToInt32(r("AppointLevelID"))
                    Appoint.CycleID = Convert.ToInt32(r("CycleID"))
                    Appoint.Subject = Convert.ToString(r("Subject"))
                    Appoint.Location = Convert.ToString(r("Location"))
                    Appoint.Description = Convert.ToString(r("Description"))
                    Appoint.AppointDate = Convert.ToDateTime(r("AppointDate"))
                    Appoint.StartTime = Convert.ToDateTime(r("StartTime"))
                    Appoint.EndTime = Convert.ToDateTime(r("EndTime"))

                    Appoint.ReminderTime = Convert.ToInt32(r("ReminderTime"))
                    Appoint.FinishedTimes = Convert.ToInt32(r("FinishedTimes"))

                    Appoint.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
                    Appoint.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    Appoint.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
                    Appoint.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))

                    If Appoint.CycleID > 0 Or Appoint.CycleTypeID <> CycleType.InitValue Then
                        Appoint.LoadCycle()
                    End If
                    Appointments.Add(Appoint)
                Next r

            End If
            Return Appointments
        End Function
        'Public Shared Function GetCycleAppointments(ByVal cycleID As Integer) As PMAppointCollection
        '    'Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), CommandType.StoredProcedure, "PM_ListEvents", projectID)
        '    Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_ListCycleAppointments", cycleID)

        '    ' Separate Data into a Collection of projects
        '    Dim Appointments As New PMAppointCollection
        '    If ds.Tables.Count > 0 Then
        '        Dim r As DataRow
        '        For Each r In ds.Tables(0).Rows
        '            Dim Appoint As New PMAppoint
        '            Appoint.AppointID = Convert.ToInt32(r("AppointID"))
        '            Appoint.ProjectID = Convert.ToInt32(r("ProjectID"))
        '            Appoint.AppointTypeID = Convert.ToInt32(r("AppointTypeID"))
        '            Appoint.AppointLevelID = Convert.ToInt32(r("AppointLevelID"))
        '            Appoint.CycleID = Convert.ToInt32(r("CycleID"))
        '            Appoint.Subject = Convert.ToString(r("Subject"))
        '            Appoint.Location = Convert.ToString(r("Location"))
        '            Appoint.Description = Convert.ToString(r("Description"))
        '            Appoint.AppointDate = Convert.ToDateTime(r("AppointDate"))
        '            Appoint.StartTime = Convert.ToDateTime(r("StartTime"))
        '            Appoint.EndTime = Convert.ToDateTime(r("EndTime"))

        '            Appoint.ReminderTime = Convert.ToInt32(r("ReminderTime"))
        '            Appoint.FinishedTimes = Convert.ToInt32(r("FinishedTimes"))

        '            Appoint.CreatedUserID = Convert.ToInt32(r("CreatedUserID"))
        '            Appoint.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
        '            Appoint.ModifiedUserID = Convert.ToInt32(r("ModifiedUserID"))
        '            Appoint.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))


        '            Appointments.Add(Appoint)
        '        Next r

        '    End If
        '    Return Appointments

        'End Function 'GetAppointments

        Public Function Load() As Boolean
            ' The Get Projects stored procedure returns a dataset with 3 tables: Projects, Members, and Categories.
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetAppoint", _appointID, _cycleID)

            If ds.Tables(0).Rows.Count < 1 Then
                Return False
            End If

            _appointID = Convert.ToInt32(ds.Tables(0).Rows(0)("AppointID"))
            _projectID = Convert.ToInt32(ds.Tables(0).Rows(0)("ProjectID"))
            _appointTypeID = Convert.ToInt32(ds.Tables(0).Rows(0)("AppointTypeID"))
            _appointLevelID = Convert.ToInt32(ds.Tables(0).Rows(0)("AppointLevelID"))
            _cycleID = Convert.ToInt32(ds.Tables(0).Rows(0)("CycleID"))
            _subject = Convert.ToString(ds.Tables(0).Rows(0)("Subject"))
            _location = Convert.ToString(ds.Tables(0).Rows(0)("Location"))
            _description = Convert.ToString(ds.Tables(0).Rows(0)("Description"))
            _appointDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("AppointDate"))
            _startTime = Convert.ToDateTime(ds.Tables(0).Rows(0)("StartTime"))
            _endTime = Convert.ToDateTime(ds.Tables(0).Rows(0)("EndTime"))

            _reminderTime = Convert.ToInt32(ds.Tables(0).Rows(0)("ReminderTime"))
            _finishedTimes = Convert.ToInt32(ds.Tables(0).Rows(0)("FinishedTimes"))

            _createdUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("CreatedUserID"))
            _createdDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("CreatedDate"))
            _modifiedUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("ModifiedUserID"))
            _modifiedDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("ModifiedDate"))


            If _cycleID > 0 Or _cycleTypeID <> CycleType.InitValue Then
                LoadCycle()
            End If

            Return True
        End Function 'Load

        Public Function LoadCycle() As Boolean
            ' The Get Projects stored procedure returns a dataset with 3 tables: Projects, Members, and Categories.
            If _cycleID <= 0 Then
                Return False
            End If

            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_GetAppointCycle", _cycleID)

            If ds.Tables(0).Rows.Count < 1 Then
                Return False
            End If


            _cycleTypeID = Convert.ToInt32(ds.Tables(0).Rows(0)("CycleTypeID"))
            _cycleSubject = Convert.ToString(ds.Tables(0).Rows(0)("Subject"))
            _cycleStartDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("CycleStartDate")).ToShortDateString
            _cycleEndDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("CycleEndDate")).ToShortDateString
            _cycleStartTime = Convert.ToDateTime(ds.Tables(0).Rows(0)("StartTime")).ToShortTimeString
            _cycleEndTime = Convert.ToDateTime(ds.Tables(0).Rows(0)("EndTime")).ToShortTimeString

            Select Case _cycleTypeID
                Case CycleType.DayCycle
                    _cycleDayTypeID = Convert.ToInt32(ds.Tables(1).Rows(0)("CycleDayTypeID"))

                Case CycleType.WeekCycle
                    _cycleWeekInterval = Convert.ToInt32(ds.Tables(2).Rows(0)("WeekInterval"))
                    _cycleWeekDays = Convert.ToString(ds.Tables(2).Rows(0)("WeekDays"))
                Case CycleType.MonthCycle
                    _cycleMonthTypeID = Convert.ToInt32(ds.Tables(3).Rows(0)("CycleMonthTypeID"))
                    _cycleMonthInterval = Convert.ToInt32(ds.Tables(3).Rows(0)("MonthInterval"))
                    _cycleMonthDay = Convert.ToInt32(ds.Tables(3).Rows(0)("MonthDay"))
                    _cycleWeekInterval = Convert.ToInt32(ds.Tables(3).Rows(0)("WeekInterval"))
                    _cycleWeekDay = Convert.ToInt32(ds.Tables(3).Rows(0)("WeekDay"))
            End Select


            Return True
        End Function 'LoadCycle
        Public Shared Sub Remove(ByVal appointID As Integer)
            SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_DeleteAppoint", appointID)
        End Sub 'Remove

        Public Shared Sub RemoveCycle(ByVal CycleID As Integer)
            SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_DeleteAppointCycle", CycleID)
        End Sub 'RemoveCycle

        Public Function Save() As Boolean
            If _appointID = 0 Then
                Return Insert()
            Else

                If _appointID > 0 Then
                    'Return Update()
                    Remove(_appointID)
                    RemoveCycle(_cycleID)
                    Return Insert()
                Else
                    _appointID = 0
                    Return False
                End If
            End If
        End Function 'Save
        '*********************************************************************
        '
        ' Inserts a project along with its members and categories into the databse
        '
        '*********************************************************************

        Private Function Insert() As Boolean

            Dim blnResult As Boolean

            'Try
            If _cycleTypeID <> CycleType.InitValue Then

                Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_AddAppointCycle", _cycleTypeID, _cycleSubject, _cycleStartDate, _cycleEndDate, _cycleStartTime, _cycleEndTime, _cycleMonthInterval, _cycleMonthDay, _cycleWeekInterval, _cycleWeekDay, _cycleWeekDays, _cycleMonthTypeID, _cycleDayTypeID, _appointTypeID, _appointLevelID, _projectID, _subject, _location, _description, _appointDate, _startTime, _endTime, _reminderTime, _finishedTimes, _modifiedUserID)
                If ds.Tables.Count > 1 Then
                    _appointID = Convert.ToInt32(ds.Tables(0).Rows(0)("AppointID"))
                    _cycleID = Convert.ToInt32(ds.Tables(ds.Tables.Count - 1).Rows(0)("CycleID"))
                End If
                ' _cycleID = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_AddAppointCycle", _cycleTypeID, _cycleSubject, _cycleStartDate, _cycleEndDate, _cycleStartTime, _cycleEndTime, _cycleMonthInterval, _cycleMonthDay, _cycleWeekInterval, _cycleWeekDay, _cycleWeekDays, _cycleMonthTypeID, _cycleDayTypeID, _appointTypeID, _appointLevelID, _projectID, _subject, _location, _description, _appointDate, _startTime, _endTime, _reminderTime, _finishedTimes, _modifiedUserID))


            Else
                _appointID = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_AddAppoint", _appointTypeID, _appointLevelID, _cycleID, _projectID, _subject, _location, _description, _appointDate, _startTime, _endTime, _reminderTime, _finishedTimes, _modifiedUserID))

            End If

            blnResult = True

            'Catch ex As Exception
            '    Throw ex
            '    blnResult = False
            'End Try


            Return blnResult

        End Function 'Insert
        '*********************************************************************
        '
        ' Updates a project along with its members and categories.
        '
        '*********************************************************************

        Private Function Update() As Boolean
            Dim blnResult As Boolean

            Try

                If _cycleTypeID <> CycleType.InitValue Then

                    Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_UpdateAppointCycle", _cycleID, _cycleTypeID, _cycleSubject, _cycleStartDate, _cycleEndDate, _cycleStartTime, _cycleEndTime, _cycleMonthInterval, _cycleMonthDay, _cycleWeekInterval, _cycleWeekDay, _cycleWeekDays, _cycleMonthTypeID, _cycleDayTypeID, _modifiedUserID)
                    'Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_AddAppointCycle", _cycleTypeID, _cycleSubject, _cycleStartDate, _cycleEndDate, _cycleStartTime, _cycleEndTime, _cycleMonthInterval, _cycleMonthDay, _cycleWeekInterval, _cycleWeekDay, _cycleWeekDays, _cycleMonthTypeID, _cycleDayTypeID, _appointTypeID, _appointLevelID, _projectID, _subject, _location, _description, _appointDate, _startTime, _endTime, _reminderTime, _finishedTimes, _modifiedUserID)

                    If ds.Tables.Count > 1 Then
                        _appointID = Convert.ToInt32(ds.Tables(0).Rows(0)("AppointID"))
                        _cycleID = Convert.ToInt32(ds.Tables(ds.Tables.Count - 1).Rows(0)("CycleID"))
                    End If
                Else

                    ' _appointID = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_AddAppoint", _appointTypeID, _appointLevelID, _cycleID, _projectID, _subject, _location, _description, _appointDate, _startTime, _endTime, _reminderTime, _finishedTimes, _modifiedUserID))
                    SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyPMConnString), "PM_UpdateAppoint", _appointID, _appointTypeID, _appointLevelID, _cycleID, _projectID, _subject, _location, _description, _appointDate, _startTime, _endTime, _reminderTime, _finishedTimes, _modifiedUserID)

                End If
                blnResult = True
            Catch ex As Exception
                Throw ex
                blnResult = False
            End Try

            Return blnResult
        End Function 'Update

        Public Shared Function GetDays() As DataTable

            Dim dt As New DataTable
            dt.Columns.Add("Day")
            dt.Columns.Add("Date")

            Dim workRow As DataRow

            '  fill table with dates and days
            Dim i As Integer
            For i = 1 To 31
                workRow = dt.NewRow()
                workRow("Day") = i
                workRow("Date") = i
                dt.Rows.Add(workRow)
            Next i
            Return dt
        End Function 'GetDays
        '*********************************************************************
        '
        ' GetWeek Static Method
        '
        ' The TimeEntry.GetWeek Method return a table of dates in a week, 
        ' given any date during the week, this table contains 2 columns, 
        ' one for Day values, and one for Date values
        ' This method is primary used in Databinding DropDowns in TimeEntry page
        ' with range of a week. 
        '
        '*********************************************************************
        Public Shared Function GetWeek(ByVal selectedDate As DateTime) As DataTable
            Dim start As DateTime = DateTime.MinValue
            Dim [end] As DateTime = DateTime.MinValue
            Dim dt As New DataTable
            dt.Columns.Add("Day")
            dt.Columns.Add("Week")

            FillCorrectStartEndDates(selectedDate, start, [end])
            Dim workRow As DataRow

            '  fill table with dates and days
            Dim i As Integer
            For i = 0 To 6
                workRow = dt.NewRow()
                workRow("Day") = Convert.ToDateTime(start.AddDays(i)).ToString("ddd")
                workRow("Week") = Weekday(start.AddDays(i))
                dt.Rows.Add(workRow)
            Next i
            Return dt
        End Function 'GetWeek
        '*********************************************************************
        '
        ' FillCorrectStartEndDates Static Method
        '
        ' The TimeEntry.FillCorrectStartEndDates Method returns the first 
        ' and last date of a week, given any date during the week.
        ' This method is primary used in TimeEntry page to calculate the range
        ' of a week.
        '
        '*********************************************************************
        Public Shared Sub FillCorrectStartEndDates(ByVal selectedDate As DateTime, ByRef startDate As DateTime, ByRef endDate As DateTime)
            Dim firstDayOfWeek As Integer = Convert.ToInt32(ConfigurationSettings.AppSettings(Web.Global.CfgKeyFirstDayOfWeek))

            ' Loop through 7 days of the week.
            Dim i As Integer
            For i = 0 To 6
                ' Match first day of the week from global setting
                If Convert.ToInt32(selectedDate.AddDays(i).DayOfWeek) = firstDayOfWeek Then
                    ' Fill in correct start and end dates
                    startDate = selectedDate.AddDays(i)
                    If i <> 0 Then
                        startDate = startDate.AddDays(-7)
                    End If
                    endDate = startDate.AddDays(6)
                    Exit For
                End If
            Next i
        End Sub 'FillCorrectStartEndDates
        Public Shared Function getWeekName(ByVal val As DayOfWeek) As String
            Dim start As DateTime = DateTime.MinValue
            Dim i As Integer
            For i = 0 To 6
                If start.AddDays(i).DayOfWeek = val Then
                    Return Convert.ToDateTime(start.AddDays(i)).ToString("ddd")
                End If
            Next
        End Function
    End Class

End Namespace

