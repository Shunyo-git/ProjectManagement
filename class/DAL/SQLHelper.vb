'*********************************************************************
' Microsoft Data Access Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/daab-rm.asp
'
' SQLHelper.cs
'
' This file contains the implementations of the SqlHelper and SqlHelperParameterCache
' classes.
'
' For more information see the Data Access Application Block Implementation Overview. 
' 
'*********************************************************************
' Copyright (C) 2000-2001 Microsoft Corporation
' All rights reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
' OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
' LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR
' FITNESS FOR A PARTICULAR PURPOSE.
'*********************************************************************
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections
Imports HsinYi.EIP.ProjectManagement.BusinessLogicLayer

Namespace HsinYi.EIP.ProjectManagement.DataAccessLayer
    _
    '*********************************************************************
    '
    ' The SqlHelper class is intended to encapsulate high performance, scalable best practices for 
    ' common uses of SqlClient.
    '
    '*********************************************************************
    Public NotInheritable Class SqlHelper

        '*********************************************************************
        '
        ' Since this class provides only static methods, make the default constructor private to prevent 
        ' instances from being created with "new SqlHelper()".
        '
        '*********************************************************************
        Private Sub New()
        End Sub 'New

        '*********************************************************************
        '
        ' This method is used to attach array of SqlParameters to a SqlCommand.
        ' 
        ' This method will assign a value of DbNull to any parameter with a direction of
        ' InputOutput and a value of null.  
        ' 
        ' This behavior will prevent default values from being used, but
        ' this will be the less common case than an intended pure output parameter (derived as InputOutput)
        ' where the user provided no input value.
        ' 
        ' param name="command" The command to which the parameters will be added
        ' param name="commandParameters" an array of SqlParameters tho be added to command
        '
        '*********************************************************************
        Private Shared Sub AttachParameters(ByVal command As SqlCommand, ByVal commandParameters() As SqlParameter)
            Dim p As SqlParameter
            For Each p In commandParameters
                'check for derived output value with no value assigned
                If p.Direction = ParameterDirection.InputOutput And p.Value Is Nothing Then
                    p.Value = Nothing
                End If
                command.Parameters.Add(p)
            Next p
        End Sub 'AttachParameters


        '*********************************************************************
        '
        ' This method assigns an array of values to an array of SqlParameters.
        ' 
        ' param name="commandParameters" array of SqlParameters to be assigned values
        ' param name="parameterValues" array of objects holding the values to be assigned
        '
        '*********************************************************************
        Private Shared Sub AssignParameterValues(ByVal commandParameters() As SqlParameter, ByVal parameterValues() As Object)

            Dim i As Short
            Dim j As Short

            If (commandParameters Is Nothing) And (parameterValues Is Nothing) Then
                'do nothing if we get no data
                Return
            End If

            ' we must have the same number of values as we pave parameters to put them in
            If commandParameters.Length <> parameterValues.Length Then
                Throw New ArgumentException("Parameter count does not match Parameter Value count.")
            End If

            'value array
            j = commandParameters.Length - 1
            For i = 0 To j
                commandParameters(i).Value = parameterValues(i)
            Next

        End Sub 'AssignParameterValues

        '*********************************************************************
        '
        ' This method opens (if necessary) and assigns a connection, transaction, command type and parameters 
        ' to the provided command.
        ' 
        ' param name="command" the SqlCommand to be prepared
        ' param name="connection" a valid SqlConnection, on which to execute this command
        ' param name="transaction" a valid SqlTransaction, or 'null'
        ' param name="commandType" the CommandType (stored procedure, text, etc.)
        ' param name="commandText" the stored procedure name or T-SQL command
        ' param name="commandParameters" an array of SqlParameters to be associated with the command or 'null' if no parameters are required
        '
        '*********************************************************************
        Private Shared Sub PrepareCommand(ByVal command As SqlCommand, _
                                          ByVal connection As SqlConnection, _
                                          ByVal transaction As SqlTransaction, _
                                          ByVal commandType As CommandType, _
                                          ByVal commandText As String, _
                                          ByVal commandParameters() As SqlParameter)

            'if the provided connection is not open, we will open it
            If connection.State <> ConnectionState.Open Then
                connection.Open()
            End If

            'associate the connection with the command
            command.Connection = connection

            'set the command text (stored procedure name or SQL statement)
            command.CommandText = commandText

            'if we were provided a transaction, assign it.
            If Not (transaction Is Nothing) Then
                command.Transaction = transaction
            End If

            'set the command type
            command.CommandType = commandType

            'attach the command parameters if they are provided
            If Not (commandParameters Is Nothing) Then
                AttachParameters(command, commandParameters)
            End If

            Return
        End Sub 'PrepareCommand


        '*********************************************************************
        '
        ' Execute a SqlCommand (that returns no resultset) against the database specified in the connection string 
        ' using the provided parameters.
        '
        ' e.g.:  
        '  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        '
        ' param name="connectionString" a valid connection string for a SqlConnection
        ' param name="commandType" the CommandType (stored procedure, text, etc.)
        ' param name="commandText" the stored procedure name or T-SQL command
        ' param name="commandParameters" an array of SqlParamters used to execute the command
        ' returns an int representing the number of rows affected by the command
        '
        '*********************************************************************
        Public Overloads Shared Function ExecuteNonQuery(ByVal connectionString As String, _
                                                         ByVal commandType As CommandType, _
                                                         ByVal commandText As String, _
                                                         ByVal ParamArray commandParameters() As SqlParameter) As Integer
            'create & open a SqlConnection, and dispose of it after we are done.
            Dim cn As New SqlConnection(connectionString)
            Try
                cn.Open()

                'call the overload that takes a connection in place of the connection string
                Return ExecuteNonQuery(cn, commandType, commandText, commandParameters)
            Finally
                cn.Dispose()
            End Try
        End Function 'ExecuteNonQuery


        '*********************************************************************
        '
        ' Execute a stored procedure via a SqlCommand (that returns no resultset) against the database specified in 
        ' the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
        ' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
        ' 
        ' This method provides no access to output parameters or the stored procedure's return value parameter.
        ' 
        ' e.g.:  
        '  int result = ExecuteNonQuery(connString, "PublishOrders", 24, 36);
        '
        ' param name="connectionString" a valid connection string for a SqlConnection
        ' param name="spName" the name of the stored prcedure
        ' param name="parameterValues" an array of objects to be assigned as the input values of the stored procedure
        ' returns an int representing the number of rows affected by the command
        '
        '*********************************************************************
        Public Overloads Shared Function ExecuteNonQuery(ByVal connectionString As String, _
                                                         ByVal spName As String, _
                                                         ByVal ParamArray parameterValues() As Object) As Integer
            Dim commandParameters As SqlParameter()

            'if we receive parameter values, we need to figure out where they go
            If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
                'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)

                commandParameters = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)

                'assign the provided values to these parameters based on parameter order
                AssignParameterValues(commandParameters, parameterValues)

                'call the overload that takes an array of SqlParameters
                Return ExecuteNonQuery(connectionString, CommandType.StoredProcedure, spName, commandParameters)
                'otherwise we can just call the SP without params
            Else
                Return ExecuteNonQuery(connectionString, CommandType.StoredProcedure, spName)
            End If
        End Function 'ExecuteNonQuery

        '*********************************************************************
        '
        ' Execute a SqlCommand (that returns no resultset) against the specified SqlConnection 
        ' using the provided parameters.
        ' 
        ' e.g.:  
        '  int result = ExecuteNonQuery(conn, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        ' 
        ' param name="connection" a valid SqlConnection 
        ' param name="commandType" the CommandType (stored procedure, text, etc.) 
        ' param name="commandText" the stored procedure name or T-SQL command 
        ' param name="commandParameters" an array of SqlParamters used to execute the command 
        ' returns an int representing the number of rows affected by the command
        '
        '*********************************************************************
        Public Overloads Shared Function ExecuteNonQuery(ByVal connection As SqlConnection, _
                                                        ByVal commandType As CommandType, _
                                                        ByVal commandText As String, _
                                                        ByVal ParamArray commandParameters() As SqlParameter) As Integer

            'create a command and prepare it for execution
            Dim cmd As New SqlCommand
            Dim retval As Integer

            PrepareCommand(cmd, connection, CType(Nothing, SqlTransaction), commandType, commandText, commandParameters)

            'finally, execute the command.
            retval = cmd.ExecuteNonQuery()

            'detach the SqlParameters from the command object, so they can be used again
            cmd.Parameters.Clear()

            Return retval

        End Function 'ExecuteNonQuery


        '*********************************************************************
        '
        ' Execute a SqlCommand (that returns a resultset) against the database specified in the connection string 
        ' using the provided parameters.
        ' 
        ' e.g.:  
        '  DataSet ds = ExecuteDataset(connString, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24));
        ' 
        ' param name="connectionString" a valid connection string for a SqlConnection 
        ' param name="commandType" the CommandType (stored procedure, text, etc.) 
        ' param name="commandText" the stored procedure name or T-SQL command 
        ' param name="commandParameters" an array of SqlParamters used to execute the command 
        ' returns a dataset containing the resultset generated by the command
        '
        '*********************************************************************
        Public Overloads Shared Function ExecuteDataset(ByVal connectionString As String, _
                                                        ByVal commandType As CommandType, _
                                                        ByVal commandText As String, _
                                                        ByVal ParamArray commandParameters() As SqlParameter) As DataSet
            'create & open a SqlConnection, and dispose of it after we are done.
            Dim cn As New SqlConnection(connectionString)
            Try
                cn.Open()

                'call the overload that takes a connection in place of the connection string
                Return ExecuteDataset(cn, commandType, commandText, commandParameters)
            Finally
                cn.Dispose()
            End Try
        End Function 'ExecuteDataset


        '*********************************************************************
        '
        ' Execute a stored procedure via a SqlCommand (that returns a resultset) against the database specified in 
        ' the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
        ' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
        ' 
        ' This method provides no access to output parameters or the stored procedure's return value parameter.
        ' 
        ' e.g.:  
        '  DataSet ds = ExecuteDataset(connString, "GetOrders", 24, 36);
        ' 
        ' param name="connectionString" a valid connection string for a SqlConnection
        ' param name="spName" the name of the stored procedure
        ' param name="parameterValues" an array of objects to be assigned as the input values of the stored procedure
        ' returns a dataset containing the resultset generated by the command
        '
        '*********************************************************************
        Public Overloads Shared Function ExecuteDataset(ByVal connectionString As String, _
                                                        ByVal spName As String, _
                                                        ByVal ParamArray parameterValues() As Object) As DataSet

            Dim commandParameters As SqlParameter()

            'if we receive parameter values, we need to figure out where they go
            If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
                'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
                commandParameters = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)

                'assign the provided values to these parameters based on parameter order
                AssignParameterValues(commandParameters, parameterValues)

                'call the overload that takes an array of SqlParameters
                Return ExecuteDataset(connectionString, CommandType.StoredProcedure, spName, commandParameters)
                'otherwise we can just call the SP without params
            Else
                Return ExecuteDataset(connectionString, CommandType.StoredProcedure, spName)
            End If
        End Function 'ExecuteDataset

        '*********************************************************************
        '
        ' Execute a SqlCommand (that returns a resultset) against the specified SqlConnection 
        ' using the provided parameters.
        ' 
        ' e.g.:  
        '  DataSet ds = ExecuteDataset(conn, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24));
        '
        ' param name="connection" a valid SqlConnection
        ' param name="commandType" the CommandType (stored procedure, text, etc.)
        ' param name="commandText" the stored procedure name or T-SQL command
        ' param name="commandParameters" an array of SqlParamters used to execute the command
        ' returns a dataset containing the resultset generated by the command
        '
        '*********************************************************************
        Public Overloads Shared Function ExecuteDataset(ByVal connection As SqlConnection, _
                                                        ByVal commandType As CommandType, _
                                                        ByVal commandText As String, _
                                                        ByVal ParamArray commandParameters() As SqlParameter) As DataSet

            'create a command and prepare it for execution
            Dim cmd As New SqlCommand
            Dim ds As New DataSet
            Dim da As SqlDataAdapter

            PrepareCommand(cmd, connection, CType(Nothing, SqlTransaction), commandType, commandText, commandParameters)

            'create the DataAdapter & DataSet
            da = New SqlDataAdapter(cmd)

            'fill the DataSet using default values for DataTable names, etc.
            da.Fill(ds)

            'detach the SqlParameters from the command object, so they can be used again
            cmd.Parameters.Clear()

            'return the dataset
            Return ds

        End Function 'ExecuteDataset


        '*********************************************************************
        '
        ' Execute a SqlCommand (that returns a 1x1 resultset) against the database specified in the connection string 
        ' using the provided parameters.
        ' 
        ' e.g.:  
        '  int orderCount = (int)ExecuteScalar(connString, CommandType.StoredProcedure, "GetOrderCount", new SqlParameter("@prodid", 24));
        ' 
        ' param name="connectionString" a valid connection string for a SqlConnection 
        ' param name="commandType" the CommandType (stored procedure, text, etc.) 
        ' param name="commandText" the stored procedure name or T-SQL command 
        ' param name="commandParameters" an array of SqlParamters used to execute the command 
        ' returns an object containing the value in the 1x1 resultset generated by the command
        '
        '*********************************************************************
        Public Overloads Shared Function ExecuteScalar(ByVal connectionString As String, _
                                                       ByVal commandType As CommandType, _
                                                       ByVal commandText As String, _
                                                       ByVal ParamArray commandParameters() As SqlParameter) As Object
            'create & open a SqlConnection, and dispose of it after we are done.
            Dim cn As New SqlConnection(connectionString)
            Try
                cn.Open()

                'call the overload that takes a connection in place of the connection string
                Return ExecuteScalar(cn, commandType, commandText, commandParameters)
            Finally
                cn.Dispose()
            End Try
        End Function 'ExecuteScalar



        '*********************************************************************
        '
        ' Execute a stored procedure via a SqlCommand (that returns a 1x1 resultset) against the database specified in 
        ' the connection string using the provided parameter values.  This method will query the database to discover the parameters for the 
        ' stored procedure (the first time each stored procedure is called), and assign the values based on parameter order.
        ' 
        ' This method provides no access to output parameters or the stored procedure's return value parameter.
        ' 
        ' e.g.:  
        '  int orderCount = (int)ExecuteScalar(connString, "GetOrderCount", 24, 36);
        ' 
        ' param name="connectionString" a valid connection string for a SqlConnection 
        ' param name="spName" the name of the stored procedure 
        ' param name="parameterValues" an array of objects to be assigned as the input values of the stored procedure 
        ' returns an object containing the value in the 1x1 resultset generated by the command
        '
        '*********************************************************************
        Public Overloads Shared Function ExecuteScalar(ByVal connectionString As String, _
                                                       ByVal spName As String, _
                                                       ByVal ParamArray parameterValues() As Object) As Object
            Dim commandParameters As SqlParameter()

            'if we receive parameter values, we need to figure out where they go
            If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
                'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
                commandParameters = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)

                'assign the provided values to these parameters based on parameter order
                AssignParameterValues(commandParameters, parameterValues)

                'call the overload that takes an array of SqlParameters
                Return ExecuteScalar(connectionString, CommandType.StoredProcedure, spName, commandParameters)
                'otherwise we can just call the SP without params
            Else
                Return ExecuteScalar(connectionString, CommandType.StoredProcedure, spName)
            End If
        End Function 'ExecuteScalar


        '*********************************************************************
        '
        ' Execute a SqlCommand (that returns a 1x1 resultset) against the specified SqlConnection 
        ' using the provided parameters.
        ' 
        ' e.g.:  
        '  int orderCount = (int)ExecuteScalar(conn, CommandType.StoredProcedure, "GetOrderCount", new SqlParameter("@prodid", 24));
        ' 
        ' param name="connection" a valid SqlConnection 
        ' param name="commandType" the CommandType (stored procedure, text, etc.) 
        ' param name="commandText" the stored procedure name or T-SQL command 
        ' param name="commandParameters" an array of SqlParamters used to execute the command 
        ' returns an object containing the value in the 1x1 resultset generated by the command
        '
        '*********************************************************************
        Public Overloads Shared Function ExecuteScalar(ByVal connection As SqlConnection, _
                                                       ByVal commandType As CommandType, _
                                                       ByVal commandText As String, _
                                                       ByVal ParamArray commandParameters() As SqlParameter) As Object
            'create a command and prepare it for execution
            Dim cmd As New SqlCommand
            Dim retval As Object

            PrepareCommand(cmd, connection, CType(Nothing, SqlTransaction), commandType, commandText, commandParameters)

            'execute the command & return the results
            retval = cmd.ExecuteScalar()

            'detach the SqlParameters from the command object, so they can be used again
            cmd.Parameters.Clear()

            Return retval

        End Function 'ExecuteScalar
    End Class 'SqlHelper
    _

    '*********************************************************************
    '
    ' SqlHelperParameterCache provides functions to leverage a static cache of procedure parameters, and the
    ' ability to discover parameters for stored procedures at run-time.
    '
    '*********************************************************************
    Public NotInheritable Class SqlHelperParameterCache

        '*********************************************************************
        '
        ' Since this class provides only static methods, make the default constructor private to prevent 
        ' instances from being created with "new SqlHelperParameterCache()".
        '
        '*********************************************************************
        Private Sub New()
        End Sub 'New 

        Private Shared paramCache As Hashtable = Hashtable.Synchronized(New Hashtable)

        '*********************************************************************
        '
        ' resolve at run time the appropriate set of SqlParameters for a stored procedure
        ' 
        ' param name="connectionString" a valid connection string for a SqlConnection 
        ' param name="spName" the name of the stored procedure 
        ' param name="includeReturnValueParameter" whether or not to include their return value parameter 
        '
        '*********************************************************************
        Private Shared Function DiscoverSpParameterSet(ByVal connectionString As String, _
                                                       ByVal spName As String, _
                                                       ByVal includeReturnValueParameter As Boolean, _
                                                       ByVal ParamArray parameterValues() As Object) As SqlParameter()

            Dim cn As New SqlConnection(connectionString)
            Dim cmd As SqlCommand = New SqlCommand(spName, cn)
            Dim discoveredParameters() As SqlParameter

            Try
                cn.Open()
                cmd.CommandType = CommandType.StoredProcedure
                SqlCommandBuilder.DeriveParameters(cmd)
                If Not includeReturnValueParameter Then
                    cmd.Parameters.RemoveAt(0)
                End If

                discoveredParameters = New SqlParameter(cmd.Parameters.Count - 1) {}
                cmd.Parameters.CopyTo(discoveredParameters, 0)
            Finally
                cmd.Dispose()
                cn.Dispose()

            End Try

            Return discoveredParameters

        End Function 'DiscoverSpParameterSet

        'deep copy of cached SqlParameter array
        Private Shared Function CloneParameters(ByVal originalParameters() As SqlParameter) As SqlParameter()

            Dim i As Short
            Dim j As Short = originalParameters.Length - 1
            Dim clonedParameters(j) As SqlParameter

            For i = 0 To j
                clonedParameters(i) = CType(CType(originalParameters(i), ICloneable).Clone, SqlParameter)
            Next

            Return clonedParameters
        End Function 'CloneParameters


        '*********************************************************************
        '
        ' add parameter array to the cache
        '
        ' param name="connectionString" a valid connection string for a SqlConnection 
        ' param name="commandText" the stored procedure name or T-SQL command 
        ' param name="commandParameters" an array of SqlParamters to be cached 
        '
        '*********************************************************************
        Public Shared Sub CacheParameterSet(ByVal connectionString As String, _
                                            ByVal commandText As String, _
                                            ByVal ParamArray commandParameters() As SqlParameter)
            Dim hashKey As String = connectionString + ":" + commandText

            paramCache(hashKey) = commandParameters
        End Sub 'CacheParameterSet


        '*********************************************************************
        '
        ' Retrieve a parameter array from the cache
        ' 
        ' param name="connectionString" a valid connection string for a SqlConnection 
        ' param name="commandText" the stored procedure name or T-SQL command 
        ' returns an array of SqlParamters
        '
        '*********************************************************************
        Public Shared Function GetCachedParameterSet(ByVal connectionString As String, ByVal commandText As String) As SqlParameter()
            Dim hashKey As String = connectionString + ":" + commandText
            Dim cachedParameters As SqlParameter() = CType(paramCache(hashKey), SqlParameter())

            If cachedParameters Is Nothing Then
                Return Nothing
            Else
                Return CloneParameters(cachedParameters)
            End If
        End Function 'GetCachedParameterSet

        '*********************************************************************
        '
        ' Retrieves the set of SqlParameters appropriate for the stored procedure
        ' 
        ' This method will query the database for this information, and then store it in a cache for future requests.
        ' 
        ' param name="connectionString" a valid connection string for a SqlConnection 
        ' param name="spName" the name of the stored procedure 
        ' returns an array of SqlParameters
        '
        '*********************************************************************
        Public Overloads Shared Function GetSpParameterSet(ByVal connectionString As String, ByVal spName As String) As SqlParameter()
            Return GetSpParameterSet(connectionString, spName, False)
        End Function 'GetSpParameterSet 


        '*********************************************************************
        '
        ' Retrieves the set of SqlParameters appropriate for the stored procedure
        ' 
        ' This method will query the database for this information, and then store it in a cache for future requests.
        ' 
        ' param name="connectionString" a valid connection string for a SqlConnection 
        ' param name="spName" the name of the stored procedure 
        ' param name="includeReturnValueParameter" a bool value indicating whether the return value parameter should be included in the results 
        ' returns an array of SqlParameters
        '
        '*********************************************************************
        Public Overloads Shared Function GetSpParameterSet(ByVal connectionString As String, _
                                                           ByVal spName As String, _
                                                           ByVal includeReturnValueParameter As Boolean) As SqlParameter()

            Dim cachedParameters() As SqlParameter
            Dim hashKey As String

            hashKey = connectionString + ":" + spName + IIf(includeReturnValueParameter = True, ":include ReturnValue Parameter", "")

            cachedParameters = CType(paramCache(hashKey), SqlParameter())

            If (cachedParameters Is Nothing) Then
                paramCache(hashKey) = DiscoverSpParameterSet(connectionString, spName, includeReturnValueParameter)
                cachedParameters = CType(paramCache(hashKey), SqlParameter())

            End If

            Return CloneParameters(cachedParameters)

        End Function 'GetSpParameterSet
    End Class 'SqlHelperParameterCache

End Namespace 'ASPNET.StarterKit.TimeTracker.DataAccessLayer