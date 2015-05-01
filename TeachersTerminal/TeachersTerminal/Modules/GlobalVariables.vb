Imports System.Data.Sql
Imports System.Data.SqlClient
Module GlobalVariables
    Public ConString As String = "Data Source =" & System.AppDomain.CurrentDomain.BaseDirectory & "\Database\Database.sdf;File Mode = Shared Read;Persist Security Info = False"
    Public ActiveUser As String = String.Empty
    Public ActiveSLRefNum As String = String.Empty
    Public ActiveSched As Integer = 0
    Public ActiveTab As Integer = 0
    Public ActiveReferral As Integer = 0
End Module

Module Conn
    Public conn As SqlConnection
    Public Function GetConnect()
        conn = New SqlConnection("server=<IP ADDRESS>;uid=sa;pwd=<password>;database=<database name>")
        Return conn
    End Function
End Module
