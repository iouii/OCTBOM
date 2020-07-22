Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Threading.Thread
Imports System.Globalization
Module config
    Public conStr As String = "Data Source=" & ConfigurationManager.AppSettings("DBServerOCT") & ";" & _
       "Connect Timeout=30;" & _
       "Initial Catalog=" & ConfigurationManager.AppSettings("DBNameOCT") & ";" & _
       "Persist Security Info=True;User ID=" & ConfigurationManager.AppSettings("DBUserIdOCT") & ";" & _
       "PassWord=" & ConfigurationManager.AppSettings("DBPwdOCT") & "; "

    Public conn As New SqlConnection(conStr)
End Module
