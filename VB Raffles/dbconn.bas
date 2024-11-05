Attribute VB_Name = "dbconn"
Option Explicit
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Function DBstr() As String
    On Error GoTo ErrorHandler
    DBstr = "Driver={MySQL ODBC 8.0 Unicode Driver};" _
            & "Server=127.0.0.1;" _
            & "Port=3306;" _
            & "Database=rafflesys;" _
            & "User=root;" _
            & "Password=usbw;" _
            & "Option=0;"
    Exit Function ' Add this line to exit the function if there is no error
    
ErrorHandler:
    MsgBox "Database Connection Error: " & Err.Description, vbCritical, "Error"
    End ' Terminate the program in case of an error
End Function


