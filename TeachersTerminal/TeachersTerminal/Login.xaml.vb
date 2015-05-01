Option Explicit On
Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
'Imports System.Xml

Public Class Login
    Dim con As New SqlCeConnection(ConString)
    Dim command As SqlCeCommand = con.CreateCommand
    Private Sub btnLogin_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnLogin.Click
        Dim aVerify As Integer = 0
        Dim StringPassCheck As String = String.Empty
        Dim ds As New DataSet
        Dim da As New SqlCeDataAdapter("SELECT * FROM Faculty WHERE (EmpNum='" & tbEmployeeNum.Text & "')", con)
        Dim cb As New SqlCeCommandBuilder(da)
        If tbEmployeeNum.Text <> "Employee Number" And tbPassword.Password <> "Password" Then
            command.CommandText = "SELECT COALESCE(COUNT(*),0) FROM Faculty WHERE (EmpNum='" & tbEmployeeNum.Text & "' AND Type= 'Professor')"
            con.Open()
            aVerify = command.ExecuteScalar()
            con.Close()
            If aVerify = 1 Then
                con.Open()

                da.Fill(ds, "Faculty")
                command.CommandText = "SELECT Department.Name FROM Faculty INNER JOIN Department ON Faculty.DeptCode=Department.DeptCode WHERE EmpNum='" & tbEmployeeNum.Text & "'"
                Dim affiliation As String = command.ExecuteScalar().ToString()
                con.Close()
                StringPassCheck = ds.Tables("Faculty").Rows(0).Item("Password").ToString()

                If tbPassword.Password.Equals(StringPassCheck) Then
                    Me.Hide()
                    GlobalVariables.ActiveUser = tbEmployeeNum.Text
                    Dim x As New MainWindow
                    x.lblEmpNum.Content = ds.Tables(0).Rows(0).Item("EmpNum").ToString
                    x.lblName.Content = ds.Tables(0).Rows(0).Item("LastName").ToString + ", " + ds.Tables(0).Rows(0).Item("FirstName").ToString + " " + ds.Tables(0).Rows(0).Item("MiddleName")
                    x.lblDesignation.Content = ds.Tables(0).Rows(0).Item("Type").ToString
                    x.lblaffiliation.Content = affiliation

                    Try
                        x.lblEmail.Content = ds.Tables(0).Rows(0).Item("Email").ToString.Replace("_", "__")
                        x.lblContact.Content = ds.Tables(0).Rows(0).Item("Contact").ToString
                        x.imgProfile.Source = New BitmapImage(New Uri(ds.Tables(0).Rows(0).Item("PhotoFilePath").ToString))
                    Catch ex As Exception
                        Console.WriteLine(ex.ToString)
                    End Try
                    x.ShowDialog()
                    Me.Show()
                    With tbPassword
                        .Password = String.Empty
                    End With
                    With tbEmployeeNum
                        .Text = String.Empty
                        .Focus()
                    End With

                Else
                    MessageBox.Show("Password Mismatch.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                    With tbPassword
                        .Password = String.Empty
                        .Focus()
                    End With
                End If
            Else
                MessageBox.Show("Employee Number entered DOES NOT EXIST NOR MEET THE REQUIREMENTS to access this system.", "Error 2", MessageBoxButton.OK, MessageBoxImage.Warning)
                With tbPassword
                    .Password = String.Empty
                End With
                With tbEmployeeNum
                    .Text = String.Empty
                    .Focus()
                End With
            End If
        Else
            tbEmployeeNum.Focus()
        End If
        ds.Dispose()
        da.Dispose()
        cb.Dispose()
    End Sub


    Private Sub tbPassword_KeyDown(sender As Object, e As System.Windows.Input.KeyEventArgs) Handles tbPassword.KeyDown
        If e.Key = Key.Enter Then
            e.Handled = True
            btnLogin_Click(btnLogin, e)
        End If
    End Sub

    Private Sub Login_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        'ReadfromXML()
        'MySettingsChanger.SetConnectionString(cnString)
    End Sub
    'Dim ServerName As String
    'Dim DatabaseName As String
    'Public Shared cnString As String
    'Public Sub ReadfromXML()
    '   If (IO.File.Exists("databaseconfig.xml")) Then
    'Dim document As XmlReader = New XmlTextReader("databaseconfig.xml")
    '       While (document.Read)
    'Dim type = document.NodeType
    '           If (type = XmlNodeType.Element) Then
    '              If (document.Name = "Server") Then
    '                 ServerName = document.ReadInnerXml.ToString()
    '            End If
    '           If (document.Name = "Database") Then
    '              DatabaseName = document.ReadInnerXml.ToString()
    '         End If
    '    End If
    '        End While
    '        MessageBox.Show("Error: databaseconfig.xml file not found!", "Missing databaseconfig.xml", MessageBoxButton.OK, MessageBoxImage.Error)
    '    End If
    '    MsgBox(ServerName)
    '    cnString = "Data Source=" & ServerName & ";Initial Catalog=" & DatabaseName & ";Integrated Security=True"

    'Dim CN As SqlConnection
    '   CN = New SqlConnection

    '    Try
    '      With CN
    '           If .State = ConnectionState.Open Then .Close()

    '        .ConnectionString = cnString
    '        .Open()
    '   End With
    '    Catch ex As Exception
    '       If Err.Number = 5 Then
    '          MsgBox("Cannot connect to server. Make sure that the server is running. " & vbCrLf & vbCrLf & "Otherwise please check for the configuration.", MsgBoxStyle.Exclamation)

    'Dim DBPath As New frmDBPath

    '           DBPath.ShowDialog()
    '          ReadfromXML()
    '       End If
    '   Finally

    '        CN.Close()
    '    End Try
    'End Sub
End Class

'Public Class MySettingsChanger
'Public Shared Sub SetConnectionString(ByVal cnnString As String)
'   My.Settings.RunTimeConnectionString = cnnString
'End Sub
'End Class
