Imports System.Data.SqlServerCe
Public Class dialogFacultyConcerns
    Dim con As New SqlCeConnection(ConString)
    Dim command As SqlCeCommand = con.CreateCommand()
    Private Sub btnCancel_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnSubmit_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnSubmit.Click
        MainWindow.timer.Stop()
        con.Close()
            con.Open()
            command.CommandText = "UPDATE Referral SET Concerns = '" & tbConcerns.Text & "' WHERE TraceNo=" & ActiveReferral
            command.ExecuteNonQuery()
            MessageBox.Show("Successfully Added Faculty Concerns to Referral " & ActiveReferral, "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        con.Close()
        Me.Close()
        MainWindow.timer.Start()
    End Sub

    Private Sub tbConcerns_TextChanged(sender As Object, e As System.Windows.Controls.TextChangedEventArgs) Handles tbConcerns.TextChanged
        tbConcerns.Text = tbConcerns.Text.Replace("'", "")
        If tbConcerns.Text.Length > 200 Then
            tbConcerns.Text = tbConcerns.Text.Remove(200)
        End If
        lblNumChars.Content = Val(200 - tbConcerns.Text.Length).ToString
        If Val(lblNumChars.Content) < 20 Then
            lblNumChars.Foreground = System.Windows.Media.Brushes.Red
        Else
            lblNumChars.Foreground = System.Windows.Media.Brushes.Black
        End If
        tbConcerns.Select(tbConcerns.Text.Length, 0)
    End Sub
End Class
