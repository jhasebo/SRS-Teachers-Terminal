Option Explicit On
Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
Imports System.Windows.Media.Animation
Imports System.Windows.Threading
Imports System.Windows.Controls.Primitives

Public Class MainWindow
    Dim con As New SqlCeConnection(ConString)
    Dim command As SqlCeCommand = con.CreateCommand
    Public Shared timer As DispatcherTimer = New DispatcherTimer

    'Onload Events
    Private Sub MainWindow_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        'For Classes
        Dim daMW As New SqlCeDataAdapter("Select Section.Time as [TIME], Section.SecCode as [SECTION], Section.SubjCode as [SUBJECT], Subject.Description as [Description], Section.BLDG as [BLDG], Section.Room as [Room] FROM Section Inner join Subject on Section.SubjCode = Subject.SubjCode WHERE Section.EmpNum='" & ActiveUser & "' AND Section.Days=1 Order by Section.Time Asc", con)
        Dim daTTH As New SqlCeDataAdapter("Select Section.Time as [TIME], Section.SecCode as [SECTION], Section.SubjCode as [SUBJECT], Subject.Description as [Description], Section.BLDG as [BLDG], Section.Room as [Room] FROM Section Inner join Subject on Section.SubjCode = Subject.SubjCode WHERE Section.EmpNum='" & ActiveUser & "' AND Section.Days=2 Order by Section.Time Asc", con)
        Dim daFS As New SqlCeDataAdapter("Select Section.Time as [TIME], Section.SecCode as [SECTION], Section.SubjCode as [SUBJECT], Subject.Description as [Description], Section.BLDG as [BLDG], Section.Room as [Room] FROM Section Inner join Subject on Section.SubjCode = Subject.SubjCode WHERE Section.EmpNum='" & ActiveUser & "' AND Section.Days=3 Order by Section.Time Asc", con)
        Dim cbMW As New SqlCeCommandBuilder(daMW)
        Dim cbTTH As New SqlCeCommandBuilder(daTTH)
        Dim cbFS As New SqlCeCommandBuilder(daFS)
        Try
            Dim dsMW, dsTTH, dsFS As New DataSet
            con.Close()
            con.Open()
            daMW.Fill(dsMW)
            daTTH.Fill(dsTTH)
            daFS.Fill(dsFS)
            dgMonWed.DataContext = dsMW.Tables(0)
            dgMonWed.ItemsSource = dsMW.Tables(0).DefaultView
            dgTueThu.DataContext = dsTTH.Tables(0)
            dgTueThu.ItemsSource = dsTTH.Tables(0).DefaultView
            dgFriSat.DataContext = dsFS.Tables(0)
            dgFriSat.ItemsSource = dsFS.Tables(0).DefaultView
        Catch ex As Exception
            MessageBox.Show("Unable to populate classes: " & ex.ToString)
        Finally
            con.Close()
        End Try
        If DateAndTime.Now.DayOfWeek = DayOfWeek.Monday Or DateAndTime.Now.DayOfWeek = DayOfWeek.Wednesday Then
            expMW.IsExpanded = vbTrue
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Tuesday Or DateAndTime.Now.DayOfWeek = DayOfWeek.Thursday Then
            expTTH.IsExpanded = True
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Friday Or DateAndTime.Now.DayOfWeek = DayOfWeek.Saturday Then
            expFS.IsExpanded = True
        End If

        timer.Interval = TimeSpan.FromSeconds(5)
        AddHandler timer.Tick, AddressOf timer_tick
        timer.Start()

        'For Reports
        fillrecents()
        'For Notifications

    End Sub

    Private Sub timer_tick()
        command.CommandText = "Select Count(*) FROM FeedbackRequest INNER JOIN Referral ON FeedbackRequest.TraceNo=Referral.TraceNo INNER JOIN StudentList ON Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Section ON StudentList.SecCode=Section.SecCode WHERE (Section.EmpNum='" & ActiveUser & "' AND Finished is null)"
        con.Close()
        con.Open()
        Dim ctr As Integer = command.ExecuteScalar()
        con.Close()
        If ctr = 0 Then
            cNotify.Visibility = Windows.Visibility.Hidden
            dgNotifications.Visibility = Windows.Visibility.Hidden
            tbNoNewNotif.Visibility = Windows.Visibility.Visible
        Else
            cNotify.Visibility = Windows.Visibility.Visible
            tbDisplayNotifCtr.Text = ctr
            dgNotifications.Visibility = Windows.Visibility.Visible
            tbNoNewNotif.Visibility = Windows.Visibility.Hidden
            fillnotifications()
        End If


    End Sub

    Private Sub fillnotifications()
        timer.Stop()
        Dim adapter As New SqlCeDataAdapter("SELECT FeedbackRequest.RequestNo AS [No], FeedbackRequest.TraceNo as [Ref No], StudentList.SecCode as [SECTION], StudentList.StudentNo as [Student Number], Student.LastName as [Last Name], Student.FirstName as [First Name], Student.MiddleName as [Middle Name] FROM FeedbackRequest INNER JOIN Referral ON FeedbackRequest.TraceNo=Referral.TraceNo INNER JOIN StudentList ON Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section ON StudentList.SecCode=Section.SecCode WHERE (FeedbackRequest.Finished is null AND Section.EmpNum='" & ActiveUser & "') ORDER BY FeedbackRequest.RequestNo ASC", con)
        Dim cbuilder As New SqlCeCommandBuilder(adapter)
        Try
            Dim dataset As New DataSet
            con.Close()
            con.Open()
            adapter.Fill(dataset)
            dgNotifications.ItemsSource = dataset.Tables(0).DefaultView
            dgNotifications.DataContext = dataset.Tables(0)
        Catch ex As Exception
        Finally
            con.Close()
        End Try
        timer.Start()
    End Sub

    Private Sub fillrecents()
        timer.Stop()
        Dim daRA As New SqlCeDataAdapter("SELECT Attendance.Date as [Date], Attendance.AttendanceRefNum as [ID], StudentList.SecCode as [Section], StudentList.StudentNo as [Student Number], Student.LastName as [Last Name], Student.FirstName as [First Name], Student.MiddleName as [Middle Name] FROM Attendance INNER JOIN StudentList ON Attendance.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section on StudentList.SecCode=Section.SecCode WHERE (Attendance.Absent=1 AND Attendance.ActionTaken is null AND Section.EmpNum='" & ActiveUser & "') ORDER BY Attendance.Date Desc", con)
        Dim daRR As New SqlCeDataAdapter("SELECT TOP(30) Referral.Date as [Date], Referral.TraceNo as [Ref No], StudentList.SecCode as [Section], StudentList.StudentNo as [Student Number], Student.LastName as [Last Name], Student.FirstName as [First Name], Student.MiddleName as [Middle Name] FROM Referral INNER JOIN StudentList ON Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section on StudentList.SecCode=Section.SecCode WHERE (Section.EmpNum='" & ActiveUser & "') ORDER BY Referral.Date DESC", con)
        Dim cbRA As New SqlCeCommandBuilder(daRA)
        Dim cbRR As New SqlCeCommandBuilder(daRR)
        Try
            Dim dsRA, dsRR As New DataSet
            con.Close()
            con.Open()
            daRA.Fill(dsRA)
            daRR.Fill(dsRR)
            dsRA.Tables(0).Columns.Add(New DataColumn("Excused", GetType(Boolean)))
            dsRA.Tables(0).Columns("Excused").SetOrdinal(0)
            Dim x As New DataSet
            Dim y As New SqlCeDataAdapter("Select Coalesce(Attendance.Excused,0) as [E],Attendance.Date as [Date], Attendance.AttendanceRefNum as [ID], StudentList.SecCode as [Section], StudentList.StudentNo as [Student Number], Student.LastName as [Last Name], Student.FirstName as [First Name], Student.MiddleName as [Middle Name] FROM Attendance INNER JOIN StudentList ON Attendance.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section on StudentList.SecCode=Section.SecCode WHERE (Attendance.Absent=1 AND Attendance.ActionTaken is null AND Section.EmpNum='" & ActiveUser & "') ORDER BY Attendance.Date Desc", con)
            Dim z As New SqlCeCommandBuilder(y)
            y.Fill(x)
            Dim i = 0
            While i < dsRA.Tables(0).Rows.Count
                If x.Tables(0).Rows(i).Item("E") = 1 Then
                    dsRA.Tables(0).Rows(i).Item("Excused") = True
                Else
                    dsRA.Tables(0).Rows(i).Item("Excused") = False
                End If

                i = i + 1
            End While
            dgActions.ItemsSource = dsRA.Tables(0).DefaultView
            dgActions.DataContext = dsRA.Tables(0)
            dgReferrals.ItemsSource = dsRR.Tables(0).DefaultView
            dgReferrals.DataContext = dsRR.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Unable to populate Recents...", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            con.Close()
        End Try
        timer.Start()
    End Sub

    Private Sub fillStudentList(adapter As SqlCeDataAdapter)
        con.Close()
        con.Open()
        Try
            Dim dsStudents As New DataSet
            con.Close()
            con.Open()
            adapter.Fill(dsStudents)
            dsStudents.Tables(0).Columns.Add(New DataColumn("Absent", GetType(Boolean)))
            dsStudents.Tables(0).Columns("Absent").SetOrdinal(0)
            dsStudents.Tables(0).Columns.Add(New DataColumn("Late", GetType(Boolean)))
            dsStudents.Tables(0).Columns("Late").SetOrdinal(1)
            Dim i As Integer = 0
            While i < dsStudents.Tables(0).Rows.Count
                command.CommandText = "SELECT SLRefNum from StudentList where SecCode='" & lblSubjSec.Content & "' and StudentNo='" & dsStudents.Tables(0).Rows(i).Item("STUDENT NUMBER").ToString() & "'"
                Try
                    con.Close()
                    con.Open()
                    ActiveSLRefNum = command.ExecuteScalar
                    command.CommandText = "SELECT Coalesce(Count(*),0) FROM Attendance WHERE (SLRefNum='" & ActiveSLRefNum & "' and Date={d '" & dpRefDate.SelectedDate.Value.ToString("yyyy-MM-dd") & "'} and Absent=1)"
                    Dim count As Integer = command.ExecuteScalar
                    If count > 0 Then
                        dsStudents.Tables(0).Rows(i)("Absent") = True
                    Else
                        dsStudents.Tables(0).Rows(i)("Absent") = False
                    End If
                    command.CommandText = "SELECT Coalesce(Count(*),0) FROM Attendance WHERE (SLRefNum='" & ActiveSLRefNum & "' and Date={d '" & dpRefDate.SelectedDate.Value.ToString("yyyy-MM-dd") & "'} and Late=1)"
                    count = command.ExecuteScalar
                    If count > 0 Then
                        dsStudents.Tables(0).Rows(i)("Late") = True
                    Else
                        dsStudents.Tables(0).Rows(i)("Late") = False
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                Finally
                    con.Close()
                End Try

                i = i + 1
            End While
            dgStudentList.ItemsSource = dsStudents.Tables(0).DefaultView
            dgStudentList.DataContext = dsStudents.Tables(0)
            dsStudents.Dispose()
        Catch ex As Exception
            MessageBox.Show("Unable to populate the student list: " & ex.ToString)
        Finally
            con.Close()
        End Try
    End Sub

    'Transition Events
    Private Sub btnUser_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnUser.Click
        ActiveTab = 1
    End Sub

    Private Sub btnClass_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnClass.Click
        ActiveTab = 2
    End Sub

    Private Sub btnNotifier_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnNotifier.Click
        ActiveTab = 4
    End Sub

    Private Sub btnReports_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnReports.Click
        ActiveTab = 3
    End Sub

    Private Sub gridMenu_LeftButtonUp(sender As System.Object, e As System.Windows.Input.MouseButtonEventArgs) Handles gridMenu.PreviewMouseLeftButtonUp
        Dim daW As New DoubleAnimation
        With daW
            .From = 800
            .To = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With
        Dim daFW As New DoubleAnimation
        With daFW
            .From = 700
            .To = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With
        Dim daOP As New DoubleAnimation
        With daOP
            .From = 0.6
            .To = 1
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With

        If ActiveTab = 1 Then
            gridUser.BeginAnimation(Grid.WidthProperty, daW)
            gridMenu.BeginAnimation(Grid.OpacityProperty, daOP)
        ElseIf ActiveTab = 2 Then
            gridClasses.BeginAnimation(Grid.WidthProperty, daW)
            gridMenu.BeginAnimation(Grid.OpacityProperty, daOP)
        ElseIf ActiveTab = 3 Then
            gridReport.BeginAnimation(Grid.WidthProperty, daW)
            gridMenu.BeginAnimation(Grid.OpacityProperty, daOP)
        ElseIf ActiveTab = 4 Then
            gridNotify.BeginAnimation(Grid.WidthProperty, daW)
            gridMenu.BeginAnimation(Grid.OpacityProperty, daOP)
        ElseIf ActiveTab = 5 Then
            gridClasses.BeginAnimation(Grid.WidthProperty, daW)
            gridStudentList.BeginAnimation(Grid.WidthProperty, daFW)
            gridMenu.BeginAnimation(Grid.OpacityProperty, daOP)
        ElseIf ActiveTab = 6 Then
            gridNotify.BeginAnimation(Grid.WidthProperty, daW)
            gridNotificationDetails.BeginAnimation(Grid.WidthProperty, daFW)
            gridMenu.BeginAnimation(Grid.OpacityProperty, daOP)
        End If
        ActiveTab = 0
    End Sub

    Private Sub Clickable_LeftButtonUp(sender As System.Object, e As System.Windows.Input.MouseButtonEventArgs) Handles Clickable.MouseLeftButtonUp
        If ActiveTab = 5 Then
            Dim daFW As New DoubleAnimation
            With daFW
                .From = 700
                .To = 0
                .Duration = New Duration(TimeSpan.FromSeconds(0.1))
            End With
            Dim da As New DoubleAnimation
            With da
                .To = 0
                .From = 70
                .Duration = New Duration(TimeSpan.FromSeconds(0))
            End With
            Clickable.BeginAnimation(Grid.WidthProperty, da)
            gridStudentList.BeginAnimation(Grid.WidthProperty, daFW)
            ActiveTab = 2
        ElseIf ActiveTab = 6 Then
            Dim dafw As New DoubleAnimation
            With dafw
                .From = 700
                .To = 0
                .Duration = New Duration(TimeSpan.FromSeconds(0.1))
            End With
            Dim da As New DoubleAnimation
            With da
                .To = 0
                .From = 70
                .Duration = New Duration(TimeSpan.FromSeconds(0))
            End With
            Clickable.BeginAnimation(Grid.WidthProperty, da)
            gridNotificationDetails.BeginAnimation(Grid.WidthProperty, dafw)
            ActiveTab = 4
            timer.Start()
        End If
    End Sub

    Private Sub setClickable()
        Dim da As New DoubleAnimation
        With da
            .To = 100
            .From = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0))
        End With
        Clickable.BeginAnimation(Grid.WidthProperty, da)
    End Sub

    'Code for Profile Tab
    Private Sub btnUpdateContact_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnUpdateContact.Click
        Dim daUC As DoubleAnimation = New DoubleAnimation()
        With daUC
            .From = 0
            .To = 120
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        tbUpdateContact.BeginAnimation(TextBox.WidthProperty, daUC)
        setContact.Width = 50
        CancelC.Width = 50
        btnUpdateContact.Width = 0
        tbUpdateContact.Focus()
    End Sub

    Private Sub setContact_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles setContact.Click
        timer.Stop()
        Try
            con.Close()
            con.Open()
            With command
                .CommandText = "Update Faculty Set Contact = '" & tbUpdateContact.Text & "' Where EmpNum='" & ActiveUser & "'"
                .ExecuteNonQuery()
            End With
            lblContact.Content = tbUpdateContact.Text()

        Catch ex As Exception
            MessageBox.Show("Unable to Update Contact: " & ex.ToString)
        Finally
            con.Close()
            tbUpdateContact.Text = String.Empty
            Dim daUC As DoubleAnimation = New DoubleAnimation()
            With daUC
                .From = 120
                .To = 0
                .Duration = New Duration(TimeSpan.FromSeconds(0.1))
            End With
            tbUpdateContact.BeginAnimation(TextBox.WidthProperty, daUC)
            btnUpdateContact.Width = 62
            setContact.Width = 0
            CancelC.Width = 0
            btnUpdateContact.Focus()
        End Try
        timer.Start()
    End Sub

    Private Sub CancelC_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles CancelC.Click
        Dim daUC As DoubleAnimation = New DoubleAnimation()
        With daUC
            .From = 120
            .To = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With
        tbUpdateContact.BeginAnimation(TextBox.WidthProperty, daUC)
        btnUpdateContact.Width = 62
        setContact.Width = 0
        CancelC.Width = 0
        btnUpdateContact.Focus()
    End Sub

    Private Sub btnUpdateEmail_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnUpdateEmail.Click
        Dim daUC As DoubleAnimation = New DoubleAnimation()
        With daUC
            .From = 0
            .To = 120
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        tbUpdateEmail.BeginAnimation(TextBox.WidthProperty, daUC)
        setEmail.Width = 50
        CancelE.Width = 50
        btnUpdateEmail.Width = 0
        tbUpdateEmail.Focus()
    End Sub

    Private Sub setEmail_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles setEmail.Click
        timer.Stop()
        Try
            con.Close()
            con.Open()
            With command
                .CommandText = "Update Faculty Set Email = '" & tbUpdateEmail.Text & "' Where EmpNum='" & ActiveUser & "'"
                .ExecuteNonQuery()
            End With
            lblEmail.Content = tbUpdateEmail.Text.Replace("_", "__")

        Catch ex As Exception
            MessageBox.Show("Unable to Update Email: " & ex.ToString)
        Finally
            con.Close()
            tbUpdateEmail.Text = String.Empty
            Dim daUC As DoubleAnimation = New DoubleAnimation()
            With daUC
                .From = 120
                .To = 0
                .Duration = New Duration(TimeSpan.FromSeconds(0.1))
            End With
            tbUpdateEmail.BeginAnimation(TextBox.WidthProperty, daUC)
            btnUpdateEmail.Width = 62
            setEmail.Width = 0
            CancelE.Width = 0
            btnUpdateEmail.Focus()
        End Try
        timer.Start()
    End Sub

    Private Sub CancelE_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles CancelE.Click
        Dim daUC As DoubleAnimation = New DoubleAnimation()
        With daUC
            .From = 120
            .To = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With
        tbUpdateEmail.BeginAnimation(TextBox.WidthProperty, daUC)
        btnUpdateEmail.Width = 62
        setEmail.Width = 0
        CancelE.Width = 0
        btnUpdateEmail.Focus()
    End Sub

    Private Sub btnPicture_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnPicture.Click
        timer.Stop()
        Dim openpic As New Microsoft.Win32.OpenFileDialog()
        openpic.FileName = "Profile Photo"
        openpic.DefaultExt = ".jpg"
        openpic.Filter = "Image Files(*.jpg,*.png,*.bmp)|*.jpg;*.png;*.bmp|All Files (*.*)|*.*"
        Dim result As Boolean = openpic.ShowDialog()
        If result = True Then
            Try
                con.Close()
                con.Open()
                With command
                    .CommandText = "Update Faculty Set PhotoFilePath = '" & openpic.FileName & "' Where EmpNum = '" & ActiveUser & "'"
                    .ExecuteNonQuery()
                End With
                imgProfile.Source = New BitmapImage(New Uri(openpic.FileName))
            Catch ex As Exception
                MessageBox.Show("Unable to Update Profile Image: " & ex.ToString)
            Finally
                con.Close()
            End Try
        End If
        timer.Start()
    End Sub

    'Code for Classes Tab
    Private Sub dgMonWed_MouseDoubleClick(sender As Object, e As System.Windows.Input.MouseButtonEventArgs) Handles dgMonWed.MouseDoubleClick
        timer.Stop()
        ActiveTab = 5
        ActiveSched = 1
        Dim row As DataRowView = dgMonWed.SelectedItems(0)
        lblSubjSec.Content = row("SECTION")
        lblDescription.Content = row("Description")
        Dim daSL As New SqlCeDataAdapter("Select Student.LastName as [LAST NAME], Student.FirstName as [FIRST NAME], Student.MiddleName as [MIDDLE NAME], StudentList.StudentNo as [STUDENT NUMBER] FROM StudentList INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo WHERE SecCode='" & row("SECTION") & "' ORDER BY Student.LastName Asc", con)
        Dim cbSL As New SqlCeCommandBuilder(daSL)
        dpRefDate.SelectedDate = DateAndTime.Now.Date
        If DateAndTime.Now.DayOfWeek = DayOfWeek.Tuesday Or DateAndTime.Now.DayOfWeek = DayOfWeek.Thursday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -1, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Friday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -2, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Saturday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -3, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Sunday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -4, DateAndTime.Now)
        End If
        fillStudentList(daSL)
        setClickable()
        Dim da As New DoubleAnimation
        With da
            .To = 700
            .From = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        gridStudentList.BeginAnimation(Grid.WidthProperty, da)
        timer.Start()
    End Sub

    Private Sub dgMonWed_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles dgMonWed.SelectionChanged
        timer.Stop()
        ActiveTab = 5
        ActiveSched = 1
        Dim row As DataRowView = dgMonWed.SelectedItems(0)
        lblSubjSec.Content = row("SECTION")
        lblDescription.Content = row("Description")
        Dim daSL As New SqlCeDataAdapter("Select Student.LastName as [LAST NAME], Student.FirstName as [FIRST NAME], Student.MiddleName as [MIDDLE NAME], StudentList.StudentNo as [STUDENT NUMBER] FROM StudentList INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo WHERE SecCode='" & row("SECTION") & "' ORDER BY Student.LastName Asc", con)
        Dim cbSL As New SqlCeCommandBuilder(daSL)
        dpRefDate.SelectedDate = DateAndTime.Now.Date
        If DateAndTime.Now.DayOfWeek = DayOfWeek.Tuesday Or DateAndTime.Now.DayOfWeek = DayOfWeek.Thursday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -1, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Friday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -2, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Saturday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -3, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Sunday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -4, DateAndTime.Now)
        End If
        fillStudentList(daSL)
        setClickable()
        Dim da As New DoubleAnimation
        With da
            .To = 700
            .From = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        gridStudentList.BeginAnimation(Grid.WidthProperty, da)
        timer.Start()
    End Sub

    Private Sub dgTueThu_MouseDoubleClick(sender As Object, e As System.Windows.Input.MouseButtonEventArgs) Handles dgTueThu.MouseDoubleClick
        timer.Stop()
        ActiveTab = 5
        ActiveSched = 2
        Dim row As DataRowView = dgTueThu.SelectedItems(0)
        lblSubjSec.Content = row("SECTION")
        lblDescription.Content = row("Description")
        Dim daSL As New SqlCeDataAdapter("Select Student.LastName as [LAST NAME], Student.FirstName as [FIRST NAME], Student.MiddleName as [MIDDLE NAME], StudentList.StudentNo as [STUDENT NUMBER] FROM StudentList INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo WHERE SecCode='" & row("SECTION") & "' ORDER BY Student.LastName Asc", con)
        Dim cbSL As New SqlCeCommandBuilder(daSL)
        dpRefDate.SelectedDate = DateAndTime.Now.Date
        If DateAndTime.Now.DayOfWeek = DayOfWeek.Monday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -4, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Wednesday Or DateAndTime.Now.DayOfWeek = DayOfWeek.Friday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -1, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Saturday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -2, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Sunday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -3, DateAndTime.Now)
        End If
        fillStudentList(daSL)
        setClickable()
        Dim da As New DoubleAnimation
        With da
            .To = 700
            .From = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        gridStudentList.BeginAnimation(Grid.WidthProperty, da)
        timer.Start()
    End Sub

    Private Sub dgTueThu_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles dgTueThu.SelectionChanged
        timer.Stop()
        ActiveTab = 5
        ActiveSched = 2
        Dim row As DataRowView = dgTueThu.SelectedItems(0)
        lblSubjSec.Content = row("SECTION")
        lblDescription.Content = row("Description")
        Dim daSL As New SqlCeDataAdapter("Select Student.LastName as [LAST NAME], Student.FirstName as [FIRST NAME], Student.MiddleName as [MIDDLE NAME], StudentList.StudentNo as [STUDENT NUMBER] FROM StudentList INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo WHERE SecCode='" & row("SECTION") & "' ORDER BY Student.LastName Asc", con)
        Dim cbSL As New SqlCeCommandBuilder(daSL)
        dpRefDate.SelectedDate = DateAndTime.Now.Date
        If DateAndTime.Now.DayOfWeek = DayOfWeek.Monday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -4, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Wednesday Or DateAndTime.Now.DayOfWeek = DayOfWeek.Friday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -1, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Saturday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -2, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Sunday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -3, DateAndTime.Now)
        End If
        fillStudentList(daSL)
        setClickable()
        Dim da As New DoubleAnimation
        With da
            .To = 700
            .From = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        gridStudentList.BeginAnimation(Grid.WidthProperty, da)
        timer.Start()
    End Sub

    Private Sub dgFriSat_MouseDoubleClick(sender As Object, e As System.Windows.Input.MouseButtonEventArgs) Handles dgFriSat.MouseDoubleClick
        timer.Stop()
        ActiveTab = 5
        ActiveSched = 3
        Dim row As DataRowView = dgFriSat.SelectedItems(0)
        lblSubjSec.Content = row("SECTION")
        lblDescription.Content = row("Description")
        Dim daSL As New SqlCeDataAdapter("Select Student.LastName as [LAST NAME], Student.FirstName as [FIRST NAME], Student.MiddleName as [MIDDLE NAME], StudentList.StudentNo as [STUDENT NUMBER] FROM StudentList INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo WHERE SecCode='" & row("SECTION") & "' ORDER BY Student.LastName Asc", con)
        Dim cbSL As New SqlCeCommandBuilder(daSL)
        dpRefDate.SelectedDate = DateAndTime.Now.Date
        If DateAndTime.Now.DayOfWeek = DayOfWeek.Monday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -2, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Tuesday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -3, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Wednesday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -4, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Thursday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -5, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Sunday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -1, DateAndTime.Now)
        End If
        fillStudentList(daSL)
        setClickable()
        Dim da As New DoubleAnimation
        With da
            .To = 700
            .From = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        gridStudentList.BeginAnimation(Grid.WidthProperty, da)
        timer.Start()
    End Sub

    Private Sub dgFriSat_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles dgFriSat.SelectionChanged
        timer.Stop()
        ActiveTab = 5
        ActiveSched = 3
        Dim row As DataRowView = dgFriSat.SelectedItems(0)
        lblSubjSec.Content = row("SECTION")
        lblDescription.Content = row("Description")
        Dim daSL As New SqlCeDataAdapter("Select Student.LastName as [LAST NAME], Student.FirstName as [FIRST NAME], Student.MiddleName as [MIDDLE NAME], StudentList.StudentNo as [STUDENT NUMBER] FROM StudentList INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo WHERE SecCode='" & row("SECTION") & "' ORDER BY Student.LastName Asc", con)
        Dim cbSL As New SqlCeCommandBuilder(daSL)
        dpRefDate.SelectedDate = DateAndTime.Now.Date
        If DateAndTime.Now.DayOfWeek = DayOfWeek.Monday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -2, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Tuesday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -3, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Wednesday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -4, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Thursday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -5, DateAndTime.Now)
        ElseIf DateAndTime.Now.DayOfWeek = DayOfWeek.Sunday Then
            dpRefDate.SelectedDate = DateAdd(DateInterval.Day, -1, DateAndTime.Now)
        End If
        setClickable()
        fillStudentList(daSL)
        Dim da As New DoubleAnimation
        With da
            .To = 700
            .From = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        gridStudentList.BeginAnimation(Grid.WidthProperty, da)
        timer.Start()
    End Sub

    Private Sub dgStudentList_CellEditEnding(sender As Object, e As System.Windows.Controls.DataGridCellEditEndingEventArgs) Handles dgStudentList.CellEditEnding
        timer.Stop()
        command.CommandText = "SELECT SLRefNum from StudentList where SecCode='" & lblSubjSec.Content & "' and StudentNo='" & e.Row.Item("STUDENT NUMBER") & "'"
        Try
            con.Close()
            con.Open()
            ActiveSLRefNum = command.ExecuteScalar()
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        Finally
            con.Close()
        End Try
        If e.Column.DisplayIndex = 0 Then
            Dim c As CheckBox = e.Column.GetCellContent(e.Row)
            con.Close()
            con.Open()
            command.CommandText = "SELECT Coalesce(Count(*),0) FROM Attendance WHERE (SLRefNum='" & ActiveSLRefNum & "' and Date={d '" & dpRefDate.SelectedDate.Value.ToString("yyyy-MM-dd") & "'})"
            Dim ctr As Integer = command.ExecuteScalar()

            If c.IsChecked = True Then
                If ctr > 0 Then
                    command.CommandText = "UPDATE Attendance SET Absent=1 WHERE (SLRefNum='" & ActiveSLRefNum & "' and Date={d '" & dpRefDate.SelectedDate.Value.ToString("yyyy-MM-dd") & "'})"
                    command.ExecuteNonQuery()
                Else
                    command.CommandText = "INSERT INTO Attendance(SLRefNum, Date, Absent, Excused, Late) VALUES('" & ActiveSLRefNum & "',{d '" & dpRefDate.SelectedDate.Value.ToString("yyyy-MM-dd") & "'} ,1,null,null)"
                    command.ExecuteNonQuery()
                End If
            Else
                command.CommandText = "UPDATE Attendance SET Absent=0 WHERE (SLRefNum='" & ActiveSLRefNum & "' and Date={d '" & dpRefDate.SelectedDate.Value.ToString("yyyy-MM-dd") & "'})"
                command.ExecuteNonQuery()
            End If
            con.Close()
        ElseIf e.Column.DisplayIndex = 1 Then
            Dim c As CheckBox = e.Column.GetCellContent(e.Row)
            con.Close()
            con.Open()
            command.CommandText = "SELECT Coalesce(Count(*),0) FROM Attendance WHERE (SLRefNum='" & ActiveSLRefNum & "' and Date={d '" & dpRefDate.SelectedDate.Value.ToString("yyyy-MM-dd") & "'})"
            Dim ctr As Integer = command.ExecuteScalar()

            If c.IsChecked = True Then
                If ctr > 0 Then
                    command.CommandText = "UPDATE Attendance SET Late=1 WHERE (SLRefNum='" & ActiveSLRefNum & "' and Date={d '" & dpRefDate.SelectedDate.Value.ToString("yyyy-MM-dd") & "'})"
                    command.ExecuteNonQuery()
                Else
                    command.CommandText = "INSERT INTO Attendance(SLRefNum, Date, Absent, Excused, Late) VALUES('" & ActiveSLRefNum & "',{d '" & dpRefDate.SelectedDate.Value.ToString("yyyy-MM-dd") & "'} ,null,null,1)"
                    command.ExecuteNonQuery()
                End If
            Else
                command.CommandText = "UPDATE Attendance SET Late=0 WHERE (SLRefNum='" & ActiveSLRefNum & "' and Date={d '" & dpRefDate.SelectedDate.Value.ToString("yyyy-MM-dd") & "'})"
                command.ExecuteNonQuery()
            End If
        End If
        ReferralReporter(ActiveSLRefNum)
        fillrecents()
        timer.Start()
    End Sub

    Private Sub dpRefDate_SelectedDateChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles dpRefDate.SelectedDateChanged
        timer.Stop()
        Dim daSL As New SqlCeDataAdapter("Select Student.LastName as [LAST NAME], Student.FirstName as [FIRST NAME], Student.MiddleName as [MIDDLE NAME], StudentList.StudentNo as [STUDENT NUMBER] FROM StudentList INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo WHERE SecCode='" & lblSubjSec.Content & "' ORDER BY Student.LastName Asc", con)
        Dim cbSL As New SqlCeCommandBuilder(daSL)
        fillStudentList(daSL)
        If ActiveSched = 1 Then
            If dpRefDate.SelectedDate.Value.DayOfWeek = DayOfWeek.Monday Or dpRefDate.SelectedDate.Value.DayOfWeek = DayOfWeek.Wednesday Then
                dgStudentList.IsEnabled = True
            Else
                dgStudentList.IsEnabled = False
            End If
        ElseIf ActiveSched = 2 Then
            If dpRefDate.SelectedDate.Value.DayOfWeek = DayOfWeek.Tuesday Or dpRefDate.SelectedDate.Value.DayOfWeek = DayOfWeek.Thursday Then
                dgStudentList.IsEnabled = True
            Else
                dgStudentList.IsEnabled = False
            End If
        ElseIf ActiveSched = 3 Then
            If dpRefDate.SelectedDate.Value.DayOfWeek = DayOfWeek.Friday Or dpRefDate.SelectedDate.Value.DayOfWeek = DayOfWeek.Saturday Then
                dgStudentList.IsEnabled = True
            Else
                dgStudentList.IsEnabled = False
            End If
        End If
        timer.Start()
    End Sub

    Private Sub ReferralReporter(SLRefNum As Integer)
        timer.Stop()
        Dim totalAbsences As Integer = 0

        command.CommandText = "Select Coalesce(Count(*),0) FROM Attendance WHERE ((Late=1 AND SLRefNum=" & SLRefNum & ") AND ActionTaken is Null)"
        con.Close()
        con.Open()
        Dim lates As Integer = command.ExecuteScalar
        totalAbsences = totalAbsences + (lates / 3)
        command.CommandText = "Select Coalesce(Count(*),0) FROM Attendance WHERE ((Absent=1 AND SLRefNum=" & SLRefNum & ") AND ((Excused is Null) AND ActionTaken is Null))"
        Dim absences As Integer = command.ExecuteScalar
        totalAbsences = totalAbsences + absences

        con.Close()

        If totalAbsences >= 2 Then
            con.Close()
            con.Open()
            command.CommandText = "INSERT INTO Referral(Type, SLRefNum, Concerns, Feedback, ActionTaken, ATby, Date) Values(1," & SLRefNum & ",null,null,null,null,{d '" & DateAndTime.Now.ToString("yyyy-MM-dd") & "'})"
            command.ExecuteNonQuery()
            command.CommandText = "SELECT MAX(TraceNo) FROM Referral"
            Dim trace As Integer = command.ExecuteScalar()
            Dim dsL, dsA As New DataSet
            Dim daL As New SqlCeDataAdapter("SELECT AttendanceRefNum as [REFERENCE] FROM Attendance WHERE ((Late=1 AND SLRefNum=" & SLRefNum & ") AND ActionTaken is null)", con)
            Dim daA As New SqlCeDataAdapter("SELECT AttendanceRefNum as [REFERENCE] FROM Attendance WHERE ((Absent=1 AND SLRefNum=" & SLRefNum & ") AND ActionTaken is null)", con)
            Dim cbL As New SqlCeCommandBuilder(daL)
            Dim cbA As New SqlCeCommandBuilder(daA)
            daL.Fill(dsL)
            daA.Fill(dsA)
            daL.Dispose()
            daA.Dispose()
            cbL.Dispose()
            cbA.Dispose()
            Dim i As Integer = 0
            While i < dsL.Tables(0).Rows.Count
                command.CommandText = "INSERT INTO ReferralDates(TraceNo, AttendanceRefNum) VALUES(" & trace & "," & dsL.Tables(0).Rows(i).Item(0).ToString() & ")"
                command.ExecuteNonQuery()
                command.CommandText = "UPDATE Attendance SET ActionTaken=1 WHERE AttendanceRefNum=" & dsL.Tables(0).Rows(i).Item(0).ToString()
                command.ExecuteNonQuery()
                i = i + 1
            End While
            dsL.Dispose()
            i = 0
            While i < dsA.Tables(0).Rows.Count
                command.CommandText = "INSERT INTO ReferralDates(TraceNo, AttendanceRefNum) VALUES(" & trace & "," & dsA.Tables(0).Rows(i).Item(0).ToString() & ")"
                command.ExecuteNonQuery()
                command.CommandText = "UPDATE Attendance SET ActionTaken=1 WHERE AttendanceRefNum=" & dsA.Tables(0).Rows(i).Item(0).ToString()
                command.ExecuteNonQuery()
                i = i + 1
            End While
            dsA.Dispose()
            Dim daD As New SqlCeDataAdapter("SELECT Student.LastName, Student.FirstName, Student.MiddleName, StudentList.SecCode FROM StudentList INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo WHERE StudentList.SLRefNum=" & SLRefNum, con)
            Dim dsD As New DataSet
            Dim cbD As New SqlCeCommandBuilder(daD)
            daD.Fill(dsD)
            Dim result As MessageBoxResult = MessageBox.Show(dsD.Tables(0).Rows(0).Item(0).ToString & ", " & dsD.Tables(0).Rows(0).Item(1).ToString & " " & dsD.Tables(0).Rows(0).Item(2) & " from " & dsD.Tables(0).Rows(0).Item(3) & " has been referred to the CATC for excessive absences. Would you prefer to fill out your concerns?", "Faculty Concerns", MessageBoxButton.YesNo, MessageBoxImage.Question)
            daD.Dispose()
            cbD.Dispose()
            dsD.Dispose()
            con.Close()
            If result = MessageBoxResult.Yes Then
                ActiveReferral = trace
                Dim x As New dialogFacultyConcerns()
                x.ShowDialog()
            End If
            con.Close()

        End If
        timer.Start()
    End Sub

    'Code for Recents Tab
    Private Sub dgActions_AutoGeneratingColumn(sender As Object, e As System.Windows.Controls.DataGridAutoGeneratingColumnEventArgs) Handles dgActions.AutoGeneratingColumn
        If e.PropertyType = GetType(System.DateTime) Then
            TryCast(e.Column, DataGridTextColumn).Binding.StringFormat = "dd/MM/yyyy"
        End If
    End Sub

    Private Sub dgReferrals_AutoGeneratingColumn(sender As Object, e As System.Windows.Controls.DataGridAutoGeneratingColumnEventArgs) Handles dgReferrals.AutoGeneratingColumn
        If e.PropertyType = GetType(System.DateTime) Then
            TryCast(e.Column, DataGridTextColumn).Binding.StringFormat = "dd/MM/yyyy"
        End If
    End Sub

    Private Sub dgActions_CellEditEnding(sender As Object, e As System.Windows.Controls.DataGridCellEditEndingEventArgs) Handles dgActions.CellEditEnding
        timer.Stop()
        If e.Column.DisplayIndex = 0 Then
            Dim c As CheckBox = e.Column.GetCellContent(e.Row)
            If c.IsChecked Then
                command.CommandText = "Update Attendance SET Excused=1 WHERE AttendanceRefNum=" & e.Row.Item("ID").ToString
                Try
                    con.Close()
                    con.Open()
                    command.ExecuteNonQuery()
                Catch ex As Exception
                    Console.WriteLine(ex.ToString)
                Finally
                    con.Close()
                End Try
            Else
                command.CommandText = "Update Attendance SET Excused=null WHERE AttendanceRefNum=" & e.Row.Item("ID").ToString
                Try
                    con.Close()
                    con.Open()
                    command.ExecuteNonQuery()
                Catch ex As Exception
                    Console.WriteLine(ex.ToString)
                Finally
                    con.Close()
                End Try
            End If

        End If
        timer.Start()
    End Sub

    'Code for Notifications Tab
    Private Sub expNotifActionTaken_Expanded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles expNotifActionTaken.Expanded
        expNotifFeedback.IsExpanded = False
    End Sub

    Private Sub expNotifFeedback_Expanded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles expNotifFeedback.Expanded
        expNotifActionTaken.IsExpanded = False
    End Sub

    Private Sub dgNotifications_MouseDoubleClick(sender As Object, e As System.Windows.Input.MouseButtonEventArgs) Handles dgNotifications.MouseDoubleClick
        timer.Stop()
        ActiveTab = 6
        setClickable()
        con.Close()
        con.Open()
        Dim row As DataRowView = dgNotifications.SelectedItems(0)
        lblRefRefNum.Content = row("Ref No")
        lblRefRefNum.Foreground = Brushes.White
        Dim ds As New DataSet
        Dim dsd As New DataSet
        Dim adapter As New SqlCeDataAdapter("Select Student.LastName as [Last Name], Student.FirstName as [First Name], Student.MiddleName as [Middle Name], Section.SubjCode as [Subject Code], Subject.Description as [Description], StudentList.SecCode as [Section], Referral.Concerns as [Concerns], Referral.ActionTaken as [Action Taken] FROM Referral INNER JOIN StudentList on Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section ON StudentList.SecCode=Section.SecCode INNER JOIN Subject ON Section.SubjCode=Subject.SubjCode WHERE Referral.TraceNo=" & lblRefRefNum.Content, con)
        Dim adapterd As New SqlCeDataAdapter("Select Attendance.Date as [Date] FROM ReferralDates INNER JOIN Attendance ON ReferralDates.AttendanceRefNum=Attendance.AttendanceRefNum WHERE TraceNo=" & lblRefRefNum.Content & "ORDER BY Attendance.Date ASC", con)
        Dim cbuilder As New SqlCeCommandBuilder(adapter)
        Dim cbuilderd As New SqlCeCommandBuilder(adapterd)
        adapter.Fill(ds)
        adapterd.Fill(dsd)
        Dim i As Integer = 0
        Dim dateholder As String
        tbRefDates.Text = String.Empty
        While i < dsd.Tables(0).Rows.Count
            dateholder = dsd.Tables(0).Rows(i).Item(0).ToString()
            tbRefDates.Text = tbRefDates.Text & Date.Parse(dateholder.ToString()).ToString("dd/MM/yyyy") & ", "
            i = i + 1
        End While
        lblRefName.Content = ds.Tables(0).Rows(0).Item("Last Name").ToString & ", " & ds.Tables(0).Rows(0).Item("First Name") & " " & ds.Tables(0).Rows(0).Item("Middle Name")
        lblRefSubject.Content = ds.Tables(0).Rows(0).Item("Subject Code").ToString & " - " & ds.Tables(0).Rows(0).Item("Description").ToString
        lblRefSection.Content = ds.Tables(0).Rows(0).Item("Section").ToString
        If ds.Tables(0).Rows(0).Item("Concerns").ToString().Equals(String.Empty) Then
            tbRefConcerns.Text = """No faculty concerns were set..."""
        Else
            tbRefConcerns.Text = ds.Tables(0).Rows(0).Item("Concerns").ToString()
        End If
        tbActionTaken.Text = ds.Tables(0).Rows(0).Item("Action Taken").ToString()
        con.Close()
        Dim da As New DoubleAnimation
        With da
            .To = 700
            .From = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        gridNotificationDetails.BeginAnimation(Grid.WidthProperty, da)
        expNotifFeedback.IsExpanded = True
    End Sub

    Private Sub dgNotifications_SelectionChanged(sender As System.Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles dgNotifications.SelectionChanged
        timer.Stop()
        ActiveTab = 6
        setClickable()
        con.Close()
        con.Open()
        Dim row As DataRowView = dgNotifications.SelectedItems(0)
        lblRefRefNum.Content = row("Ref No")
        lblRefRefNum.Foreground = Brushes.White
        Dim ds As New DataSet
        Dim dsd As New DataSet
        Dim adapter As New SqlCeDataAdapter("Select Student.LastName as [Last Name], Student.FirstName as [First Name], Student.MiddleName as [Middle Name], Section.SubjCode as [Subject Code], Subject.Description as [Description], StudentList.SecCode as [Section], Referral.Concerns as [Concerns], Referral.ActionTaken as [Action Taken] FROM Referral INNER JOIN StudentList on Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section ON StudentList.SecCode=Section.SecCode INNER JOIN Subject ON Section.SubjCode=Subject.SubjCode WHERE Referral.TraceNo=" & lblRefRefNum.Content, con)
        Dim adapterd As New SqlCeDataAdapter("Select Attendance.Date as [Date] FROM ReferralDates INNER JOIN Attendance ON ReferralDates.AttendanceRefNum=Attendance.AttendanceRefNum WHERE TraceNo=" & lblRefRefNum.Content & "ORDER BY Attendance.Date ASC", con)
        Dim cbuilder As New SqlCeCommandBuilder(adapter)
        Dim cbuilderd As New SqlCeCommandBuilder(adapterd)
        adapter.Fill(ds)
        adapterd.Fill(dsd)
        Dim i As Integer = 0
        Dim dateholder As String
        tbRefDates.Text = String.Empty
        While i < dsd.Tables(0).Rows.Count
            dateholder = dsd.Tables(0).Rows(i).Item(0).ToString()
            tbRefDates.Text = tbRefDates.Text & Date.Parse(dateholder.ToString()).ToString("dd/MM/yyyy") & ", "
            i = i + 1
        End While
        lblRefName.Content = ds.Tables(0).Rows(0).Item("Last Name").ToString & ", " & ds.Tables(0).Rows(0).Item("First Name") & " " & ds.Tables(0).Rows(0).Item("Middle Name")
        lblRefSubject.Content = ds.Tables(0).Rows(0).Item("Subject Code").ToString & " - " & ds.Tables(0).Rows(0).Item("Description").ToString
        lblRefSection.Content = ds.Tables(0).Rows(0).Item("Section").ToString
        If ds.Tables(0).Rows(0).Item("Concerns").ToString().Equals(String.Empty) Then
            tbRefConcerns.Text = """No faculty concerns were set..."""
        Else
            tbRefConcerns.Text = ds.Tables(0).Rows(0).Item("Concerns").ToString()
        End If
        tbActionTaken.Text = ds.Tables(0).Rows(0).Item("Action Taken").ToString()
        con.Close()
        Dim da As New DoubleAnimation
        With da
            .To = 700
            .From = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        gridNotificationDetails.BeginAnimation(Grid.WidthProperty, da)
        expNotifFeedback.IsExpanded = True
    End Sub

    Private Sub btnRefClear_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnRefClear.Click
        tbFeedback.Text = String.Empty
    End Sub

    Private Sub tbFeedback_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles tbFeedback.TextChanged
        tbFeedback.Text = tbFeedback.Text.Replace("'", "")
        If tbFeedback.Text.Length > 200 Then
            tbFeedback.Text = tbFeedback.Text.Remove(200)
        End If
        lblRefFeedCtr.Content = Val(200 - tbFeedback.Text.Length).ToString
        If Val(lblRefFeedCtr.Content) < 20 Then
            lblRefFeedCtr.Foreground = System.Windows.Media.Brushes.Red
        Else
            lblRefFeedCtr.Foreground = System.Windows.Media.Brushes.White
        End If
        tbFeedback.Select(tbFeedback.Text.Length, 0)
    End Sub

    Private Sub btnRefSubmit_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnRefSubmit.Click
        con.Close()
        con.Open()
        command.CommandText = "UPDATE Referral SET Feedback = '" & tbFeedback.Text & "' WHERE TraceNo=" & lblRefRefNum.Content
        command.ExecuteNonQuery()
        MessageBox.Show("Successfully submitted feedback to Referral " & lblRefRefNum.Content & ".", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        command.CommandText = "UPDATE FeedbackRequest SET Finished=1 WHERE TraceNo=" & lblRefRefNum.Content
        command.ExecuteNonQuery()
        con.Close()
        fillnotifications()
    End Sub

    Private Sub Search_TextChanged(sender As Object, e As System.Windows.Controls.TextChangedEventArgs) Handles Search.TextChanged
        timer.Stop()
        Dim da As New SqlCeDataAdapter("Select Student.LastName as [LAST NAME], Student.FirstName as [FIRST NAME], Student.MiddleName as [MIDDLE NAME], StudentList.StudentNo as [STUDENT NUMBER] FROM StudentList INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo WHERE (SecCode='" & lblSubjSec.Content & "' AND (Student.LastName LIKE '%" & Search.Text & "%' OR Student.FirstName LIKE '%" & Search.Text & "%' OR Student.MiddleName LIKE '%" & Search.Text & "%' OR Student.StudentNo LIKE '%" & Search.Text & "%'))  ORDER BY Student.LastName Asc", con)
        Dim cb As New SqlCeCommandBuilder(da)
        fillStudentList(da)
        timer.Start()
    End Sub
    'Code for Settings
    Private Sub btnSettings_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnSettings.Click
        Dim da As New DoubleAnimation
        With da
            .From = 0
            .To = 90
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        settings.BeginAnimation(StackPanel.HeightProperty, da)
    End Sub

    Private Sub btnChangePassword_MouseEnter(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnChangePassword.MouseEnter
        btnChangePassword.Foreground = System.Windows.Media.Brushes.Black
    End Sub

    Private Sub btnChangePassword_MouseLeave(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnChangePassword.MouseLeave
        btnChangePassword.Foreground = System.Windows.Media.Brushes.White
    End Sub

    Private Sub btnAbout_MouseEnter(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnAbout.MouseEnter
        btnAbout.Foreground = System.Windows.Media.Brushes.Black
    End Sub

    Private Sub btnAbout_MouseLeave(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnAbout.MouseLeave
        btnAbout.Foreground = System.Windows.Media.Brushes.White
    End Sub


    Private Sub btnHelp_MouseEnter(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnHelp.MouseEnter
        btnHelp.Foreground = System.Windows.Media.Brushes.Black
    End Sub

    Private Sub btnHelp_MouseLeave(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnHelp.MouseLeave
        btnHelp.Foreground = System.Windows.Media.Brushes.White
    End Sub

  
    Private Sub settings_MouseLeave(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles settings.MouseLeave
        Dim da As New DoubleAnimation
        With da
            .From = 90
            .To = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With
        settings.BeginAnimation(StackPanel.HeightProperty, da)
    End Sub
End Class

