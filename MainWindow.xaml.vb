Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Class WifiCrendialsInformation
    Public Property ID As Integer
    Public Property SSID As String
    Public Property Password As String
End Class

Class MainWindow
    Dim g_ssid, h_ssid As String
    Dim g_pwd, h_pwd As String
    Dim o_id As Integer = 0
    Public Property WifiCredentials As ObservableCollection(Of WifiCrendialsInformation)

    Private ReadOnly worker As BackgroundWorker = New BackgroundWorker()


    Private Sub TopWindow_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TopWindow.MouseDown
        If e.ChangedButton = MouseButton.Left Then DragMove()
    End Sub
    Private Sub Close_MouseEnter(sender As Object, e As MouseEventArgs) Handles Close.MouseEnter
        Close.Foreground = New SolidColorBrush(Colors.Red)
    End Sub

    Private Sub Close_MouseLeave(sender As Object, e As MouseEventArgs) Handles Close.MouseLeave
        Close.Foreground = New SolidColorBrush(Colors.Firebrick)
    End Sub
    Private Sub Close_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Close.MouseLeftButtonDown
        MyBase.Close()
    End Sub

    Private Sub Window_Closing(sender As Object, e As CancelEventArgs)
        MySettings.Default.Save()
    End Sub


    Private Sub InitWifiCredentialsCollection(ars As Array, arp As Array)

        WifiCredentials = New ObservableCollection(Of WifiCrendialsInformation)()
        Dim oe As Integer = 0
        For Each element As String In ars

            WifiCredentials.Add(New WifiCrendialsInformation() With {
                          .ID = oe,
                          .SSID = element,
                          .Password = arp(oe)})
            oe += 1
        Next

    End Sub

    Private Sub worker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)

        Dim arr_ssid(1) As String
        Dim arr_pwd(1) As String



        startprgrmwifi("cmd /c netsh wlan show profiles")

        Dim pattern As String = "(: )(.*)"
        Dim options As RegexOptions = RegexOptions.Multiline

        Dim input_s As String = g_ssid
        Dim o_id As Integer = 1

        For Each m As Match In Regex.Matches(input_s, pattern, options)

            m = m.NextMatch()
            h_ssid = m.Groups(2).Value
            rgx_wifi(m.Groups(2).Value)
            ReDim Preserve arr_pwd(o_id)
            ReDim Preserve arr_ssid(o_id)

            arr_ssid(o_id) = h_ssid
            arr_pwd(o_id) = h_pwd

            o_id += 1
        Next

        InitWifiCredentialsCollection(arr_ssid, arr_pwd)


        WifiCredentials.Remove(WifiCredentials.First)
        WifiCredentials.Remove(WifiCredentials.Last)

    End Sub
    Private Sub worker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)
        DataContext = Me
        label_PleaseWait.Visibility = Visibility.Hidden
        GetBtn.IsEnabled = True
    End Sub

    Private Sub GetWifiPass_Click(sender As Object, e As RoutedEventArgs)

        AddHandler worker.DoWork, AddressOf worker_DoWork
        AddHandler worker.RunWorkerCompleted, AddressOf worker_RunWorkerCompleted
        worker.RunWorkerAsync()

        label_PleaseWait.Visibility = Visibility.Visible
        GetBtn.IsEnabled = False


    End Sub


    Private Sub rgx_wifi(g_input)
        startprgrmwifi("cmd /c netsh wlan show profiles " & Chr(34) & g_input & Chr(34) & " key=clear |find /I " & Chr(34) & "Conte" & Chr(34))

        Dim pattern As String = "(: )(.*)"
        Dim options As RegexOptions = RegexOptions.Multiline

        Dim input_p As String = g_pwd
        For Each m As Match In Regex.Matches(input_p, pattern, options)
            h_pwd = m.Groups(2).Value
        Next

    End Sub
    Private Sub startprgrmwifi(ByVal bat_command As String)

        Dim file_path As String = IO.Path.GetTempPath

        Dim FILE_NAME As String = file_path & "\000-wifi.bat"

        Dim fs As FileStream = File.Create(FILE_NAME)
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(bat_command)
        fs.Write(info, 0, info.Length)
        fs.Close()
        Dim start_info As New ProcessStartInfo(FILE_NAME)
        start_info.UseShellExecute = False
        start_info.CreateNoWindow = True
        start_info.RedirectStandardOutput = True
        start_info.RedirectStandardError = True

        Dim proc As New Process()
        proc.StartInfo = start_info
        proc.Start()

        Dim std_out As StreamReader = proc.StandardOutput()
        Dim std_err As StreamReader = proc.StandardError()

        g_ssid = std_out.ReadToEnd()
        g_pwd = g_ssid

        std_out.Close()
        std_err.Close()
        proc.Close()
        My.Computer.FileSystem.DeleteFile(FILE_NAME)
    End Sub

End Class
