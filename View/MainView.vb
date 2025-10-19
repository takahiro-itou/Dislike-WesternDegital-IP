Public Class MainView

Private Sub RunDiskAccess()
''--------------------------------------------------------------------
''    指定されたディスクアクセスを実行する。
''--------------------------------------------------------------------
Dim encUtf As System.Text.Encoding
Dim outText As String

    outText = $"書き込み時刻：{DateTime.Now:yyyy/MM/dd HH:mm:ss}"
    encUtf = System.Text.Encoding.UTF8

    Try
        Using sw As New System.IO.StreamWriter("F:\\Work\\DisWdIp.txt", True, encUtf)
            sw.WriteLine(outText)
        End Using
    Catch e As Exception
        txtOutput.Text += $"ファイルにアクセスできません：{e.Message}{Environment.NewLine}"
    End Try

    Try
        Using sw As New System.IO.StreamWriter("I:\\Work\\DisWdIp.txt", True, encUtf)
            sw.WriteLine(outText)
        End Using
    Catch e As Exception
        txtOutput.Text += $"ファイルにアクセスできません：{e.Message}{Environment.NewLine}"
    End Try

    txtOutput.Text += $"{outText}{Environment.NewLine}"

End Sub


Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles _
            btnRun.Click
''--------------------------------------------------------------------
''    「実行」ボタンのクリックイベントハンドラ。
''
''    入力したコマンドを実行する。
''--------------------------------------------------------------------

    RunDiskAccess()
End Sub

Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles _
            mnuFileExit.Click
''--------------------------------------------------------------------
''    メニュー「ファイル」－「終了」
''--------------------------------------------------------------------

    Application.Exit()
End Sub


Private Sub mnuRunCommand_Click(sender As Object, e As EventArgs) Handles _
            mnuRunCommand.Click
''--------------------------------------------------------------------
''    メニュー「実行」－「コマンドを実行」
''--------------------------------------------------------------------

    RunDiskAccess()
End Sub


Private Sub tmrDisk_Tick(sender As Object, e As EventArgs) Handles _
            tmrDisk.Tick
''--------------------------------------------------------------------
''    「タイマー」のイベントハンドラ
''--------------------------------------------------------------------

    RunDiskAccess()
End Sub

End Class
