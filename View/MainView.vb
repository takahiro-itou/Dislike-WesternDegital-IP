Public Class MainView

Private Sub RunDiskAccess()
''--------------------------------------------------------------------
''    指定されたディスクアクセスを実行する。
''--------------------------------------------------------------------
Dim encUtf As System.Text.Encoding
Dim outText As String

    outText = $"書き込み時刻：{DateTime.Now:yyyy/MM/DD HH:mm:ss}"
    encUtf = System.Text.Encoding.UTF8

    Using sw As New System.IO.StreamWriter("F:\\DisWdIp.txt", True, encUtf)
        sw.WriteLine(outText)
    End Using

    txtOutput.Text += $"{outText}\n"

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
