Public Class MainView

Private Sub RunDiskAccess()
''--------------------------------------------------------------------
''    指定されたディスクアクセスを実行する。
''--------------------------------------------------------------------

    Using process As New System.Diagnostics.Process()
        process.StartInfo.FileName = "ipconfig.exe"
        process.StartInfo.UseShellExecute = False
        process.StartInfo.RedirectStandardInput = False
        process.StartInfo.RedirectStandardOutput = True
        process.StartInfo.RedirectStandardError = False
        process.Start()

        Dim Reader As System.IO.StreamReader = process.StandardOutput
        Dim output As String = Reader.ReadToEnd()

        txtOutput.Text = output
        process.WaitForExit()
        process.Close()
    End Using

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
