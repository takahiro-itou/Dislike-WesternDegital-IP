Public Class MainView


Private Sub InitializeWorkFiles()
''--------------------------------------------------------------------
''    作業用のファイルを初期化する。
''--------------------------------------------------------------------
Dim encUtf As System.Text.Encoding
Dim outText As String

    outText = $"{DateTime.Now:yyyy/MM/dd HH:mm:ss}  初期化"
    encUtf = System.Text.Encoding.UTF8
    txtOutput.Text = $"{Environment.NewLine}"

    Try
        Using sw As New System.IO.StreamWriter("F:\\Work\\DisWdIp.txt", False, encUtf)
            sw.WriteLine(outText)
        End Using
    Catch e As Exception
        txtOutput.Text += $"ファイルにアクセスできません：{e.Message}{Environment.NewLine}"
    End Try

    Try
        Using sw As New System.IO.StreamWriter("I:\\Work\\DisWdIp.txt", False, encUtf)
            sw.WriteLine(outText)
        End Using
    Catch e As Exception
        txtOutput.Text += $"ファイルにアクセスできません：{e.Message}{Environment.NewLine}"
    End Try

End Sub

Private Sub RunDiskAccess()
''--------------------------------------------------------------------
''    指定されたディスクアクセスを実行する。
''--------------------------------------------------------------------
Dim outText As String

    outText = $"{DateTime.Now:yyyy/MM/dd HH:mm:ss}  書き込み"

    txtOutput.Text = $"{outText}{Environment.NewLine}"
    writeToWorkFile("F:\Work\DisWdIp.txt", True, encUtf);
    writeToWorkFile("I:\Work\DisWdIp.txt", True, encUtf);
    txtOutput.Text += "完了"

End Sub

Private Sub writeToWorkFile(
        ByVal fileName As String, ByVal strText As String,
        ByVal bAppend As Boolean)
''--------------------------------------------------------------------
''    指定されたファイルに書き込みを行う
''
''  @param [in] fileName    書き込み先のファイル名
''  @param [in] strText     書き込む内容
''  @param [in] bAppend     真ならば追記を行う
''--------------------------------------------------------------------
Dim encUtf As System.Text.Encoding

    encUtf = System.Text.Encoding.UTF8

    Try
        Using sw As New System.IO.StreamWriter(fileName, bAppend, encUtf)
            sw.WriteLine(strText)
        End Using
        txtOutput.Text += $"ファイル {fileName} に書き込み成功!{Environment.NewLine}"
    Catch e As Exception
        txtOutput.Text += $"ファイルにアクセスできません：{e.Message}{Environment.NewLine}"
    End Try

End Sub

Private Sub MainView_Load(sender As Object, e As EventArgs) Handles _
            MyBase.Load
''--------------------------------------------------------------------
''    フォームのロードイベントハンドラ。
''--------------------------------------------------------------------

    InitializeWorkFiles()
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
