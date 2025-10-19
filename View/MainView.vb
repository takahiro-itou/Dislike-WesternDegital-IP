Public Class MainView


Private m_prvText As String


Private Sub initializeWorkFiles()
''--------------------------------------------------------------------
''    作業用のファイルを初期化する。
''--------------------------------------------------------------------
Dim outText As String
Dim curText As String

    outText = $"{DateTime.Now:yyyy/MM/dd HH:mm:ss}  初期化"
    curText = outText

    txtOutput.Text = $"{Environment.NewLine}"
    curText += writeToWorkFile("F:\Work\DisWdIp.txt", outText, False)
    curText += writeToWorkFile("I:\Work\DisWdIp.txt", outText, False)
    curText += "完了"

    txtOutput.Text += $"{Environment.NewLine}{curText}{Environment.NewLine}"
    Me.m_prvText = curText

End Sub

Private Sub runDiskAccess()
''--------------------------------------------------------------------
''    指定されたディスクアクセスを実行する。
''--------------------------------------------------------------------
Dim outText As String
Dim curText As String

    outText = $"{DateTime.Now:yyyy/MM/dd HH:mm:ss}  書き込み"
    curText = outText

    txtOutput.Text = $"{Me.m_prvText}{Environment.NewLine}"
    curText += writeToWorkFile("F:\Work\DisWdIp.txt", outText, True)
    curText += writeToWorkFile("I:\Work\DisWdIp.txt", outText, True)
    curText += "完了"

    txtOutput.Text += $"{Environment.NewLine}{curText}{Environment.NewLine}"
    Me.m_prvText = curText

End Sub


Private Function writeToWorkFile(
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
    writeToWorkFile = ""

    Try
        Using sw As New System.IO.StreamWriter(fileName, bAppend, encUtf)
            sw.WriteLine(strText)
        End Using
        writeToWorkFile += $"ファイル {fileName} に書き込み成功!{Environment.NewLine}"
    Catch e As Exception
        writeToWorkFile += $"ファイルにアクセスできません：{e.Message}{Environment.NewLine}"
    End Try

End Function


Private Sub MainView_Load(sender As Object, e As EventArgs) Handles _
            MyBase.Load
''--------------------------------------------------------------------
''    フォームのロードイベントハンドラ。
''--------------------------------------------------------------------

    Me.m_prvText = "ロード中"

    initializeWorkFiles()
    tmrDisk.Enabled = True
End Sub


Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles _
            btnRun.Click
''--------------------------------------------------------------------
''    「実行」ボタンのクリックイベントハンドラ。
''
''    入力したコマンドを実行する。
''--------------------------------------------------------------------

    runDiskAccess()
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

    runDiskAccess()
End Sub


Private Sub tmrDisk_Tick(sender As Object, e As EventArgs) Handles _
            tmrDisk.Tick
''--------------------------------------------------------------------
''    「タイマー」のイベントハンドラ
''--------------------------------------------------------------------

    runDiskAccess()
End Sub

End Class
