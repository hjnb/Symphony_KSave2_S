Public Class topForm

    '.iniファイルのパス
    Public iniFilePath As String = My.Application.Info.DirectoryPath & "\KSave2.ini"

    'メインとするデータベース(S:シンフォニー、A:アネックス)
    Public mainType As String = Util.getIniString("System", "MainType", iniFilePath)

    'データベースのパス
    Public mainDBFileName As String = Util.getIniString("System", "MainDBFileName", iniFilePath)
    Public dbFilePath As String = My.Application.Info.DirectoryPath & "\" & mainDBFileName
    Public DB_KSave2 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFilePath

    'エクセルのパス
    Public excelFilePass As String = My.Application.Info.DirectoryPath & "\書式.xls"

    'もう一方のデーターベースのパス
    Public dbAnotherFilePath As String = Util.getIniString("System", "AnotherDBFilePath", iniFilePath)
    Public DB_AnotherKSave2 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbAnotherFilePath

    'フォーム
    Private surveySlipForm As 認定調査票
    Private readOnlyForm As 認定調査票閲覧
    Private masterForm As マスタ
    Private writerMFrom As 記入者マスタ

    Public Sub New()
        InitializeComponent()
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        btnTarget.Visible = False
        initPrintState()
    End Sub

    Private Sub topForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'データベース、エクセル、構成ファイルの存在チェック
        If Not System.IO.File.Exists(iniFilePath) Then
            MsgBox("構成ファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        If mainType <> "S" AndAlso mainType <> "A" Then
            MsgBox("構成ファイルのMainTypeは'S'または'A'を指定して下さい。", MsgBoxStyle.Exclamation)
            Me.Close()
            Exit Sub
        End If

        If Not System.IO.File.Exists(dbFilePath) Then
            MsgBox(dbFilePath & "が存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        If Not System.IO.File.Exists(dbAnotherFilePath) Then
            MsgBox(dbAnotherFilePath & "が存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        If Not System.IO.File.Exists(excelFilePass) Then
            MsgBox("エクセルファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

    End Sub

    Private Sub btnMaster_Click(sender As System.Object, e As System.EventArgs) Handles btnMaster.Click
        btnTarget.Visible = True
        btnWriter.Visible = True
    End Sub

    Private Sub btnTarget_Click(sender As System.Object, e As System.EventArgs) Handles btnTarget.Click
        If IsNothing(masterForm) OrElse masterForm.IsDisposed Then
            masterForm = New マスタ()
            masterForm.Owner = Me
            masterForm.Show()
        End If
    End Sub

    Private Sub btnWriter_Click(sender As System.Object, e As System.EventArgs) Handles btnWriter.Click
        If IsNothing(writerMFrom) OrElse writerMFrom.IsDisposed Then
            writerMFrom = New 記入者マスタ()
            writerMFrom.Owner = Me
            writerMFrom.Show()
        End If
    End Sub

    Private Sub btnSurveySlip_Click(sender As System.Object, e As System.EventArgs) Handles btnSurveySlip.Click
        If IsNothing(surveySlipForm) OrElse surveySlipForm.IsDisposed Then
            surveySlipForm = New 認定調査票()
            surveySlipForm.Owner = Me
            surveySlipForm.Show()
        End If
    End Sub

    Private Sub initPrintState()
        Dim state As String = Util.getIniString("System", "Printer", iniFilePath)
        If state = "Y" Then
            rbtnPrint.Checked = True
        Else
            rbtnPreview.Checked = True
        End If
    End Sub

    Private Sub rbtnPreview_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtnPreview.CheckedChanged
        If rbtnPreview.Checked = True Then
            Util.putIniString("System", "Printer", "N", iniFilePath)
        End If
    End Sub

    Private Sub rbtnPrint_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtnPrint.CheckedChanged
        If rbtnPrint.Checked = True Then
            Util.putIniString("System", "Printer", "Y", iniFilePath)
        End If
    End Sub

    ''' <summary>
    ''' （閲覧用）ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnReadOnly_Click(sender As System.Object, e As System.EventArgs) Handles btnReadOnly.Click
        If IsNothing(readOnlyForm) OrElse readOnlyForm.IsDisposed Then
            readOnlyForm = New 認定調査票閲覧()
            readOnlyForm.Owner = Me
            readOnlyForm.Show()
        End If
    End Sub
End Class
