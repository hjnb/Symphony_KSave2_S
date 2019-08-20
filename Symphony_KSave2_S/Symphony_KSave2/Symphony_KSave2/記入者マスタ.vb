Imports System.Data.OleDb

Public Class 記入者マスタ

    '利用者マスタ表示用データテーブル
    Private dtWriter As DataTable

    Public Sub New()
        InitializeComponent()
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
    End Sub

    Private Sub 記入者マスタ_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If e.Control = False Then
                Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
            End If
        End If
    End Sub

    Private Sub 記入者マスタ_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '入力テキストボックス設定
        initTextBox()

        'dgv初期設定
        initDgvWriter()

        'マスタデータ表示
        displayDgvWriter()
    End Sub

    Private Sub initTextBox()
        '利用者名ボックス
        namBox.ImeMode = Windows.Forms.ImeMode.Hiragana

        'カナボックス
        kanaBox.ImeMode = Windows.Forms.ImeMode.Hiragana
    End Sub

    Private Sub displayDgvWriter()
        clearInputText()
        dgvWriter.DataSource = Nothing
        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select Nam, Num from EtcM order by Num"
        cnn.Open(topForm.DB_KSave2)
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        Dim da As OleDbDataAdapter = New OleDbDataAdapter()
        Dim ds As DataSet = New DataSet()
        da.Fill(ds, rs, "EtcM")
        dtWriter = ds.Tables(0)
        dgvWriter.DataSource = dtWriter

        '列幅、スタイル設定
        With dgvWriter
            '並び替えができないようにする
            For Each c As DataGridViewColumn In .Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            '非表示
            .Columns("Num").Visible = False

            With .Columns("Nam")
                .Width = 150
                .HeaderText = "記入者名"
            End With

        End With

        '選択解除
        If Not IsNothing(dgvWriter.CurrentCell) Then
            dgvWriter.CurrentCell.Selected = False
        End If

        'フォーカス
        namBox.Focus()
    End Sub

    Private Sub initDgvWriter()
        Util.EnableDoubleBuffering(dgvWriter)

        With dgvWriter
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .ReadOnly = True
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 18
            .RowTemplate.Height = 16
            .RowHeadersWidth = 36
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ShowCellToolTips = False
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(155, 202, 239)
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .RowHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(155, 202, 239)
            .EnableHeadersVisualStyles = False
        End With
    End Sub

    Private Sub clearInputText()
        namBox.Text = ""
        kanaBox.Text = ""
    End Sub

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        Dim nam As String = namBox.Text
        If nam = "" Then
            MsgBox("記入者名を入力して下さい。")
            namBox.Focus()
            Return
        End If

        'Numの最大値取得
        Dim numMax As Integer = 0
        If dgvWriter.Rows.Count > 0 Then
            numMax = dgvWriter("Num", dgvWriter.Rows.Count - 1).Value
        End If

        Dim cn As New ADODB.Connection()
        cn.Open(topForm.DB_KSave2)
        Dim sql As String = "select * from EtcM where Nam = '" & nam & "'"
        Dim rs As New ADODB.Recordset
        rs.Open(sql, cn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If rs.RecordCount > 0 Then
            MsgBox("既に登録されています。", MsgBoxStyle.Exclamation)
            rs.Close()
            cn.Close()
            Return
        End If
        rs.AddNew()
        rs.Fields("Nam").Value = nam
        rs.Fields("Num").Value = numMax + 1
        rs.Update()
        rs.Close()
        cn.Close()

        '再表示
        displayDgvWriter()


    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        '氏名
        Dim nam As String = namBox.Text
        If nam = "" Then
            MsgBox("選択して下さい。", MsgBoxStyle.Exclamation)
            Return
        End If

        '削除
        Dim result As DialogResult = MessageBox.Show("削除してよろしいですか？", "削除", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = Windows.Forms.DialogResult.Yes Then
            Dim cnn As New ADODB.Connection
            cnn.Open(topForm.DB_KSave2)
            Dim cmd As New ADODB.Command()
            cmd.ActiveConnection = cnn
            cmd.CommandText = "delete from EtcM where Nam = '" & nam & "'"
            cmd.Execute()
            cnn.Close()

            '再表示
            displayDgvWriter()
        End If
    End Sub

    Private Sub dgvUser_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvWriter.CellMouseClick
        If e.RowIndex >= 0 Then
            Dim nam As String = Util.checkDBNullValue(dgvWriter("Nam", e.RowIndex).Value)
            namBox.Text = nam
        End If
    End Sub

    Private Sub dgvUser_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvWriter.CellPainting
        '行ヘッダーかどうか調べる
        If e.ColumnIndex < 0 AndAlso e.RowIndex >= 0 Then
            'セルを描画する
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All)

            '行番号を描画する範囲を決定する
            'e.AdvancedBorderStyleやe.CellStyle.Paddingは無視しています
            Dim indexRect As Rectangle = e.CellBounds
            indexRect.Inflate(-2, -2)

            '選択状態を調べて文字色を変更する
            Dim forecolor As Color
            If DataGridViewElementStates.Selected = e.State Then
                forecolor = e.CellStyle.SelectionForeColor
            Else
                forecolor = e.CellStyle.ForeColor
            End If

            '行番号を描画する
            TextRenderer.DrawText(e.Graphics, _
                (e.RowIndex + 1).ToString(), _
                e.CellStyle.Font, _
                indexRect, _
                forecolor, _
                TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            '描画が完了したことを知らせる
            e.Handled = True
        End If
    End Sub
End Class