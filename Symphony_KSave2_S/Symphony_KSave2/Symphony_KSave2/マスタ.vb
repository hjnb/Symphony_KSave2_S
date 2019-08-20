Imports System.Data.OleDb

Public Class マスタ

    '利用者マスタ表示用データテーブル
    Private dtUser As DataTable

    '行ヘッダーのカレントセルを表す三角マークを非表示に設定する為のクラス。
    Public Class dgvRowHeaderCell

        'DataGridViewRowHeaderCell を継承
        Inherits DataGridViewRowHeaderCell

        'DataGridViewHeaderCell.Paint をオーバーライドして行ヘッダーを描画
        Protected Overrides Sub Paint(ByVal graphics As Graphics, ByVal clipBounds As Rectangle, _
           ByVal cellBounds As Rectangle, ByVal rowIndex As Integer, ByVal cellState As DataGridViewElementStates, _
           ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, _
           ByVal cellStyle As DataGridViewCellStyle, ByVal advancedBorderStyle As DataGridViewAdvancedBorderStyle, _
           ByVal paintParts As DataGridViewPaintParts)
            '標準セルの描画からセル内容の背景だけ除いた物を描画(-5)
            MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, _
                     formattedValue, errorText, cellStyle, advancedBorderStyle, _
                     Not DataGridViewPaintParts.ContentBackground)
        End Sub

    End Class

    Public Sub New()
        InitializeComponent()
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
    End Sub

    Private Sub マスタ_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '入力テキストボックス設定
        initTextBox()

        'dgv初期設定
        initDgvUser()

        'マスタデータ表示
        displayDgvUser()
    End Sub

    Private Sub マスタ_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If e.Control = False Then
                Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
            End If
        End If
    End Sub

    Private Sub initTextBox()
        '利用者名ボックス
        namBox.ImeMode = Windows.Forms.ImeMode.Hiragana

        'カナボックス
        kanaBox.ImeMode = Windows.Forms.ImeMode.Hiragana
    End Sub

    Private Sub displayDgvUser()
        clearInputText()
        dgvUser.DataSource = Nothing
        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select Nam, Kana, Autono from UsrM order by Kana"
        cnn.Open(topForm.DB_KSave2)
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        Dim da As OleDbDataAdapter = New OleDbDataAdapter()
        Dim ds As DataSet = New DataSet()
        da.Fill(ds, rs, "UsrM")
        dtUser = ds.Tables(0)
        dgvUser.DataSource = dtUser

        '列幅、スタイル設定
        settingDgvUserStyle()

        '選択解除
        If Not IsNothing(dgvUser.CurrentCell) Then
            dgvUser.CurrentCell.Selected = False
        End If
    End Sub

    Private Sub initDgvUser()
        Util.EnableDoubleBuffering(dgvUser)

        With dgvUser
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
            .RowTemplate.HeaderCell = New dgvRowHeaderCell() '行ヘッダの三角マークを非表示に
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ShowCellToolTips = False
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(155, 202, 239)
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .RowHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(155, 202, 239)
            .EnableHeadersVisualStyles = False
        End With
    End Sub

    Private Sub settingDgvUserStyle()
        With dgvUser
            '並び替えができないようにする
            For Each c As DataGridViewColumn In .Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            '非表示列
            .Columns("Autono").Visible = False

            With .Columns("Nam")
                .Width = 113
                .HeaderText = "利用者名"
            End With

            With .Columns("Kana")
                .Width = 157
                .HeaderText = "かな"
            End With

        End With
    End Sub

    Private Sub clearInputText()
        namBox.Text = ""
        kanaBox.Text = ""
    End Sub

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        Dim nam As String = namBox.Text
        If nam = "" Then
            MsgBox("漢字氏名を入力して下さい。")
            namBox.Focus()
            Return
        End If
        Dim kana As String = kanaBox.Text
        If kana = "" Then
            MsgBox("かな氏名を入力して下さい。")
            kanaBox.Focus()
            Return
        End If

        Dim cn As New ADODB.Connection()
        cn.Open(topForm.DB_KSave2)
        Dim cmd As New ADODB.Command()
        cmd.ActiveConnection = cn
        If Not IsNothing(dgvUser.CurrentCell) AndAlso dgvUser.CurrentCell.Selected Then
            '変更登録
            Dim result As DialogResult = MessageBox.Show("変更登録してよろしいですか？", "登録", MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.Yes Then
                Dim autono As Integer = dgvUser("Autono", dgvUser.CurrentRow.Index).Value
                cmd.CommandText = "update UsrM Set Nam=?, Kana=? where Autono=?"
                cmd.Execute(Parameters:={nam, kana, autono})
                cn.Close()

                '再表示
                displayDgvUser()
            Else
                cn.Close()
            End If
        Else
            '新規登録
            Dim result As DialogResult = MessageBox.Show("新規登録してよろしいですか？", "登録", MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.Yes Then
                cmd.CommandText = "insert into UsrM (Nam, Kana) values ('" & nam & "', '" & kana & "')"
                cmd.Execute()
                cn.Close()

                '再表示
                displayDgvUser()
            Else
                cn.Close()
            End If
        End If

    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        If Not IsNothing(dgvUser.CurrentCell) AndAlso dgvUser.CurrentCell.Selected Then
            Dim cn As New ADODB.Connection()
            cn.Open(topForm.DB_KSave2)
            Dim cmd As New ADODB.Command()
            cmd.ActiveConnection = cn
            Dim result As DialogResult = MessageBox.Show("削除してよろしいですか？", "削除", MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.Yes Then
                Dim autono As Integer = dgvUser("Autono", dgvUser.CurrentRow.Index).Value
                cmd.CommandText = "Delete from UsrM where Autono=?"
                cmd.Parameters.Refresh()
                cmd.Execute(Parameters:=autono)
                cn.Close()

                '再表示
                displayDgvUser()
            Else
                cn.Close()
            End If
        Else
            MsgBox("選択されていません。")
            Return
        End If
    End Sub

    Private Sub dgvUser_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvUser.CellMouseClick
        If e.RowIndex >= 0 Then
            Dim autono As Integer = dgvUser("Autono", e.RowIndex).Value
            Dim nam As String = Util.checkDBNullValue(dgvUser("Nam", e.RowIndex).Value)
            Dim kana As String = Util.checkDBNullValue(dgvUser("Kana", e.RowIndex).Value)

            namBox.Text = nam
            kanaBox.Text = kana
        End If
    End Sub

    Private Sub dgvUser_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvUser.CellPainting
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