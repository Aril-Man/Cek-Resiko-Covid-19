Public Class Form1
    Dim sql As String

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If CheckBox1.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox2.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox3.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox4.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox5.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox6.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox7.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox8.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox9.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox10.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox11.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox12.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox13.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox14.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox15.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox16.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox17.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If
        If CheckBox18.Checked = True Then
            RichTextBox1.Text = Val(RichTextBox1.Text) + 1
        End If

        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
        CheckBox5.Enabled = False
        CheckBox6.Enabled = False
        CheckBox7.Enabled = False
        CheckBox8.Enabled = False
        CheckBox9.Enabled = False
        CheckBox10.Enabled = False
        CheckBox11.Enabled = False
        CheckBox12.Enabled = False
        CheckBox13.Enabled = False
        CheckBox14.Enabled = False
        CheckBox15.Enabled = False
        CheckBox16.Enabled = False
        CheckBox17.Enabled = False
        CheckBox18.Enabled = False
    End Sub

    Sub panggil()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_covid", con)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_covid")
        DataGridView1.DataSource = DS.Tables("tb_covid")
        DataGridView1.Enabled = True
    End Sub

    Sub jalan()
        Dim dataku As New System.Data.OleDb.OleDbCommand
        Call konek()
        dataku.Connection = con
        dataku.CommandType = CommandType.Text
        dataku.CommandText = sql
        dataku.ExecuteNonQuery()
        dataku.Dispose()
        Tnama.Text = ""
        Tnis.Text = ""
        RichTextBox1.Text = ""
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call panggil()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RichTextBox1.Text >= 0 Then
            Resiko.Text = "Rendah"
        End If

        If RichTextBox1.Text >= 8 Then
            Resiko.Text = "Sendang"
        End If

        If RichTextBox1.Text >= 15 Then
            Resiko.Text = "Tinggi"
        End If
        sql = "insert into tb_covid (Nama, NIS, Point,Resiko) values ('" & Tnama.Text & "','" & Tnis.Text & "','" & RichTextBox1.Text & "','" & Resiko.Text & "')"
        Call jalan()
        MsgBox("Data Berhasil Tersimpan")
        Call panggil()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_covid where Nama like '%" & Tnama.Text & "%'", con)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_covid")
        DataGridView1.DataSource = DS.Tables("tb_covid")
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        Tnama.Text = DataGridView1.Item(0, i).Value
        Tnis.Text = DataGridView1.Item(1, i).Value
        RichTextBox1.Text = DataGridView1.Item(2, i).Value
        Resiko.Text = DataGridView1.Item(3, i).Value
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        sql = "DELETE FROM tb_covid where Nama='" & Tnama.Text & "'"
        Call jalan()
        MsgBox("Data Berhasil Terhapus")
        Call panggil()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        CheckBox6.Enabled = True
        CheckBox7.Enabled = True
        CheckBox8.Enabled = True
        CheckBox9.Enabled = True
        CheckBox10.Enabled = True
        CheckBox11.Enabled = True
        CheckBox12.Enabled = True
        CheckBox13.Enabled = True
        CheckBox14.Enabled = True
        CheckBox15.Enabled = True
        CheckBox16.Enabled = True
        CheckBox17.Enabled = True
        CheckBox18.Enabled = True

        Tnama.Text = " "
        Tnis.Text = " "
        Resiko.Text = ""
        RichTextBox1.Text = " "
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
        End
    End Sub
End Class
