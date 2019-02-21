Imports word = Microsoft.Office.Interop.Word
Public Class permohonan_kompensasi
    Sub reset()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        ComboBox1.Text = ""
    End Sub
    Sub word()
        Dim ObjAppWord As New word.Application
        Dim ObjDocWord As New word.Document
        Dim namafile As String

        ObjDocWord = ObjAppWord.Documents.Open("D:\APP\Document\permohonan kompensasi.docx")
        ObjDocWord.Bookmarks("no").Select()
        ObjAppWord.Selection.TypeText(TextBox1.Text)
        ObjDocWord.Bookmarks("lampiran").Select()
        ObjAppWord.Selection.TypeText(TextBox2.Text)
        ObjDocWord.Bookmarks("nama").Select()
        ObjAppWord.Selection.TypeText(TextBox3.Text)
        ObjDocWord.Bookmarks("tl").Select()
        ObjAppWord.Selection.TypeText(TextBox4.Text)
        ObjDocWord.Bookmarks("tgl").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker1.Text)
        ObjDocWord.Bookmarks("jk").Select()
        ObjAppWord.Selection.TypeText(ComboBox1.Text)
        ObjDocWord.Bookmarks("pk").Select()
        ObjAppWord.Selection.TypeText(TextBox5.Text)
        ObjDocWord.Bookmarks("nik").Select()
        ObjAppWord.Selection.TypeText(TextBox6.Text)
        ObjDocWord.Bookmarks("alamat").Select()
        ObjAppWord.Selection.TypeText(TextBox7.Text)
        ObjDocWord.Bookmarks("nop").Select()
        ObjAppWord.Selection.TypeText(TextBox8.Text)
        ObjDocWord.Bookmarks("nop1").Select()
        ObjAppWord.Selection.TypeText(TextBox9.Text)
        ObjDocWord.Bookmarks("an").Select()
        ObjAppWord.Selection.TypeText(TextBox10.Text)
        ObjDocWord.Bookmarks("nobook").Select()
        ObjAppWord.Selection.TypeText(TextBox11.Text)
        ObjDocWord.Bookmarks("tgls").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker2.Text)
        ObjDocWord.Bookmarks("pemohon").Select()
        ObjAppWord.Selection.TypeText(TextBox3.Text)

        namafile = "D:\APP\NewDocument\Permohonan Kompensasi\" & TextBox3.Text & "-" & DateTimePicker2.Text & ".docx"
        ObjDocWord.SaveAs(namafile)
        ObjDocWord.Close()
        ObjAppWord.Quit()
    End Sub
    Sub open()
        Dim namafile As String
        Dim word As New Microsoft.Office.Interop.Word.Application
        Dim doc As New Microsoft.Office.Interop.Word.Document
        namafile = "D:\APP\NewDocument\Permohonan Kompensasi\" & TextBox3.Text & "-" & DateTimePicker2.Text & ".docx"
        doc = word.Documents.Open(namafile)
        doc.Activate()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        word()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        reset()
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Environment.Exit(10)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        word()
    End Sub

    Private Sub Panel3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel2.Paint

    End Sub
End Class