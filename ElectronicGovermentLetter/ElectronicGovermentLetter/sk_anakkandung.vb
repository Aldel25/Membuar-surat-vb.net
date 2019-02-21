Imports word = Microsoft.Office.Interop.Word
Public Class sk_anakkandung
    Sub reset()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
    End Sub

    Private Sub sk_anakkandung_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Islam")
        ComboBox1.Items.Add("Budha")
        ComboBox1.Items.Add("Hindu")
        ComboBox1.Items.Add("Katholik")
        ComboBox1.Items.Add("Kristen")

        ComboBox2.Items.Add("Laki-Laki")
        ComboBox2.Items.Add("Perempuan")
    End Sub
    Sub word()
        Dim ObjAppWord As New word.Application
        Dim ObjDocWord As New word.Document
        Dim namafile As String

        ObjDocWord = ObjAppWord.Documents.Open("D:\APP\Document\keterangan anak kandung.docx")
        ObjDocWord.Bookmarks("no_surat").Select()
        ObjAppWord.Selection.TypeText(TextBox1.Text)
        ObjDocWord.Bookmarks("nama").Select()
        ObjAppWord.Selection.TypeText(TextBox2.Text)
        ObjDocWord.Bookmarks("jk").Select()
        ObjAppWord.Selection.TypeText(ComboBox2.Text)
        ObjDocWord.Bookmarks("tmptlahir").Select()
        ObjAppWord.Selection.TypeText(TextBox3.Text)
        ObjDocWord.Bookmarks("tgllahir").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker1.Text)
        ObjDocWord.Bookmarks("agama").Select()
        ObjAppWord.Selection.TypeText(ComboBox2.Text)
        ObjDocWord.Bookmarks("namaayah").Select()
        ObjAppWord.Selection.TypeText(TextBox4.Text)
        ObjDocWord.Bookmarks("namaibu").Select()
        ObjAppWord.Selection.TypeText(TextBox5.Text)
        ObjDocWord.Bookmarks("tglsurat").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker2.Text)

        namafile = "D:\APP\NewDocument\Surat Keterangan Anak Kandung\" & TextBox2.Text & "-" & TextBox1.Text & "-" & DateTimePicker2.Text & ".docx"
        ObjDocWord.SaveAs(namafile)
        ObjDocWord.Close()
        ObjAppWord.Quit()
    End Sub

    Sub open()
        Dim namafile As String
        Dim word As New Microsoft.Office.Interop.Word.Application
        Dim doc As New Microsoft.Office.Interop.Word.Document
        namafile = "D:\APP\NewDocument\Surat Keterangan Anak Kandung\" & TextBox2.Text & "-" & TextBox1.Text & "-" & DateTimePicker2.Text & ".docx"
        doc = word.Documents.Open(namafile)
        doc.Activate()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        open()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        reset()
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Environment.Exit(10)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        word()
    End Sub
End Class