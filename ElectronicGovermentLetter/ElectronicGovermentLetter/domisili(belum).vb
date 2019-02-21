Imports word = Microsoft.Office.Interop.Word
Public Class domisili_belum_
    Sub reset()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox3.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        ComboBox1.Text = ""
    End Sub
    Sub word()
        Dim ObjAppWord As New word.Application
        Dim ObjDocWord As New word.Document
        Dim namafile As String

        ObjDocWord = ObjAppWord.Documents.Open("D:\APP\Document\domisili(belum).docx")
        ObjDocWord.Bookmarks("no_surat").Select()
        ObjAppWord.Selection.TypeText(TextBox1.Text)
        ObjDocWord.Bookmarks("nama").Select()
        ObjAppWord.Selection.TypeText(TextBox2.Text)
        ObjDocWord.Bookmarks("tl").Select()
        ObjAppWord.Selection.TypeText(TextBox3.Text)
        ObjDocWord.Bookmarks("tgll").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker1.Text)
        ObjDocWord.Bookmarks("agama").Select()
        ObjAppWord.Selection.TypeText(ComboBox1.Text)
        ObjDocWord.Bookmarks("jk").Select()
        ObjAppWord.Selection.TypeText(ComboBox2.Text)
        ObjDocWord.Bookmarks("sp").Select()
        ObjAppWord.Selection.TypeText(ComboBox3.Text)
        ObjDocWord.Bookmarks("alamat1").Select()
        ObjAppWord.Selection.TypeText(TextBox5.Text)
        ObjDocWord.Bookmarks("tgls").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker2.Text)
        ObjDocWord.Bookmarks("alamat2").Select()
        ObjAppWord.Selection.TypeText(TextBox6.Text)

        namafile = "D:\APP\NewDocument\Domisili(Belum ada surat pindah)\" & TextBox5.Text & "-" & DateTimePicker2.Text & ".docx"
        ObjDocWord.SaveAs(namafile)
        ObjDocWord.Close()
        ObjAppWord.Quit()
    End Sub
    Sub open()
        Dim namafile As String
        Dim word As New Microsoft.Office.Interop.Word.Application
        Dim doc As New Microsoft.Office.Interop.Word.Document
        namafile = "D:\APP\NewDocument\Domisili(Belum ada surat pindah)\" & TextBox5.Text & "-" & DateTimePicker2.Text & ".docx"
        doc = word.Documents.Open(namafile)
        doc.Activate()
    End Sub
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        open()
    End Sub

    Private Sub domisili_belum__Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Islam")
        ComboBox1.Items.Add("Budha")
        ComboBox1.Items.Add("Hindu")
        ComboBox1.Items.Add("Katholik")
        ComboBox1.Items.Add("Kristen")

        ComboBox2.Items.Add("Laki-Laki")
        ComboBox2.Items.Add("Perempuan")

        ComboBox3.Items.Add("Kawin")
        ComboBox3.Items.Add("Belum Kawin")
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        reset()
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Environment.Exit(10)
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        word()
    End Sub
End Class