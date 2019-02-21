Imports word = Microsoft.Office.Interop.Word
Public Class s_pendistribusian
    Private Sub DomisiliToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DomisiliToolStripMenuItem.Click
        If domisili_rt_.Enabled = True Then
            domisili_rt_.Show()
            Me.Hide()
        End If
    End Sub
    Private Sub PengantarIPPTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PengantarIPPTToolStripMenuItem.Click
        If pengantarippt.Enabled = True Then
            pengantarippt.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub PermohonanKomponesasiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PermohonanKomponesasiToolStripMenuItem.Click
        If permohonan_kompensasi.Enabled = True Then
            permohonan_kompensasi.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub SuratKuasaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuratKuasaToolStripMenuItem.Click
        If suratkuasa.Enabled = True Then
            suratkuasa.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub SuratKeteranganAnakKandungToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuratKeteranganAnakKandungToolStripMenuItem.Click
        If sk_anakkandung.Enabled = True Then
            sk_anakkandung.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub BedaIdentitasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BedaIdentitasToolStripMenuItem.Click
        If skbedaidentitas.Enabled = True Then
            skbedaidentitas.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub DomisiliAsliToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DomisiliAsliToolStripMenuItem.Click
        If domisili_asli_.Enabled = True Then
            domisili_asli_.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub MengambilBPKBMotorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MengambilBPKBMotorToolStripMenuItem.Click
        If SKKM.Enabled = True Then
            SKKM.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub KSMToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KSMToolStripMenuItem.Click
        If domsili_ksm.Enabled = True Then
            domsili_ksm.Show()
            Me.Hide()
        End If

    End Sub
    Sub reset()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
    End Sub
    Sub word()
        Dim ObjAppWord As New word.Application
        Dim ObjDocWord As New word.Document
        Dim namafile As String

        ObjDocWord = ObjAppWord.Documents.Open("D:\APP\Document\SURAT PERMOHONAN PENDISTRIBUSIAN RASTRA.docx")
        ObjDocWord.Bookmarks("no_surat").Select()
        ObjAppWord.Selection.TypeText(TextBox1.Text)
        ObjDocWord.Bookmarks("nama").Select()
        ObjAppWord.Selection.TypeText(TextBox2.Text)
        ObjDocWord.Bookmarks("jabatan").Select()
        ObjAppWord.Selection.TypeText(TextBox3.Text)
        ObjDocWord.Bookmarks("alamat").Select()
        ObjAppWord.Selection.TypeText(TextBox4.Text)
        ObjDocWord.Bookmarks("peru").Select()
        ObjAppWord.Selection.TypeText(TextBox1.Text)
        ObjDocWord.Bookmarks("kg").Select()
        ObjAppWord.Selection.TypeText(TextBox2.Text)
        ObjDocWord.Bookmarks("kpm").Select()
        ObjAppWord.Selection.TypeText(TextBox3.Text)
        ObjDocWord.Bookmarks("tgls").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker2.Text)

        namafile = "D:\APP\NewDocument\Permohonan Pendistribusian\" & TextBox4.Text & "-" & DateTimePicker2.Text & ".docx"
        ObjDocWord.SaveAs(namafile)
        ObjDocWord.Close()
        ObjAppWord.Quit()
    End Sub
    Sub open()
        Dim namafile As String
        Dim word As New Microsoft.Office.Interop.Word.Application
        Dim doc As New Microsoft.Office.Interop.Word.Document
        namafile = "D:\APP\NewDocument\Permohonan Pendistribusian\" & TextBox4.Text & "-" & DateTimePicker2.Text & ".docx"
        doc = word.Documents.Open(namafile)
        doc.Activate()
    End Sub
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        open()
    End Sub
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Environment.Exit(10)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        word()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        reset()
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub SuratBalikNamaTanahToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuratBalikNamaTanahToolStripMenuItem.Click
        If sk_baliknamatanah.Enabled = True Then
            sk_baliknamatanah.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub CeraiMatiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CeraiMatiToolStripMenuItem.Click
        If sk_ceraimati.Enabled = True Then
            sk_ceraimati.Show()
            Me.Hide()
        End If
    End Sub
End Class