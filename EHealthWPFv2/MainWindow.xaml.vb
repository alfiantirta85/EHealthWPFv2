Imports System.Collections.ObjectModel
Imports System.Globalization
Imports System.IO
Imports System.Text

Class MainWindow
    Private dataPasien As New ObservableCollection(Of Pasien)

    Public Class Pasien
        Public Property ID As String
        Public Property Nama As String
        Public Property TanggalLahir As Date
        Public Property JenisKelamin As String
        Public Property BeratBadan As Double
        Public Property TinggiBadan As Double
        Public Property Alamat As String
        Public Property Telepon As String
        Public Property Keluhan As String
        Public Property Diagnosa As String
        Public Property TanggalDaftar As Date

        Public ReadOnly Property TanggalLahirStr As String
            Get
                Return TanggalLahir.ToString("dd/MM/yyyy")
            End Get
        End Property

        Public ReadOnly Property TanggalDaftarStr As String
            Get
                Return TanggalDaftar.ToString("dd/MM/yyyy")
            End Get
        End Property

        Public ReadOnly Property UmurStr As String
            Get
                Return HitungUmur().ToString()
            End Get
        End Property

        Public ReadOnly Property BeratBadanStr As String
            Get
                Return BeratBadan.ToString()
            End Get
        End Property
        Public ReadOnly Property TinggiBadanStr As String
            Get
                Return TinggiBadan.ToString()
            End Get
        End Property

        Public Function HitungUmur() As String
            Dim umur As Integer = Date.Now.Year - TanggalLahir.Year
            Dim suffix As String
            If Date.Now < TanggalLahir.AddYears(umur) Then
                umur -= 1
            End If
            suffix = "Tahun"

            If umur = 0 Then
                umur =
                    (Today.Year - TanggalLahir.Year) * 12 + Today.Month - TanggalLahir.Month
                If Today.Day < TanggalLahir.Day Then
                    umur -= 1
                End If
                suffix = "Bulan"
            End If
            Return String.Join(" ", umur, suffix)
        End Function

        Public Sub New(
            id As String,
            nama As String,
            tglLahir As Date,
            jk As String,
            bb As Double,
            tb As Double,
            alamat As String,
            telp As String,
            keluhan As String,
            diagnosa As String,
            tglDaftar As Date
        )
            Me.ID = id
            Me.Nama = nama
            TanggalLahir = tglLahir
            JenisKelamin = jk
            TinggiBadan = tb
            BeratBadan = bb
            Me.Alamat = alamat
            Telepon = telp
            Me.Keluhan = keluhan
            Me.Diagnosa = diagnosa
            TanggalDaftar = tglDaftar
        End Sub

    End Class

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        dtpTanggalDaftar.SelectedDate = Nothing
        dtpTanggalLahir.DisplayDateEnd = Date.Today
        dtpTanggalLahir.SelectedDate = Nothing

        btnEdit.IsEnabled = False
        btnHapus.IsEnabled = False

        dgvPasien.ItemsSource = dataPasien

        ScrollViewer.SetHorizontalScrollBarVisibility(dgvPasien, ScrollBarVisibility.Auto)

        MuatDataDariFile()
        UpdateJumlahPasien()
    End Sub

    Function IsValidPhoneNumber(phone As String) As Boolean
        phone = phone.Trim()

        For Each ch As Char In phone
            If Not Char.IsDigit(ch) AndAlso
           ch <> "+"c AndAlso
           ch <> "-"c AndAlso
           ch <> "("c AndAlso
           ch <> ")"c Then
                Return False
            End If
        Next

        Return True
    End Function

    Private Function ValidasiInput(ByRef pesanError As String, ByVal isEdit As Boolean) As Boolean
        If dataPasien.Any(Function(p) p.ID = txtID.Text.Trim()) AndAlso Not isEdit Then
            pesanError = "ID Pasien sudah digunakan!"
            txtID.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtID.Text) Then
            pesanError = "ID Pasien harus diisi!"
            txtID.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtNama.Text) Then
            pesanError = "Nama pasien harus diisi!"
            txtNama.Focus()
            Return False
        End If

        If Not dtpTanggalLahir.SelectedDate.HasValue Then
            pesanError = "Tanggal lahir harus diisi!"
            dtpTanggalLahir.Focus()
            Return False
        End If

        If dtpTanggalLahir.SelectedDate.Value > Date.Today Then
            pesanError = "Tanggal lahir tidak valid! (Tanggal dari masa depan)"
            dtpTanggalLahir.Focus()
            Return False
        End If

        If cboJenisKelamin.SelectedIndex = -1 Then
            pesanError = "Jenis kelamin harus di isi!"
            cboJenisKelamin.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtBerat.Text) Then
            pesanError = "Berat badan Pasien harus diisi!"
            txtBerat.Focus()
            Return False
        End If

        For i As Integer = 0 To txtBerat.Text.Length - 1
            If Not Char.IsDigit(txtBerat.Text.Chars(i)) Then
                pesanError = "Berat badan tidak boleh diisi huruf!"
                txtBerat.Focus()
                Return False
            End If
        Next

        If String.IsNullOrWhiteSpace(txtTinggi.Text) Then
            pesanError = "Tinggi badan Pasien harus diisi!"
            txtTinggi.Focus()
            Return False
        End If

        For i As Integer = 0 To txtTinggi.Text.Length - 1
            If Not Char.IsDigit(txtTinggi.Text.Chars(i)) Then
                pesanError = "Tinggi badan tidak boleh diisi huruf!"
                txtTinggi.Focus()
                Return False
            End If
        Next

        If String.IsNullOrWhiteSpace(txtAlamat.Text) Then
            pesanError = "Alamat harus diisi!"
            txtAlamat.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtTelepon.Text) Then
            pesanError = "Nomor telepon harus diisi!"
            txtTelepon.Focus()
            Return False
        End If

        If Not IsValidPhoneNumber(txtTelepon.Text) Then
            pesanError = "Nomor telepon hanya boleh diisi Angka, Plus (+), Minus(-), dan Tanda kurung ()!"
            txtTelepon.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtKeluhan.Text) Then
            pesanError = "Keluhan harus diisi!"
            txtKeluhan.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtDiagnosa.Text) Then
            pesanError = "Diagnosa harus diisi!"
            txtDiagnosa.Focus()
            Return False
        End If

        If Not dtpTanggalDaftar.SelectedDate.HasValue Then
            pesanError = "Tanggal daftar harus diisi!"
            dtpTanggalDaftar.Focus()
            Return False
        End If

        Return True
    End Function

    Private Sub BtnTambah_Click(sender As Object, e As RoutedEventArgs) Handles btnTambah.Click
        Try
            Dim pesanError As String = ""

            If Not ValidasiInput(pesanError, False) Then
                MessageBox.Show(pesanError, "Validasi Error",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                Exit Sub
            End If

            Dim pasienBaru As New Pasien(
                txtID.Text.Trim(),
                txtNama.Text.Trim(),
                dtpTanggalLahir.SelectedDate.Value,
                CType(cboJenisKelamin.SelectedItem, ComboBoxItem).Content.ToString(),
                Double.Parse(txtBerat.Text),
                Double.Parse(txtTinggi.Text),
                txtAlamat.Text.Trim(),
                txtTelepon.Text.Trim(),
                txtKeluhan.Text,
                txtDiagnosa.Text.Trim(),
                dtpTanggalDaftar.SelectedDate.Value
            )

            dataPasien.Add(pasienBaru)

            MessageBox.Show("Data pasien berhasil ditambahkan!", "Sukses",
                MessageBoxButton.OK, MessageBoxImage.Information)

            SimpanDataKeFile()
            UpdateJumlahPasien()
            BersihkanForm()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As RoutedEventArgs) Handles btnEdit.Click
        Try
            If dgvPasien.SelectedItem Is Nothing Then
                MessageBox.Show("Pilih data yang akan diedit!", "Peringatan",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Dim pesanError As String = ""
            If Not ValidasiInput(pesanError, True) Then
                txtID.IsEnabled = False
                MessageBox.Show(pesanError, "Validasi Error",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                Exit Sub
            End If

            Dim selectedPatient As Pasien = CType(dgvPasien.SelectedItem, Pasien)

            With selectedPatient
                .Nama = txtNama.Text
                .TanggalLahir = dtpTanggalLahir.SelectedDate.Value
                .BeratBadan = Double.Parse(txtBerat.Text)
                .TinggiBadan = Double.Parse(txtTinggi.Text)
                .Alamat = txtAlamat.Text
                .Telepon = txtTelepon.Text
                .Keluhan = txtKeluhan.Text
                .Diagnosa = txtDiagnosa.Text
                .TanggalDaftar = dtpTanggalDaftar.SelectedDate.Value
            End With

            For Each item As ComboBoxItem In cboJenisKelamin.Items
                If item.Content.ToString() = cboJenisKelamin.SelectedItem.ToString() Then
                    selectedPatient.JenisKelamin = item.Content.ToString()
                    Exit For
                End If
            Next

            MessageBox.Show("Data pasien berhasil diperbarui!", "Sukses",
                MessageBoxButton.OK, MessageBoxImage.Information)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try

        SimpanDataKeFile()
        UpdateJumlahPasien()
        BersihkanForm()
        MuatDataDariFile()

        txtID.IsEnabled = True
    End Sub

    Private Sub BtnHapus_Click(sender As Object, e As RoutedEventArgs) Handles btnHapus.Click
        Try
            If dgvPasien.SelectedItem Is Nothing Then
                MessageBox.Show("Pilih data yang akan dihapus!", "Peringatan",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Dim hasil As MessageBoxResult = MessageBox.Show(
                "Apakah Anda yakin ingin menghapus data ini?",
                "Konfirmasi Hapus",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question)

            If hasil = MessageBoxResult.Yes Then
                Dim pasienDipilih As Pasien = CType(dgvPasien.SelectedItem, Pasien)
                dataPasien.Remove(pasienDipilih)

                SimpanDataKeFile()
                UpdateJumlahPasien()
                BersihkanForm()

                MessageBox.Show("Data berhasil dihapus!", "Sukses",
                              MessageBoxButton.OK, MessageBoxImage.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnCari_Click(sender As Object, e As RoutedEventArgs) Handles btnCari.Click
        Try
            Dim keyword As String = txtCari.Text.Trim().ToLower()

            If String.IsNullOrWhiteSpace(keyword) Then
                dgvPasien.ItemsSource = dataPasien
                UpdateJumlahPasien()
                Return
            End If

            Dim hasilCari = From p In dataPasien
                            Where p.ID.ToLower().Contains(keyword) OrElse
                                  p.Nama.ToLower().Contains(keyword) OrElse
                                  p.Diagnosa.ToLower().Contains(keyword)
                            Select p

            dgvPasien.ItemsSource = hasilCari.ToList()
            lblJumlah.Text = "Ditemukan: " & hasilCari.Count() & " data"

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnBersihkan_Click(sender As Object, e As RoutedEventArgs) Handles btnBersihkan.Click
        txtID.IsEnabled = True
        dgvPasien.UnselectAll()
        dgvPasien.UnselectAllCells()
        BersihkanForm()
        txtCari.Text = ""
        dgvPasien.ItemsSource = dataPasien
        UpdateJumlahPasien()
    End Sub

    Private Sub BtnCetak_Click(sender As Object, e As RoutedEventArgs) Handles btnCetak.Click
        Try
            Dim namaFile As String = "Laporan_Pasien_" &
                                    Date.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
            Dim direktori As String = AppDomain.CurrentDomain.BaseDirectory & "Laporan\"

            If Not Directory.Exists(direktori) Then
                Directory.CreateDirectory(direktori)
            End If

            Dim pathFile As String = Path.Combine(direktori, namaFile)

            Dim sb As New StringBuilder()
            sb.AppendLine("=" & New String("="c, 80))
            sb.AppendLine("            LAPORAN DATA PASIEN E-HEALTH")
            sb.AppendLine("=" & New String("="c, 80))
            sb.AppendLine("Tanggal Cetak: " & Date.Now.ToString("dd MMMM yyyy HH:mm:ss"))
            sb.AppendLine("Total Pasien: " & dataPasien.Count)
            sb.AppendLine("=" & New String("="c, 80))
            sb.AppendLine()


            Dim nomer As Integer = 1
            For Each pasien As Pasien In dataPasien
                sb.AppendLine("Pasien #" & nomer)
                sb.AppendLine(New String("-"c, 80))
                sb.AppendLine("ID pasien      : " & pasien.ID)
                sb.AppendLine("Nama           : " & pasien.Nama)
                sb.AppendLine("Tanggal lahir  : " & pasien.TanggalLahir.ToString("dd/MM/yyyy"))
                sb.AppendLine("Umur           : " & pasien.HitungUmur())
                sb.AppendLine("Jenis kelamin  : " & pasien.JenisKelamin)
                sb.AppendLine("Berat badan    : " & pasien.BeratBadanStr())
                sb.AppendLine("Tinggi badan   : " & pasien.TinggiBadanStr())
                sb.AppendLine("Alamat         : " & pasien.Alamat)
                sb.AppendLine("Telepon        : " & pasien.Telepon)
                sb.AppendLine("Keluhan        : " & pasien.Keluhan)
                sb.AppendLine("Diagnosa       : " & pasien.Diagnosa)
                sb.AppendLine("Tanggal daftar : " & pasien.TanggalDaftar.ToString("dd/MM/yyyy"))
                sb.AppendLine()
                nomer += 1
            Next

            sb.AppendLine("=" & New String("="c, 80))
            sb.AppendLine("Akhir Laporan")

            File.WriteAllText(pathFile, sb.ToString())

            MessageBox.Show("Laporan berhasil dibuat!" & vbCrLf &
                          "File: " & pathFile, "Sukses",
                          MessageBoxButton.OK, MessageBoxImage.Information)

            Process.Start("explorer.exe", direktori)

        Catch ex As Exception
            MessageBox.Show("Error membuat laporan: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub UpdateJumlahPasien()
        lblJumlah.Text = "Total Pasien: " & dataPasien.Count
    End Sub

    Private Sub BersihkanForm()
        txtID.Clear()
        txtNama.Clear()
        txtBerat.Clear()
        txtTinggi.Clear()
        txtAlamat.Clear()
        txtTelepon.Clear()
        txtKeluhan.Clear()
        txtDiagnosa.Clear()

        cboJenisKelamin.SelectedIndex = -1

        dtpTanggalLahir.SelectedDate = Nothing
        dtpTanggalDaftar.SelectedDate = Nothing

        txtID.IsEnabled = True
        btnTambah.Content = "Tambah Data"

        txtNama.Focus()
    End Sub

    Private Sub SimpanDataKeFile()
        Try
            Dim namaFile As String = "data_pasien.txt"
            Dim pathFile As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, namaFile)

            Dim sb As New StringBuilder()

            For Each p As Pasien In dataPasien
                sb.AppendLine(
                        String.Join("|",
                        p.ID,
                        p.Nama,
                        p.TanggalLahir.ToString("dd/MM/yyyy"),
                        p.JenisKelamin,
                        p.BeratBadan,
                        p.TinggiBadan,
                        p.Alamat,
                        p.Telepon,
                        p.Keluhan,
                        p.Diagnosa,
                        p.TanggalDaftar.ToString("dd/MM/yyyy")
                    ))
            Next

            File.WriteAllText(pathFile, sb.ToString())

        Catch ex As Exception
            MessageBox.Show("Error menyimpan data: " & ex.Message, "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub MuatDataDariFile()
        Try
            Dim namaFile As String = "data_pasien.txt"
            Dim pathFile As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, namaFile)

            If Not File.Exists(pathFile) Then
                Return
            End If

            Dim lines() As String = File.ReadAllLines(pathFile)

            dataPasien.Clear()

            For Each line As String In lines
                If Not String.IsNullOrWhiteSpace(line) Then
                    Dim data() As String = line.Split("|"c)

                    If data.Length <> 11 Then Continue For
                    If data.Length >= 11 Then
                        Dim pasien As New Pasien(
                            data(0),                                    ' ID
                            data(1),                                    ' Nama
                            Date.ParseExact(
                                data(2),
                                "dd/MM/yyyy",
                                CultureInfo.InvariantCulture),          ' TanggalLahir
                            data(3),                                    ' JenisKelamin
                            Double.Parse(data(4)),                      ' BeratBadan
                            Double.Parse(data(5)),                      ' TinggiBadan
                            data(6),                                    ' Alamat
                            data(7),                                    ' Telepon
                            data(8),                                    ' Keluhan
                            data(9),                                    ' Diagnosa
                            Date.ParseExact(
                                data(10),
                                "dd/MM/yyyy",
                                CultureInfo.InvariantCulture)           ' TanggalDaftar
                        )

                        dataPasien.Add(pasien)
                    End If
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error memuat data: " & ex.Message, "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub dgvPasien_SelectionChanged(sender As Object, e As EventArgs) Handles dgvPasien.SelectionChanged
        If dgvPasien.SelectedItem Is Nothing Then Exit Sub

        Dim dataPasien As Pasien = CType(dgvPasien.SelectedItem, Pasien)

        btnEdit.IsEnabled = True
        btnHapus.IsEnabled = True
        txtID.IsEnabled = False

        txtID.Text = dataPasien.ID
        txtNama.Text = dataPasien.Nama
        dtpTanggalLahir.SelectedDate = dataPasien.TanggalLahir
        txtBerat.Text = dataPasien.BeratBadan.ToString()
        txtTinggi.Text = dataPasien.TinggiBadan.ToString()
        txtAlamat.Text = dataPasien.Alamat
        txtTelepon.Text = dataPasien.Telepon.ToString()
        txtKeluhan.Text = dataPasien.Keluhan
        txtDiagnosa.Text = dataPasien.Diagnosa
        dtpTanggalDaftar.SelectedDate = dataPasien.TanggalDaftar

        For Each item As ComboBoxItem In cboJenisKelamin.Items
            If item.Content.ToString() = dataPasien.JenisKelamin Then
                cboJenisKelamin.SelectedItem = item
                Exit For
            End If
        Next

        MainTab.SelectedItem = Form_Tab
    End Sub
End Class