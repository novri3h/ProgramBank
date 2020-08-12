Attribute VB_Name = "MdlBank"

Public Conn As New adodb.Connection
Public RSKasir As adodb.Recordset
Public RSAtm As adodb.Recordset
Public RSNasabah As adodb.Recordset
Public RSTagihan As adodb.Recordset
Public RSTransaksi As adodb.Recordset
Public Lokasi As String


Public Sub BukaDB()
Set Conn = New adodb.Connection
Set RSKasir = New adodb.Recordset
Set RSAtm = New adodb.Recordset
Set RSNasabah = New adodb.Recordset
Set RSTagihan = New adodb.Recordset
Set RSTransaksi = New adodb.Recordset
Lokasi = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ADOBank.mdb"
Conn.Open Lokasi
End Sub

Public Sub Auto()
Call BukaDB
RSTransaksi.Open "select * from Transaksi Where Notransaksi In(Select Max(Notransaksi)From Transaksi)Order By Notransaksi Desc", Conn
RSTransaksi.Requery
    Dim Urutan As String * 9
    Dim Hitung As Long
    Dim Nomor As Double
    With RSTransaksi
        If .EOF Then
            Urutan = Format(Date, "ymmdd") + "0001"
            Nomor = Urutan
        Else
            If Left(!NoTransaksi, 5) <> Format(Date, "ymmdd") Then
                Urutan = Format(Date, "ymmdd") + "0001"
            Else
                Hitung = (!NoTransaksi) + 1
                Urutan = Format(Date, "ymmdd") + Right("0000" & Hitung, 4)
            End If
        End If
        Nomor = Urutan
    End With
    Nasabah.Nomor = Nomor
    Ambil2.Nomor = Nomor
    AmbilAtm.Nomor = Nomor
    Pembayaran.Nomor = Nomor
    Pengambilan.Nomor = Nomor
    Setoran.Nomor = Nomor
    Transfer.Nomor = Nomor
End Sub
