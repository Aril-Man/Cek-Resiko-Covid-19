Imports System.Data
Imports System.Data.OleDb
Module Module1
    Public con As OleDbConnection
    Public cmd As OleDbCommand
    Public DS As New DataSet
    Public DA As OleDbDataAdapter
    Public RD As OleDbDataReader
    Public lokasi As String
    Public Sub konek()
        lokasi = "provider=microsoft.jet.oledb.4.0;data source=Data_covid.mdb"
        con = New OleDbConnection(lokasi)
        If con.State = ConnectionState.Closed Then
            con.Open()
        End If
    End Sub
End Module
