Imports System.Data.Odbc
Module connection
    Public command As OdbcCommand
    Public d_set As New DataSet
    Public adapter As OdbcDataAdapter
    Public reader As OdbcDataReader
    Public tabel As New DataTable

    Public storage As String = "Driver={MySQL ODBC 3.51 Driver}; Database=data_kas; server=localhost; uid=root"
    Public CONN As OdbcConnection = New OdbcConnection(storage)

    Public Sub mulaikoneksi()
        If CONN.State = ConnectionState.Closed Then
            Try
                CONN.Open()
            Catch ex As Exception
                MsgBox("Koneksi Gagal" & ex.ToString)
            End Try
        End If
    End Sub

    Public Sub stopkoneksi()
        If CONN.State = ConnectionState.Open Then
            Try
                CONN.Close()
            Catch ex As Exception
                MsgBox("Gagal Menutup Koneksi" & ex.ToString)
            End Try
        End If
    End Sub
End Module
