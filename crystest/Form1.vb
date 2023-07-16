Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Windows.Forms

Public Class Form1
    Dim dset As New DataSet
    Dim comm As New OleDbCommand
    Dim da As New OleDbDataAdapter(comm)
    Dim dt As New DataTable

    Dim strReportPath As String

    Private Sub populate()
        ' Loads the Updated Data when Operations was used.

        da = New OleDbDataAdapter("Select * from ppltable", conn)
        dset = New DataSet
        da.Fill(dset, "ppltable")

        dgrid.DataSource = dset.Tables("ppltable").DefaultView
    End Sub

    Private Sub verifyCR()
        strReportPath = "C:\Users\Roseliza Lacao\source\repos\crystest\crystest\CrystalReport1.rpt"

        If Not IO.File.Exists(strReportPath) Then
            MessageBox.Show("Unable to locate report file:" &
              vbCrLf & strReportPath)
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        connect()
        populate()
        verifyCR()

        Try
            ' Load the Crystal Report .rpt file and pass onto Datatable
            Dim cr As New ReportDocument

            cr.Load(strReportPath)
            cr.SetDataSource(dset.Tables("ppltable"))

            ' Set the CrystalReportViewer's appearance and set the ReportSource:
            CrystalReportViewer1.ShowRefreshButton = False
            CrystalReportViewer1.ShowCloseButton = False
            CrystalReportViewer1.ShowGroupTreeButton = False

            CrystalReportViewer1.ReportSource = cr

        Catch ex As Exception
            MessageBox.Show("An Error Occured.")
        Finally
            conn.Close()
        End Try
    End Sub
End Class
