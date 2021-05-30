Imports System.Data.OleDb
Imports System.Data

Public Class Form2
    Private Sub btn_register_Click(sender As Object, e As EventArgs) Handles btn_register.Click
        If txt_username.Text = "" Or txt_password.Text = "" Or txt_cpassword.Text = "" Then
            MsgBox("All fields are required")
        ElseIf txt_password.Text <> txt_cpassword.Text Then
            MsgBox("Password should be the same")
        Else
            Try
                Dim conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\inventory.accdb")
                Dim insert As String = "INSERT INTO users(username, password) VALUES('" & txt_username.Text & "','" & txt_password.Text & "');"
                Dim cmd As New OleDbCommand(insert, conn)
                conn.Open()
                cmd.ExecuteNonQuery()
                MsgBox("create success")
                Me.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub txt_cpassword_TextChanged(sender As Object, e As EventArgs)

    End Sub
End Class