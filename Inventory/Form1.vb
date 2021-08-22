Imports System.ComponentModel
Imports System.Data.OleDb

Public Class Form1
    Dim con As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\inventory\db\inventory.accdb")
    Dim current_role As String = ""
    Dim purchase_price As String = ""

    Sub DisplayData(ByVal table_name As String, ByVal data_grid_name As DataGridView, ByVal sort_item As String)
        Try
            Dim sql As String
            Dim cmd As New OleDb.OleDbCommand
            Dim dt As New DataTable
            Dim da As New OleDb.OleDbDataAdapter
            sql = "Select * from " + table_name + " ORDER BY " + sort_item
            cmd.Connection = con
            cmd.CommandText = sql
            da.SelectCommand = cmd

            da.Fill(dt)
            data_grid_name.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            data_grid_name.DataSource = dt
            data_grid_name.Columns(0).Visible = False
            If (table_name = "Transactions") Then
                data_grid_name.Columns(7).Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Sub DisplayDataByDate(ByVal table_name As String, ByVal data_grid_name As DataGridView)
        Try
            Dim sql As String
            Dim cmd As New OleDb.OleDbCommand
            Dim dt As New DataTable
            Dim da As New OleDb.OleDbDataAdapter
            'sql = "Select * from " + table_name
            sql = "SELECT * FROM " + table_name + " WHERE DATE_CREATED BETWEEN 
			#" & start_dt.Value.ToString("MM/dd/yyyy") & "# AND #" & end_dt.Value.ToString("MM/dd/yyyy") & "#"
            cmd.Connection = con
            cmd.CommandText = sql
            da.SelectCommand = cmd

            da.Fill(dt)
            data_grid_name.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            data_grid_name.DataSource = dt
            data_grid_name.Columns(0).Visible = False
            If (table_name = "Transactions") Then
                data_grid_name.Columns(7).Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Sub DisplayDataByType(ByVal table_name As String, ByVal type As String, ByVal data_grid_name As DataGridView)
        Try
            Dim sql As String
            Dim cmd As New OleDb.OleDbCommand
            Dim dt As New DataTable
            Dim da As New OleDb.OleDbDataAdapter
            sql = "Select * from " + table_name + " where TRANSACTION_TYPE = '" & type & "' and DATE_CREATED BETWEEN 
			#" & start_dt.Value.ToString("MM/dd/yyyy") & "# AND #" & end_dt.Value.ToString("MM/dd/yyyy") & "#"
            cmd.Connection = con
            cmd.CommandText = sql
            da.SelectCommand = cmd

            da.Fill(dt)
            data_grid_name.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            data_grid_name.DataSource = dt
            data_grid_name.Columns(0).Visible = False
            If (table_name = "Transactions") Then
                data_grid_name.Columns(7).Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Sub getPriceOfSelectedItem(ByVal product_name As String)
        Try
            con.Open()
            Dim cm As New OleDbCommand("Select * from Products where PRODUCT_NAME='" & product_name & "'", con)
            Dim dr As OleDbDataReader = cm.ExecuteReader
            While dr.Read
                If cb_transaction.SelectedItem = "Sale" Then
                    txt_rate.Text = dr("SALE_PRICE").ToString
                    purchase_price = dr("PURCHASE_PRICE").ToString
                ElseIf cb_transaction.SelectedItem = "Purchase" Then
                    txt_rate.Text = dr("PURCHASE_PRICE").ToString
                Else
                    MsgBox("Please select transaction type")
                    txt_rate.Text = ""
                End If

            End While
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Sub InsertProductInStock(ByVal product As String)
        Try
            Dim sql As String
            Dim cmd As New OleDb.OleDbCommand
            Dim qty As Decimal = 0
            con.Open()
            sql = "INSERT INTO Stock (PRODUCT, QUANTITY) VALUES ('" & product & "', '" & qty & "');"
            cmd.Connection = con
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Function checkProductQuantity(ByVal product As String, ByVal qty As Integer, ByVal transaction As String)
        Dim cmd As OleDbCommand
        Dim sql = "Select * FROM stock WHERE PRODUCT ='" & product & "'"
        Dim q As Integer
        cmd = New OleDbCommand(sql, con)
        con.Open()

        Dim sdr As OleDbDataReader = cmd.ExecuteReader()
        If (sdr.Read() = True) Then
            q = sdr("QUANTITY")
            con.Close()
        End If
        If transaction = "Sale" Then
            If q >= qty Then
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If
    End Function

    Function updateProductInStocks(ByVal product As String, ByVal tr_type As String, ByVal qty As String, ByVal q As String)
        Dim sql As String
        Dim quantity As Decimal
        Dim i As Integer
        Dim cmd As New OleDb.OleDbCommand

        If tr_type = "Sale" Then
            quantity = CInt(q) - CInt(qty)
        ElseIf tr_type = "Purchase" Then
            quantity = CInt(q) + CInt(qty)
        End If
        sql = "UPDATE Stock SET QUANTITY = '" & quantity & "' WHERE PRODUCT ='" & product & "'"
        cmd.Connection = con
        cmd.CommandText = sql
        i = cmd.ExecuteNonQuery()
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Function updateProductName(ByVal product As String)
        Dim sql As String
        Dim i As Integer
        Dim cmd As New OleDb.OleDbCommand

        sql = "UPDATE Stock SET PRODUCT = '" & product & "' WHERE PRODUCT ='" & product_change.Text & "'"
        cmd.Connection = con
        cmd.CommandText = sql
        i = cmd.ExecuteNonQuery()
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Function DeleteProductInStock()
        Dim sql As String
        Dim i As Integer
        Dim cmd As New OleDb.OleDbCommand

        sql = "DELETE * FROM Stock WHERE PRODUCT ='" & product_change.Text & "'"
        cmd.Connection = con
        cmd.CommandText = sql
        i = cmd.ExecuteNonQuery()
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Function updateProductInStockByUpdatingTrans(ByVal product As String, ByVal quantity As String)
        Dim sql As String
        Dim i As Integer
        Dim cmd As New OleDb.OleDbCommand
        con.Open()
        sql = "UPDATE Stock SET QUANTITY = '" & Val(quantity) & "' WHERE PRODUCT ='" & product & "'"
        cmd.Connection = con
        cmd.CommandText = sql
        i = cmd.ExecuteNonQuery()
        If i > 0 Then
            Return True
        Else
            Return False
        End If
        con.Close()
    End Function

    Sub updateProductInStock(ByVal product As String, ByVal tr_type As String, ByVal qty As String)
        Try
            Dim cmd As OleDbCommand

            Dim sql = "Select * FROM Stock WHERE PRODUCT ='" & product & "'"
            cmd = New OleDbCommand(sql, con)
            con.Open()
            Dim sdr As OleDbDataReader = cmd.ExecuteReader()

            If (sdr.Read() = True) Then
                updateProductInStocks(product, tr_type, qty, CInt(sdr("QUANTITY")))
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Function checkDuplicate(ByVal product As String)
        Dim sqlQRY As String = "SELECT COUNT(*) AS numRows FROM Products WHERE PRODUCT_NAME = '" & product & "'"
        Dim queryResult As Integer

        con.Open()

        Dim com As New OleDb.OleDbCommand(sqlQRY, con)
        queryResult = com.ExecuteScalar()

        con.Close()

        If queryResult > 0 Then
            MsgBox("Product is already exist")
            Return False
        Else
            Return True
        End If
    End Function

    Sub getTotalQuantity()
        Try
            Dim sqlQRY As String = "SELECT SUM(QUANTITY) FROM Stock"
            Dim queryResult As Integer

            con.Open()

            Dim com As New OleDb.OleDbCommand(sqlQRY, con)
            queryResult = com.ExecuteScalar()
            'MsgBox(queryResult)
            tb_quantity.Text = queryResult
        Catch ex As Exception
            'MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Sub Print_Profit()
        Try
            Dim sqlQRY As String = "SELECT SUM(PROFIT) FROM Transactions WHERE DATE_CREATED BETWEEN 
			#" & start_dt.Value.ToString("MM/dd/yyyy") & "# AND #" & end_dt.Value.ToString("MM/dd/yyyy") & "#"
            Dim queryResult As Integer

            con.Open()

            Dim com As New OleDb.OleDbCommand(sqlQRY, con)
            queryResult = com.ExecuteScalar()
            tb_profit.Text = queryResult
        Catch ex As Exception
            'MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Sub Print_Total_Stock()
        Try
            Dim sqlQRY As String = "select sum(Stock.quantity*Products.SALE_PRICE) from Products inner join" +
                " Stock on Products.PRODUCT_NAME = Stock.PRODUCT"
            Dim queryResult As Integer

            con.Open()

            Dim com As New OleDb.OleDbCommand(sqlQRY, con)
            queryResult = com.ExecuteScalar()
            tb_amount.Text = queryResult
        Catch ex As Exception
            'MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Sub Print_Sale()
        Try
            Dim sqlQRY As String = "SELECT SUM(TOTAL) FROM Transactions WHERE TRANSACTION_TYPE='Sale' AND DATE_CREATED BETWEEN 
			#" & start_dt.Value.ToString("MM/dd/yyyy") & "# AND #" & end_dt.Value.ToString("MM/dd/yyyy") & "#"
            Dim queryResult As Integer

            con.Open()

            Dim com As New OleDb.OleDbCommand(sqlQRY, con)
            queryResult = com.ExecuteScalar()
            tb_sale.Text = queryResult
        Catch ex As Exception
            'MsgBox(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub

    Sub Print_Purchase()
        Try
            Dim sqlQRY As String = "SELECT SUM(TOTAL) FROM Transactions WHERE TRANSACTION_TYPE='Purchase' AND DATE_CREATED BETWEEN 
			#" & start_dt.Value.ToString("MM/dd/yyyy") & "# AND #" & end_dt.Value.ToString("MM/dd/yyyy") & "#"
            Dim queryResult As Integer

            con.Open()

            Dim com As New OleDb.OleDbCommand(sqlQRY, con)
            queryResult = com.ExecuteScalar()
            tb_purchase.Text = queryResult
        Catch ex As Exception
            'MsgBox(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub

    Sub createAutocomplete(ByVal au_col As String, ByVal au_table As String, ByVal au_field As TextBox)
        'Dim cm As New OleDbCommand("Select PRODUCT_NAME from Products", con)
        Dim cm As New OleDbCommand("Select " + au_col + " from " + au_table + "", con)
        Dim da As New OleDbDataAdapter(cm)
        Dim dt As New DataSet
        da.Fill(dt)
        Dim column1 As New AutoCompleteStringCollection
        Dim i As Integer
        For i = 0 To dt.Tables(0).Rows.Count - 1
            column1.Add(dt.Tables(0).Rows(i)(au_col).ToString())
        Next
        au_field.AutoCompleteSource = AutoCompleteSource.CustomSource
        au_field.AutoCompleteCustomSource = column1
        au_field.AutoCompleteMode = AutoCompleteMode.Suggest
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        pn_login.Visible = True
        pn_register.Visible = False
        pn_users.Visible = False
        pn_help.Visible = False

        'Call DisplayDataByDate("Transactions", DataGridView2)
        Call DisplayData("Products", DataGridView1, "PRODUCT_NAME")
        'Call DisplayData("Stock", DataGridView3, "PRODUCT")

        cb_transaction.Items.Clear()
        cb_transaction.Items.Add("Sale")
        cb_transaction.Items.Add("Purchase")

        Call createAutocomplete("PRODUCT_NAME", "Products", all_product)

        Dim today As System.DateTime
        Dim answer As System.DateTime

        today = System.DateTime.Now
        'answer = today.AddDays(30)
        'answer = today.AddMinutes(2)
        answer = #09/18/2021 00:00:00#
        'MsgBox("Today: " + today + "\nAnswer: " + answer)
        Dim result As Integer = DateTime.Compare(today, answer)
        'MsgBox(answer.ToString("MM\/dd\/yyyy"))

        Dim regDate As Date = Date.Now()
        Dim strDate As String = regDate.ToString("MM\/dd\/yyyy")
        'MsgBox(strDate)
        If result > 0 Then
            pn_login.Visible = False
            pn_register.Visible = False
            pn_users.Visible = False
            pn_help.Visible = False
            pn_expired.Visible = True
            pn_expense.Visible = False
            pn_transaction.Visible = False
            MenuStrip1.Visible = False
        End If
    End Sub

    Private Sub btn_back_Click(sender As Object, e As EventArgs) Handles btn_back.Click
        pn_login.Visible = True
        pn_register.Visible = False
    End Sub

    Private Sub btn_register_Click(sender As Object, e As EventArgs) Handles btn_register.Click
        If txt_cpassword.Text = "" Or txt_username.Text = "" Or txt_password.Text = "" Then
            MsgBox("Please all fields are required", MsgBoxStyle.Exclamation)
        ElseIf txt_password.Text <> txt_cpassword.Text Then
            MsgBox("The password should be the same", MsgBoxStyle.Exclamation)
        Else
            Try
                Dim sql As String
                Dim cmd As New OleDb.OleDbCommand
                Dim i As Integer
                con.Open()
                sql = "INSERT INTO users (username,[password]) VALUES ('" & txt_username.Text & "', '" & txt_password.Text & "');"
                cmd.Connection = con
                cmd.CommandText = sql
                i = cmd.ExecuteNonQuery
                If i > 0 Then
                    MsgBox("New record has been inserted successfully!")
                    ''clear fields
                    txt_password.Text = ""
                    txt_cpassword.Text = ""
                    txt_username.Text = ""
                Else
                    MsgBox("No record has been inserted successfully!")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If txt_user.Text = "" Or txt_pass.Text = "" Then
            MsgBox("Please all fields are required")
        Else
            Try
                Dim cmd As OleDbCommand
                Dim sql = "Select * FROM users WHERE username ='" & txt_user.Text & "' AND password ='" & txt_pass.Text & "'"
                cmd = New OleDbCommand(sql, con)
                con.Open()

                Dim sdr As OleDbDataReader = cmd.ExecuteReader()
                If (sdr.Read() = True) Then
                    current_role = sdr("role")
                    If (sdr("role") = "user") Then
                        btn_add.Visible = False
                        btn_update.Visible = False
                        btn_delete.Visible = False
                        update_trans_btn.Visible = False
                    End If
                    txt_user.Text = ""
                    txt_pass.Text = ""
                    pn_login.Visible = False
                    pn_users.Visible = True
                    pn_register.Visible = False
                Else
                    MessageBox.Show("Invalid username or password!")
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()
            End Try
        End If
    End Sub

    Private Sub btn_add_Click(sender As Object, e As EventArgs) Handles btn_add.Click

        If txt_name.Text = "" Or txt_purchase.Text = "" Or txt_sale.Text = "" Then
            MsgBox("Please all fields are required", MsgBoxStyle.Exclamation)
        ElseIf checkDuplicate(txt_name.Text) Then

            Try
                Dim sql As String
                Dim cmd As New OleDb.OleDbCommand
                Dim i As Integer
                con.Open()
                sql = "INSERT INTO Products (PRODUCT_NAME, PURCHASE_PRICE, SALE_PRICE) VALUES ('" & txt_name.Text & "', '" & txt_purchase.Text & "', '" & txt_sale.Text & "');"
                cmd.Connection = con
                cmd.CommandText = sql
                i = cmd.ExecuteNonQuery
                If i > 0 Then
                    MsgBox("New product has been inserted successfully!", MsgBoxStyle.Information)
                    Call DisplayData("Products", DataGridView1, "PRODUCT_NAME")
                    Call InsertProductInStock(txt_name.Text)
                    ''clear fields
                    txt_name.Text = ""
                    txt_purchase.Text = ""
                    txt_sale.Text = ""
                    Call DisplayData("Stock", DataGridView3, "PRODUCT")
                Else
                    MsgBox("No product has been inserted successfully!")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()

            End Try
        End If
    End Sub

    Private Sub HelpToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        pn_login.Visible = False
        pn_register.Visible = False
        pn_users.Visible = False
        pn_help.Visible = True
    End Sub

    Private Sub ProductToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ProductToolStripMenuItem.Click
        pn_login.Visible = False
        pn_register.Visible = False
        pn_users.Visible = True
        pn_help.Visible = False
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        txt_id.Text = DataGridView1.CurrentRow.Cells(0).Value
        txt_name.Text = DataGridView1.CurrentRow.Cells(1).Value
        txt_purchase.Text = DataGridView1.CurrentRow.Cells(2).Value
        txt_sale.Text = DataGridView1.CurrentRow.Cells(3).Value
        product_change.Text = DataGridView1.CurrentRow.Cells(1).Value
    End Sub

    Private Sub btn_delete_Click(sender As Object, e As EventArgs) Handles btn_delete.Click
        If txt_id.Text = "" Then
            MsgBox("Please select item you want to delete!")
        Else
            Dim n As String = MsgBox("Do you really want to delete the product?", MsgBoxStyle.YesNo)
            If n = vbYes Then
                Try
                    Dim sql As String
                    Dim i As Integer
                    Dim cmd As New OleDb.OleDbCommand
                    con.Open()
                    sql = "Delete * from Products WHERE ID=" & Val(txt_id.Text) & ""
                    cmd.Connection = con
                    cmd.CommandText = sql

                    i = cmd.ExecuteNonQuery
                    If i > 0 Then
                        Call DeleteProductInStock()
                        MsgBox("Product has been deleted successfully!")
                        Call DisplayData("Products", DataGridView1, "PRODUCT_NAME")
                        ''clear fields
                        txt_name.Text = ""
                        txt_purchase.Text = ""
                        txt_sale.Text = ""
                        txt_id.Text = ""
                    Else
                        MsgBox("No product has been deleted!")
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message)
                Finally
                    con.Close()

                End Try
            Else
                txt_name.Text = ""
                txt_purchase.Text = ""
                txt_sale.Text = ""
                txt_id.Text = ""
            End If

        End If

    End Sub

    Private Sub txt_sale_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_sale.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txt_purchase_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_purchase.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txt_quantity_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_quantity.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub btn_update_Click(sender As Object, e As EventArgs) Handles btn_update.Click
        If txt_name.Text = "" Or txt_purchase.Text = "" Or txt_sale.Text = "" Then
            MsgBox("Please select product you want to edit!")
        Else
            Try
                Dim sql As String
                Dim i As Integer
                Dim cmd As New OleDb.OleDbCommand
                con.Open()
                sql = "UPDATE Products SET PRODUCT_NAME='" & txt_name.Text & "', PURCHASE_PRICE='" & txt_purchase.Text & "', " &
             " SALE_PRICE='" & txt_sale.Text & "' WHERE ID=" & Val(txt_id.Text) & ""
                cmd.Connection = con
                cmd.CommandText = sql

                i = cmd.ExecuteNonQuery
                If i > 0 Then
                    Call updateProductName(txt_name.Text)
                    MsgBox("Product has been updated successfully!")
                    Call DisplayData("Products", DataGridView1, "PRODUCT_NAME")
                    txt_name.Text = ""
                    txt_purchase.Text = ""
                    txt_sale.Text = ""
                    txt_id.Text = ""
                    product_change.Text = ""
                Else
                    MsgBox("No product has been UPDATED!")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()
            End Try
        End If
    End Sub

    Private Sub btn_logout_Click(sender As Object, e As EventArgs)
        Dim n As String = MsgBox("Do you really want to exit?", MsgBoxStyle.YesNo)
        If n = vbYes Then
            pn_login.Visible = True

        End If
        pn_login.Visible = True
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogoutToolStripMenuItem.Click
        Dim n As String = MsgBox("Do you really want to exit?", MsgBoxStyle.YesNo)
        If n = vbYes Then
            current_role = ""
            Me.Close()
        End If
    End Sub

    Private Sub TransactionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransactionToolStripMenuItem.Click
        pn_transaction.Visible = True
        pn_login.Visible = False
        pn_register.Visible = False
        pn_users.Visible = False
        pn_help.Visible = False

        all_product.Text = ""
        cb_transaction.Text = ""
        txt_quantity.Text = ""
        txt_rate.Text = ""

        'Call addProductsToComboBox()
        Call getTotalQuantity()
        Call Print_Profit()
        Call Print_Sale()
        Call Print_Purchase()
        Call Print_Total_Stock()

        Call DisplayDataByDate("Transactions", DataGridView2)
        Call DisplayData("Stock", DataGridView3, "PRODUCT")
        Call createAutocomplete("PRODUCT_NAME", "Products", all_product)
    End Sub

    Private Sub cb_product_SelectedValueChanged(sender As Object, e As EventArgs) Handles cb_product.SelectedValueChanged
        getPriceOfSelectedItem(cb_product.SelectedItem)
    End Sub

    Private Sub cb_transaction_SelectedValueChanged(sender As Object, e As EventArgs) Handles cb_transaction.SelectedValueChanged
        all_product.Text = ""
    End Sub

    Private Sub add_trans_btn_Click(sender As Object, e As EventArgs) Handles add_trans_btn.Click
        Dim Amt As Decimal
        Dim Profit As Decimal
        If txt_quantity.Text <> "" Or txt_rate.Text <> "" Or all_product.Text <> "" Or cb_transaction.Text <> "" Then

            Try
                If checkProductQuantity(all_product.Text, Val(txt_quantity.Text), cb_transaction.Text) Then

                    Amt = CDec(txt_rate.Text) * CDec(txt_quantity.Text)
                    If (cb_transaction.Text = "Sale") Then
                        Profit = Amt - (CDec(purchase_price) * CDec(txt_quantity.Text))
                    Else
                        Profit = Amt * 0
                    End If
                    Dim n As String = MsgBox("Are you sure you want to add this transaction?", MsgBoxStyle.YesNo)


                    If n = vbYes Then
                        'Try
                        Dim sql As String
                        Dim cmd As New OleDb.OleDbCommand
                        Dim i As Integer
                        con.Open()
                        sql = "INSERT INTO Transactions (PRODUCT, TRANSACTION_TYPE, QUANTITY, RATE, TOTAL, DATE_CREATED, PROFIT)
						VALUES ('" & all_product.Text & "', '" & cb_transaction.Text & "', '" & txt_quantity.Text & "', '" & txt_rate.Text & "', '" & Amt & "', '" & transaction_date.Value.ToShortDateString & "', '" & Profit & "');"
                        cmd.Connection = con
                        cmd.CommandText = sql
                        i = cmd.ExecuteNonQuery
                        If i > 0 Then
                            MsgBox("New Transaction has been added successfully!")
                            Call DisplayDataByDate("Transactions", DataGridView2)
                            Call updateProductInStock(all_product.Text, cb_transaction.Text, txt_quantity.Text)
                            Call DisplayData("Stock", DataGridView3, "PRODUCT")
                            Call getTotalQuantity()
                            Call Print_Profit()
                            Call Print_Sale()
                            Call Print_Purchase()
                            Call Print_Total_Stock()
                            ''clear fields
                            all_product.Text = ""
                            cb_transaction.Text = ""
                            txt_quantity.Text = ""
                            txt_rate.Text = ""
                        Else
                            MsgBox("No Transaction has been added successfully!")
                        End If

                    Else
                        txt_quantity.Text = ""
                    End If
                Else
                    con.Close()
                    MsgBox("Please, purchase <" + all_product.Text + "> before sale.")
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()
            End Try
        Else
            MsgBox("Please all fields are required")
        End If
    End Sub

    Private Sub radio_btn_all_Click(sender As Object, e As EventArgs) Handles radio_btn_all.Click
        Call DisplayDataByDate("Transactions", DataGridView2)
    End Sub

    Private Sub radio_btn_purchase_Click(sender As Object, e As EventArgs) Handles radio_btn_purchase.Click
        Call DisplayDataByType("Transactions", "Purchase", DataGridView2)
    End Sub

    Private Sub radio_btn_sale_Click(sender As Object, e As EventArgs) Handles radio_btn_sale.Click
        Call DisplayDataByType("Transactions", "Sale", DataGridView2)
    End Sub

    Private Sub btn_filter_Click(sender As Object, e As EventArgs) Handles btn_filter.Click
        Call Print_Profit()
        Call Print_Sale()
        Call Print_Purchase()
        Call DisplayDataByDate("Transactions", DataGridView2)
        'MsgBox(start_dt.Value.ToString("MM/dd/yyyy"))
    End Sub

    Private Sub update_trans_btn_Click(sender As Object, e As EventArgs) Handles update_trans_btn.Click
        Dim Amt As Decimal
        Dim Profit As Decimal
        Dim sql As String
        Dim i As Integer
        Dim cmd As New OleDb.OleDbCommand
        If txt_quantity.Text = "" Or txt_rate.Text = "" Or all_product.Text = "" Or cb_transaction.Text = "" Then
            MsgBox("Please all fields are required", MsgBoxStyle.Exclamation)
        Else
            Amt = CDec(txt_rate.Text) * CDec(txt_quantity.Text)
            If (cb_transaction.Text = "Sale") Then
                Profit = Amt - (CDec(purchase_price) * CDec(txt_quantity.Text))
            Else
                Profit = Amt * 0
            End If
            Dim n As String = MsgBox("Are you sure you want to replace the existing quantity?", MsgBoxStyle.YesNo)
            If n = vbYes Then
                Try
                    con.Open()
                    sql = "UPDATE Transactions SET PRODUCT='" & all_product.Text & "', TRANSACTION_TYPE='" & cb_transaction.Text & "', " &
                 "QUANTITY='" & txt_quantity.Text & "', RATE='" & txt_rate.Text & "', TOTAL='" & Amt & "', " &
                 "DATE_CREATED='" & transaction_date.Value.ToString("MM/dd/yyyy") & "', PROFIT='" & Profit & "' WHERE ID=" & Val(tb_id.Text) & ""
                    cmd.Connection = con
                    cmd.CommandText = sql

                    i = cmd.ExecuteNonQuery
                    MsgBox("Product has been updated successfully!")
                    Call DisplayDataByDate("Transactions", DataGridView2)
                    Call updateProductInStockByUpdatingTrans(all_product.Text, txt_quantity.Text)
                    Call DisplayData("Stock", DataGridView3, "PRODUCT")
                    Call getTotalQuantity()
                    Call Print_Profit()
                    Call Print_Sale()
                    Call Print_Purchase()
                    Call Print_Total_Stock()
                    ''clear fields
                    all_product.Text = ""
                    cb_transaction.Text = ""
                    txt_quantity.Text = ""
                    txt_rate.Text = ""
                    'If i > 0 Then

                    'Else
                    'MsgBox("No product has been updated!")
                    'End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                Finally
                    con.Close()
                End Try
            Else
                all_product.Text = ""
                cb_transaction.Text = ""
                txt_quantity.Text = ""
                txt_rate.Text = ""
            End If
        End If
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If current_role = "admin" Then
            tb_id.Text = DataGridView2.CurrentRow.Cells(0).Value
            all_product.Text = DataGridView2.CurrentRow.Cells(1).Value
            cb_transaction.Text = DataGridView2.CurrentRow.Cells(2).Value
            txt_quantity.Text = DataGridView2.CurrentRow.Cells(3).Value
            txt_rate.Text = DataGridView2.CurrentRow.Cells(4).Value
        End If
    End Sub

    Private Sub all_product_TextChanged(sender As Object, e As EventArgs) Handles all_product.TextChanged
        getPriceOfSelectedItem(all_product.Text)
    End Sub

    Private Sub ExpenseToolStripMenuItem_Click(sender As Object, e As EventArgs)
        pn_login.Visible = False
        pn_register.Visible = False
        pn_users.Visible = False
        pn_help.Visible = False
    End Sub

    Sub searchProduct(ByVal table_name As String, ByVal data_grid_name As DataGridView, ByVal product_name As String, ByVal product_field As String)
        Try
            Dim sql As String
            Dim cmd As New OleDb.OleDbCommand
            Dim dt As New DataTable
            Dim da As New OleDb.OleDbDataAdapter
            sql = "select * from " + table_name + " where " + product_field + " like '%" & product_name & "%' order by " + product_field
            cmd.Connection = con
            cmd.CommandText = sql
            da.SelectCommand = cmd
            dt.Clear()

            da.Fill(dt)
            data_grid_name.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            data_grid_name.DataSource = dt
            'data_grid_name.Columns(0).Visible = False
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub stock_search_txt_TextChanged(sender As Object, e As EventArgs) Handles stock_search_txt.TextChanged
        Call searchProduct("Stock", DataGridView3, stock_search_txt.Text, "PRODUCT")
    End Sub

    Private Sub stock_search_txt_GotFocus(sender As Object, e As EventArgs) Handles stock_search_txt.GotFocus
        If stock_search_txt.Text = "Search..." Then
            stock_search_txt.Text = ""
            Call DisplayData("Stock", DataGridView3, "PRODUCT")
        End If
    End Sub

    Private Sub stock_search_txt_LostFocus(sender As Object, e As EventArgs) Handles stock_search_txt.LostFocus
        If stock_search_txt.Text = "" Then
            stock_search_txt.Text = "Search..."
            Call DisplayData("Stock", DataGridView3, "PRODUCT")
        End If
    End Sub

    Private Sub product_search_txt_TextChanged(sender As Object, e As EventArgs) Handles product_search_txt.TextChanged
        Call searchProduct("Products", DataGridView1, product_search_txt.Text, "PRODUCT_NAME")
    End Sub

    Private Sub product_search_txt_GotFocus(sender As Object, e As EventArgs) Handles product_search_txt.GotFocus
        If product_search_txt.Text = "search" Then
            product_search_txt.Text = ""
            Call DisplayData("Products", DataGridView1, "PRODUCT_NAME")
        End If
    End Sub

    Private Sub product_search_txt_LostFocus(sender As Object, e As EventArgs) Handles product_search_txt.LostFocus
        If product_search_txt.Text = "" Then
            product_search_txt.Text = "search"
            Call DisplayData("Products", DataGridView1, "PRODUCT_NAME")
        End If
    End Sub

End Class
