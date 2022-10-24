Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Linq.Expressions
Imports System.Net.Security
Imports System.Runtime.CompilerServices

Public Class Form1

    Dim conexion As SqlConnection
    Dim comando As SqlCommand

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conexion = New SqlConnection("server=DESKTOP-91OTMS3\SQLEXPRESS; database=trabajo; integrated security=true")

        Me.ToolTip1.IsBalloon = True
        Me.ToolTip1.SetToolTip(btnguardar, "GUARDAR DATOS")
        Me.ToolTip1.SetToolTip(btneditar, "EDITAR DATOS")
        Me.ToolTip1.SetToolTip(btnactualizar, "ACTUALIZAR DATOS")
        Me.ToolTip1.SetToolTip(btneliminar, "ELIMINAR DATOS")
        Me.ToolTip1.SetToolTip(btnbuscar, "BUSCAR")
    End Sub

    Private Sub btnguardar_Click(sender As Object, e As EventArgs) Handles btnguardar.Click

        If textid.Text.Equals("") = False And textnombre.Text.Equals("") = False And textapellido.Text.Equals("") = False And textfecha.Text.Equals("") = False And textcelu.Text.Equals("") = False And textdirec.Text.Equals("") = False Then
            conexion.Open()
            Dim consulta As String = "insert into registro (Id, Nombre, Apellido, Fecha_nacimiento, Celular, Direccion, Telefono, Email) values(" & textid.Text & ",'" & textnombre.Text & "','" & textapellido.Text & "','" & textfecha.Text & "','" & textcelu.Text & "','" & textdirec.Text & "','" & texttele.Text & "','" & textemail.Text & "')"
            comando = New SqlCommand(consulta, conexion)
            comando.ExecuteNonQuery()
            MsgBox("Los datos se han ingresado correctamente")
            conexion.Close()
        Else
            MsgBox("No se puede guardar existen campos vacios")
        End If
        textid.Clear()
        textnombre.Clear()
        textapellido.Clear()
        textfecha.Clear()
        textcelu.Clear()
        textdirec.Clear()
        texttele.Clear()
        textemail.Clear()

        conexion.Close()

    End Sub

    Private Sub btnactualizar_Click(sender As Object, e As EventArgs) Handles btnactualizar.Click
        Try
            Dim mostrar As New SqlDataAdapter("select * from registro", conexion)
            Dim almacena As New DataSet

            mostrar.Fill(almacena)
            DataGridView1.DataSource = almacena.Tables(0)

        Catch ex As Exception
            MsgBox("Ha ocurrido en error")
        End Try
        conexion.Close()
    End Sub

    Private Sub btneditar_Click(sender As Object, e As EventArgs) Handles btneditar.Click
        Try
            conexion.Open()
            Dim consulta As String = "update registro set Nombre='" + textnombre.Text + "',Apellido='" + textapellido.Text + "',Fecha_nacimiento='" + textfecha.Text + "',Celular='" + textcelu.Text + "',Direccion='" + textdirec.Text + "',Telefono='" + texttele.Text + "',Email='" + textemail.Text + "' where Id='" + textid.Text + "'"
            comando = New SqlCommand(consulta, conexion)
            comando.ExecuteNonQuery()
            MsgBox("Los datos se han editado correctamente")
        Catch ex As Exception
            MsgBox("Los datos no se han editado correctamente")
        End Try

        textid.Clear()
        textnombre.Clear()
        textapellido.Clear()
        textfecha.Clear()
        textcelu.Clear()
        textdirec.Clear()
        texttele.Clear()
        textemail.Clear()

        conexion.Close()
    End Sub

    Private Sub btneliminar_Click(sender As Object, e As EventArgs) Handles btneliminar.Click
        Try
            conexion.Open()
            Dim consulta As String = "delete from registro where ID='" + textid.Text + "'"
            comando = New SqlCommand(consulta, conexion)
            comando.ExecuteNonQuery()
            MsgBox("Los datos se han eliminado correctamente")
        Catch ex As Exception
            MsgBox("Los datos no se han eliminado correctamente")
        End Try
        textid.Clear()
        textnombre.Clear()
        textapellido.Clear()
        textfecha.Clear()
        textcelu.Clear()
        textdirec.Clear()
        texttele.Clear()
        textemail.Clear()

        conexion.Close()
    End Sub

    Private Sub btnbuscar_Click(sender As Object, e As EventArgs) Handles btnbuscar.Click
        Try
            Dim mostrar As New SqlDataAdapter("select * from registro where Apellido='" & textbuscar.Text & "'", conexion)
            Dim almacena As New DataSet

            mostrar.Fill(almacena)
            DataGridView1.DataSource = almacena.Tables(0)
        Catch ex As Exception
            MsgBox("No se ha encontrado al alumno")
        End Try

        conexion.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer = DataGridView1.CurrentRow.Index
        textid.Text = DataGridView1.Item(0, i).Value
        textnombre.Text = DataGridView1.Item(1, i).Value
        textapellido.Text = DataGridView1.Item(2, i).Value
        textfecha.Text = DataGridView1.Item(3, i).Value
        textcelu.Text = DataGridView1.Item(4, i).Value
        textdirec.Text = DataGridView1.Item(5, i).Value
        texttele.Text = DataGridView1.Item(6, i).Value
        textemail.Text = DataGridView1.Item(7, i).Value


    End Sub

    Private Sub textid_MouseHover(sender As Object, e As EventArgs) Handles textid.MouseHover
        ToolTip2.SetToolTip(textid, "Ingrese 8 caracteres")
        ToolTip2.ToolTipTitle = "ID o DNI"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub textnombre_MouseHover(sender As Object, e As EventArgs) Handles textnombre.MouseHover
        ToolTip2.SetToolTip(textnombre, "Ingrese aqui sus nombres")
        ToolTip2.ToolTipTitle = "Nombres del alumno"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub textapellido_MouseHover(sender As Object, e As EventArgs) Handles textapellido.MouseHover
        ToolTip2.SetToolTip(textapellido, "Ingrese aqui sus apellidos")
        ToolTip2.ToolTipTitle = "Apellidos del alumno"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub textfecha_MouseHover(sender As Object, e As EventArgs) Handles textfecha.MouseHover
        ToolTip2.SetToolTip(textfecha, "Ingrese aqui dia/mes/año")
        ToolTip2.ToolTipTitle = "Fecha de nacimiento"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub textcelu_MouseHover(sender As Object, e As EventArgs) Handles textcelu.MouseHover
        ToolTip2.SetToolTip(textcelu, "Ingrese aqui su celular")
        ToolTip2.ToolTipTitle = "Numero de celular"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub textdirec_MouseHover(sender As Object, e As EventArgs) Handles textdirec.MouseHover
        ToolTip2.SetToolTip(textdirec, "Ingrese aqui su dirección")
        ToolTip2.ToolTipTitle = "Dirección"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub texttele_MouseHover(sender As Object, e As EventArgs) Handles texttele.MouseHover
        ToolTip2.SetToolTip(texttele, "Ingrese aqui su Teléfono")
        ToolTip2.ToolTipTitle = "Teléfono"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub textemail_MouseHover(sender As Object, e As EventArgs) Handles textemail.MouseHover
        ToolTip2.SetToolTip(textemail, "Ingrese aqui su correo electronico")
        ToolTip2.ToolTipTitle = "Email"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub textbuscar_MouseHover(sender As Object, e As EventArgs) Handles textbuscar.MouseHover
        ToolTip2.SetToolTip(textbuscar, "Ingrese aqui su apellidos")
        ToolTip2.ToolTipTitle = "Buscar Apellidos"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info
    End Sub
End Class
