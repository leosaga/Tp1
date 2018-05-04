
Public Class Form1

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not Char.IsNumber(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "/" And Not e.KeyChar = "-" And Not e.KeyChar = "." And Not e.KeyChar = "," Then
            e.Handled = True
            Exit Sub

        End If

        Dim pos As Integer = TextBox1.SelectionStart
        If e.KeyChar = "/" And pos <> 2 And pos <> 5 Then
            e.Handled = True
            Exit Sub
        End If
        If Char.IsNumber(e.KeyChar) And (pos = 2 Or pos = 5) Then
            e.Handled = True
            Exit Sub
        End If
        If e.KeyChar = "." Or e.KeyChar = "," Or e.KeyChar = "-" Then
            SendKeys.Send("/")
            e.Handled = (True)
        End If
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus
        Dim dia As Integer = CInt(TextBox1.Text.Substring(0, 2))
        Dim mes As Integer = CInt(TextBox1.Text.Substring(3, 2))
        Dim año As Integer = CInt(TextBox1.Text.Substring(6, 4))

        'si el año es menor a 1980 o el año es mayor a 2018 esta mal
        If año < 1900 Or año > 2018 Then
            MsgBox("año no valido")
            TextBox1.Focus()
            Exit Sub
        End If

        If mes < 1 Or mes > 12 Then
            MsgBox("mes no valido")
            TextBox1.Focus()
            Exit Sub

        End If
        Select mes
            Case 1, 3, 5, 7, 8, 10, 12
                If dia < 1 Or dia > 31 Then
                    MsgBox("dia no valido")
                    TextBox1.Focus()
                    Exit Sub
                End If
            Case 4, 6, 9, 11
                If dia < 1 Or dia > 30 Then

                    MsgBox("dia no valido")
                    TextBox1.Focus()
                    Exit Sub
                End If
            Case 2
                If año Mod 4 = 0 Then
                    If dia < 1 Or dia > 29 Then
                        MsgBox("dia no valido")
                        TextBox1.Focus()
                    End If
                Else
                    If dia < 1 Or dia > 28 Then

                        MsgBox("dia no valido")
                        TextBox1.Focus()
                        Exit Sub
                    End If
                End If
        End Select
    End Sub
End Class

