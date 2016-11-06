Public Class Frm1
    Dim Tablero(6, 6) As Integer
    Dim bomb As New ArrayList
    Dim pbMain As New PictureBox
    Dim pbList As New ArrayList
    Dim flags As Integer
    Dim count_win As Integer
    Dim Tabpb(6, 6) As PictureBox
    Dim clicks As Integer = 0
    ' Tablero: 0 sin tocar, 1 con bomba, 2 con bandera, 3 descubierto
    ' Pbs Tag: 0 sin tocar, 1 con Bomba, 3 descubierto
    ' Pbs Text: 2 con bandera
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i, j, num1, num2 As Integer
        'Dim lblcf As New Label

        ' Asignar bombas
        Randomize()
        For x = 0 To 6
            Do
                num1 = Int(7 * Rnd())
                num2 = Int(7 * Rnd())
            Loop Until Tablero(num1, num2) <> 1
            Tablero(num1, num2) = 1
            ' MsgBox(num1 & num2)
        Next
        ' finaliza asignar bombas, con este código introduzco las bombas en el tablero (array)

        'lblcf.Size = New Drawing.Size(35, 35)
        'lblcf.Font = Label.DefaultFont

        ' con este código dibujo la carita smile de arriba
        pbMain.Size = New Drawing.Size(35, 35)
        pbMain.Location = New Drawing.Point(108, 5)
        pbMain.Name = "pbMain"
        pbMain.BorderStyle = BorderStyle.FixedSingle
        pbMain.ImageLocation = Application.StartupPath & "/img/smile.png"
        pbMain.SizeMode = PictureBoxSizeMode.StretchImage
        pbMain.Tag = 0
        Me.Controls.Add(pbMain) ' aquí lo añado a los controles
        AddHandler pbMain.MouseUp, AddressOf smile_MouseClick ' aquí le asigno comportamiento

        ' aquí dibujo el tablero
        For i = 0 To 6 ' la i es vertical
            For j = 0 To 6 'la j sería el horizontal
                Dim pb As New PictureBox
                pb.Size = New Drawing.Size(35, 35)
                pb.Location = New Drawing.Point(3 + (35 * j), 50 + (35 * i))
                pb.Name = "pbCelda" & i & j
                pb.BorderStyle = BorderStyle.FixedSingle
                pb.ImageLocation = Application.StartupPath & "/img/boton.png"
                pb.SizeMode = PictureBoxSizeMode.StretchImage
                pb.Tag = 0 ' todas las casillas están a 0 inicialmente
                ' pero con este if pongo las bombas a 1 en los picture box
                If Tablero(i, j) = 1 Then
                    pb.Tag = 1
                    bomb.Add(pb) ' y los añado al array de bombas
                End If
                pbList.Add(pb) ' aquí los añado al array de objetos picture box
                Tabpb(i, j) = pb
                Me.Controls.Add(pb)
                AddHandler pb.MouseUp, AddressOf PictureBox_MouseClick
            Next
        Next
    End Sub
    Private Sub PictureBox_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        Select Case e.Button
            Case MouseButtons.Left
                ' MsgBox(sender.Tag)
                If sender.Tag = 0 Then
                    clicks += 1
                    Try
                        sender.Image = Nothing
                    Catch
                        msgFin()
                    End Try
                    sender.Backcolor = Color.DarkGray
                    sender.Tag = 3
                    numbers(sender)
                ElseIf (sender.Tag = 1) And (clicks = 0) Then
                    ' esto permite cambiar si la primera es bomba
                    clicks += 1
                    ' MsgBox("Hubieras pulsado en bomba")
                    Dim nombre As String = sender.name
                    Dim f, c, num1, num2 As Integer

                    Try
                        f = nombre.Substring(7, 1)
                        c = nombre.Substring(8, 1)
                    Catch
                        msgFin()
                    End Try
                    sender.Tag = 0
                    Tablero(f, c) = 0
                    bomb.Remove(sender)
                    For x = 1 To 1
                        Do
                            num1 = Int(7 * Rnd())
                            num2 = Int(7 * Rnd())
                        Loop Until Tablero(num1, num2) <> 1
                        Tablero(num1, num2) = 1
                        Tabpb(num1, num2).Tag = 1
                        bomb.Add(Tabpb(num1, num2))
                        ' MsgBox(num1 & num2)
                    Next
                    ' MsgBox("Bomba cambiada de sitio")
                ElseIf sender.Tag = 1 Then
                    terminar(sender)
                End If
                'If (sender.Tag = 1) And (clicks = 0) Then
                '    ' esto permite cambiar si la primera es bomba
                '    clicks += 1
                '    Dim nombre As String = sender.name
                '    Dim f, c, num1, num2 As Integer

                '    Try
                '        f = nombre.Substring(7, 1)
                '        c = nombre.Substring(8, 1)
                '    Catch
                '        msgFin()
                '    End Try
                '    sender.Tag = 0
                '    Tablero(f, c) = 0
                '    For x = 1 To 1
                '        Do
                '            num1 = Int(7 * Rnd())
                '            num2 = Int(7 * Rnd())
                '        Loop Until Tablero(num1, num2) <> 1
                '        Tablero(num1, num2) = 1
                '        Tabpb(num1, num2).Tag = 1
                '        ' MsgBox(num1 & num2)
                '    Next
                'Else
                '    terminar(sender)
                'End If
            Case MouseButtons.Right
                Select Case sender.Tag
                    Case 0
                        If sender.Text = "2" Then
                            flags -= 1
                            sender.ImageLocation = Application.StartupPath & "/img/boton.png"
                            sender.Text = ""
                            'MsgBox(flags)
                            'MsgBox(count_win)
                        Else
                            flags += 1
                            sender.ImageLocation = Application.StartupPath & "/img/bandera.png"
                            sender.Text = "2"
                        End If
                    Case 1
                        If sender.Text = "2" Then
                            flags -= 1
                            count_win -= 1
                            sender.ImageLocation = Application.StartupPath & "/img/boton.png"
                            sender.Text = ""
                            'MsgBox(flags)
                            'MsgBox(count_win)
                        Else
                            flags += 1
                            count_win += 1
                            sender.ImageLocation = Application.StartupPath & "/img/bandera.png"
                            sender.Text = "2"
                        End If
                        If (count_win = 7) And (flags = 7) Then
                            win()
                        End If
                End Select
        End Select
    End Sub
    Private Sub terminar(sender As Object)
        ' Esto se aplica si haces click izquierdo en una bomba
        For Each pb In bomb
            pb.Backcolor = Color.DarkGray
            pb.ImageLocation = Application.StartupPath & "/img/bomb2.png"
        Next
        sender.ImageLocation = Application.StartupPath & "/img/bomb.png"
        pbMain.ImageLocation = Application.StartupPath & "/img/sad.png"
        pbMain.Tag = 1
        For Each pb In pbList
            pb.Enabled = False
        Next
    End Sub
    Private Sub smile_MouseClick(sender As Object, e As EventArgs)
        Application.Restart()
        'Select Case sender.Tag
        '    Case 0
        '        Application.Restart()
        '    Case 1
        '        Dim res As Integer
        '        res = MsgBox("Terminado. ¿Desea reiniciar?", vbYesNo)
        '        If res = 6 Then
        '            Application.Restart()
        '        Else
        '            End
        '        End If
        '    Case 2
        '        Application.Restart()
        'End Select
    End Sub
    Private Sub win()
        ' esta función salta cuando se gana la partida
        MsgBox("Has ganado!!")
        For Each pb In pbList
            pb.Enabled = False
        Next
        pbMain.ImageLocation = Application.StartupPath & "/img/win.png"
    End Sub
    Private Sub numbers(sender As Object)
        Dim nombre As String = sender.name
        Dim f, c As Integer
        Dim suma As Integer = 0
        Dim i, j As Integer

        Try
            f = nombre.Substring(7, 1)
            c = nombre.Substring(8, 1)
        Catch
            msgFin()
        End Try

        For i = compruebaFilainf(f) To compruebaFilasup(f)
            For j = compruebaColinf(c) To compruebaFilasup(c)
                ' MsgBox((CStr(i) & CStr(j)))
                If Tablero(i, j) = 1 Then
                    suma += 1
                End If
            Next
        Next

        ' MsgBox(suma)
        If suma > 0 Then
            ' MsgBox(suma)
            Try
                sender.ImageLocation = Application.StartupPath & "/img/" & suma & ".png"
            Catch ex As Exception

            End Try
        End If
        suma = 0
    End Sub
    Private Function compruebaFilainf(f As Integer)
        If f = 0 Then
            Return 0
        Else
            Return (f - 1)
        End If
    End Function

    Private Function compruebaColinf(c As Integer)
        If c = 0 Then
            Return 0
        Else
            Return (c - 1)
        End If
    End Function
    Private Function compruebaFilasup(f As Integer)
        If f = 6 Then
            Return 6
        Else
            Return (f + 1)
        End If
    End Function

    Private Function compruebaColsup(c As Integer)
        If c = 6 Then
            Return 6
        Else
            Return (c + 1)
        End If
    End Function
    Private Sub msgFin()
        Dim res As Integer
        res = MsgBox("Terminado. ¿Desea reiniciar?", vbYesNo)
        If res = 6 Then
            Application.Restart()
        Else
            End
        End If
    End Sub
End Class
