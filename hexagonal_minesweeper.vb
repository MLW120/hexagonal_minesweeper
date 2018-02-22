Imports System.Drawing
Public Class Form1
    Public tauler(26, 19) As Integer
    Public obert(26, 19) As Boolean
    Public BClicked(26, 19) As Boolean

    Private Sub Final()
        For x = 0 To 26
            For y = 0 To 19
                If tauler(x, y) = 99 And BClicked(x, y) = False Then
                    Dim p As Point
                    p.Y = (y + 1) * 30 - 7
                    p.X = 37 + 34 * x - 7
                    If y Mod 2 = 1 Then
                        p.X = p.X + 17
                    End If
                    Me.CreateGraphics.DrawString("M", New Font("Arial", 11, FontStyle.Bold), Brushes.Black, p)
                End If
                If tauler(x, y) <> 99 And BClicked(x, y) = True Then
                    Dim p As Point
                    p.Y = (y + 1) * 30 - 7
                    p.X = 37 + 34 * x - 7
                    If y Mod 2 = 1 Then
                        p.X = p.X + 17
                    End If
                    Me.CreateGraphics.DrawString("//", New Font("Arial", 11, FontStyle.Bold), Brushes.Red, p)
                End If
                If tauler(x, y) = 99 And obert(x, y) = True Then
                    Dim p As Point
                    p.Y = (y + 1) * 30 - 7
                    p.X = 37 + 34 * x - 7
                    If y Mod 2 = 1 Then
                        p.X = p.X + 17
                    End If
                    Me.CreateGraphics.DrawString("M", New Font("Arial", 11, FontStyle.Bold), Brushes.Red, p)
                End If
            Next
        Next

        Threading.Thread.Sleep(3000)
        MsgBox("¡Intentelo de nuevo!")
    End Sub

    Private Sub CrearTauler()
        Randomize()
        For i = 0 To 75
            Dim x As Integer = CInt(Int(26 * Rnd()))
            Dim y As Integer = CInt(Int(19 * Rnd()))
            If tauler(x, y) = 99 Then
                i = i - 1
            Else
                tauler(x, y) = 99
            End If
        Next
        For x1 = 0 To 26
            For y1 = 0 To 19
                If tauler(x1, y1) <> 99 Then
                    Dim n As Integer = 0
                    Dim p As parelles
                    p.a = x1
                    p.b = y1
                    Dim v() As parelles = Veins(p)
                    For i = 0 To 5
                        Dim x As Integer = v(i).a
                        Dim y As Integer = v(i).b
                        If x >= 0 And y >= 0 And x <= 26 And y <= 19 Then
                            If tauler(x, y) = 99 Then
                                n = n + 1
                            Else
                                n = n
                            End If
                        End If
                    Next
                    tauler(p.a, p.b) = n
                End If
                obert(x1, y1) = False
                BClicked(x1, y1) = False
            Next
        Next
    End Sub

    Private Sub Escriure(ByVal p As parelles, ByVal right As Boolean)
        Dim c As Brush = Brushes.Black
        If right = True Then
            Dim x As Integer = p.a
            Dim y As Integer = p.b
            If x < 0 Or y < 0 Or x > 26 Or y > 19 Then Exit Sub
            If obert(x, y) = True Or BClicked(x, y) = True Then Exit Sub
            Dim s As String = Nothing
            Dim p1 As Point
            p1.Y = (y + 1) * 30 - 7
            p1.X = 37 + 34 * x - 7
            If y Mod 2 = 1 Then
                p1.X = p1.X + 17
            End If
            obert(x, y) = True
            If x >= 0 And y >= 0 And x <= 26 And y <= 19 Then
                Select Case tauler(x, y)
                    Case 0
                        s = "0"
                        c = Brushes.DarkViolet
                        Dim v() As parelles = Veins(p)
                        For i = 0 To 5
                            If v(i).a >= 0 And v(i).b >= 0 And v(i).a <= 26 And v(i).b <= 19 Then
                                If obert(v(i).a, v(i).b) = False Then
                                    Escriure(v(i), True)
                                End If
                            End If
                        Next
                    Case 1
                        s = "1"
                        c = Brushes.Blue
                    Case 2
                        s = "2"
                        c = Brushes.Green
                    Case 3
                        s = "3"
                        c = Brushes.Red
                    Case 4
                        s = "4"
                        c = Brushes.DarkBlue
                    Case 5
                        s = "5"
                        c = Brushes.Brown
                    Case 6
                        s = "6"
                        c = Brushes.Turquoise
                    Case 99
                        s = "M"
                        Final()
                End Select
            End If
            Me.CreateGraphics.DrawString(s, New Font("Arial", 11, FontStyle.Bold), c, p1)
        Else
            Dim x As Integer = p.a
            Dim y As Integer = p.b
            If x < 0 Or y < 0 Or x > 26 Or y > 19 Then Exit Sub
            Dim s As String = "B"
            Dim p1 As Point
            p1.Y = (y + 1) * 30 - 7
            p1.X = 37 + 34 * x - 7
            If y Mod 2 = 1 Then
                p1.X = p1.X + 17
            End If
            If obert(x, y) = False And BClicked(x, y) = False Then
                BClicked(x, y) = True
                Me.CreateGraphics.DrawString(s, New Font("Arial", 11, FontStyle.Bold), c, p1)
            Else
                If obert(x, y) = False And BClicked(x, y) = True Then
                    BClicked(x, y) = False
                    s = "B"
                    Me.CreateGraphics.DrawString(s, New Font("Arial", 11, FontStyle.Bold), Brushes.LightGray, p1)
                End If
            End If
        End If
    End Sub

    Private Function Veins(ByVal p As parelles) As parelles()
        Dim r(6) As parelles
        Dim FP As Boolean
        Dim FS As Boolean
        If p.b Mod 2 = 0 Then
            FP = True
        Else
            FS = True
        End If
        If FP = True Then
            r(0).a = p.a - 1
            r(0).b = p.b - 1
            r(1).a = p.a - 1
            r(1).b = p.b
            r(2).a = p.a - 1
            r(2).b = p.b + 1
            r(3).a = p.a
            r(3).b = p.b + 1
            r(4).a = p.a + 1
            r(4).b = p.b
            r(5).a = p.a
            r(5).b = p.b - 1
        End If
        If FS = True Then
            r(0).a = p.a
            r(0).b = p.b - 1
            r(1).a = p.a - 1
            r(1).b = p.b
            r(2).a = p.a
            r(2).b = p.b + 1
            r(3).a = p.a + 1
            r(3).b = p.b + 1
            r(4).a = p.a + 1
            r(4).b = p.b
            r(5).a = p.a + 1
            r(5).b = p.b - 1
        End If
        Return r
    End Function

    Public Structure parelles
        Dim a As Integer
        Dim b As Integer
    End Structure

    Private Sub Form1_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
        Dim p As Point = e.Location
        Dim pa As parelles = WhereIs(p)
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Escriure(pa, True)
        End If
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Escriure(pa, False)
        End If
        Dim i As Integer = 0
        Dim m As Integer = 75
        For x = 0 To 26
            For y = 0 To 19
                If obert(x, y) = True Or tauler(x, y) = 99 Then
                    i = i + 1
                End If
                If BClicked(x, y) = True Then
                    m -= 1
                End If
            Next
        Next
        txtMQQ.Text = m.ToString
        If i = 540 Then
            MsgBox("¡¡¡Felicidades!!!")
            End
        End If
    End Sub

    Public Function WhereIs(ByVal p As Point) As parelles
        Dim x As Integer = p.X
        Dim y As Integer = p.Y
        Dim HX As Integer
        Dim HY As Integer
        Dim r As parelles
        r.a = 50
        r.b = 50
        Dim filaS As Boolean = False
        Dim filaP As Boolean = False
        Dim filaF1 As Boolean = False
        Dim filaF2 As Boolean = False
        If y <= 10 Or y >= 620 Then
            r.a = -1
            r.b = -1
            Return r
        End If
        If x <= 20 Or x >= 955 Then ' 982
            r.a = -1
            r.b = -1
            Return r
        End If
        If (y - 20) / 20 Mod 3 >= 0 And (y - 20) / 20 Mod 3 <= 1 Then
            filaS = True
        End If
        If (y - 50) / 20 Mod 3 >= 0 And (y - 50) / 20 Mod 3 <= 1 Then
            filaP = True
        End If
        If filaS = False And filaP = False Then
            If (y - 10) / 10 Mod 6 >= 0 And (y - 10) / 10 Mod 6 <= 1 Then

                filaF1 = True
                'MsgBox("1")
            End If
            If (y - 40) / 10 Mod 6 >= 0 And (y - 40) / 10 Mod 6 <= 1 Then
                filaF2 = True
                'MsgBox("2")
            End If
        End If

        If filaS = True Then
            If x < 20 Or x > 938 Then ' 965
                r.a = -1
                r.b = -1
                Return r
            Else
                HX = Int((x - 20) / 34)   '35
                HY = Int((y - 20) / 60) * 2
                r.a = HX
                r.b = HY
            End If
        End If

        If filaP = True Then
            If x < 37 Or x > 955 Then   ' 982
                r.a = -1
                r.b = -1
                Return r
            Else
                HX = Int((x - 37) / 34)   ' 35
                HY = Int((y - 50) / 60) * 2 + 1
                r.a = HX
                r.b = HY
            End If
        End If

        If filaF1 = True Then
            Dim sobre As Boolean
            Dim sota As Boolean
            If ((x - 20) Mod 34) / 34 >= 0 And ((x - 20) Mod 34) / 34 <= 0.5 Then      '35 !!!
                If ((y - 10) Mod 10) > ((-0.58823529 * (x - 20) Mod 10) + 10) Then   'Mod 10 ???
                    sota = True
                    Dim pN As Point
                    pN.X = p.X
                    pN.Y = p.Y + 11
                    r = WhereIs(pN)

                Else
                    sobre = True
                    Dim pN As Point
                    pN.X = p.X
                    pN.Y = p.Y - 11
                    r = WhereIs(pN)

                End If
            End If
            If ((x - 20) Mod 34) / 34 > 0.5 And ((x - 20) Mod 34) / 34 <= 1 Then      '35 !!!
                If ((y - 10) Mod 10) > ((0.58823529 * (x - 20) Mod 10)) Then   'Mod 10 ???
                    sota = True
                    Dim pN As Point
                    pN.X = p.X
                    pN.Y = p.Y + 11
                    r = WhereIs(pN)

                Else
                    sobre = True
                    Dim pN As Point
                    pN.X = p.X
                    pN.Y = p.Y - 11
                    r = WhereIs(pN)

                End If
            End If
        End If

        If filaF2 = True Then
            Dim sobre As Boolean
            Dim sota As Boolean
            If ((x - 20) Mod 34) / 34 >= 0 And ((x - 20) Mod 34) / 34 <= 0.5 Then      '35 !!!
                If ((y - 40) Mod 10) > ((0.58823529 * (x - 20) Mod 10)) Then   'Mod 10 ???
                    sota = True
                    Dim pN As Point
                    pN.X = p.X
                    pN.Y = p.Y + 11
                    r = WhereIs(pN)

                Else
                    sobre = True
                    Dim pN As Point
                    pN.X = p.X
                    pN.Y = p.Y - 11
                    r = WhereIs(pN)

                End If
            End If
            If ((x - 20) Mod 34) / 34 > 0.5 And ((x - 20) Mod 34) / 34 <= 1 Then      '35 !!!
                If ((y - 40) Mod 10) > ((-0.58823529 * (x - 20) Mod 10 + 10)) Then   'Mod 10 ???
                    sota = True
                    Dim pN As Point
                    pN.X = p.X
                    pN.Y = p.Y + 11
                    r = WhereIs(pN)

                Else
                    sobre = True
                    Dim pN As Point
                    pN.X = p.X
                    pN.Y = p.Y - 11
                    r = WhereIs(pN)

                End If
            End If
        End If

        Return r
    End Function

    Public Sub Columna(ByVal p As Point)
        Dim x As Integer = p.X
        Dim y As Integer = p.Y
        Dim ph As Point = p
        For i As Integer = 1 To 10
            Fila(ph)
            ph.X = ph.X + 17
            ph.Y = ph.Y + 30
            Fila(ph)
            ph.X = ph.X - 17
            ph.Y = ph.Y + 30
        Next
    End Sub

    Public Sub Fila(ByVal p As Point)
        Dim x As Integer = p.X
        Dim y As Integer = p.Y
        Dim ph As Point = p
        For i As Integer = 1 To 27
            Hexagon(ph)
            ph.X = ph.X + 34     '35
        Next
    End Sub

    Public Sub Hexagon(ByVal p As Point)
        Dim x As Integer = p.X
        Dim y As Integer = p.Y
        Me.CreateGraphics.DrawLine(Drawing.Pens.Black, x, y, x, y + 20)
        Me.CreateGraphics.DrawLine(Drawing.Pens.Black, x + 17, y + 30, x, y + 20)
        Me.CreateGraphics.DrawLine(Drawing.Pens.Black, x + 17, y + 30, x + 34, y + 20)  '35
        Me.CreateGraphics.DrawLine(Drawing.Pens.Black, x + 34, y + 20, x + 34, y)  '35
        Me.CreateGraphics.DrawLine(Drawing.Pens.Black, x + 34, y, x + 17, y - 10)  '35
        Me.CreateGraphics.DrawLine(Drawing.Pens.Black, x + 17, y - 10, x, y)
    End Sub


    Private Sub Form1_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Pintar()
        'Presentar()
    End Sub

    Public Sub Pintar()
        Dim p1 As Point
        p1.X = 20
        p1.Y = 20
        Columna(p1)
    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        CrearTauler()
    End Sub

    Private Function Pens() As Drawing.Font
        Throw New NotImplementedException
    End Function

    Private Sub btnReiniciar_Click(sender As System.Object, e As System.EventArgs) Handles btnReiniciar.Click
        Application.Restart()
    End Sub

    Public Sub Presentar()
        For x As Integer = 0 To 26
            For y As Integer = 0 To 19
                If (obert(x, y) = True) Or (BClicked(x, y) = True) Then
                    obert(x, y) = False
                    BClicked(x, y) = False
                    Dim p As parelles
                    p.a = x
                    p.b = y
                    Escriure(p, True)
                End If
            Next
        Next
    End Sub
End Class
