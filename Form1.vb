Public Class Form1

    Inherits System.Windows.Forms.Form

    Private Declare Function midiOutGetNumDevs Lib "winmm" () As Short
    Private Declare Function midiOutOpen Lib "winmm.dll" (ByRef lphMidiOut As Integer, ByVal uDeviceID As Integer, ByVal dwCallback As Integer, ByVal dwInstance As Integer, ByVal dwFlags As Integer) As Integer
    Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Integer, ByVal dwMsg As Integer) As Integer
    Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Integer) As Integer
    Private Declare Function GetTickCount Lib "kernel32" () As Integer
    Private hMidiOut As Integer

    Private Sub Form1_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Closed

        Dim ret As Integer

        ret = midiOutClose(hMidiOut)

        End

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim n As Short
        Dim ret As Integer

        n = midiOutGetNumDevs()

        If n = 0 Then
            MessageBox.Show("MIDI音源を利用できません", "デジタルピアノ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ret = midiOutOpen(hMidiOut, -1, 0, 0, 0)

        tickSelect.SelectedIndex = 1

    End Sub


    Private Sub pctKey0_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pctKeym1.MouseDown, pctKeym2.MouseDown, _
    pctKeym3.MouseDown, pctKeym4.MouseDown, pctKeym5.MouseDown, pctKeym6.MouseDown, pctKeym7.MouseDown, pctKeym8.MouseDown, pctKeym9.MouseDown, pctKeym10.MouseDown, pctKeym11.MouseDown, pctKeym12.MouseDown, _
        pctKey0.MouseDown, pctKey1.MouseDown, pctKey2.MouseDown, pctKey3.MouseDown, pctKey4.MouseDown, pctKey5.MouseDown, pctKey6.MouseDown, pctKey7.MouseDown, pctKey8.MouseDown, pctKey9.MouseDown, pctKey10.MouseDown, _
        pctKey11.MouseDown, pctKey12.MouseDown, pctKey13.MouseDown, pctKey14.MouseDown, pctKey15.MouseDown, pctKey16.MouseDown, pctKey17.MouseDown, pctKey18.MouseDown, pctKey19.MouseDown, pctKey20.MouseDown, _
        pctKey21.MouseDown, pctKey22.MouseDown, pctKey23.MouseDown, pctKey24.MouseDown

        Call color_on(sender.Tag)
        Call midi_on(sender.Tag)

    End Sub

    Private Sub pctKey0_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pctKeym1.MouseUp, pctKeym2.MouseUp, _
        pctKeym3.MouseUp, pctKeym4.MouseUp, pctKeym5.MouseUp, pctKeym6.MouseUp, pctKeym7.MouseUp, pctKeym8.MouseUp, pctKeym9.MouseUp, pctKeym10.MouseUp, pctKeym11.MouseUp, pctKeym12.MouseUp, _
        pctKey0.MouseUp, pctKey1.MouseUp, pctKey2.MouseUp, pctKey3.MouseUp, pctKey4.MouseUp, pctKey5.MouseUp, pctKey6.MouseUp, pctKey7.MouseUp, pctKey8.MouseUp, pctKey9.MouseUp, pctKey10.MouseUp, _
        pctKey11.MouseUp, pctKey12.MouseUp, pctKey13.MouseUp, pctKey14.MouseUp, pctKey15.MouseUp, pctKey16.MouseUp, pctKey17.MouseUp, pctKey18.MouseUp, pctKey19.MouseUp, pctKey20.MouseUp, _
        pctKey21.MouseUp, pctKey22.MouseUp, pctKey23.MouseUp, pctKey24.MouseUp

        Call color_off(sender.Tag)
        Call midi_off(sender.Tag)

    End Sub

    Private Sub color_on(ByVal n As Integer)

        Select Case n
            Case -12 : pctKeym12.BackColor = Color.LemonChiffon
            Case -11 : pctKeym11.BackColor = Color.LemonChiffon
            Case -10 : pctKeym10.BackColor = Color.LemonChiffon
            Case -9 : pctKeym9.BackColor = Color.LemonChiffon
            Case -8 : pctKeym8.BackColor = Color.LemonChiffon
            Case -7 : pctKeym7.BackColor = Color.LemonChiffon
            Case -6 : pctKeym6.BackColor = Color.LemonChiffon
            Case -5 : pctKeym5.BackColor = Color.LemonChiffon
            Case -4 : pctKeym4.BackColor = Color.LemonChiffon
            Case -3 : pctKeym3.BackColor = Color.LemonChiffon
            Case -2 : pctKeym2.BackColor = Color.LemonChiffon
            Case -1 : pctKeym1.BackColor = Color.LemonChiffon
            Case 0 : pctKey0.BackColor = Color.LemonChiffon
            Case 1 : pctKey1.BackColor = Color.LemonChiffon
            Case 2 : pctKey2.BackColor = Color.LemonChiffon
            Case 3 : pctKey3.BackColor = Color.LemonChiffon
            Case 4 : pctKey4.BackColor = Color.LemonChiffon
            Case 5 : pctKey5.BackColor = Color.LemonChiffon
            Case 6 : pctKey6.BackColor = Color.LemonChiffon
            Case 7 : pctKey7.BackColor = Color.LemonChiffon
            Case 8 : pctKey8.BackColor = Color.LemonChiffon
            Case 9 : pctKey9.BackColor = Color.LemonChiffon
            Case 10 : pctKey10.BackColor = Color.LemonChiffon
            Case 11 : pctKey11.BackColor = Color.LemonChiffon
            Case 12 : pctKey12.BackColor = Color.LemonChiffon
            Case 13 : pctKey13.BackColor = Color.LemonChiffon
            Case 14 : pctKey14.BackColor = Color.LemonChiffon
            Case 15 : pctKey15.BackColor = Color.LemonChiffon
            Case 16 : pctKey16.BackColor = Color.LemonChiffon
            Case 17 : pctKey17.BackColor = Color.LemonChiffon
            Case 18 : pctKey18.BackColor = Color.LemonChiffon
            Case 19 : pctKey19.BackColor = Color.LemonChiffon
            Case 20 : pctKey20.BackColor = Color.LemonChiffon
            Case 21 : pctKey21.BackColor = Color.LemonChiffon
            Case 22 : pctKey22.BackColor = Color.LemonChiffon
            Case 23 : pctKey23.BackColor = Color.LemonChiffon
            Case 24 : pctKey24.BackColor = Color.LemonChiffon
        End Select

    End Sub

    Private Sub color_off(ByVal n As Integer)

        Select Case n
            Case -12 : pctKeym12.BackColor = Color.White
            Case -11 : pctKeym11.BackColor = Color.Black
            Case -10 : pctKeym10.BackColor = Color.White
            Case -9 : pctKeym9.BackColor = Color.Black
            Case -8 : pctKeym8.BackColor = Color.White
            Case -7 : pctKeym7.BackColor = Color.White
            Case -6 : pctKeym6.BackColor = Color.Black
            Case -5 : pctKeym5.BackColor = Color.White
            Case -4 : pctKeym4.BackColor = Color.Black
            Case -3 : pctKeym3.BackColor = Color.White
            Case -2 : pctKeym2.BackColor = Color.Black
            Case -1 : pctKeym1.BackColor = Color.White
            Case 0 : pctKey0.BackColor = Color.White
            Case 1 : pctKey1.BackColor = Color.Black
            Case 2 : pctKey2.BackColor = Color.White
            Case 3 : pctKey3.BackColor = Color.Black
            Case 4 : pctKey4.BackColor = Color.White
            Case 5 : pctKey5.BackColor = Color.White
            Case 6 : pctKey6.BackColor = Color.Black
            Case 7 : pctKey7.BackColor = Color.White
            Case 8 : pctKey8.BackColor = Color.Black
            Case 9 : pctKey9.BackColor = Color.White
            Case 10 : pctKey10.BackColor = Color.Black
            Case 11 : pctKey11.BackColor = Color.White
            Case 12 : pctKey12.BackColor = Color.White
            Case 13 : pctKey13.BackColor = Color.Black
            Case 14 : pctKey14.BackColor = Color.White
            Case 15 : pctKey15.BackColor = Color.Black
            Case 16 : pctKey16.BackColor = Color.White
            Case 17 : pctKey17.BackColor = Color.White
            Case 18 : pctKey18.BackColor = Color.Black
            Case 19 : pctKey19.BackColor = Color.White
            Case 20 : pctKey20.BackColor = Color.Black
            Case 21 : pctKey21.BackColor = Color.White
            Case 22 : pctKey22.BackColor = Color.Black
            Case 23 : pctKey23.BackColor = Color.White
            Case 24 : pctKey24.BackColor = Color.White
        End Select

    End Sub

    Private Sub midi_on(ByVal n As Integer)

        Dim ret As Integer
        Dim dwMsg As Integer

        dwMsg = &H7F3C90 + (n * 256)
        ret = midiOutShortMsg(hMidiOut, dwMsg)

    End Sub

    Private Sub midi_off(ByVal n As Integer)

        Dim ret As Integer
        Dim dwMsg As Integer

        dwMsg = &H3C90 + (n * 256)
        ret = midiOutShortMsg(hMidiOut, dwMsg)

    End Sub

    Private Sub btnAuto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAuto.Click

        Dim fileNo As Integer
        Dim tempo As Single
        Dim tick As Double
        Dim no As String
        Dim no2 As String
        Dim no3 As String
        Dim no4 As String
        Dim no5 As String
        Dim no6 As String
        Dim s As String
        Dim s2 As String
        Dim s3 As String
        Dim s4 As String
        Dim s5 As String
        Dim s6 As String

        If IsNumeric(txtTempo.Text) = False Then
            MessageBox.Show("テンポは0以外の正の数値を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        ElseIf txtTempo.Text = 0 Then
            MessageBox.Show("テンポは0以外の正の数値を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        ElseIf txtTempo.Text < 0 Then
            MessageBox.Show("テンポは0以外の正の数値を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        Else
            tempo = txtTempo.Text
        End If

        If tickSelect.SelectedItem = "32分音符" Then
            tick = 0.125
        ElseIf tickSelect.SelectedItem = "16分音符" Then
            tick = 0.25
        ElseIf tickSelect.SelectedItem = "8分音符" Then
            tick = 0.5
        ElseIf tickSelect.SelectedItem = "4分音符" Then
            tick = 1
        ElseIf tickSelect.SelectedItem = "2分音符" Then
            tick = 2
        ElseIf tickSelect.SelectedItem = "全音符" Then
            tick = 4
        Else
            MessageBox.Show("最小単位は正しい値を選択してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        dlgOpen.InitialDirectory = Application.StartupPath & "\.."

        If dlgOpen.ShowDialog() = DialogResult.Cancel Then
            txtTitle.Text = ""
            Exit Sub
        End If

        txtTitle.Text = dlgOpen.FileName

        fileNo = FreeFile()
        FileOpen(fileNo, dlgOpen.FileName, OpenMode.Input)

        Call sleep(500)

        no = 99
        no2 = 99
        no3 = 99
        no4 = 99
        no5 = 99
        no6 = 99
        s = 99
        s2 = 99
        s3 = 99
        s4 = 99
        s5 = 99
        s6 = 99

        Input(fileNo, no)
        Input(fileNo, no2)
        Input(fileNo, no3)
        Input(fileNo, no4)
        Input(fileNo, no5)
        Input(fileNo, no6)

        If Not no = "p" Then
            Call color_on(no)
            Call midi_on(no)
            s = no
        End If

        If Not no2 = "p" Then
            Call color_on(no2)
            Call midi_on(no2)
            s2 = no2
        End If

        If Not no3 = "p" Then
            Call color_on(no3)
            Call midi_on(no3)
            s3 = no3
        End If

        If Not no4 = "p" Then
            Call color_on(no4)
            Call midi_on(no4)
            s4 = no4
        End If

        If Not no5 = "p" Then
            Call color_on(no5)
            Call midi_on(no5)
            s5 = no5
        End If

        If Not no6 = "p" Then
            Call color_on(no6)
            Call midi_on(no6)
            s6 = no6
        End If

        Call sleep(tempo * tick)

        Do Until EOF(fileNo)
            Input(fileNo, no)
            Input(fileNo, no2)
            Input(fileNo, no3)
            Input(fileNo, no4)
            Input(fileNo, no5)
            Input(fileNo, no6)
            If Not no = "p" Then
                Call color_off(s)
                Call midi_off(s)
                Call color_on(no)
                Call midi_on(no)
                s = no
            End If
            If Not no2 = "p" Then
                Call color_off(s2)
                Call midi_off(s2)
                Call color_on(no2)
                Call midi_on(no2)
                s2 = no2
            End If
            If Not no3 = "p" Then
                Call color_off(s3)
                Call midi_off(s3)
                Call color_on(no3)
                Call midi_on(no3)
                s3 = no3
            End If
            If Not no4 = "p" Then
                Call color_off(s4)
                Call midi_off(s4)
                Call color_on(no4)
                Call midi_on(no4)
                s4 = no4
            End If
            If Not no5 = "p" Then
                Call color_off(s5)
                Call midi_off(s5)
                Call color_on(no5)
                Call midi_on(no5)
                s5 = no5
            End If
            If Not no6 = "p" Then
                Call color_off(s6)
                Call midi_off(s6)
                Call color_on(no6)
                Call midi_on(no6)
                s6 = no6
            End If
            Call sleep(tempo * tick)
        Loop

        Call midi_off(s)
        Call color_off(s)
        Call midi_off(s2)
        Call color_off(s2)
        Call midi_off(s3)
        Call color_off(s3)
        Call midi_off(s4)
        Call color_off(s4)
        Call midi_off(s5)
        Call color_off(s5)
        Call midi_off(s6)
        Call color_off(s6)

        FileClose(fileNo)

    End Sub

    Private Sub sleep(ByVal n As Double)

        Dim t As Double

        t = GetTickCount()

        Do Until GetTickCount() > t + n
            Application.DoEvents()
        Loop

    End Sub

End Class
