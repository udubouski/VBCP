'Imports Microsoft.Office.Interop.Excel
Public Class Form1
    Dim R(8), Tsumi(8) As Integer
    Dim gamma, Tgamma, Treq1, Treq2 As Double
    Dim n As Integer = 8

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        R(0) = 3
        R(1) = 4
        R(2) = 0
        R(3) = 0
        R(4) = 2
        R(5) = 3
        R(6) = 2
        R(7) = 4

        Tsumi(0) = 285
        Tsumi(1) = 286
        Tsumi(2) = 286
        Tsumi(3) = 333
        Tsumi(4) = 290
        Tsumi(5) = 345
        Tsumi(6) = 336
        Tsumi(7) = 301
        If input_box_for_E.Text <> "" Or input_box_for_m.Text <> "" Then

            btn_execute.Enabled = True
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        R(0) = 3
        R(1) = 4
        R(2) = 0
        R(3) = 0
        R(4) = 2
        R(5) = 3
        R(6) = 2
        R(7) = 4

        Tsumi(0) = 285
        Tsumi(1) = 286
        Tsumi(2) = 286
        Tsumi(3) = 333
        Tsumi(4) = 290
        Tsumi(5) = 345
        Tsumi(6) = 336
        Tsumi(7) = 301

        If input_box_for_E.Text <> "" Or input_box_for_m.Text <> "" Then

            btn_execute.Enabled = True
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        R(0) = 3
        R(1) = 4
        R(2) = 0
        R(3) = 0
        R(4) = 2
        R(5) = 3
        R(6) = 2
        R(7) = 4

        Tsumi(0) = 285
        Tsumi(1) = 286
        Tsumi(2) = 286
        Tsumi(3) = 333
        Tsumi(4) = 290
        Tsumi(5) = 345
        Tsumi(6) = 336
        Tsumi(7) = 301

        If input_box_for_E.Text <> "" Or input_box_for_m.Text <> "" Then

            btn_execute.Enabled = True
        End If

    End Sub

    Private Sub inputBoxTextChanged(sender As Object, e As EventArgs) Handles input_box_for_E.TextChanged, input_box_for_m.TextChanged
        If RadioButton1.Checked Or RadioButton2.Checked Or RadioButton3.Checked Then
            If input_box_for_E.Text <> "" Or input_box_for_m.Text <> "" Then

                btn_execute.Enabled = True
            End If
        End If
    End Sub
   


    Private Sub btn_execute_Handler(sender As Object, e As EventArgs) Handles btn_execute.Click
        Dim T, Tsum(8), Ti(8), Tlast_sumi(8), Tsumj(50001), Eps, F, Ez As Double
        Dim i, j, s, m, m_max As Integer
        j = 0
        If input_box_for_E.Text <> "" Then
            Ez = Val(input_box_for_E.Text)
        Else
            Ez = 1
        End If
        If Ez = 0 Then
            Ez = 1
        End If
        If input_box_for_m.Text <> "" Then
            m_max = Val(input_box_for_m.Text)
        Else
            m_max = 50000
        End If

        T = 0
        gamma = 0.9
        Tgamma = 1.645
        Treq1 = 150
        Treq2 = 400
        fill_initial_forms()

        'REM Ez - заданное значение точности моделирования
        'REM Tij -  массив значений 
        'REM Tsumi - входной массив T 

        For i = 0 To n - 1
            If R(i) > 0 Then
                Ti(i) = Tsumi(i) / R(i)
            Else
                Ti(i) = Tsumi(i) * 1.44
            End If
            T = T + 1 / Ti(i)
        Next
        T = 1 / T
        Tsumj(0) = T
        fill_forms_after_first_interation(Ti, T)
        Do
            j += 1
            For i = 0 To n - 1

                If R(i) = 0 Then
                    Tlast_sumi(i) = Ti(i)

                Else
                    s = 1
                    F = 0
                    Do

                        Randomize()
                        F += Math.Log(CDbl(Rnd()))
                        s = s + 1

                    Loop While s < R(i)

                    Tlast_sumi(i) = (-1) * Ti(i) / F
                End If

            Next
            For i = 0 To n - 1
                Tsumj(j) += 1 / Tlast_sumi(i)
            Next
            Tsumj(j) = 1 / Tsumj(j)
            If j Mod 100 = 0 Then
                Dim Temp1, Temp2 As Double
                Temp1 = Temp2 = 0
                For i = 1 To j
                    Temp1 += Tsumj(i)
                    Temp2 += Tsumj(i) * Tsumj(i)
                Next

                Eps = Tgamma * Math.Sqrt((Math.Pow(Temp1, -2) * Temp2 - 1 / j) * (j / (j - 1)))

            End If
        Loop While Eps <= Ez And j < m_max
        m = j
        T = 0
        For s = 0 To n - 1
            T += 1 / Tlast_sumi(s)
        Next
        T = 1 / T
        fill_forms_after_last_iteration(Tlast_sumi, T)

        Array.Sort(Tsumj, 0, m)
        Dim gamma1, gamma2 As Double
        gamma1 = (1 - gamma) / 2
        gamma2 = (1 + gamma) / 2
        Dim bottom As Integer = (m * gamma1)
        Dim top As Integer = (m * gamma2)
        Dim Tb, Tt As Double
        Tb = Tsumj(bottom)
        Tt = Tsumj(top)
        REM  заполнение оставшихся форм
        box_for_gamma1.Text = gamma1
        box_for_gamma2.Text = gamma2
        box_for_Tmin.Text = Tsumj(0)
        box_for_Tmax.Text = Tsumj.Max()
        box_for_Tb.Text = Tb
        box_for_Ttop.Text = Tt
        box_M_after_ex.Text = m
        box_E_after_ex.Text = Eps
        REM добавляем таблицу в эксэль
        Dim stepT As Double = (Tsumj.Max() - Tsumj(0)) / 20
        Dim T_diagram(20) As Double
        T_diagram(0) = Tsumj(0) + stepT
        For i = 1 To 19
            T_diagram(i) = T_diagram(i - 1) + stepT
        Next
        Dim count(20) As Integer
        j = 0
        For i = 0 To Tsumj.ToList().IndexOf(Tsumj.Max())
            If Tsumj(i) > T_diagram(j) Then
                j += 1
            End If
            count(j) += 1
        Next



        insert_to_Excel(T_diagram, count)
    End Sub

    Private Sub insert_to_Excel(mas1() As Double, mas2() As Integer)
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object
        Dim filePath As String = My.Application.Info.DirectoryPath + "\GRAFIK.xlsx"
        If My.Computer.FileSystem.FileExists(filePath) Then
            My.Computer.FileSystem.DeleteFile(filePath)
        End If
        'Открыть новую книгу Excel
        oExcel = CreateObject("Excel.Application")
        oBook = oExcel.Workbooks.Add

        'Создать массив с 3 столбцами и 100 строками
        Dim DataArray(0 To 19, 0 To 1) As Double
        Dim r As Integer
        For r = 0 To 19
            DataArray(r, 0) = mas1(r)
            DataArray(r, 1) = mas2(r)
        Next

        'Добавить заголовки в строку 1
        oSheet = oBook.Worksheets(1)
        oSheet.Range("A1").Value = "Верхняя граница интервала"
        oSheet.Range("B1").Value = "Число реализаций"

        'Передать массив на лист, начиная с ячейки A2
        oSheet.Range("A2").Resize(20, 2).Value = DataArray
        'Добавляем график
        '        Dim myChart As ChartObject

        '        myChart = oSheet.ChartObjects.Add(100, 50, 400, 200)
        '  With myChart
        '.Chart.SetSourceData(oSheet.Range("A2:B19"))
        '  .Chart.ChartType = XlChartType.xlXYScatterLines
        '   End With
        'Сохранить книгу и закрыть Excel
        oBook.SaveAs(filePath)
        oExcel.Quit()

        oExcel.WorkBooks.open(filePath)
        oExcel.Visible = True


    End Sub
    Private Sub fill_forms_after_last_iteration(Tlast() As Double, T As Double)
        box_for_Ts1_after_ex.Text = Tlast(0)
        box_for_Ts2_after_ex.Text = Tlast(1)
        box_for_Ts3_after_ex.Text = Tlast(2)
        box_for_Ts4_after_ex.Text = Tlast(3)
        box_for_Ts5_after_ex.Text = Tlast(4)
        box_for_Ts6_after_ex.Text = Tlast(5)
        box_for_Ts7_after_ex.Text = Tlast(6)
        box_for_Ts8_after_ex.Text = Tlast(7)
        box_for_T_after_ex.Text = T
    End Sub

    Private Sub fill_initial_forms()
        box_for_n.Text = 8
        box_for_Treq1.Text = Treq1
        box_for_Treq2.Text = Treq2
        box_for_y.Text = gamma
        box_for_Ty.Text = Tgamma

        box_for_R1.Text = R(0)
        box_for_R2.Text = R(1)
        box_for_R3.Text = R(2)
        box_for_R4.Text = R(3)
        box_for_R5.Text = R(4)
        box_for_R6.Text = R(5)
        box_for_R7.Text = R(6)
        box_for_R8.Text = R(7)

        box_for_T1.Text = Tsumi(0)
        box_for_T2.Text = Tsumi(1)
        box_for_T3.Text = Tsumi(2)
        box_for_T4.Text = Tsumi(3)
        box_for_T5.Text = Tsumi(4)
        box_for_T6.Text = Tsumi(5)
        box_for_T7.Text = Tsumi(6)
        box_for_T8.Text = Tsumi(7)
    End Sub
    Private Sub fill_forms_after_first_interation(Ti() As Double, T As Double)
        box_for_Ti1.Text = Ti(0)
        box_for_Ti2.Text = Ti(1)
        box_for_Ti3.Text = Ti(2)
        box_for_Ti4.Text = Ti(3)
        box_for_Ti5.Text = Ti(4)
        box_for_Ti6.Text = Ti(5)
        box_for_Ti7.Text = Ti(6)
        box_for_Ti8.Text = Ti(7)

        box_for_T.Text = T
        Dim Ts As Double
        For I As Integer = 0 To 7
            Ts += Ti(I)
        Next
        box_for_Ts.Text = Ts / 8
    End Sub
End Class
