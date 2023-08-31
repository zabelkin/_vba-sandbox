

'10-значный ИНН
'Вычислить сумму произведений цифр ИНН (с 1-й по 9-ю) на следующие коэффициенты — 2, 4, 10, 3, 5, 9, 4, 6, 8 (т.е. 2 * ИНН[1] + 4 * ИНН[2] + ...).
'Вычислить остаток от деления полученной суммы на 11.
'Сравнить младший разряд полученного остатка от деления с младшим разрядом ИНН. Если они равны, то ИНН верный.
'12-значный ИНН
'Вычислить 1-ю контрольную цифру:
'Вычислить сумму произведений цифр ИНН (с 1-й по 10-ю) на следующие коэффициенты — 7, 2, 4, 10, 3, 5, 9, 4, 6, 8 (т.е. 7 * ИНН[1] + 2 * ИНН[2] + ...).
'Вычислить младший разряд остатка от деления полученной суммы на 11.
'Вычислить 2-ю контрольную цифру:
'Вычислить сумму произведений цифр ИНН (с 1-й по 11-ю) на следующие коэффициенты — 3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8 (т.е. 3 * ИНН[1] + 7 * ИНН[2] + ...).
'Вычислить младший разряд остатка от деления полученной суммы на 11.
'Сравнить 1-ю контрольную цифру с 11-й цифрой ИНН и сравнить 2-ю контрольную цифру с 12-й цифрой ИНН. Если они равны, то ИНН верный.


Public Function CheckINNValidity(INN_str As String) As Boolean

    Dim quot10_arr() As Variant: quot10_arr = Array(2, 4, 10, 3, 5, 9, 4, 6, 8)
    Dim quot12_1_arr() As Variant: quot12_1_arr = Array(7, 2, 4, 10, 3, 5, 9, 4, 6, 8)
    Dim quot12_2_arr() As Variant: quot12_2_arr = Array(3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8)
    Dim control_sum As Integer, idx As Integer, control_sum1 As Integer, control_sum2 As Integer
    
    CheckINNValidity = False
    '10-значный ИНН
    If Len(INN_str) = 10 Then
        For idx = 1 To 9
            control_sum = control_sum + Int(Mid(INN_str, idx, 1)) * quot10_arr(idx - 1)
        Next idx
        control_sum = control_sum Mod 11
        If IIf(control_sum = 10, 0, control_sum) = Int(Mid(INN_str, 10, 1)) Then CheckINNValidity = True
    End If
    '12-значный ИНН
    If Len(INN_str) = 12 Then
        For idx = 1 To 10
            control_sum1 = control_sum1 + Int(Mid(INN_str, idx, 1)) * quot12_1_arr(idx - 1)
        Next idx
        control_sum1 = control_sum1 Mod 11
        For idx = 1 To 11
            control_sum2 = control_sum2 + Int(Mid(INN_str, idx, 1)) * quot12_2_arr(idx - 1)
        Next idx
        control_sum2 = control_sum2 Mod 11
        If IIf(control_sum1 = 10, 0, control_sum1) = Int(Mid(INN_str, 11, 1)) And IIf(control_sum2 = 10, 0, control_sum2) = Int(Mid(INN_str, 12, 1)) _
            Then CheckINNValidity = True
    End If

End Function