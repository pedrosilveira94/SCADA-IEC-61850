' Link 2 - Delay Attack Detection

' Variáveis iniciais
Dim last_timestamp_Link2
Dim max_delay_Link2
Dim LTMS_TmSyn_Link2
Dim time_difference_Link2
Dim current_time_Link2

' Inicialização
Set last_timestamp_Link2 = CreateObject("Scripting.Dictionary")
max_delay_Link2 = 2 ' Defina o limite de atraso permitido em segundos

Sub Delay_Attack_Detection_Link2()
    ' Obtenha o status de sincronização de tempo
    LTMS_TmSyn_Link2 = Application.GetObject("LTMS_TmSyn_Link2").Value
    current_time_Link2 = Now
    
    Dim GoID
    GoID = "Link2"
    
    If LTMS_TmSyn_Link2 <> "synchronized" Then
        Application.GetObject("Delay_Detected_Link2").Value = True
    End If
    
    If last_timestamp_Link2.Exists(GoID) Then
        time_difference_Link2 = DateDiff("s", last_timestamp_Link2(GoID), current_time_Link2)
        If time_difference_Link2 > max_delay_Link2 Then
            Application.GetObject("Delay_Detected_Link2").Value = True
        End If
    End If
    
    last_timestamp_Link2(GoID) = current_time_Link2
End Sub
