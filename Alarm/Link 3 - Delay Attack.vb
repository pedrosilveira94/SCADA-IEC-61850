' Link 3 - Delay Attack Detection

' Variáveis iniciais
Dim last_timestamp_Link3
Dim max_delay_Link3
Dim LTMS_TmSyn_Link3
Dim time_difference_Link3
Dim current_time_Link3

' Inicialização
Set last_timestamp_Link3 = CreateObject("Scripting.Dictionary")
max_delay_Link3 = 2 ' Defina o limite de atraso permitido em segundos

Sub Delay_Attack_Detection_Link3()
    ' Obtenha o status de sincronização de tempo
    LTMS_TmSyn_Link3 = Application.GetObject("LTMS_TmSyn_Link3").Value
    current_time_Link3 = Now
    
    Dim GoID
    GoID = "Link3"
    
    If LTMS_TmSyn_Link3 <> "synchronized" Then
        Application.GetObject("Delay_Detected_Link3").Value = True
    End If
    
    If last_timestamp_Link3.Exists(GoID) Then
        time_difference_Link3 = DateDiff("s", last_timestamp_Link3(GoID), current_time_Link3)
        If time_difference_Link3 > max_delay_Link3 Then
            Application.GetObject("Delay_Detected_Link3").Value = True
        End If
    End If
    
    last_timestamp_Link3(GoID) = current_time_Link3
End Sub
