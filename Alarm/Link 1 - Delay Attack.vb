' Link 1 - Delay Attack Detection

' Variáveis iniciais
Dim last_timestamp_Link1
Dim max_delay_Link1
Dim LTMS_TmSyn_Link1
Dim time_difference_Link1
Dim current_time_Link1

' Inicialização
Set last_timestamp_Link1 = CreateObject("Scripting.Dictionary")
max_delay_Link1 = 2 ' Defina o limite de atraso permitido em segundos

Sub Delay_Attack_Detection_Link1()
    ' Obtenha o status de sincronização de tempo
    LTMS_TmSyn_Link1 = Application.GetObject("LTMS_TmSyn_Link1").Value
    current_time_Link1 = Now
    
    Dim GoID
    GoID = "Link1"
    
    If LTMS_TmSyn_Link1 <> "synchronized" Then
        Application.GetObject("Delay_Detected_Link1").Value = True
    End If
    
    If last_timestamp_Link1.Exists(GoID) Then
        time_difference_Link1 = DateDiff("s", last_timestamp_Link1(GoID), current_time_Link1)
        If time_difference_Link1 > max_delay_Link1 Then
            Application.GetObject("Delay_Detected_Link1").Value = True
        End If
    End If
    
    last_timestamp_Link1(GoID) = current_time_Link1
End Sub
