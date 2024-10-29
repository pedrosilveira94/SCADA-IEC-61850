' Link 3 - DoS High Flooding Attack Detection

' Variáveis iniciais
Dim message_count_Link3
Dim previous_t_Link3
Dim current_t_Link3
Dim detection_window_Link3
Dim threshold_rate_Link3
Dim min_time_Link3
Dim TTL_Link3
Dim txCnt_IED_PUB_Link3
Dim rxCnt_IED_SUB_Link3
Dim outOv_IED_PUB_Link3
Dim inOv_IED_SUB_Link3
Dim ferCh_IED_SUB_Link3

' Inicialização
message_count_Link3 = 0
previous_t_Link3 = Now
detection_window_Link3 = 1 ' Em segundos
threshold_rate_Link3 = 100 ' Máximo de mensagens por segundo
min_time_Link3 = 0.01 ' Tempo mínimo entre mensagens (em segundos)
TTL_Link3 = 2 ' Defina conforme necessário (em segundos)

Sub DoS_Attack_Detection_Link3()
    ' Atualize as variáveis conforme necessário
    current_t_Link3 = Now ' Timestamp atual
    message_count_Link3 = message_count_Link3 + 1
    
    ' Obtenha os valores de txCnt, rxCnt, inOv, outOv, ferCh via tags do Elipse E3
    txCnt_IED_PUB_Link3 = Application.GetObject("txCnt_IED_PUB_Link3").Value
    rxCnt_IED_SUB_Link3 = Application.GetObject("rxCnt_IED_SUB_Link3").Value
    outOv_IED_PUB_Link3 = Application.GetObject("outOv_IED_PUB_Link3").Value
    inOv_IED_SUB_Link3 = Application.GetObject("inOv_IED_SUB_Link3").Value
    ferCh_IED_SUB_Link3 = Application.GetObject("ferCh_IED_SUB_Link3").Value
    
    ' Verificação de DoS Attack
    If message_count_Link3 > threshold_rate_Link3 Or (DateDiff("s", previous_t_Link3, current_t_Link3) < min_time_Link3) Or (DateDiff("s", previous_t_Link3, current_t_Link3) > TTL_Link3) Then
        Application.GetObject("DoS_Detected_Link3").Value = True
        ' Ação adicional, como descartar mensagens
    ElseIf inOv_IED_SUB_Link3 > 0 Then
        Application.GetObject("DoS_Detected_Link3").Value = True
    ElseIf ferCh_IED_SUB_Link3 > 0 Then
        Application.GetObject("DoS_Detected_Link3").Value = True
    End If
    
    ' Verifica discrepâncias de contagem
    If rxCnt_IED_SUB_Link3 < txCnt_IED_PUB_Link3 Then
        Application.GetObject("DoS_Detected_Link3").Value = True
    ElseIf rxCnt_IED_SUB_Link3 > txCnt_IED_PUB_Link3 Then
        Application.GetObject("DoS_Detected_Link3").Value = True
    ElseIf outOv_IED_PUB_Link3 > 0 Then
        Application.GetObject("DoS_Detected_Link3").Value = True
    End If
    
    previous_t_Link3 = current_t_Link3
End Sub
