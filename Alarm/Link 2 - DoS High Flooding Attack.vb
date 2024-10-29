' Link 2 - DoS High Flooding Attack Detection

' Variáveis iniciais
Dim message_count_Link2
Dim previous_t_Link2
Dim current_t_Link2
Dim detection_window_Link2
Dim threshold_rate_Link2
Dim min_time_Link2
Dim TTL_Link2
Dim txCnt_IED_PUB_Link2
Dim rxCnt_IED_SUB_Link2
Dim outOv_IED_PUB_Link2
Dim inOv_IED_SUB_Link2
Dim ferCh_IED_SUB_Link2

' Inicialização
message_count_Link2 = 0
previous_t_Link2 = Now
detection_window_Link2 = 1 ' Em segundos
threshold_rate_Link2 = 100 ' Máximo de mensagens por segundo
min_time_Link2 = 0.01 ' Tempo mínimo entre mensagens (em segundos)
TTL_Link2 = 2 ' Defina conforme necessário (em segundos)

Sub DoS_Attack_Detection_Link2()
    ' Atualize as variáveis conforme necessário
    current_t_Link2 = Now ' Timestamp atual
    message_count_Link2 = message_count_Link2 + 1
    
    ' Obtenha os valores de txCnt, rxCnt, inOv, outOv, ferCh via tags do Elipse E3
    txCnt_IED_PUB_Link2 = Application.GetObject("txCnt_IED_PUB_Link2").Value
    rxCnt_IED_SUB_Link2 = Application.GetObject("rxCnt_IED_SUB_Link2").Value
    outOv_IED_PUB_Link2 = Application.GetObject("outOv_IED_PUB_Link2").Value
    inOv_IED_SUB_Link2 = Application.GetObject("inOv_IED_SUB_Link2").Value
    ferCh_IED_SUB_Link2 = Application.GetObject("ferCh_IED_SUB_Link2").Value
    
    ' Verificação de DoS Attack
    If message_count_Link2 > threshold_rate_Link2 Or (DateDiff("s", previous_t_Link2, current_t_Link2) < min_time_Link2) Or (DateDiff("s", previous_t_Link2, current_t_Link2) > TTL_Link2) Then
        Application.GetObject("DoS_Detected_Link2").Value = True
        ' Ação adicional, como descartar mensagens
    ElseIf inOv_IED_SUB_Link2 > 0 Then
        Application.GetObject("DoS_Detected_Link2").Value = True
    ElseIf ferCh_IED_SUB_Link2 > 0 Then
        Application.GetObject("DoS_Detected_Link2").Value = True
    End If
    
    ' Verifica discrepâncias de contagem
    If rxCnt_IED_SUB_Link2 < txCnt_IED_PUB_Link2 Then
        Application.GetObject("DoS_Detected_Link2").Value = True
    ElseIf rxCnt_IED_SUB_Link2 > txCnt_IED_PUB_Link2 Then
        Application.GetObject("DoS_Detected_Link2").Value = True
    ElseIf outOv_IED_PUB_Link2 > 0 Then
        Application.GetObject("DoS_Detected_Link2").Value = True
    End If
    
    previous_t_Link2 = current_t_Link2
End Sub
