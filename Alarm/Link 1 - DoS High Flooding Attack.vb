' Link 1 - DoS High Flooding Attack Detection

' Variáveis iniciais
Dim message_count_Link1
Dim previous_t_Link1
Dim current_t_Link1
Dim detection_window_Link1
Dim threshold_rate_Link1
Dim min_time_Link1
Dim TTL_Link1
Dim txCnt_IED_PUB_Link1
Dim rxCnt_IED_SUB_Link1
Dim outOv_IED_PUB_Link1
Dim inOv_IED_SUB_Link1
Dim ferCh_IED_SUB_Link1

' Inicialização
message_count_Link1 = 0
previous_t_Link1 = Now
detection_window_Link1 = 1 ' Em segundos
threshold_rate_Link1 = 100 ' Máximo de mensagens por segundo
min_time_Link1 = 0.01 ' Tempo mínimo entre mensagens (em segundos)
TTL_Link1 = 2 ' Defina conforme necessário (em segundos)

Sub DoS_Attack_Detection_Link1()
    ' Atualize as variáveis conforme necessário
    current_t_Link1 = Now ' Timestamp atual
    message_count_Link1 = message_count_Link1 + 1
    
    ' Obtenha os valores de txCnt, rxCnt, inOv, outOv, ferCh via tags do Elipse E3
    txCnt_IED_PUB_Link1 = Application.GetObject("txCnt_IED_PUB_Link1").Value
    rxCnt_IED_SUB_Link1 = Application.GetObject("rxCnt_IED_SUB_Link1").Value
    outOv_IED_PUB_Link1 = Application.GetObject("outOv_IED_PUB_Link1").Value
    inOv_IED_SUB_Link1 = Application.GetObject("inOv_IED_SUB_Link1").Value
    ferCh_IED_SUB_Link1 = Application.GetObject("ferCh_IED_SUB_Link1").Value
    
    ' Verificação de DoS Attack
    If message_count_Link1 > threshold_rate_Link1 Or (DateDiff("s", previous_t_Link1, current_t_Link1) < min_time_Link1) Or (DateDiff("s", previous_t_Link1, current_t_Link1) > TTL_Link1) Then
        Application.GetObject("DoS_Detected_Link1").Value = True
        ' Ação adicional, como descartar mensagens
    ElseIf inOv_IED_SUB_Link1 > 0 Then
        Application.GetObject("DoS_Detected_Link1").Value = True
    ElseIf ferCh_IED_SUB_Link1 > 0 Then
        Application.GetObject("DoS_Detected_Link1").Value = True
    End If
    
    ' Verifica discrepâncias de contagem
    If rxCnt_IED_SUB_Link1 < txCnt_IED_PUB_Link1 Then
        Application.GetObject("DoS_Detected_Link1").Value = True
    ElseIf rxCnt_IED_SUB_Link1 > txCnt_IED_PUB_Link1 Then
        Application.GetObject("DoS_Detected_Link1").Value = True
    ElseIf outOv_IED_PUB_Link1 > 0 Then
        Application.GetObject("DoS_Detected_Link1").Value = True
    End If
    
    previous_t_Link1 = current_t_Link1
End Sub
