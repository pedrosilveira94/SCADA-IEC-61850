' Link 3 - High stNum Attack Detection

' Variáveis iniciais
Dim previous_stnum_Link3
Dim last_timestamp_Link3
Dim max_diff_Link3
Dim max_interval_Link3
Dim TTL_valid_Link3
Dim stNumIED_PUB_GOOSE_Link3
Dim stNumIED_PUB_MMS_Link3
Dim statusIED_PUB_GOOSE_Link3
Dim statusIED_PUB_MMS_Link3
Dim statusIED_SUB_GOOSE_Link3
Dim stNumIED_SUB_GOOSE_Link3
Dim timestamp_Link3

' Inicialização
Set previous_stnum_Link3 = CreateObject("Scripting.Dictionary")
Set last_timestamp_Link3 = CreateObject("Scripting.Dictionary")
max_diff_Link3 = 10 ' Defina conforme necessário
max_interval_Link3 = 1 ' Em segundos
TTL_valid_Link3 = 2 ' Defina conforme necessário (em segundos)

Sub HighStNum_Attack_Detection_Link3()
    ' Obtenha valores das mensagens GOOSE e via MMS
    stNumIED_PUB_GOOSE_Link3 = Application.GetObject("stNumIED_PUB_GOOSE_Link3").Value
    stNumIED_PUB_MMS_Link3 = Application.GetObject("stNumIED_PUB_MMS_Link3").Value
    statusIED_PUB_GOOSE_Link3 = Application.GetObject("statusIED_PUB_GOOSE_Link3").Value
    statusIED_PUB_MMS_Link3 = Application.GetObject("statusIED_PUB_MMS_Link3").Value
    statusIED_SUB_GOOSE_Link3 = Application.GetObject("statusIED_SUB_GOOSE_Link3").Value
    stNumIED_SUB_GOOSE_Link3 = Application.GetObject("stNumIED_SUB_GOOSE_Link3").Value
    timestamp_Link3 = Now

    ' Supondo GoID seja "Link3"
    Dim GoID
    GoID = "Link3"
    
    If previous_stnum_Link3.Exists(GoID) Then
        If (stNumIED_PUB_GOOSE_Link3 - previous_stnum_Link3(GoID) <= max_diff_Link3) And (DateDiff("s", last_timestamp_Link3(GoID), timestamp_Link3) <= max_interval_Link3) And (TTL_valid_Link3 >= 0) Then
            If statusIED_PUB_GOOSE_Link3 <> statusIED_PUB_MMS_Link3 Then
                Application.GetObject("HighStNum_Detected_Link3").Value = True
                ' Descartar mensagem
            End If
            If stNumIED_PUB_GOOSE_Link3 <> stNumIED_PUB_MMS_Link3 Then
                Application.GetObject("HighStNum_Detected_Link3").Value = True
                ' Descartar mensagem
            End If
            If (statusIED_PUB_GOOSE_Link3 <> statusIED_SUB_GOOSE_Link3) Or (stNumIED_PUB_GOOSE_Link3 <> stNumIED_SUB_GOOSE_Link3) Then
                Application.GetObject("HighStNum_Detected_Link3").Value = True
                ' Descartar mensagem
            End If
        ElseIf (stNumIED_PUB_GOOSE_Link3 - previous_stnum_Link3(GoID) > max_diff_Link3) Or (DateDiff("s", last_timestamp_Link3(GoID), timestamp_Link3) > max_interval_Link3) Then
            Application.GetObject("HighStNum_Detected_Link3").Value = True
            ' Descartar mensagem
        End If
    Else
        ' Processar mensagem normalmente
    End If
    
    previous_stnum_Link3(GoID) = stNumIED_PUB_GOOSE_Link3
    last_timestamp_Link3(GoID) = timestamp_Link3
End Sub
