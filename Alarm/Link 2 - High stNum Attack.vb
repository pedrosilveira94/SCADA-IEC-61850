' Link 2 - High stNum Attack Detection

' Variáveis iniciais
Dim previous_stnum_Link2
Dim last_timestamp_Link2
Dim max_diff_Link2
Dim max_interval_Link2
Dim TTL_valid_Link2
Dim stNumIED_PUB_GOOSE_Link2
Dim stNumIED_PUB_MMS_Link2
Dim statusIED_PUB_GOOSE_Link2
Dim statusIED_PUB_MMS_Link2
Dim statusIED_SUB_GOOSE_Link2
Dim stNumIED_SUB_GOOSE_Link2
Dim timestamp_Link2

' Inicialização
Set previous_stnum_Link2 = CreateObject("Scripting.Dictionary")
Set last_timestamp_Link2 = CreateObject("Scripting.Dictionary")
max_diff_Link2 = 10 ' Defina conforme necessário
max_interval_Link2 = 1 ' Em segundos
TTL_valid_Link2 = 2 ' Defina conforme necessário (em segundos)

Sub HighStNum_Attack_Detection_Link2()
    ' Obtenha valores das mensagens GOOSE e via MMS
    stNumIED_PUB_GOOSE_Link2 = Application.GetObject("stNumIED_PUB_GOOSE_Link2").Value
    stNumIED_PUB_MMS_Link2 = Application.GetObject("stNumIED_PUB_MMS_Link2").Value
    statusIED_PUB_GOOSE_Link2 = Application.GetObject("statusIED_PUB_GOOSE_Link2").Value
    statusIED_PUB_MMS_Link2 = Application.GetObject("statusIED_PUB_MMS_Link2").Value
    statusIED_SUB_GOOSE_Link2 = Application.GetObject("statusIED_SUB_GOOSE_Link2").Value
    stNumIED_SUB_GOOSE_Link2 = Application.GetObject("stNumIED_SUB_GOOSE_Link2").Value
    timestamp_Link2 = Now

    ' Supondo GoID seja "Link2"
    Dim GoID
    GoID = "Link2"
    
    If previous_stnum_Link2.Exists(GoID) Then
        If (stNumIED_PUB_GOOSE_Link2 - previous_stnum_Link2(GoID) <= max_diff_Link2) And (DateDiff("s", last_timestamp_Link2(GoID), timestamp_Link2) <= max_interval_Link2) And (TTL_valid_Link2 >= 0) Then
            If statusIED_PUB_GOOSE_Link2 <> statusIED_PUB_MMS_Link2 Then
                Application.GetObject("HighStNum_Detected_Link2").Value = True
                ' Descartar mensagem
            End If
            If stNumIED_PUB_GOOSE_Link2 <> stNumIED_PUB_MMS_Link2 Then
                Application.GetObject("HighStNum_Detected_Link2").Value = True
                ' Descartar mensagem
            End If
            If (statusIED_PUB_GOOSE_Link2 <> statusIED_SUB_GOOSE_Link2) Or (stNumIED_PUB_GOOSE_Link2 <> stNumIED_SUB_GOOSE_Link2) Then
                Application.GetObject("HighStNum_Detected_Link2").Value = True
                ' Descartar mensagem
            End If
        ElseIf (stNumIED_PUB_GOOSE_Link2 - previous_stnum_Link2(GoID) > max_diff_Link2) Or (DateDiff("s", last_timestamp_Link2(GoID), timestamp_Link2) > max_interval_Link2) Then
            Application.GetObject("HighStNum_Detected_Link2").Value = True
            ' Descartar mensagem
        End If
    Else
        ' Processar mensagem normalmente
    End If
    
    previous_stnum_Link2(GoID) = stNumIED_PUB_GOOSE_Link2
    last_timestamp_Link2(GoID) = timestamp_Link2
End Sub
