' Link 1 - High stNum Attack Detection

' Variáveis iniciais
Dim previous_stnum_Link1
Dim last_timestamp_Link1
Dim max_diff_Link1
Dim max_interval_Link1
Dim TTL_valid_Link1
Dim stNumIED_PUB_GOOSE_Link1
Dim stNumIED_PUB_MMS_Link1
Dim statusIED_PUB_GOOSE_Link1
Dim statusIED_PUB_MMS_Link1
Dim statusIED_SUB_GOOSE_Link1
Dim stNumIED_SUB_GOOSE_Link1
Dim timestamp_Link1

' Inicialização
Set previous_stnum_Link1 = CreateObject("Scripting.Dictionary")
Set last_timestamp_Link1 = CreateObject("Scripting.Dictionary")
max_diff_Link1 = 10 ' Defina conforme necessário
max_interval_Link1 = 1 ' Em segundos
TTL_valid_Link1 = 2 ' Defina conforme necessário (em segundos)

Sub HighStNum_Attack_Detection_Link1()
    ' Obtenha valores das mensagens GOOSE e via MMS
    stNumIED_PUB_GOOSE_Link1 = Application.GetObject("stNumIED_PUB_GOOSE_Link1").Value
    stNumIED_PUB_MMS_Link1 = Application.GetObject("stNumIED_PUB_MMS_Link1").Value
    statusIED_PUB_GOOSE_Link1 = Application.GetObject("statusIED_PUB_GOOSE_Link1").Value
    statusIED_PUB_MMS_Link1 = Application.GetObject("statusIED_PUB_MMS_Link1").Value
    statusIED_SUB_GOOSE_Link1 = Application.GetObject("statusIED_SUB_GOOSE_Link1").Value
    stNumIED_SUB_GOOSE_Link1 = Application.GetObject("stNumIED_SUB_GOOSE_Link1").Value
    timestamp_Link1 = Now

    ' Supondo GoID seja "Link1"
    Dim GoID
    GoID = "Link1"
    
    If previous_stnum_Link1.Exists(GoID) Then
        If (stNumIED_PUB_GOOSE_Link1 - previous_stnum_Link1(GoID) <= max_diff_Link1) And (DateDiff("s", last_timestamp_Link1(GoID), timestamp_Link1) <= max_interval_Link1) And (TTL_valid_Link1 >= 0) Then
            If statusIED_PUB_GOOSE_Link1 <> statusIED_PUB_MMS_Link1 Then
                Application.GetObject("HighStNum_Detected_Link1").Value = True
                ' Descartar mensagem
            End If
            If stNumIED_PUB_GOOSE_Link1 <> stNumIED_PUB_MMS_Link1 Then
                Application.GetObject("HighStNum_Detected_Link1").Value = True
                ' Descartar mensagem
            End If
            If (statusIED_PUB_GOOSE_Link1 <> statusIED_SUB_GOOSE_Link1) Or (stNumIED_PUB_GOOSE_Link1 <> stNumIED_SUB_GOOSE_Link1) Then
                Application.GetObject("HighStNum_Detected_Link1").Value = True
                ' Descartar mensagem
            End If
        ElseIf (stNumIED_PUB_GOOSE_Link1 - previous_stnum_Link1(GoID) > max_diff_Link1) Or (DateDiff("s", last_timestamp_Link1(GoID), timestamp_Link1) > max_interval_Link1) Then
            Application.GetObject("HighStNum_Detected_Link1").Value = True
            ' Descartar mensagem
        End If
    Else
        ' Processar mensagem normalmente
    End If
    
    previous_stnum_Link1(GoID) = stNumIED_PUB_GOOSE_Link1
    last_timestamp_Link1(GoID) = timestamp_Link1
End Sub
