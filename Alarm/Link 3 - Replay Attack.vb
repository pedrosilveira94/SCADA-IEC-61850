' Link 3 - Replay Attack Detection

' Variáveis iniciais
Dim previous_stnum_Link3
Dim last_timestamp_Link3
Dim max_stNum_Link3
Dim TTL_Link3
Dim stNumIED_PUB_GOOSE_Link3
Dim stNumIED_PUB_MMS_Link3
Dim statusIED_PUB_GOOSE_Link3
Dim statusIED_PUB_MMS_Link3
Dim message_timestamp_Link3

' Inicialização
Set previous_stnum_Link3 = CreateObject("Scripting.Dictionary")
Set last_timestamp_Link3 = CreateObject("Scripting.Dictionary")
max_stNum_Link3 = 4294967295 ' Valor máximo de 32 bits
TTL_Link3 = 2 ' Defina conforme necessário (em segundos)

Sub Replay_Attack_Detection_Link3()
    ' Obtenha valores das mensagens GOOSE e via MMS
    stNumIED_PUB_GOOSE_Link3 = Application.GetObject("stNumIED_PUB_GOOSE_Link3").Value
    statusIED_PUB_GOOSE_Link3 = Application.GetObject("statusIED_PUB_GOOSE_Link3").Value
    statusIED_PUB_MMS_Link3 = Application.GetObject("statusIED_PUB_MMS_Link3").Value
    message_timestamp_Link3 = Now
    
    Dim GoID
    GoID = "Link3"
    
    If previous_stnum_Link3.Exists(GoID) Then
        If DateDiff("s", message_timestamp_Link3, Now) > TTL_Link3 Then
            If stNumIED_PUB_GOOSE_Link3 < previous_stnum_Link3(GoID) Then
                Application.GetObject("Replay_Detected_Link3").Value = True
            End If
        Else
            If stNumIED_PUB_GOOSE_Link3 < previous_stnum_Link3(GoID) Then
                If (previous_stnum_Link3(GoID) - stNumIED_PUB_GOOSE_Link3) > (max_stNum_Link3 / 2) Then
                    ' Rollover legítimo
                    stNumIED_PUB_GOOSE_Link3 = stNumIED_PUB_GOOSE_Link3 + max_stNum_Link3 + 1
                Else
                    Application.GetObject("Replay_Detected_Link3").Value = True
                End If
            End If
        End If
        
        If statusIED_PUB_MMS_Link3 <> statusIED_PUB_GOOSE_Link3 Then
            Application.GetObject("Replay_Detected_Link3").Value = True
        End If
        
        If message_timestamp_Link3 < last_timestamp_Link3(GoID) Then
            Application.GetObject("Replay_Detected_Link3").Value = True
        End If
    Else
        ' Processar mensagem
    End If
    
    previous_stnum_Link3(GoID) = stNumIED_PUB_GOOSE_Link3
    last_timestamp_Link3(GoID) = message_timestamp_Link3
End Sub
