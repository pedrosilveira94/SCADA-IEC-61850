' Link 2 - Replay Attack Detection

' Variáveis iniciais
Dim previous_stnum_Link2
Dim last_timestamp_Link2
Dim max_stNum_Link2
Dim TTL_Link2
Dim stNumIED_PUB_GOOSE_Link2
Dim stNumIED_PUB_MMS_Link2
Dim statusIED_PUB_GOOSE_Link2
Dim statusIED_PUB_MMS_Link2
Dim message_timestamp_Link2

' Inicialização
Set previous_stnum_Link2 = CreateObject("Scripting.Dictionary")
Set last_timestamp_Link2 = CreateObject("Scripting.Dictionary")
max_stNum_Link2 = 4294967295 ' Valor máximo de 32 bits
TTL_Link2 = 2 ' Defina conforme necessário (em segundos)

Sub Replay_Attack_Detection_Link2()
    ' Obtenha valores das mensagens GOOSE e via MMS
    stNumIED_PUB_GOOSE_Link2 = Application.GetObject("stNumIED_PUB_GOOSE_Link2").Value
    statusIED_PUB_GOOSE_Link2 = Application.GetObject("statusIED_PUB_GOOSE_Link2").Value
    statusIED_PUB_MMS_Link2 = Application.GetObject("statusIED_PUB_MMS_Link2").Value
    message_timestamp_Link2 = Now
    
    Dim GoID
    GoID = "Link2"
    
    If previous_stnum_Link2.Exists(GoID) Then
        If DateDiff("s", message_timestamp_Link2, Now) > TTL_Link2 Then
            If stNumIED_PUB_GOOSE_Link2 < previous_stnum_Link2(GoID) Then
                Application.GetObject("Replay_Detected_Link2").Value = True
            End If
        Else
            If stNumIED_PUB_GOOSE_Link2 < previous_stnum_Link2(GoID) Then
                If (previous_stnum_Link2(GoID) - stNumIED_PUB_GOOSE_Link2) > (max_stNum_Link2 / 2) Then
                    ' Rollover legítimo
                    stNumIED_PUB_GOOSE_Link2 = stNumIED_PUB_GOOSE_Link2 + max_stNum_Link2 + 1
                Else
                    Application.GetObject("Replay_Detected_Link2").Value = True
                End If
            End If
        End If
        
        If statusIED_PUB_MMS_Link2 <> statusIED_PUB_GOOSE_Link2 Then
            Application.GetObject("Replay_Detected_Link2").Value = True
        End If
        
        If message_timestamp_Link2 < last_timestamp_Link2(GoID) Then
            Application.GetObject("Replay_Detected_Link2").Value = True
        End If
    Else
        ' Processar mensagem
    End If
    
    previous_stnum_Link2(GoID) = stNumIED_PUB_GOOSE_Link2
    last_timestamp_Link2(GoID) = message_timestamp_Link2
End Sub
