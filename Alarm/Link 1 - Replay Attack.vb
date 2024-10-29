' Link 1 - Replay Attack Detection

' Variáveis iniciais
Dim previous_stnum_Link1
Dim last_timestamp_Link1
Dim max_stNum_Link1
Dim TTL_Link1
Dim stNumIED_PUB_GOOSE_Link1
Dim stNumIED_PUB_MMS_Link1
Dim statusIED_PUB_GOOSE_Link1
Dim statusIED_PUB_MMS_Link1
Dim message_timestamp_Link1

' Inicialização
Set previous_stnum_Link1 = CreateObject("Scripting.Dictionary")
Set last_timestamp_Link1 = CreateObject("Scripting.Dictionary")
max_stNum_Link1 = 4294967295 ' Valor máximo de 32 bits
TTL_Link1 = 2 ' Defina conforme necessário (em segundos)

Sub Replay_Attack_Detection_Link1()
    ' Obtenha valores das mensagens GOOSE e via MMS
    stNumIED_PUB_GOOSE_Link1 = Application.GetObject("stNumIED_PUB_GOOSE_Link1").Value
    statusIED_PUB_GOOSE_Link1 = Application.GetObject("statusIED_PUB_GOOSE_Link1").Value
    statusIED_PUB_MMS_Link1 = Application.GetObject("statusIED_PUB_MMS_Link1").Value
    message_timestamp_Link1 = Now
    
    Dim GoID
    GoID = "Link1"
    
    If previous_stnum_Link1.Exists(GoID) Then
        If DateDiff("s", message_timestamp_Link1, Now) > TTL_Link1 Then
            If stNumIED_PUB_GOOSE_Link1 < previous_stnum_Link1(GoID) Then
                Application.GetObject("Replay_Detected_Link1").Value = True
            End If
        Else
            If stNumIED_PUB_GOOSE_Link1 < previous_stnum_Link1(GoID) Then
                If (previous_stnum_Link1(GoID) - stNumIED_PUB_GOOSE_Link1) > (max_stNum_Link1 / 2) Then
                    ' Rollover legítimo
                    stNumIED_PUB_GOOSE_Link1 = stNumIED_PUB_GOOSE_Link1 + max_stNum_Link1 + 1
                Else
                    Application.GetObject("Replay_Detected_Link1").Value = True
                End If
            End If
        End If
        
        If statusIED_PUB_MMS_Link1 <> statusIED_PUB_GOOSE_Link1 Then
            Application.GetObject("Replay_Detected_Link1").Value = True
        End If
        
        If message_timestamp_Link1 < last_timestamp_Link1(GoID) Then
            Application.GetObject("Replay_Detected_Link1").Value = True
        End If
    Else
        ' Processar mensagem
    End If
    
    previous_stnum_Link1(GoID) = stNumIED_PUB_GOOSE_Link1
    last_timestamp_Link1(GoID) = message_timestamp_Link1
End Sub
