' Link 3 - Configuration Attack Detection

' Variáveis iniciais
Dim previous_GoCBRef_Link3
Dim previous_SimSt_Link3
Dim previous_NdsCom_Link3
Dim previous_confRev_Link3
Dim previous_numDatSetEntries_Link3
Dim GoCBRef_Link3, SimSt_Link3, NdsCom_Link3, confRev_Link3, numDatSetEntries_Link3
Dim GoCBRef_MMS_Link3, NdsCom_MMS_Link3, confRev_MMS_Link3, numDatSetEntries_MMS_Link3
Dim TmSyn_Link3, ChLiv_Link3, PhyHealth_Link3

' Inicialização
Set previous_GoCBRef_Link3 = CreateObject("Scripting.Dictionary")
Set previous_SimSt_Link3 = CreateObject("Scripting.Dictionary")
Set previous_NdsCom_Link3 = CreateObject("Scripting.Dictionary")
Set previous_confRev_Link3 = CreateObject("Scripting.Dictionary")
Set previous_numDatSetEntries_Link3 = CreateObject("Scripting.Dictionary")

Sub Configuration_Attack_Detection_Link3()
    ' Obtenha os valores atuais via GOOSE
    GoCBRef_Link3 = Application.GetObject("GoCBRef_Link3").Value
    NdsCom_Link3 = Application.GetObject("NdsCom_Link3").Value
    confRev_Link3 = Application.GetObject("confRev_Link3").Value
    numDatSetEntries_Link3 = Application.GetObject("numDatSetEntries_Link3").Value
    SimSt_Link3 = Application.GetObject("SimSt_Link3").Value
    
    ' Obtenha valores via MMS
    GoCBRef_MMS_Link3 = Application.GetObject("GoCBRef_MMS_Link3").Value
    NdsCom_MMS_Link3 = Application.GetObject("NdsCom_MMS_Link3").Value
    confRev_MMS_Link3 = Application.GetObject("confRev_MMS_Link3").Value
    numDatSetEntries_MMS_Link3 = Application.GetObject("numDatSetEntries_MMS_Link3").Value
    
    ' Comparações entre GOOSE e MMS
    If GoCBRef_Link3 <> GoCBRef_MMS_Link3 Or NdsCom_Link3 <> NdsCom_MMS_Link3 Or confRev_Link3 <> confRev_MMS_Link3 Or numDatSetEntries_Link3 <> numDatSetEntries_MMS_Link3 Then
        Application.GetObject("Configuration_Detected_Link3").Value = True
    End If
    
    ' Supondo GoID seja "Link3"
    Dim GoID
    GoID = "Link3"
    
    ' Verificação de discrepâncias entre mensagens consecutivas
    If previous_GoCBRef_Link3.Exists(GoID) Then
        If GoCBRef_Link3 <> previous_GoCBRef_Link3(GoID) Or SimSt_Link3 <> previous_SimSt_Link3(GoID) Or NdsCom_Link3 <> previous_NdsCom_Link3(GoID) Or confRev_Link3 <> previous_confRev_Link3(GoID) Or numDatSetEntries_Link3 <> previous_numDatSetEntries_Link3(GoID) Then
            Application.GetObject("Configuration_Detected_Link3").Value = True
        End If
    End If
    
    ' Verifica saúde do sistema
    TmSyn_Link3 = Application.GetObject("TmSyn_Link3").Value
    ChLiv_Link3 = Application.GetObject("ChLiv_Link3").Value
    PhyHealth_Link3 = Application.GetObject("PhyHealth_Link3").Value
    
    If TmSyn_Link3 <> "synchronized" Then
        Application.GetObject("Configuration_Detected_Link3").Value = True
    End If
    
    If ChLiv_Link3 <> True Then
        Application.GetObject("Configuration_Detected_Link3").Value = True
    End If
    
    If PhyHealth_Link3 <> "good" Then
        Application.GetObject("Configuration_Detected_Link3").Value = True
    End If
    
    ' Atualiza valores anteriores
    previous_GoCBRef_Link3(GoID) = GoCBRef_Link3
    previous_SimSt_Link3(GoID) = SimSt_Link3
    previous_NdsCom_Link3(GoID) = NdsCom_Link3
    previous_confRev_Link3(GoID) = confRev_Link3
    previous_numDatSetEntries_Link3(GoID) = numDatSetEntries_Link3
End Sub
