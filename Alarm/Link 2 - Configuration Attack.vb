' Link 2 - Configuration Attack Detection

' Variáveis iniciais
Dim previous_GoCBRef_Link2
Dim previous_SimSt_Link2
Dim previous_NdsCom_Link2
Dim previous_confRev_Link2
Dim previous_numDatSetEntries_Link2
Dim GoCBRef_Link2, SimSt_Link2, NdsCom_Link2, confRev_Link2, numDatSetEntries_Link2
Dim GoCBRef_MMS_Link2, NdsCom_MMS_Link2, confRev_MMS_Link2, numDatSetEntries_MMS_Link2
Dim TmSyn_Link2, ChLiv_Link2, PhyHealth_Link2

' Inicialização
Set previous_GoCBRef_Link2 = CreateObject("Scripting.Dictionary")
Set previous_SimSt_Link2 = CreateObject("Scripting.Dictionary")
Set previous_NdsCom_Link2 = CreateObject("Scripting.Dictionary")
Set previous_confRev_Link2 = CreateObject("Scripting.Dictionary")
Set previous_numDatSetEntries_Link2 = CreateObject("Scripting.Dictionary")

Sub Configuration_Attack_Detection_Link2()
    ' Obtenha os valores atuais via GOOSE
    GoCBRef_Link2 = Application.GetObject("GoCBRef_Link2").Value
    NdsCom_Link2 = Application.GetObject("NdsCom_Link2").Value
    confRev_Link2 = Application.GetObject("confRev_Link2").Value
    numDatSetEntries_Link2 = Application.GetObject("numDatSetEntries_Link2").Value
    SimSt_Link2 = Application.GetObject("SimSt_Link2").Value
    
    ' Obtenha valores via MMS
    GoCBRef_MMS_Link2 = Application.GetObject("GoCBRef_MMS_Link2").Value
    NdsCom_MMS_Link2 = Application.GetObject("NdsCom_MMS_Link2").Value
    confRev_MMS_Link2 = Application.GetObject("confRev_MMS_Link2").Value
    numDatSetEntries_MMS_Link2 = Application.GetObject("numDatSetEntries_MMS_Link2").Value
    
    ' Comparações entre GOOSE e MMS
    If GoCBRef_Link2 <> GoCBRef_MMS_Link2 Or NdsCom_Link2 <> NdsCom_MMS_Link2 Or confRev_Link2 <> confRev_MMS_Link2 Or numDatSetEntries_Link2 <> numDatSetEntries_MMS_Link2 Then
        Application.GetObject("Configuration_Detected_Link2").Value = True
    End If
    
    ' Supondo GoID seja "Link2"
    Dim GoID
    GoID = "Link2"
    
    ' Verificação de discrepâncias entre mensagens consecutivas
    If previous_GoCBRef_Link2.Exists(GoID) Then
        If GoCBRef_Link2 <> previous_GoCBRef_Link2(GoID) Or SimSt_Link2 <> previous_SimSt_Link2(GoID) Or NdsCom_Link2 <> previous_NdsCom_Link2(GoID) Or confRev_Link2 <> previous_confRev_Link2(GoID) Or numDatSetEntries_Link2 <> previous_numDatSetEntries_Link2(GoID) Then
            Application.GetObject("Configuration_Detected_Link2").Value = True
        End If
    End If
    
    ' Verifica saúde do sistema
    TmSyn_Link2 = Application.GetObject("TmSyn_Link2").Value
    ChLiv_Link2 = Application.GetObject("ChLiv_Link2").Value
    PhyHealth_Link2 = Application.GetObject("PhyHealth_Link2").Value
    
    If TmSyn_Link2 <> "synchronized" Then
        Application.GetObject("Configuration_Detected_Link2").Value = True
    End If
    
    If ChLiv_Link2 <> True Then
        Application.GetObject("Configuration_Detected_Link2").Value = True
    End If
    
    If PhyHealth_Link2 <> "good" Then
        Application.GetObject("Configuration_Detected_Link2").Value = True
    End If
    
    ' Atualiza valores anteriores
    previous_GoCBRef_Link2(GoID) = GoCBRef_Link2
    previous_SimSt_Link2(GoID) = SimSt_Link2
    previous_NdsCom_Link2(GoID) = NdsCom_Link2
    previous_confRev_Link2(GoID) = confRev_Link2
    previous_numDatSetEntries_Link2(GoID) = numDatSetEntries_Link2
End Sub
