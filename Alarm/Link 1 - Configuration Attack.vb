' Link 1 - Configuration Attack Detection

' Variáveis iniciais
Dim previous_GoCBRef_Link1
Dim previous_SimSt_Link1
Dim previous_NdsCom_Link1
Dim previous_confRev_Link1
Dim previous_numDatSetEntries_Link1
Dim GoCBRef_Link1, SimSt_Link1, NdsCom_Link1, confRev_Link1, numDatSetEntries_Link1
Dim GoCBRef_MMS_Link1, NdsCom_MMS_Link1, confRev_MMS_Link1, numDatSetEntries_MMS_Link1
Dim TmSyn_Link1, ChLiv_Link1, PhyHealth_Link1

' Inicialização
Set previous_GoCBRef_Link1 = CreateObject("Scripting.Dictionary")
Set previous_SimSt_Link1 = CreateObject("Scripting.Dictionary")
Set previous_NdsCom_Link1 = CreateObject("Scripting.Dictionary")
Set previous_confRev_Link1 = CreateObject("Scripting.Dictionary")
Set previous_numDatSetEntries_Link1 = CreateObject("Scripting.Dictionary")

Sub Configuration_Attack_Detection_Link1()
    ' Obtenha os valores atuais via GOOSE
    GoCBRef_Link1 = Application.GetObject("GoCBRef_Link1").Value
    NdsCom_Link1 = Application.GetObject("NdsCom_Link1").Value
    confRev_Link1 = Application.GetObject("confRev_Link1").Value
    numDatSetEntries_Link1 = Application.GetObject("numDatSetEntries_Link1").Value
    SimSt_Link1 = Application.GetObject("SimSt_Link1").Value
    
    ' Obtenha valores via MMS
    GoCBRef_MMS_Link1 = Application.GetObject("GoCBRef_MMS_Link1").Value
    NdsCom_MMS_Link1 = Application.GetObject("NdsCom_MMS_Link1").Value
    confRev_MMS_Link1 = Application.GetObject("confRev_MMS_Link1").Value
    numDatSetEntries_MMS_Link1 = Application.GetObject("numDatSetEntries_MMS_Link1").Value
    
    ' Comparações entre GOOSE e MMS
    If GoCBRef_Link1 <> GoCBRef_MMS_Link1 Or NdsCom_Link1 <> NdsCom_MMS_Link1 Or confRev_Link1 <> confRev_MMS_Link1 Or numDatSetEntries_Link1 <> numDatSetEntries_MMS_Link1 Then
        Application.GetObject("Configuration_Detected_Link1").Value = True
    End If
    
    ' Supondo GoID seja "Link1"
    Dim GoID
    GoID = "Link1"
    
    ' Verificação de discrepâncias entre mensagens consecutivas
    If previous_GoCBRef_Link1.Exists(GoID) Then
        If GoCBRef_Link1 <> previous_GoCBRef_Link1(GoID) Or SimSt_Link1 <> previous_SimSt_Link1(GoID) Or NdsCom_Link1 <> previous_NdsCom_Link1(GoID) Or confRev_Link1 <> previous_confRev_Link1(GoID) Or numDatSetEntries_Link1 <> previous_numDatSetEntries_Link1(GoID) Then
            Application.GetObject("Configuration_Detected_Link1").Value = True
        End If
    End If
    
    ' Verifica saúde do sistema
    TmSyn_Link1 = Application.GetObject("TmSyn_Link1").Value
    ChLiv_Link1 = Application.GetObject("ChLiv_Link1").Value
    PhyHealth_Link1 = Application.GetObject("PhyHealth_Link1").Value
    
    If TmSyn_Link1 <> "synchronized" Then
        Application.GetObject("Configuration_Detected_Link1").Value = True
    End If
    
    If ChLiv_Link1 <> True Then
        Application.GetObject("Configuration_Detected_Link1").Value = True
    End If
    
    If PhyHealth_Link1 <> "good" Then
        Application.GetObject("Configuration_Detected_Link1").Value = True
    End If
    
    ' Atualiza valores anteriores
    previous_GoCBRef_Link1(GoID) = GoCBRef_Link1
    previous_SimSt_Link1(GoID) = SimSt_Link1
    previous_NdsCom_Link1(GoID) = NdsCom_Link1
    previous_confRev_Link1(GoID) = confRev_Link1
    previous_numDatSetEntries_Link1(GoID) = numDatSetEntries_Link1
End Sub
