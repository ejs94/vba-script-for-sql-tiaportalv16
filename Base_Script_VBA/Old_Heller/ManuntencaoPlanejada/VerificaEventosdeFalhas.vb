Sub VerificaEventosdeFalhas()
'VERIFICA SE EXISTEM OCORRÊNCIAS DE MANUTENÇÃO PLANEJADA

'DECLARACAO DE TAGs
Dim Falhas_MP, Reset_MP
Dim ResetAtuado


On Error Resume Next

'STD_Log "SmartTags(Falhas_ManutPl)=" & SmartTags("Falhas_ManutPl")
'STD_Log "SmartTags(ResetAtuado)=" & SmartTags("ResetAtuado")
'STD_Log "SmartTags(StartMSGManutPlanejada)=" & SmartTags("StartMSGManutPlanejada")

If Not SmartTags("Falhas_ManutPl") And Not SmartTags("ResetAtuado") And SmartTags("StartMSGManutPlanejada") Then Exit Sub
'STD_Log "VerificaEventosFalhas Linha 13 - FalhasMatPlan=" & SmartTags("Falhas_ManutPl")


SmartTags("Falhas_ManutPl")=False
If SmartTags("SinalDeVida") Then SmartTags("StartMSGManutPlanejada")=True


If SmartTags("DB_0121_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_0221_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_0321_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_0421_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_0421_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_0521_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_0621_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_0721_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_0921_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_1021_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_1121_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_1221_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True Or _
SmartTags("DB_1321_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=True _
Then Falhas_MP = True Else Falhas_MP = False




'*************** LOG Manutenção Planejada Solicitada **********************
'STD_Log "DB0121 = " & SmartTags("DB_0121_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger") & _
'", DB0221 = " & SmartTags("DB_0221_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")& vbNewLine & _
'"DB0321 = " & SmartTags("DB_0321_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")  & _
'", DB0421 = " & SmartTags("DB_0421_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")  & vbNewLine & _
'"DB0521 = " & SmartTags("DB_0521_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger") & _
'", DB0621 = " & SmartTags("DB_0621_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger") & vbNewLine & _
'"DB0721 = " & SmartTags("DB_0721_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")





'**** Atualiza TAGs Falhas ****
SmartTags("Falhas_ManutPl") = Falhas_MP



'**************** Verifica Botão Reset Pressionado **************************
Reset_MP = SmartTags("Reset_ManutPl")


'***************************** Reset Falhas *********************************
If Reset_MP Then
	ResetAtuado=True
	Reset_MP=False
	SmartTags("ID_MSG")=0
	SmartTags("Falhas_ManutPl")=False
	SmartTags("DB_0121_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	SmartTags("DB_0221_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	SmartTags("DB_0321_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	SmartTags("DB_0421_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	SmartTags("DB_0521_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	SmartTags("DB_0621_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	SmartTags("DB_0721_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	SmartTags("DB_0921_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	SmartTags("DB_1021_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	SmartTags("DB_1121_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	SmartTags("DB_1221_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	SmartTags("DB_1321_ControleManutencaoPlanejada_IN_InterfaceIPC_Trigger")=False
	
	'Reset Alarmes
	SmartTags("DB_0121_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	SmartTags("DB_0221_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	SmartTags("DB_0321_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	SmartTags("DB_0421_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	SmartTags("DB_0521_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	SmartTags("DB_0621_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	SmartTags("DB_0721_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	SmartTags("DB_0921_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	SmartTags("DB_1021_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	SmartTags("DB_1121_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	SmartTags("DB_1221_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	SmartTags("DB_1321_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=True
	
Else
	SmartTags("DB_0121_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=False
	SmartTags("DB_0221_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=False
	SmartTags("DB_0321_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=False
	SmartTags("DB_0421_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=False
	SmartTags("DB_0521_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=False
	SmartTags("DB_0621_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=False
	SmartTags("DB_0721_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=False
	SmartTags("DB_0921_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=False
	SmartTags("DB_1021_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=False
	SmartTags("DB_1121_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=False
	SmartTags("DB_1221_ControleManutencaoPlanejada_IN_InterfaceIPC_ResetAlarme")=False
End If




SmartTags("Reset_ManutPl")=Reset_MP


If ResetAtuado Then SmartTags("ResetAtuado")=True Else SmartTags("ResetAtuado")=False

End Sub