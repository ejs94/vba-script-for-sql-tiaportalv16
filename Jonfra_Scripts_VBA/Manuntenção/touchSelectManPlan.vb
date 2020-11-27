Sub touchSelectManPlan(ByRef pFieldNumber)

On Error Resume Next


SmartTags("edit_descr") = SmartTags("Descricao_Field_" & pFieldNumber)
SmartTags("edit_dt_mant") = SmartTags("Dia_Reservado_" & pFieldNumber)
SmartTags("edit_h_plan") = SmartTags("Hrs_Planejada_" & pFieldNumber)
SmartTags("edit_respons") = SmartTags("Responsavel_" & pFieldNumber)
SmartTags("Edit_ManPlan_ID") = SmartTags("ID_Manuntecao_" & pFieldNumber)

Select Case SmartTags("Maquina_Field_" & pFieldNumber)
    Case "Outro"
        SmartTags("edit_maqEqu") = 0
    Case "Est. Entrad."
        SmartTags("edit_maqEqu") = 1
    Case "Est. Saida."
        SmartTags("edit_maqEqu") = 2
    Case "MCH250"
        SmartTags("edit_maqEqu") = 3
    Case "MCH350"
        SmartTags("edit_maqEqu") = 4
    Case "G704"
        SmartTags("edit_maqEqu") = 5
    Case "G516"
        SmartTags("edit_maqEqu") = 6
    Case Else
        SmartTags("edit_maqEqu") = 0
End Select

Select Case SmartTags("TipoMaquina_Field_" & pFieldNumber)
    Case "Outro"
        SmartTags("edit_TipoManuten") = 0
    Case "Mecânico"
        SmartTags("edit_TipoManuten") = 1
    Case "Elétrico"
        SmartTags("edit_TipoManuten") = 2
    Case "Software"
        SmartTags("edit_TipoManuten") = 3   
    Case Else
        SmartTags("edit_TipoManuten") = 0
End Select

Select Case SmartTags("Prioridade_Field_" & pFieldNumber)
    Case "Baixa"
        SmartTags("edit_prior") = 0
    Case "Alta"
        SmartTags("edit_prior") = 1
    Case Else
        SmartTags("edit_prior") = 0
End Select

End Sub