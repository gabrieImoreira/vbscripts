'call CalculaDiaUtilAnterior("28/02/2022") ' simula data

Function CalculaDiaUtilAnterior(var)

	On Error Resume Next
		arrayDS = array("Domingo","Segunda-Feira","Terça-Feira","Quarta-Feira","Quinta-Feira","Sexta-Feira","Sábado")
		dataInput = CDate(var)
		if(weekday(dataInput) = 1)then
			diadasemana = "Domingo"
		else
			diadasemana = arrayDS(weekday(dataInput-1))
		end if
	if(diadasemana ="Segunda-Feira")then
		data = dataInput - 3
	elseif (diadasemana ="Sábado")then
		data = dataInput - 1
	elseif (diadasemana ="Domingo")then
		data = dataInput - 2
	else
		data = dataInput - 1
	end if

    	If Err.Number <> 0 Then
        	Dim res
        	res = "ERRO, número do erro:" & CStr(Err.Number) & ", Descrição do erro:" & CStr(Err.Description)
        	CalculaDiaUtilAnterior = res
    	Else
    	'msgbox(data)
	CalculaDiaUtilAnterior = data
	end if
End Function

Function isDiaUtil(var)
	
	On Error Resume Next
		arrayDS = array("Domingo","Segunda-Feira","Terça-Feira","Quarta-Feira","Quinta-Feira","Sexta-Feira","Sábado")
		dataInput = CDate(var)
		if(weekday(dataInput) = 1)then
			diadasemana = "Domingo"
		else
			diadasemana = arrayDS(weekday(dataInput-1))
		end if
		if(diadasemana ="Sábado" OR diadasemana ="Domingo")then
			isDiaUtil = 0
		else
			isDiaUtil = 1
		end if
	
    	If Err.Number <> 0 Then
        	Dim res
        	res = "ERRO, número do erro:" & CStr(Err.Number) & ", Descrição do erro:" & CStr(Err.Description)
        	isDiaUtil = res
	end if
End Function

Function isMonday(var)

	On Error Resume Next
		arrayDS = array("Domingo","Segunda-Feira","TerÃ§a-Feira","Quarta-Feira","Quinta-Feira","Sexta-Feira","SÃ¡bado")
		dataInput = CDate(var)
		if(weekday(dataInput) = 1)then
			diadasemana = "Domingo"
		else
			diadasemana = arrayDS(weekday(dataInput-1))
		end if
		if(diadasemana ="Segunda-Feira")then
			isMonday = 1
		else
			isMonday = 0
		end if

    	If Err.Number <> 0 Then
        	Dim res
        	res = "ERRO, nÃºmero do erro:" & CStr(Err.Number) & ", DescriÃ§Ã£o do erro:" & CStr(Err.Description)
        	isMonday = res
	end if
	msgbox(isMonday)
End Function
