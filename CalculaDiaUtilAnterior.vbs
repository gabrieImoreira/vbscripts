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
