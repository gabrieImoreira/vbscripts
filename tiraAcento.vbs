Function TiraAcento(Palavra)
 CAcento = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
 SAcento = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
 Texto = ""
 If Palavra <> "" Then
  For X = 1 To Len(Palavra)
    Letra = Mid(Palavra, X, 1)
    Pos_Acento = InStr(CAcento, Letra)
    If Pos_Acento > 0 Then
    Letra = Mid(SAcento, Pos_Acento, 1)
    End If
    Texto = Texto & Letra
   Next
  TiraAcento = Texto
 End If
