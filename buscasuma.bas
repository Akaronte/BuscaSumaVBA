Sub buscaSuma()
    
' Declaramos las variables'

   Dim gz As Boolean
   gz = True
   
   Dim obj As Double
   Dim monto As Double
     
   Dim f As Variant
   ReDim f(1 To 1)
   
   Dim c As Integer

   Dim c2 As Integer
   c2 = 1
   Dim cc As Integer
   cc = 1
    
   Dim n As Integer


   Dim nombreHoja As String
   nombreHoja = "RESULTADO"
   CrearHoja (nombreHoja)
   
   
   n = Sheets("hoja1").Range("A" & Rows.Count).End(xlUp).Row
   c = n - 1
   
   Dim unidades As Variant
   ReDim unidades(1 To n) As Double
   
   Dim tmp As Variant
   ReDim tmp(1 To n) As Double
   
   
   Dim cicle As Variant
   ReDim cicle(1 To c) As Double
   
   Dim tmpcicle As Variant
   ReDim tmpcicle(1 To c) As Double
   
   counter = 1
      For counter = 1 To n
        unidades(counter) = Worksheets("Hoja1").Cells(counter, 1).Value
   Next counter

   counter = 1
   While counter <= UBound(unidades)
      counter = counter + 1
   Wend

   obj = Worksheets("Hoja1").Cells(1, 2).Value

'Comanzamos a coger cada elemto y por cada posion rotado de la posion relativa de los demas elemtos'
For k = 1 To n
f = unidades(k)
cc = 1
    For m = 1 To n
        If m = k Then
        Else
        cicle(cc) = unidades(m)
        cc = cc + 1
        End If
    Next m
    
For p = 1 To n

    'creamos el array ciclo para rotarlo a cada posion respecto al elemento k en cada posicion p'
    tmp(p) = f
    c2 = 1
    For rc = 1 To c
        If p = rc Then
        Else
        tmp(rc) = cicle(c2)
        c2 = c2 + 1
        End If
    Next rc
    For pp = 1 To n
                    monto = 0
                    For ppp = 1 To n
                    monto = monto + tmp(ppp)
                    If monto = obj Then
                    'Se ha encontrado el objetivo'
                            gz = False
                            monto = 0
                            For enc = 1 To n
                            monto = monto + tmp(enc)
                            If monto > obj Then
                            Worksheets(nombreHoja).Cells(enc, 1).Value = tmp(enc)
                            Else
                            Worksheets(nombreHoja).Cells(enc, 1).Interior.ColorIndex = 5
                            Worksheets(nombreHoja).Cells(enc, 1).Value = tmp(enc)
                            End If
                            Worksheets(nombreHoja).Cells(enc, 2).Value = "=sum(" & Col_Letter(1) & "1:" & Col_Letter(1) & enc & ")"
                            Next enc

                    End If
'Worksheets(nombreHoja).Cells(fila, colum).Value = tmp(ppp)'
'Worksheets(nombreHoja).Cells(fila, colum + 1).Value = "=sum(" & Col_Letter(colum) & "1:" & Col_Letter(colum) & fila & ")"'
                    'fila = fila + 1'
                    Next ppp
                    'colum = colum + 1 + 1'
                    'fila = 1'
                    
                    'Rotamos el ciclo para la siguiente posion'
                    For rrcc = 1 To c
                        If rrcc < c Then
                        tmpcicle(rrcc) = cicle(rrcc + 1)
                        End If
                        If rrcc = c Then
                        tmpcicle(rrcc) = cicle(1)
                        End If
                    Next rrcc
                    For rrcc2 = 1 To c
                    cicle(rrcc2) = tmpcicle(rrcc2)
                    Next rrcc2
                    'recostruimos el array a comprobar'
                    c2 = 1
                    For rc3 = 1 To n
                    If p = rc3 Then
                    
                    Else
                    tmp(rc3) = cicle(c2)
                    c2 = c2 + 1
                    End If
                    Next rc3
                Next pp
    Next p


Next k

If gz Then
MsgBox "No se ha encontrado el objetivo"
Else
MsgBox "Se ha encontrado el objetivo"
End If
  

End Sub


Function CrearHoja(nombreHoja As String) As Boolean
     
    Dim existe As Boolean
     
    On Error Resume Next
    existe = (Worksheets(nombreHoja).Name <> "")
     
    If Not existe Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = nombreHoja
    End If
     
    CrearHoja = existe
     
End Function


Function Col_Letter(ByVal lngCol As Integer) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
