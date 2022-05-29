
Private Sub btnSearch_Click()
noDatos = Hoja1.Range("B" & Rows.Count).End(xlUp).Row

aux = 0

listPerson = Clear
listPerson.RowSource = Clear

For fila = 2 To noDatos
nombre = Hoja1.Cells(fila, 2).Value
If nombre Like "*" & Me.txtSearch & "*" Then
Me.listPerson.AddItem
Me.listPerson.List(aux, 0) = Hoja1.Cells(fila, 1).Value
Me.listPerson.List(aux, 2) = Hoja1.Cells(fila, 2).Value
Me.listPerson.List(aux, 3) = Hoja1.Cells(fila, 3).Value
Me.listPerson.List(aux, 4) = Hoja1.Cells(fila, 4).Value
Me.listPerson.List(aux, 5) = Hoja1.Cells(fila, 5).Value
aux = aux + 1
End If
Next
End Sub

Private Sub CommandButton1_Click()
Application.ScreenUpdating = False
Workbooks.Add

Range("A1").Value = "Codigo"
Range("B1").Value = "Nombre"
Range("C1").Value = "Fecha de Nacimiento"
Range("D1").Value = "Correo Electrónico"
Range("E1").Value = "Domicilio"

For i = 0 To listPerson.ListCount - 1
    Range("A" & i + 2).Value = listPerson.List(i, 0)
    Range("B" & i + 2).Value = listPerson.List(i, 1)
    Range("C" & i + 2).Value = listPerson.List(i, 2)
    Range("D" & i + 2).Value = listPerson.List(i, 3)
    Range("E" & i + 2).Value = listPerson.List(i, 4)
    
Next i

MsgBox "Exportación Completa", vbExclamation

Application.ScreenUpdating = True
End Sub

Private Sub btnAdd_Click()
If Me.namePerson = "" Then
MsgBox ("Ingrese el nombre")
ElseIf Me.dateBirth = "" Then
MsgBox ("Ingrese la fecha de nacimiento")
Else
If IsDate(Me.dateBirth) = False Then
MsgBox ("Ingrese una fecha correcta")
Else
If (valida_email_fx(Me.email.Value)) Then
Sheets("BDPersona").Range("A2").EntireRow.Insert
Sheets("BDPersona").Range("A2").Value = Sheets("BDPersona").Range("G1").Value
Sheets("BDPersona").Range("B2").Value = Me.namePerson.Value
Sheets("BDPersona").Range("C2").Value = Me.dateBirth.Value
Sheets("BDPersona").Range("D2").Value = Me.email.Value
Sheets("BDPersona").Range("E2").Value = Me.domicile.Value
'codePerson.Value = Sheets("BDPersona").Range("G1").Value
Me.namePerson.Value = Empty
Me.dateBirth.Value = Empty
Me.email.Value = Empty
Me.domicile.Value = Empty

Me.listPerson.RowSource = "Persona"
Me.listPerson.ColumnCount = 5
Else
MsgBox ("Ingrese una correo correcto")
End If
End If
End If
End Sub

Private Sub btnDelete_Click()
codeSearch = Me.codePerson.Value

Set fila = Sheets("BDPersona").Range("A:A").Find(codeSearch, lookat:=xlWhole)
linea = fila.Row
Range("A" & linea).EntireRow.Delete
End Sub

Private Sub btnUpdate_Click()
If listPerson.ListIndex = -1 Then
MsgBox ("Selecciona un registro")
Else
Dim fila As Object
Dim linea As Integer

codeSearch = Me.codePerson

Set fila = Sheets("BDPersona").Range("A:A").Find(codeSearch, lookat:=xlWhole)
linea = fila.Row
Range("B" & linea).Value = Me.namePerson.Value
Range("C" & linea).Value = Me.dateBirth.Value
Range("D" & linea).Value = Me.email.Value
Range("E" & linea).Value = Me.domicile.Value

Me.namePerson.Value = Empty
Me.dateBirth.Value = Empty
Me.email.Value = Empty
Me.domicile.Value = Empty

End If
End Sub

Private Sub codePerson_Change()
Dim codigo As Integer
codigo = codePerson.Value
Me.namePerson = Application.WorksheetFunction.VLookup(codigo, Sheets("BDPersona").Range("A:E"), 2, 0)
Me.dateBirth = Application.WorksheetFunction.VLookup(codigo, Sheets("BDPersona").Range("A:E"), 3, 0)
Me.email = Application.WorksheetFunction.VLookup(codigo, Sheets("BDPersona").Range("A:E"), 4, 0)
Me.domicile = Application.WorksheetFunction.VLookup(codigo, Sheets("BDPersona").Range("A:E"), 5, 0)
End Sub

Private Sub Label1_Click()

End Sub

Private Sub listPerson_Change()

End Sub

Private Sub listPerson_Click()
    Dim codigo As Integer
    codigo = listPerson.List(listPerson.ListIndex, 0)
    Me.codePerson.Value = codigo
    'Me.namePerson.Value = listPerson.List(listPerson.ListIndex, 1)
    'Me.dateBirth.Value = listPerson.List(listPerson.ListIndex, 2)
    'Me.email.Value = listPerson.List(listPerson.ListIndex, 3)
    'Me.domicile.Value = listPerson.List(listPerson.ListIndex, 4)
    'Me.codePerson.Enabled = True
    
End Sub


Private Sub UserForm_Activate()
'Me.codePerson.Value = Sheets("BDPersona").Range("G1").Value
Me.listPerson.RowSource = "Persona"
Me.listPerson.ColumnCount = 5
Sheets("BDPersona").Visible = True
End Sub

Private Sub UserForm_Click()

End Sub

Function valida_email_fx(email As String) As Boolean
Application.Volatile
Dim oReg As RegExp
Set oReg = New RegExp

On Error GoTo ErrorHandler

oReg.Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
valida_email_fx = oReg.Test(email)

Set oReg = Nothing

Exit Function

'Si ocurre error
ErrorHandler:
MsgBox "Ha ocurrido un error: ", vbExclamation, "EXCELeINFO"
End Function
