VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPagenda 
   Caption         =   "Pagenda"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11580
   OleObjectBlob   =   "Pagendagihub.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPagenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'codigo que abre, preenche e fecha documento word
Private Sub btn_montar_contrato_Click()

    Dim Word As Word.Application
    Dim DOC As Word.Document
    Dim i As Integer
    Dim somaum As Integer
    somaum = 1
    Set Word = CreateObject("Word.Application")
    Word.Visible = True
    
    Set DOC = Word.Documents.Open("C:\Users\AmericoMaia\Desktop\Pagenda\mod77.docx")
    
    For i = 1 To 31
    With DOC
        '*Dados entrada
        .Application.Selection.Find.Text = "#E" & i & "#"
        .Application.Selection.Find.Execute

        .Application.Selection.Range = ActiveSheet.Range("B" & i + somaum).Value
       
        End With
        
        Next i
      
       With DOC
      .SaveAs ("C:\Users\AmericoMaia\Desktop\Pagenda\mod77_a.docx")
       .Close
       End With
    
    
    
    
    
    Set DOC = Word.Documents.Open("C:\Users\AmericoMaia\Desktop\Pagenda\mod77_a.docx")
     For i = 1 To 31
        
  With DOC
        '*Dados entrada
        .Application.Selection.Find.Text = "#S" & i & "#"
        .Application.Selection.Find.Execute
      
         ActiveSheet.Range("BB" & i).Value = ActiveSheet.Range("BB" & i + somaum).Value
        .Application.Selection.Range = ActiveSheet.Range("BB" & i).Value
        
        End With
      
     
        Next i
    
     With DOC
      .SaveAs ("C:\Users\AmericoMaia\Desktop\Pagenda\mod77_b.docx")
       .Close
       End With

    
    
    Word.Quit
    Set DOC = Nothing
    Set Word = Nothing



End Sub

Private Sub cmdAnterior_Click()
ActiveCell.Offset(-1, 0).Select
If ActiveCell = "ID" Then
ActiveCell.Offset(1, 0).Select
CarregarDadosNoFormulario
MsgBox "Você encontra-se no primeiro registo", vbInformation, "Page"
Else
CarregarDadosNoFormulario
End If
End Sub

Private Sub cmdExcluir_Click()
If MsgBox("Você tem a certeza que deseja excluir este este item?", vbYesNo, "Confirmação") = vbYes Then
ActiveCell.Offset(0, 1).Value = Empty
ActiveCell.Offset(0, 2).Value = Empty
ActiveCell.Offset(0, 3).Value = Empty
ActiveCell.Offset(0, 4).Value = Empty
ActiveCell.Offset(0, 5).Value = Empty
ActiveCell.Offset(0, 52).Value = Empty
ActiveCell.Offset(0, 126).Value = Empty
ActiveCell.Offset(0, 7).Value = Empty
ActiveCell.Offset(0, 8).Value = Empty
ActiveCell.Offset(0, 9).Value = Empty
ActiveCell.Offset(0, 10).Value = Empty
ActiveCell.Offset(0, 11).Value = Empty
ActiveCell.Offset(0, 12).Value = Empty
ActiveCell.Offset(0, 13).Value = Empty
ActiveCell.Offset(0, 24).Value = Empty
ActiveCell.Offset(0, 36).Value = Empty
ActiveCell.Offset(0, 50).Value = Empty
ActiveCell.Offset(0, 51).Value = Empty
ActiveCell.Offset(0, 127).Value = Empty
ActiveCell.Offset(0, 52).Value = Empty
ActiveCell.Offset(0, 53).Value = Empty

limparCampos
ActiveWorkbook.Save
MsgBox "A exclusão foi efetuada com sucesso", vbInformation, "Confirmação"
End If

End Sub

Private Sub cmdPróximo_Click()
ActiveCell.Offset(1, 0).Select
If ActiveCell = "" Then
ActiveCell.Offset(-1, 0).Select

MsgBox "Você está no ultimo registo", vbInformation, "Pagenda"

Else
CarregarDadosNoFormulario
End If

End Sub


Private Sub cmndActualizar_Click()
ActiveCell.Offset(0, 1).Value = txtnome.Text
ActiveCell.Offset(0, 2).Value = txtDD1.Text
ActiveCell.Offset(0, 3).Value = txtTelef1.Text
ActiveCell.Offset(0, 4).Value = txtDD2.Text
ActiveCell.Offset(0, 5).Value = txtTelef2.Text
ActiveCell.Offset(0, 51).Value = txtObservacoes.Text
ActiveCell.Offset(0, 6).Value = txtEndereco.Text
Me.Image1.Picture = LoadPicture(ActiveCell.Offset(0, 6).Value)
ActiveCell.Offset(0, 7).Value = TxtLocalFactos.Text
ActiveCell.Offset(0, 8).Value = TxtNomeArguido.Text
ActiveCell.Offset(0, 9).Value = TxtMorada.Text
ActiveCell.Offset(0, 10).Value = TxtCodPostal.Text
ActiveCell.Offset(0, 11).Value = TxtLocalidade.Text
ActiveCell.Offset(0, 12).Value = TxtPais.Text
ActiveCell.Offset(0, 13).Value = TxtViolacaoSubAlineaArt.Text
ActiveCell.Offset(0, 24).Value = TxtSubAlineaArtPunivel.Text
ActiveCell.Offset(0, 36).Value = TxtDescricaoFactos.Text
ActiveCell.Offset(0, 50).Value = TxtNomeInstrutor.Text
ActiveCell.Offset(0, 52).Value = txtDia.Text
lblDia.Caption = "Dia " & ActiveCell.Offset(0, 52).Value
MsgBox "Actualização realizada com sucesso", vbInformation, "Page"
End Sub

Private Sub cmndBusca_Click()

If txtLocalizar.Text = "" Then
MsgBox "Digite um valor a pesquisar", vbInformation, "Pagenda"
Exit Sub

Else
txtLocalizar = UCase(txtLocalizar)
Range("A2").Select                       'Na folha numero1'
End If
'----------------------------------------------------------------------------------------------'
                                         'Inicio do Loop do comando Busca'
Do
If IsNumeric(ActiveCell) Then
ActiveCell.Offset(0, 1).Select
If ActiveCell.Value = txtLocalizar.Value Then
Exit Do
End If
ActiveCell.Offset(1, -1).Select
If IsEmpty(ActiveCell) Then
MsgBox "O registo da sua procura não foi encontrado!", vbInformation, "Pagenda"
Exit Do
Exit Sub
End If

Else
Exit Do
End If

Loop Until ActiveCell.Text = txtLocalizar.Text
         'FINAL DO LOOP'
'----------------------------------------------------------------------------------------------'
On Error Resume Next
ActiveCell.Offset(0, -1).Select
 CarregarDadosNoFormulario

End Sub

Private Sub cmndImagem_Click()
Dim enderecoimg As String
enderecoimg = Application.GetOpenFilename(filefilter:="Picture Filco,*.ico;*.bpm")
txtEndereco.Text = enderecoimg
'colocar if se enderecoimg=null then
Me.Image1.Picture = LoadPicture(enderecoimg)
Image1.PictureSizeMode = fmPictureSizeModeStretch
End Sub

Private Sub cmndNovo_Click()
limparCampos
txtnome.SetFocus
End Sub

Private Sub cmndSalvar_Click()
Range("A2").Select
Do
    If Not (IsEmpty(ActiveCell)) Then
    ActiveCell.Offset(1, 0).Select
    End If
Loop Until IsEmpty(ActiveCell) = True
    numeracaocontactos                       'procedimento de numeraçao automatica'
    converterCaracteres                      'procedimento converte minusculas em maiusculas.'
ActiveCell.Offset(0, 1).Value = txtnome.Text
ActiveCell.Offset(0, 2).Value = txtDD1.Text
ActiveCell.Offset(0, 3).Value = txtTelef1.Text
ActiveCell.Offset(0, 4).Value = txtDD2.Text
ActiveCell.Offset(0, 5).Value = txtTelef2.Text
Me.Image1.Picture = LoadPicture(ActiveCell.Offset(0, 6).Value)
ActiveCell.Offset(0, 51).Value = txtObservacoes.Text
txtEndereco.Text = ActiveCell.Offset(0, 126).Value
ActiveCell.Offset(0, 7).Value = TxtLocalFactos.Text
ActiveCell.Offset(0, 8).Value = TxtNomeArguido.Text
ActiveCell.Offset(0, 9).Value = TxtMorada.Text
ActiveCell.Offset(0, 10).Value = TxtCodPostal.Text
ActiveCell.Offset(0, 11).Value = TxtLocalidade.Text
ActiveCell.Offset(0, 12).Value = TxtPais.Text
ActiveCell.Offset(0, 13).Value = TxtViolacaoSubAlineaArt.Text
ActiveCell.Offset(0, 52).Value = txtDia.Text
lblDia.Caption = "Dia " & ActiveCell.Offset(0, 52).Value
Select Case ComboBox2.Value
Case Is = "Diploma VioladoI"
ActiveCell.Offset(0, 14).Value = ComboBox1.Value
Case Is = "Diploma VioladoII"
ActiveCell.Offset(0, 15).Value = ComboBox1.Value
Case Is = "Diploma VioladoIII"
ActiveCell.Offset(0, 16).Value = ComboBox1.Value
Case Is = "Diploma VioladoIV"
ActiveCell.Offset(0, 17).Value = ComboBox1.Value
Case Is = "Diploma VioladoV"
ActiveCell.Offset(0, 18).Value = ComboBox1.Value
Case Is = "Diploma VioladoVI"
ActiveCell.Offset(0, 19).Value = ComboBox1.Value
Case Is = "Diploma VioladoVII"
ActiveCell.Offset(0, 20).Value = ComboBox1.Value
Case Is = "Diploma VioladoVIII"
ActiveCell.Offset(0, 21).Value = ComboBox1.Value
Case Is = "Diploma VioladoIX"
ActiveCell.Offset(0, 22).Value = ComboBox1.Value
Case Is = "Diploma VioladoX"
ActiveCell.Offset(0, 23).Value = ComboBox1.Value
Case Else
0
End Select

ActiveCell.Offset(0, 24).Value = TxtSubAlineaArtPunivel.Text

Select Case ComboBox4.Value
Case Is = "Punivel Pelo DiplomaI"
ActiveCell.Offset(0, 25).Value = ComboBox3.Value
Case Is = "Punivel Pelo DiplomaII"
ActiveCell.Offset(0, 26).Value = ComboBox3.Value
Case Is = "Punivel Pelo DiplomaIII"
ActiveCell.Offset(0, 27).Value = ComboBox3.Value
Case Is = "Punivel Pelo DiplomaIV"
ActiveCell.Offset(0, 28).Value = ComboBox3.Value
Case Is = "Punivel Pelo DiplomaV"
ActiveCell.Offset(0, 29).Value = ComboBox3.Value
Case Is = "Punivel Pelo DiplomaVI"
ActiveCell.Offset(0, 30).Value = ComboBox3.Value
Case Is = "Punivel Pelo DiplomaVII"
ActiveCell.Offset(0, 31).Value = ComboBox3.Value
Case Is = "Punivel Pelo DiplomaVIII"
ActiveCell.Offset(0, 32).Value = ComboBox3.Value
Case Is = "Punivel Pelo DiplomaIX"
ActiveCell.Offset(0, 33).Value = ComboBox3.Value
Case Is = "Punivel Pelo DiplomaX"
ActiveCell.Offset(0, 34).Value = ComboBox3.Value
Case Else

End Select
ActiveCell.Offset(0, 35).Value = CbxCoimaAbstrato.Value
ActiveCell.Offset(0, 36).Value = TxtDescricaoFactos.Text
ActiveCell.Offset(0, 37).Value = CbxEspeciesAnimais.Value
ActiveCell.Offset(0, 38).Value = CbxDefesa.Value
ActiveCell.Offset(0, 49).Value = CbxApreciacaoInstrutor.Value
ActiveCell.Offset(0, 50).Value = TxtNomeInstrutor.Text

limparCampos
ActiveWorkbook.Save
MsgBox "Os seus daddos foram salvos com sucesso!", vbInformation, "Pagenda"
End Sub


Public Sub numeracaocontactos()
If IsNumeric(ActiveCell.Offset(-1, 0)) Then
ActiveCell = ActiveCell.Offset(-1, 0) + 1
Else
ActiveCell = 1
End If


End Sub


Public Sub limparCampos()
txtnome.Text = ""
txtDD1.Text = ""
txtTelef1.Text = ""
txtDD2.Text = ""
txtTelef2.Text = ""
txtEndereco.Text = ""
'Me.Image1.Picture = LoadPicture("")
TxtLocalFactos.Text = ""
TxtNomeArguido.Text = ""
TxtMorada.Text = ""
TxtCodPostal.Text = ""
TxtLocalidade.Text = ""
TxtPais.Text = ""
TxtViolacaoSubAlineaArt.Text = ""
TxtSubAlineaArtPunivel.Text = ""
TxtDescricaoFactos.Text = ""
TxtNomeInstrutor.Text = ""
txtObservacoes.Text = ""
txtDia.Text = ""
lblDia.Caption = ""
End Sub

Public Sub CarregarDadosNoFormulario()
txtnome.Text = ActiveCell.Offset(0, 1).Value
txtDD1.Text = ActiveCell.Offset(0, 2).Value
txtTelef1.Text = ActiveCell.Offset(0, 3).Value
txtDD2.Text = ActiveCell.Offset(0, 4).Value
txtTelef2.Text = ActiveCell.Offset(0, 5).Value
Me.Image1.Picture = LoadPicture(ActiveCell.Offset(0, 6).Value)
txtObservacoes.Text = ActiveCell.Offset(0, 51).Value
txtEndereco.Text = ActiveCell.Offset(0, 126).Value
txtDia.Text = ActiveCell.Offset(0, 52).Value
TxtLocalFactos.Text = ActiveCell.Offset(0, 7).Value
TxtNomeArguido.Text = ActiveCell.Offset(0, 8).Value
TxtMorada.Text = ActiveCell.Offset(0, 9).Value
TxtCodPostal.Text = ActiveCell.Offset(0, 10).Value
TxtLocalidade.Text = ActiveCell.Offset(0, 11).Value
TxtPais.Text = ActiveCell.Offset(0, 12).Value
TxtViolacaoSubAlineaArt.Text = ActiveCell.Offset(0, 13).Value


Image1.PictureSizeMode = fmPictureSizeModeStretch
ComboBox1.RowSource = "DD2: DD73"
ComboBox2.RowSource = "DU2: DU12"
ComboBox3.RowSource = "DD2: DD73"
ComboBox4.RowSource = "DV2: DV12"
TxtSubAlineaArtPunivel.Text = ActiveCell.Offset(0, 24).Value
CbxCoimaAbstrato.RowSource = "DE2:DE46"
CbxEspeciesAnimais.RowSource = "DF2:DF14"
CbxDefesa.RowSource = "DG2:DG4"
CbxApreciacaoInstrutor.RowSource = "DQ2:DQ6"
TxtDescricaoFactos.Text = ActiveCell.Offset(0, 36).Value
TxtNomeInstrutor.Text = ActiveCell.Offset(0, 50).Value
lblDia.Caption = "Dia " & ActiveCell.Offset(0, 52).Value
End Sub
'Este botao ´do word

Private Sub CommandButton1_Click()
btn_montar_contrato_Click
End Sub



'Estes comandos manipulam as textbox em horas e min

Private Sub txtDD1_Change()
txtDD1.Value = Format(txtDD1.Text, "HH:MM")
End Sub



Private Sub txtDD2_Change()
txtDD2.Value = Format(txtDD2.Text, "HH:MM")
End Sub


Private Sub txtnome_Change()
txtnome.Value = Format(txtnome.Text, "HH:MM")
End Sub



Private Sub txtTelef1_Change()
txtTelef1.Value = Format(txtTelef1.Text, "HH:MM")
End Sub


Private Sub txtTelef2_Change()
txtTelef2.Value = Format(txtTelef2.Text, "HH:MM")
End Sub

'Fim manipulação textbox__________________________________________________________________



Private Sub UserForm_Initialize()
Range("A2").Select
CarregarDadosNoFormulario
End Sub



'Public Sub converterCaracteres()
'txtnome = UCase(txtnome.Text)
'End Sub



