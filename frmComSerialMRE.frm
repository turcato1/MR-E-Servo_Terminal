VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmComSerialMRE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComSerialMRE.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnLimpaEnviar 
      Caption         =   "Limpa string a enviar"
      Height          =   615
      Left            =   6480
      TabIndex        =   27
      Top             =   2220
      Width           =   1515
   End
   Begin VB.TextBox txtStNumber 
      Height          =   315
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   26
      Text            =   "0"
      Top             =   1320
      Width           =   315
   End
   Begin VB.CheckBox chkStNumber 
      Caption         =   "Usar Station Number:"
      Height          =   315
      Left            =   180
      TabIndex        =   25
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtStringReceb 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   60
      Top             =   120
   End
   Begin VB.CommandButton btnLimpaReceb 
      Caption         =   "Limpa string recebida"
      Height          =   615
      Left            =   6480
      TabIndex        =   19
      Top             =   3300
      Width           =   1515
   End
   Begin VB.CommandButton btnETX 
      Caption         =   $"frmComSerialMRE.frx":0342
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   17
      Top             =   1140
      Width           =   555
   End
   Begin VB.CommandButton btnSTX 
      Caption         =   $"frmComSerialMRE.frx":034C
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4260
      TabIndex        =   16
      Top             =   1140
      Width           =   555
   End
   Begin VB.CommandButton btnSOH 
      Caption         =   $"frmComSerialMRE.frx":0356
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Top             =   1140
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Caption         =   "Exemplo de construção de strings"
      Height          =   3195
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   8175
      Begin VB.CommandButton btnRESET 
         Caption         =   "RESET"
         Height          =   735
         Left            =   420
         TabIndex        =   29
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton btnJogMode 
         Caption         =   "Jog Mode"
         Height          =   735
         Left            =   6420
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton btnSairTestMode 
         Caption         =   "Sair TEST MODE"
         Enabled         =   0   'False
         Height          =   735
         Left            =   420
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton btnSTOP 
         Caption         =   "STOP"
         Enabled         =   0   'False
         Height          =   735
         Left            =   6000
         TabIndex        =   21
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton btnPositMode 
         Caption         =   "Positioning Mode"
         Height          =   735
         Left            =   6420
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   13
         Text            =   "4000"
         Top             =   1380
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   11
         Text            =   "131072"
         Top             =   960
         Width           =   2475
      End
      Begin VB.TextBox txtVeloc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   9
         Text            =   "1500"
         Top             =   540
         Width           =   2475
      End
      Begin VB.CommandButton btnREV 
         Caption         =   "REV"
         Enabled         =   0   'False
         Height          =   735
         Left            =   4320
         TabIndex        =   8
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton btnFWD 
         Caption         =   "FWD"
         Enabled         =   0   'False
         Height          =   735
         Left            =   2700
         TabIndex        =   7
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Acel/Desac. (ms):"
         Height          =   195
         Left            =   2040
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Deslocamento (pls):"
         Height          =   195
         Left            =   1920
         TabIndex        =   12
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Velocidade (rpm):"
         Height          =   195
         Left            =   2100
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.CommandButton btnEnviar 
      Caption         =   "Enviar"
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   1500
      Width           =   1515
   End
   Begin VB.TextBox txtStringEnvio 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1740
      Width           =   3015
   End
   Begin VB.CommandButton btnFecharPorta 
      Caption         =   "Fechar porta"
      Height          =   735
      Left            =   300
      TabIndex        =   1
      Top             =   2940
      Width           =   1335
   End
   Begin VB.CommandButton btnAbrirPorta 
      Caption         =   "Abrir porta"
      Height          =   735
      Left            =   300
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin MSCommLib.MSComm mscCOM1 
      Left            =   7800
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      ParitySetting   =   2
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Porta precisa ser COM1:, deixar servo-amp em 9600bps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   24
      Top             =   480
      Width           =   8295
   End
   Begin VB.Label Label6 
      Caption         =   "<SOH> [St Number (1)] [Cmd (2)] <STX> [Data No. (depende)] <ETX> [Chksum (2)]"
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   120
      Width           =   6075
   End
   Begin VB.Label Label2 
      Caption         =   "String Receb:"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "String Envio:"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmComSerialMRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programa de comunicação com o MR-E/J2S via porta serial (23/06/08)
'Este programa tem propósito teste em situações controladas
'O autor ou essa empresa não se responsabilizam pelo uso indevido e/ou
'com fins comerciais desse software
'Autor: Thiago Turcato do Rego <turcato@melcotec.com.br>
'
' <SOH> [St Number (1)] [Cmd (2)] <STX> [Data No. (depende)] <ETX> [Chksum (2)]

'Parameter:
' Command: 84 - parameter write
' Data No.: 00~54 - Num. parametro (total 8 chars)
' Data:     X Y nnnnnn
'           X -> 3: Write RAM / 0: Write EEPROM
'           Y -> Decimal point: 0 (no dec. point) ~ 5 (lower 5th digit)
' Ex. Escrever 6 no param. 3
' Data: 84 04 3 0 00006

'Positioning:
' Command: 90 - Desabilitar entr. fisicas
' Data No.: 00 Data: 1EA5
'
' Command: 8B - Escolher modo de operacao
' Data No.: 00 Data 0002 - Positioning

' Command: A0 - Write parameters
' Data No.: 10 - Velocid (pls/seg) Data: 8 chars
'           11 - Acel/Desac. (ms) Data: 8 chars
'           13 - Deslocamento (pls) Data: 8 chars
'
' Command: 92 - Start
' Data No.: 00  Data: 00000801 - FWD / SON
'               Data: 00001001 - REV / SON
'
' Command: A0 - Write param (Stop)
' Data No.: 15 - Stop Data: 1EA5

Option Explicit

Public sStNumber As String

Public Function fChkStation() As String
    
    If (chkStNumber.Value = 1) Then
      fChkStation = txtStNumber.Text
    Else
      fChkStation = ""
    End If

End Function

Public Function fSumCheck(sString As String) As String
 Dim i, iSum As Integer
 Dim sSum, sSum2 As String

  'Calcula o check sum
  iSum = 0
  For i = 1 To Len(sString)
    iSum = iSum + Asc(Mid(sString, i, 1))
  Next
  sSum = Hex$(iSum)
  
  If (Len(sSum) > 1) Then
    For i = 0 To 1
      sSum2 = Mid(sSum, (Len(sSum) - i), 1) + sSum2
    Next
      fSumCheck = sSum2
  Else
    fSumCheck = "0" + sSum
  End If
    
End Function

Public Function fEnviaMsg(sStation As String, sCommand As String, _
                           sDataNo As String, sData As String) As String
Dim sMensagem As String

 ' Protocolo generico:
 '<SOH> [St Number(1)] [Cmd(2)] <STX> [Data No.(depende)] <ETX> [Chksum (2)]

  sMensagem = ""
  
  If (mscCOM1.PortOpen) Then
    If (sStation <> "") Then sMensagem = sStation             '[St Number]
    If (sCommand <> "") Then sMensagem = sMensagem + sCommand '[Cmd]
    sMensagem = sMensagem + Chr(2)                            '<STX>
    If (sDataNo <> "") Then sMensagem = sMensagem + sDataNo   '[Data No.]
    If (sData <> "") Then sMensagem = sMensagem + sData       '[Data]
    sMensagem = sMensagem + Chr(3)                            '<ETX>
    If (mscCOM1.PortOpen) Then
      mscCOM1.Output = Chr(1) + sMensagem + fSumCheck(sMensagem) '[Check sum]
    End If
      fEnviaMsg = Chr(1) + sMensagem + fSumCheck(sMensagem)
  Else
    fEnviaMsg = "Error 1: Port not open"
  End If
    
End Function

Private Sub btnJogMode_Click()
 Dim bTestModeOK As Boolean
 
  Timer1.Enabled = False
  bTestModeOK = True
    
  'Desabilita as entradas físicas
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "90", "00", "1EA5") + vbCrLf
                   
  If (txtStringEnvio = "Error 1: Port not open") Then bTestModeOK = False
                     
  'Muda para o modo de operação para Jog Mode
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "8B", "00", "0001") + vbCrLf
                   
  If (txtStringEnvio = "Error 1: Port not open") Then bTestModeOK = False
  
  If (bTestModeOK) Then
    btnPositMode.Enabled = False
    btnJogMode.Enabled = False
    btnSairTestMode.Enabled = True
    btnFWD.Enabled = True
    btnREV.Enabled = True
    btnSTOP.Enabled = True
  End If
    
  Timer1.Enabled = True

End Sub

Private Sub btnLimpaEnviar_Click()

  txtStringEnvio.Text = ""

End Sub

Private Sub btnLimpaReceb_Click()
  
  txtStringReceb.Text = ""

End Sub

Private Sub btnPositMode_Click()
 Dim bTestModeOK As Boolean

  Timer1.Enabled = False
  bTestModeOK = True
    
  'Desabilita as entradas físicas
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "90", "00", "1EA5") + vbCrLf
                   
  If (txtStringEnvio = "Error 1: Port not open") Then bTestModeOK = False
                     
  'Muda para o modo de operação para Positioning Mode
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "8B", "00", "0002") + vbCrLf
                   
  If (txtStringEnvio = "Error 1: Port not open") Then bTestModeOK = False
  
  If (bTestModeOK) Then
    btnPositMode.Enabled = False
    btnJogMode.Enabled = False
    btnSairTestMode.Enabled = True
    btnFWD.Enabled = True
    btnREV.Enabled = True
    btnSTOP.Enabled = True
  End If
    
  Timer1.Enabled = True

End Sub

Private Sub btnReceber_Click()
End Sub

Private Sub btnRESET_Click()

  Timer1.Enabled = False
    
  'Define Velocidade como 1500 rpm
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "82", "00", "1EA5") + vbCrLf

  Timer1.Enabled = True

End Sub

Private Sub btnREV_Click()
 Dim sMensagem As String

  Timer1.Enabled = False
    
  'Define Velocidade como 1500 rpm
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "A0", "10", "05DC") + vbCrLf
  
  'Define Acel/Desac como 100 ms
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "A0", "11", "00000FA0") + vbCrLf
    
  'Define Deslocamento como 131072 pls
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "A0", "13", "00020000") + vbCrLf
      
  'Liga comandos ST2, SON ,LSP e LSN
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "92", "00", "00001007") + vbCrLf
     
  Timer1.Enabled = True
   
End Sub

Private Sub btnSairTestMode_Click()
 Dim bTestModeOK As Boolean
   
  Timer1.Enabled = False
  
  bTestModeOK = True
     
  'Desliga entradas de comando forçadas
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "92", "00", "00000000") + vbCrLf
  
  If (txtStringEnvio = "Error 1: Port not open") Then bTestModeOK = False
     
  'Re-Habilita comando via I/O externo
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "A0", "12", "1EA5") + vbCrLf
  
  If (txtStringEnvio = "Error 1: Port not open") Then bTestModeOK = False
     
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "8B", "00", "0000") + vbCrLf
  
  If (txtStringEnvio = "Error 1: Port not open") Then bTestModeOK = False
      
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "10", "10", "1EA5") + vbCrLf
    
  If (txtStringEnvio = "Error 1: Port not open") Then bTestModeOK = False
  
  Timer1.Enabled = True
  
  If (bTestModeOK) Then
    btnPositMode.Enabled = True
    btnJogMode.Enabled = True
    btnSairTestMode.Enabled = False
    btnFWD.Enabled = False
    btnREV.Enabled = False
    btnSTOP.Enabled = False
  End If
      
End Sub

Private Sub Command1_Click()

   
End Sub

Private Sub btnEscrDados_Click()
 
End Sub

Private Sub btnETX_Click()

  txtStringEnvio.Text = txtStringEnvio.Text + Chr(3)

End Sub

Private Sub btnFWD_Click()

  Timer1.Enabled = False
  
  'Define Velocidade como 1500 rpm
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "A0", "10", "05DC") + vbCrLf
  
  'Define Acel/Desac como 100 ms
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "A0", "11", "00000FA0") + vbCrLf
    
  'Define Deslocamento como 131072 pls
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "A0", "13", "00020000") + vbCrLf
      
  'Liga comandos ST1, SON ,LSP e LSN
  txtStringEnvio = txtStringEnvio + _
                   fEnviaMsg(fChkStation, "92", "00", "00000807") + vbCrLf
       
  Timer1.Enabled = True

End Sub

Private Sub btnSOH_Click()

  txtStringEnvio.Text = txtStringEnvio.Text + Chr(1)

End Sub

Private Sub btnSTOP_Click()
   
  Timer1.Enabled = False
       
  txtStringEnvio.Text = txtStringEnvio.Text + _
                fEnviaMsg(fChkStation, "92", "00", "00000001") + vbCrLf
     
  Timer1.Enabled = True
   
End Sub

Private Sub btnSTX_Click()

  txtStringEnvio.Text = txtStringEnvio.Text + Chr(2)

End Sub

Private Sub btnAbrirPorta_Click()

  If (Not (mscCOM1.PortOpen)) Then
    mscCOM1.PortOpen = True
    btnAbrirPorta.Enabled = False
    btnFecharPorta.Enabled = True
  End If

End Sub

Private Sub btnEnviar_Click()
  
  If (mscCOM1.PortOpen) Then
    mscCOM1.Output = txtStringEnvio.Text
  Else
    MsgBox ("A porta não foi aberta!!")
  End If

End Sub

Private Sub btnFecharPorta_Click()
  
  Timer1.Enabled = False
  
  If (mscCOM1.PortOpen) Then
    mscCOM1.PortOpen = False
    btnAbrirPorta.Enabled = True
    btnFecharPorta.Enabled = False
  End If
  

End Sub

Private Sub chkStNumber_Click()

  If (chkStNumber.Value = 1) Then
    txtStNumber.Visible = True
  Else
    txtStNumber.Visible = False
  End If

End Sub

Private Sub Form_Load()

  btnFecharPorta.Enabled = False

  If (Not (mscCOM1.PortOpen)) Then
    mscCOM1.PortOpen = True
    btnAbrirPorta.Enabled = False
    btnFecharPorta.Enabled = True
  End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
 
  Timer1.Enabled = False

End Sub

Private Sub Timer1_Timer()
Dim sMensagemT As String

  If (mscCOM1.PortOpen) Then
    'Lê status do servo (alarme)
    sMensagemT = fEnviaMsg(fChkStation, "02", "00", "")
    sMensagemT = mscCOM1.Input
    If (sMensagemT <> "") Then txtStringReceb = sMensagemT
  End If
   
End Sub
