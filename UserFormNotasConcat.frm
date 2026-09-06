VERSION 5.00
Begin VB.Form UserFormNotasConcat
   Caption = "SHIP SELL - Relatório OTM"
   ClientHeight = 490
   ClientLeft = 0
   ClientTop = 0
   ClientWidth = 640
   BeginProperty Font
      Name = "Calibri"
      Size = 11
      Charset = 0
      Weight = 400
      Underline = 0
      Italic = 0
      Strikethrough = 0
   EndProperty
   LinkTopic = "Form1"
   ScaleHeight = 490
   ScaleWidth = 640
   StartUpPosition = 1  ' CenterOwner
   Begin VB.CommandButton CommandButton1
      Caption = "Continuar"
      Height = 360
      Left = 420
      TabIndex = 2
      Top = 390
      Width = 180
   End
   Begin VB.TextBox TextBoxNotasConcat
      Height = 330
      Left = 180
      MultiLine = -1  ' True
      ScrollBars = 3  ' Both
      TabIndex = 1
      Text = ""
      Top = 60
      Width = 420
   End
   Begin VB.TextBox TextBoxQTDNotas
      Alignment = 2  ' Center
      Height = 330
      Left = 20
      Locked = -1  ' True
      TabIndex = 0
      Text = "0"
      Top = 60
      Width = 140
   End
   Begin VB.Label LabelQTD
      Caption = "Quantidade de processos"
      Height = 30
      Left = 20
      Top = 20
      Width = 140
   End
   Begin VB.Label LabelNotas
      Caption = "SHIP SELLs concatenados"
      Height = 30
      Left = 180
      Top = 20
      Width = 420
   End
End
Attribute VB_Name = "UserFormNotasConcat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Unload Me
End Sub
