VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19500
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   19500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Repetir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Jugar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8400
      TabIndex        =   3
      Top             =   1680
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nombreUsuario As String
Dim eleccionUsuario, eleccionMaquina As String
Dim victorias, perdidas, empates As Integer

Private Sub Command1_Click()
    
    
    eleccionUsuario = LCase(Text1.Text)
   

    Randomize
    eleccionMaquina = Int(Rnd() * 3) + 1
    
    Select Case eleccionMaquina
        Case 1
        eleccionMaquina = "papel"
        Case 2
        eleccionMaquina = "piedra"
        Case 3
        eleccionMaquina = "tijera"
    End Select
    
    If eleccionUsuario = eleccionMaquina Then
        empates = empates + 1
    ElseIf eleccionUsuario = "piedra" And eleccionMaquina = "tijera" Or eleccionUsuario = "tijera" And eleccionMaquina = "papel" Or eleccionUsuario = "papel" And eleccionMaquina = "piedra" Then
        victorias = victorias + 1
    Else
        perdidas = perdidas + 1
    End If
    
    If eleccionUsuario <> "papel" And eleccionUsuario <> "piedra" And eleccionUsuario <> "tijera" Then
    MsgBox "Porfa capo elegi papel,piedra o tijera", vbCritical
    Label1.Caption = ""
    victorias = 0
    perdidas = 0
    empates = 0
    Text1.Text = ""
   End If

     
    Label1.Caption = "victorias: " & victorias & " / derrotas: " & perdidas & " / empates: " & empates
    
        If victorias = 3 Or perdidas = 3 Then
            MsgBox "Felicidades " & nombreUsuario & vbCrLf & "Resultado: " & victorias & " victorias" & vbCrLf & perdidas & " derrotas" & vbCrLf & empates & " empates", vbInformation
            Command1.Enabled = False
            Text1.Enabled = False
        End If
        
        
  
End Sub
Private Sub Command2_Click()

        Command1.Enabled = True
        Text1.Enabled = True
        victorias = 0
        perdidas = 0
        empates = 0
        Label1.Caption = ""
        Text1.Text = ""
        

End Sub

Private Sub Form_Activate()


    nombreUsuario = InputBox("Ingrese nombre:", "Nombre de Usuario")


End Sub


Private Sub Label3_Click()

End Sub

