VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GTExpert2"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Registro:"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Identificación:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Introduce dirección del Programa:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Introduce Fecha   dd/mm/aaaa :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gdFechaVieja As Date


Private Sub Command1_Click()

  On Error GoTo FechaVieja
  Text3 = UCase(Text3)
  If Encriptar(Label5) <> Text3 Then
    MsgBox "Número de Registro Erroneo", vbInformation, "GTExpert2"
  Else
    Date = CDate(Text1)
'    txtCDRom = GetRegValue(HKEY_CURRENT_USER, "Software\Next\GT", "CDRom")
    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Einsa Multimedia\GTExpert\1.0", "FECHA")
'    Call DeleteValue(HKEY_CURRENT_USER, "Software\Next\GT", "Version")
       
    ret = Shell(Text2)
  End If
  
  Exit Sub
  
FechaVieja:
  Date = gdFechaVieja
  Debug.Print Err.Number
  If Err.Number = 53 Or Err.Number = 5 Then
    MsgBox "Dirección de Programa Erronea o Inexistente", vbInformation, "GTExpert2"
  End If
  
End Sub

Private Sub Command2_Click()
  Unload Form1
  Date = gdFechaVieja
  End
End Sub

Private Sub Form_Load()
  gdFechaVieja = Date
  Text1 = GetSetting("GTExpert2", "Inicio", "Fecha")
  Text2 = GetSetting("GTExpert2", "Inicio", "Path")
  Text3 = UCase(GetSetting("GTExpert2", "Inicio", "Registro"))
  Text4 = UCase(GetSetting("GTExpert2", "Inicio", "Unidad"))
  If Text2 = "" Then
     Text2 = "C:\Archivos de programa\Einsa Multimedia\GtExpert\GTExpert.exe"
  End If
'  If Text4 = "" Then
'     Text4 = "D:\"
'  End If
  
  'Lee el numero del HD y comprueba
  Label5 = LeerNumeroHD(Left$(App.Path, 3))
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting "GTExpert2", "Inicio", "Fecha", Text1
  SaveSetting "GTExpert2", "Inicio", "Path", Text2
  SaveSetting "GTExpert2", "Inicio", "Registro", Text3
  Date = gdFechaVieja
End Sub

