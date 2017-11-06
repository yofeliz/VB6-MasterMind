VERSION 5.00
Begin VB.Form fMasterMind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Mind"
   ClientHeight    =   4935
   ClientLeft      =   6165
   ClientTop       =   4590
   ClientWidth     =   5175
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MasterMind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fMarco 
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2055
      Begin VB.OptionButton oFacil 
         Caption         =   "Nivel FÁCIL"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   670
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton oDificil 
         Caption         =   "Nivel DIFÍCIL"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton btnAyuda 
         Caption         =   "Ayuda"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4320
         Width           =   855
      End
      Begin VB.Timer tmrReloj 
         Interval        =   1
         Left            =   1320
         Top             =   1320
      End
      Begin VB.CommandButton btnSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton btnComprobar 
         Caption         =   "Comprobar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton btnSolucion 
         Caption         =   "Solución"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton btnComenzar 
         Caption         =   "Nuevo Juego"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Image pctPatron 
         Height          =   375
         Index           =   6
         Left            =   120
         Picture         =   "MasterMind.frx":0442
         Stretch         =   -1  'True
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image pctPatron 
         Height          =   375
         Index           =   7
         Left            =   1560
         Picture         =   "MasterMind.frx":0BFE
         Stretch         =   -1  'True
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Aleatorio"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lGenerando 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generador"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   915
      End
      Begin VB.Image pctGenerando 
         Height          =   480
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image pctSolucion 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   375
      End
      Begin VB.Image pctSolucion 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   375
      End
      Begin VB.Image pctSolucion 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   600
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   375
      End
      Begin VB.Image pctSolucion 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   375
      End
      Begin VB.Image pctPatron 
         Height          =   375
         Index           =   5
         Left            =   1080
         Picture         =   "MasterMind.frx":13AA
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   375
      End
      Begin VB.Image pctPatron 
         Height          =   375
         Index           =   4
         Left            =   600
         Picture         =   "MasterMind.frx":1B66
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   375
      End
      Begin VB.Image pctPatron 
         Height          =   375
         Index           =   3
         Left            =   1560
         Picture         =   "MasterMind.frx":2336
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image pctPatron 
         Height          =   375
         Index           =   2
         Left            =   1080
         Picture         =   "MasterMind.frx":2AF9
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image pctPatron 
         Height          =   375
         Index           =   1
         Left            =   600
         Picture         =   "MasterMind.frx":3294
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image pctPatron 
         Height          =   360
         Index           =   0
         Left            =   120
         MousePointer    =   4  'Icon
         OLEDropMode     =   1  'Manual
         Picture         =   "MasterMind.frx":3A39
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
   End
   Begin VB.Image pctArrastrar 
      Height          =   240
      Left            =   2280
      Picture         =   "MasterMind.frx":4212
      Top             =   4080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pctLuna 
      Height          =   480
      Index           =   7
      Left            =   2280
      Picture         =   "MasterMind.frx":435C
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pctLuna 
      Height          =   480
      Index           =   6
      Left            =   2280
      Picture         =   "MasterMind.frx":479E
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pctLuna 
      Height          =   480
      Index           =   5
      Left            =   2280
      Picture         =   "MasterMind.frx":4BE0
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pctLuna 
      Height          =   480
      Index           =   4
      Left            =   2280
      Picture         =   "MasterMind.frx":5022
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pctLuna 
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "MasterMind.frx":5464
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pctLuna 
      Height          =   480
      Index           =   2
      Left            =   2280
      Picture         =   "MasterMind.frx":58A6
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pctLuna 
      Height          =   480
      Index           =   1
      Left            =   2280
      Picture         =   "MasterMind.frx":5CE8
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image pctLuna 
      Height          =   480
      Index           =   0
      Left            =   2280
      Picture         =   "MasterMind.frx":612A
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlecha 
      Height          =   375
      Left            =   2280
      Picture         =   "MasterMind.frx":656C
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   36
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   39
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   38
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   37
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   35
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   34
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   33
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   32
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   31
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   30
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   29
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   28
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   27
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   26
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   25
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   24
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   23
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   22
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   21
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   20
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   19
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   18
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   17
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   16
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   15
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   14
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   13
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   12
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   11
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image pctArray 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   39
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   38
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   37
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   36
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   35
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   34
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   33
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   32
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   31
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   30
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   29
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   28
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   27
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   26
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   25
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   1560
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   24
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   1560
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   23
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   22
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   21
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   2040
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   20
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   2040
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   19
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   2760
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   18
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   2760
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   17
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   16
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   15
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   14
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   13
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   3000
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   12
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   3000
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   11
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   10
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   9
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   3480
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   8
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   3480
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   7
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   4200
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   6
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   4200
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   5
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   4
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   3
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   4680
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   2
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   4680
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   1
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   4440
      Width           =   135
   End
   Begin VB.Shape crcArray 
      Height          =   135
      Index           =   0
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   4440
      Width           =   135
   End
End
Attribute VB_Name = "fMasterMind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim semilla As Double                        'Semilla para el Randomize
Dim linea As Integer                         'Numero de línea actual
Dim combinacionMaestra(4) As New StdPicture  'Array de Picture's maestro
Dim contadorLuna As Integer                  'Contador para la animación de la Luna
Dim nuevoPatron                              'Array de colores.

Private Sub btnAyuda_Click()
    MsgBox "Ayuda en línea:" + vbCrLf + vbCrLf + _
           "   · Versión:" + vbCrLf + _
           "      - MasterMind 1.0" + vbCrLf + vbCrLf + _
           "   · Objetivo:" + vbCrLf + _
           "      - Averigua la combinación de colores secreta." + vbCrLf + vbCrLf + _
           "   · Modo de juego:" + vbCrLf + _
           "      - Arrastra los colores a las casillas cuadradas y comprueba la combinación." + vbCrLf + vbCrLf + _
           "   · Resolución:" + vbCrLf + _
           "      - Cada círculo negro indica que uno de los colores existe dentro de la combinación y que además está en la posición correcta." + vbCrLf + _
           "      - Cada círculo blanco indica que uno de los colores existe dentro de la combinación, pero está en una posición incorrecta." + vbCrLf + _
           "      - Cada círculo transparente indica que hay un color que no está dentro de la combinación." + vbCrLf + vbCrLf + _
           "   · Compilador:" + vbCrLf + _
           "      - Visual Basic 6.0" + vbCrLf + vbCrLf + _
           "   · Programador:" + vbCrLf + _
           "      - David Díaz" + vbCrLf + vbCrLf + _
           "   · Fecha:" + vbCrLf + _
           "      - Mayo 2005" + vbCrLf + vbCrLf + _
           "   · Dedicado con especial cariño a:" + vbCrLf + _
           "      - Susana, Ely y José Andrés  -  Ojén (MARBELLA)", vbInformation, "Ayuda"
End Sub

Private Sub btnComenzar_Click()
    linea = 1                    'Inicializa la linea a 1
    activarPatron                'Habilita el patron de imagenes para arrastrar
    desactivarLineas             'Deshabilita todas las lineas del pctArray
    activarLinea                 'Habilita la linea en curso (en este caso, siempre 1)
    borrarImagenes               'Limpia el contenido de pctArray(x).Picture
    borrarBanderas               'Limpia el contenido de crcArray(x).BackStyle
    borrarSolucion               'Limpia el contenido de pctSolucion(x).Picture
    generarMaestro               'Se genera la línea a descubrir (master)
    inicializarFlecha            'Muestra la flecha y la posiciona
    btnComenzar.Enabled = False  'Se deshabilita a sí mismo
    btnSolucion.Enabled = True   'Habilita el botón Solución
    btnComprobar.Enabled = True  'Habilita el botón Comprobar
    btnComprobar.SetFocus        'Pasa el foco al botón Comprobar
    oFacil.Enabled = False       'Deshabilita la opcion FÁCIL.
    oDificil.Enabled = False     'Deshabilita la opción DIFÍCIL.
    tmrReloj.Enabled = False     'Para el contador de semilla
End Sub

Private Sub btnComprobar_Click()
    Dim c As Integer
    Dim vacio As Boolean
    vacio = False
    For c = 0 To UBound(combinacionMaestra) - 1  'Comprueba que las casillas de cada línea estén completas
        If pctArray(c + (linea - 1) * 4).Picture = 0 Then
            vacio = True
        End If
    Next c
    If vacio = False Then
        If comprobarAciertos = True Then  'Comprueba los aciertos
            MsgBox "¡ENHORABUENA! Has ganado.", vbInformation, "Resultado del Juego"
            btnSolucion.Enabled = False   'Se deshabilita a sí mismo
            btnComprobar.Enabled = False  'Deshabilita el botón Comprobar
            btnComenzar.Enabled = True    'Habilita el botón Comenzar
            btnComenzar.SetFocus          'Pasa el foco al botón Comenzar
            desactivarLineas              'Deshabilita las lineas para poder arrastrar
            desactivarPatron              'Deshabilita el patron para poder arrastrar
            deshabilitarFlecha            'Oculta la flecha
            tmrReloj.Enabled = True       'Reactiva el contador de semilla
            oFacil.Enabled = True         'Deshabilita la opcion FÁCIL.
            oDificil.Enabled = True       'Deshabilita la opción DIFÍCIL.
            For c = 0 To UBound(combinacionMaestra) - 1
                pctSolucion(c).Picture = combinacionMaestra(c)
            Next c
        Else
            linea = linea + 1   'Suma en una unidad la linea en curso
            If linea > 10 Then
                MsgBox "¡MALA SUERTE! Has perdido.", vbInformation, "Resultado del Juego"
                btnSolucion.Enabled = False   'Se deshabilita a sí mismo
                btnComprobar.Enabled = False  'Deshabilita el botón Comprobar
                btnComenzar.Enabled = True    'Habilita el botón Comenzar
                btnComenzar.SetFocus          'Pasa el foco al botón Comenzar
                desactivarLineas              'Deshabilita las lineas para poder arrastrar
                desactivarPatron              'Deshabilita el patron para poder arrastrar
                deshabilitarFlecha            'Oculta la flecha
                tmrReloj.Enabled = True       'Reactiva el contador de semilla
                For c = 0 To UBound(combinacionMaestra) - 1
                    pctSolucion(c).Picture = combinacionMaestra(c)
                Next c
            Else
                desactivarLineas    'Deshabilita todas las líneas del pctArray
                activarLinea        'Habilita la siguiente linea
                desplazarFlecha     'Mueve la flecha una línea
            End If
        End If
    Else
        MsgBox "¡ERROR! No puede haber casillas vacías.", vbCritical, "Error"
    End If
End Sub

Private Sub btnSolucion_Click()
    If MsgBox("¿Está seguro de que quiere ver la solución?", vbYesNo, "Solución") = vbYes Then
        Dim c As Integer
        btnSolucion.Enabled = False   'Se deshabilita a sí mismo
        btnComprobar.Enabled = False  'Deshabilita el botón Comprobar
        btnComenzar.Enabled = True    'Habilita el botón Comenzar
        btnComenzar.SetFocus          'Pasa el foco al botón Comenzar
        desactivarLineas              'Deshabilita las lineas para poder arrastrar
        desactivarPatron              'Deshabilita el patron para poder arrastrar
        deshabilitarFlecha            'Oculta la flecha
        oFacil.Enabled = True         'Deshabilita la opcion FÁCIL.
        oDificil.Enabled = True       'Deshabilita la opción DIFÍCIL.
        tmrReloj.Enabled = True       'Reactiva el contador de semilla
        For c = 0 To UBound(combinacionMaestra) - 1
            pctSolucion(c).Picture = combinacionMaestra(c)
        Next c
    End If
End Sub

Private Sub btnSalir_Click()
    End  'Sale de programa
End Sub

Private Sub Form_Load()
    linea = 1    'Se inicializa la linea a la primera
    If oFacil.Value = True Then 'Se crea un nuevo patrón con 6 casillas.
        nuevoPatron = Array(pctPatron(0), pctPatron(1), pctPatron(2), _
                            pctPatron(3), pctPatron(4), pctPatron(5))
        pctPatron(6).Visible = False
        pctPatron(7).Visible = False
    End If
End Sub

Private Sub oDificil_Click()
    If oDificil.Value = True Then 'Se crea un nuevo patron con 8 casillas.
        nuevoPatron = Array(pctPatron(0), pctPatron(1), pctPatron(2), _
                            pctPatron(3), pctPatron(4), pctPatron(5), _
                            pctPatron(6), pctPatron(7))
        pctPatron(6).Visible = True
        pctPatron(7).Visible = True
    End If
End Sub

Private Sub oFacil_Click()
    If oFacil.Value = True Then 'Se crea un nuevo patrón con 6 casillas.
        nuevoPatron = Array(pctPatron(0), pctPatron(1), pctPatron(2), _
                            pctPatron(3), pctPatron(4), pctPatron(5))
        pctPatron(6).Visible = False
        pctPatron(7).Visible = False
    End If
End Sub

Private Sub pctArray_DragDrop(Indice As Integer, Source As Control, X As Single, Y As Single)
    pctArray(Indice).Picture = Source.Picture  'Asigna el Picture al Source.Picture
    Source.DragIcon = LoadPicture()            'Cambia el puntero
End Sub

Private Sub pctArray_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    If State = 2 Then
        Source.DragIcon = pctArrastrar.Picture
    End If
    If State = 1 Then
        Source.DragIcon = LoadPicture()
    End If
End Sub

Private Sub tmrReloj_Timer()
    semilla = semilla + 1  'Incrementa en 1 la semilla
    If semilla = 1E+300 Then
        semilla = 0
    End If
    If (semilla Mod 5) = 0 Then
        pctGenerando.Picture = pctLuna(contadorLuna).Picture
        contadorLuna = contadorLuna + 1
        If contadorLuna = 8 Then
            contadorLuna = 0
        End If
    End If
End Sub

Private Sub activarPatron()  'Recorre las casillas para arrastrar y las habilita
    Dim c As Integer
    'For c = 0 To pctPatron.UBound
    For c = 0 To UBound(nuevoPatron)
        pctPatron(c).DragMode = 1
    Next c
End Sub

Private Sub desactivarPatron()  'Deshabilita el patron para arrastrar
    Dim c As Integer
    'For c = 0 To pctPatron.ubound
    For c = 0 To UBound(nuevoPatron)
        pctPatron(c).DragMode = 0
    Next c
End Sub

Private Sub activarLinea()  'En funcion de la linea,
    Dim c As Integer        'habilita las 4 Image's correspondientes
    desactivarLineas
    Select Case linea
        Case 1
            For c = 0 To 3
                pctArray(c).Enabled = True
            Next c
        Case 2
            For c = 4 To 7
                pctArray(c).Enabled = True
            Next c
        Case 3
            For c = 8 To 11
                pctArray(c).Enabled = True
            Next c
        Case 4
            For c = 12 To 15
                pctArray(c).Enabled = True
            Next c
        Case 5
            For c = 16 To 19
                pctArray(c).Enabled = True
            Next c
        Case 6
            For c = 20 To 23
                pctArray(c).Enabled = True
            Next c
        Case 7
            For c = 24 To 27
                pctArray(c).Enabled = True
            Next c
        Case 8
            For c = 28 To 31
                pctArray(c).Enabled = True
            Next c
        Case 9
            For c = 32 To 35
                pctArray(c).Enabled = True
            Next c
        Case 10
            For c = 36 To 39
                pctArray(c).Enabled = True
            Next c
    End Select
End Sub

Private Sub desactivarLineas()  'Recorre el pctArray y lo deshabilita
    Dim c As Integer
    For c = 0 To pctArray.UBound
        pctArray(c).Enabled = False
    Next c
End Sub

Private Sub borrarImagenes()  'Borra las imágenes de la pctArray
    Dim c As Integer
    For c = 0 To pctArray.UBound
        pctArray(c).Picture = LoadPicture()
    Next c
End Sub

Private Sub borrarBanderas()  'Borra los rellenos de los resultados
    Dim c As Integer
    For c = 0 To crcArray.UBound
        crcArray(c).BackStyle = 0
    Next c
End Sub

Private Sub borrarSolucion()  'Borra las imágenes de la solución
    Dim c As Integer
    For c = 0 To pctSolucion.UBound
        pctSolucion(c).Picture = LoadPicture()
    Next c
End Sub

Private Sub generarMaestro()  'Genera la combinacion maestra a descubrir
    Dim c As Integer
    Dim temp As StdPicture
    Erase combinacionMaestra  'Se borra el contenido
    Randomize semilla
    For c = 0 To UBound(combinacionMaestra) - 1
        'Set temp = pctPatron(Int(Rnd * pctPatron.UBound)).Picture  'Se genera el numero aleatorio
        Set temp = pctPatron(Int(Rnd * (UBound(nuevoPatron) + 1))).Picture 'Se genera el numero aleatorio
        Do While existeImagen(temp) = True  'Mientras exista la imagen
            'Set temp = pctPatron(Int(Rnd * pctPatron.UBound)).Picture  'Se regenera el número
            Set temp = pctPatron(Int(Rnd * (UBound(nuevoPatron) + 1))).Picture 'Se regenera el número
        Loop
        Set combinacionMaestra(c) = temp  'Se asigna la imagen
    Next c
End Sub

Private Function existeImagen(imagen As StdPicture) As Boolean 'Devuelve si existe o no
    Dim c As Integer                                           'una imagen en la
    Dim existe As Boolean                                      'combinación maestra
    existe = False
    For c = 0 To UBound(combinacionMaestra) - 1
        If imagen = combinacionMaestra(c) Then
            existe = True
        End If
    Next c
    existeImagen = existe
End Function

Private Sub inicializarFlecha()  'Posición inicial = 4440
    imgFlecha.Top = 4440
    imgFlecha.Visible = True
End Sub

Private Sub desplazarFlecha()  'Desplazamiento de 480
    imgFlecha.Top = imgFlecha.Top - 480
End Sub

Private Sub deshabilitarFlecha()  'Oculta la flecha
    imgFlecha.Visible = False
End Sub

Private Function comprobarAciertos() As Boolean  'Se comprueban los aciertos exactos e inexactos
    Dim aciertosExactos As Integer               'Los vacíos son la diferencia hasta 4
    Dim aciertosInexactos As Integer             'Devuelve True si se han hecho 4 aciertos exactos
    Dim c, i As Integer
    aciertosExactos = 0
    aciertosInexactos = 0
    For c = 0 To UBound(combinacionMaestra) - 1
        If pctArray(c + (linea - 1) * 4).Picture = combinacionMaestra(c) Then
            aciertosExactos = aciertosExactos + 1
        Else
            If existeImagen(pctArray(c + (linea - 1) * 4).Picture) = True Then
                aciertosInexactos = aciertosInexactos + 1
            End If
        End If
    Next c
    If aciertosExactos = 4 Then
        comprobarAciertos = True
    Else
        comprobarAciertos = False
    End If
    pintarAciertos aciertosExactos, aciertosInexactos
End Function

Private Sub pintarAciertos(aciertosExactos As Integer, aciertosInexactos As Integer)
    Dim c As Integer  'Se pintan las banderas en funcion de los aciertos
    For c = 0 To UBound(combinacionMaestra) - 1
        If aciertosExactos <> 0 Then
            With crcArray(c + (linea - 1) * 4)
                .BackStyle = 1
                .BackColor = &H0&
            End With
            aciertosExactos = aciertosExactos - 1
        Else
            If aciertosInexactos <> 0 Then
                With crcArray(c + (linea - 1) * 4)
                    .BackStyle = 1
                    .BackColor = &HFFFFFF
                End With
                aciertosInexactos = aciertosInexactos - 1
            End If
        End If
    Next c
End Sub
