VERSION 5.00
Object = "{B6369C6C-9300-11D4-A9E0-4481F8C00000}#1.0#0"; "KNOBOCX.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KnobOCX"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   1695
   StartUpPosition =   3  'Windows Default
   Begin KNOBOCX.Knob Knob5 
      Height          =   975
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
   End
   Begin KNOBOCX.Knob Knob3 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
   End
   Begin KNOBOCX.Knob Knob2 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin KNOBOCX.Knob Knob1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin KNOBOCX.Knob Knob4 
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Knob1_Changed()
Me.Caption = Knob1.KnobValue
End Sub

Private Sub Knob2_Changed()
Me.Caption = Knob2.KnobValue
End Sub

Private Sub Knob3_Changed()
Me.Caption = Knob3.KnobValue
End Sub

Private Sub Knob4_Changed()
Me.Caption = Knob4.KnobValue
End Sub

Private Sub Knob5_Changed()
Me.Caption = Knob5.KnobValue
End Sub
