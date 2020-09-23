VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   443
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7800
      Top             =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Conservation of Momentum
' Hou Xiong

' A smaller ball with a smaller mass collides with
' a larger ball with a larger mass.  The reaction
' occurs just as expected using Conservation of
' Momentum.  Play around with the variables in Form_Load().

Dim x1 As Single    'position
Dim r1 As Single    'radius
Dim v1 As Single    'velocity
Dim m1 As Single    'mass

Dim x2 As Single
Dim r2 As Single
Dim v2 As Single
Dim m2 As Single

Private Sub Form_Load()
    x1 = 50
    r1 = 16
    v1 = 5
    m1 = 2
    
    x2 = 400
    r2 = 32
    v2 = -2
    m2 = 8
End Sub

Private Sub Timer1_Timer()
    Dim vt1 As Single, vt2 As Single
    
    'update position
    x1 = x1 + v1
    x2 = x2 + v2
    
    'store temporary velocities
    vt1 = v1
    vt2 = v2
    
    'the two balls have collided
    If x1 + r1 > x2 - r2 Then
        'calculate velocities after collision (conservation of momentum)
        v1 = ((m1 - m2) / (m1 + m2)) * vt1 + ((2 * m2) / (m1 + m2)) * vt2
        v2 = ((2 * m1) / (m1 + m2)) * vt1 + ((m2 - m1) / (m1 + m2)) * vt2
    End If
    
    Me.Refresh
    Me.Circle (x1, Me.ScaleHeight / 2), r1
    Me.Circle (x2, Me.ScaleHeight / 2), r2
End Sub
