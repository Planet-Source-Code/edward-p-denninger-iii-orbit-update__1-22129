VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orbit - [epd3]"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Options"
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   0
      Width           =   2055
      Begin VB.TextBox txtMoonOrbitSpeed 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Text            =   "5"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPlanetOrbitSpeed 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Text            =   "1"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtMoonOrbitRadius 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Text            =   "20"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtPlanetOrbitRadius 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "80"
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkFollowMouse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Follow Mouse"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Moon Orbit Speed:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Planet Orbit Speed:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   1920
         X2              =   120
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Moon Orbit Radius:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Planet Orbit Radius:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Orbit"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image imgMoon 
      Height          =   135
      Left            =   4560
      Picture         =   "frmMain.frx":0442
      Top             =   3360
      Width           =   135
   End
   Begin VB.Image imgPlanet 
      Height          =   135
      Left            =   3000
      Picture         =   "frmMain.frx":04B4
      Top             =   2760
      Width           =   135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
' Project:  Orbit                                            *
' Filename: frmMain.frm                                      *
' Author:   Edward P. Denninger III                          *
' Date:     3/22/2001                                        *
' Copyright © 2001 Edward P. Denninger III                   *
'*************************************************************
'*                         NOTICE                            *
'*************************************************************
' You may use and freely distribute this porject and source  *
' at your own leisure as long as I am given credit for my    *
' work.  If you have any comments or ideas for improvement,  *
' you can reach me at: edward3@optonline.net                 *
'*************************************************************

Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private m_MousePos      As POINTAPI     ' Mouse pos, also the center of the planet's obit
Private m_MoonPos       As POINTAPI     ' Moon pos
Private m_PlanetAng     As Single       ' Planet angle
Private m_MoonAng       As Single       ' Moon angle
Private m_PlanetDist    As Single       ' Planet orbit radius
Private m_MoonDist      As Single       ' Moon orbit radius
Private m_Running       As Boolean      ' If the loop is still running
Private m_OTime         As Long         ' Time when the loop starts


Private Sub cmdStart_Click()
    
    ' Start the loop
    OrbitLoop
End Sub

Private Sub cmdStop_Click()
    
    ' End the loop
    m_Running = False
End Sub

Private Sub Form_Load()
    
    ' Set some variables
    m_PlanetAng = 0
    m_MoonAng = 0
    m_PlanetDist = Val(txtPlanetOrbitRadius.Text)
    m_MoonDist = Val(txtMoonOrbitRadius.Text)
End Sub

Private Sub OrbitLoop()
    
    ' Make sure the loop will run
    m_Running = True
    
    Do While m_Running
        
        ' Get the time
        m_OTime = GetTickCount
    
        '--------------------------------------------
        ' Planet Orbit
        '--------------------------------------------
        
        ' Reset the angle
        If m_PlanetAng >= 360 Then m_PlanetAng = 0
        
        ' Increment the angle
        m_PlanetAng = m_PlanetAng + Val(txtPlanetOrbitSpeed.Text)
        
        If chkFollowMouse Then
            
            ' Get the cursor pos
            GetCursorPos m_MousePos
            
            ' Adjust the cursor pos for the X & Y of frmMain
            m_MousePos.x = m_MousePos.x - (Left / 15)
            m_MousePos.y = m_MousePos.y - (Top / 15)
        Else
            
            ' Center the orbit
            m_MousePos.x = ScaleWidth / 2
            m_MousePos.y = ScaleHeight / 2
        End If
        
        ' Calculate the new pos and move the planet
        imgPlanet.Left = (m_MousePos.x + Cos(m_PlanetAng * 0.017453295) * m_PlanetDist) - (imgPlanet.Width / 2)
        imgPlanet.Top = (m_MousePos.y + Sin(m_PlanetAng * 0.017453295) * m_PlanetDist) - (imgPlanet.Height / 2)
        
        
        '--------------------------------------------
        ' Moon Orbit
        '--------------------------------------------
        
        ' Reset the angle
        If m_MoonAng >= 360 Then m_MoonAng = 0
        
        ' Increment the angle
        m_MoonAng = m_MoonAng + Val(txtMoonOrbitSpeed.Text)
        
        ' Get the planet pos
        m_MoonPos.x = imgPlanet.Left + (imgPlanet.Width / 2)
        m_MoonPos.y = imgPlanet.Top + (imgPlanet.Height / 2)
        
        ' Calculate the new pos and move the moon
        imgMoon.Left = (m_MoonPos.x + Cos(m_MoonAng * 0.017453295) * m_MoonDist) - (imgMoon.Width / 2)
        imgMoon.Top = (m_MoonPos.y + Sin(m_MoonAng * 0.017453295) * m_MoonDist) - (imgMoon.Width / 2)
        
        ' Pause for a bit
        Do While GetTickCount < m_OTime + 5
        Loop
        DoEvents
    Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' End the loop
    m_Running = False
End Sub

Private Sub txtMoonOrbitRadius_Change()
    
    m_MoonDist = Val(txtMoonOrbitRadius.Text)
End Sub

Private Sub txtPlanetOrbitRadius_Change()
    
    m_PlanetDist = Val(txtPlanetOrbitRadius)
End Sub
