VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The smart cirlce"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Frame Frame2 
         Height          =   3855
         Left            =   3480
         TabIndex        =   2
         Top             =   120
         Width           =   1695
         Begin VB.CommandButton CmdExit 
            Cancel          =   -1  'True
            Caption         =   "Exit"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton CmdStop 
            Caption         =   "Stop"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton CmdMove 
            Caption         =   "Move"
            Default         =   -1  'True
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.PictureBox Path 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H80000008&
         ForeColor       =   &H00C0C0FF&
         Height          =   3735
         Left            =   120
         Picture         =   "Moving Object.frx":0000
         ScaleHeight     =   3705
         ScaleWidth      =   3225
         TabIndex        =   1
         Top             =   240
         Width           =   3255
         Begin VB.Shape Obj 
            BorderColor     =   &H000000FF&
            FillColor       =   &H000000FF&
            Height          =   120
            Left            =   360
            Shape           =   3  'Circle
            Top             =   3240
            Width           =   120
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Designed by: Yehia Muhsen
'Date       : 4-19-2003
'Description: In this program, a red circle will try to find it's way
'             in the path. Before the movement, there must be a search
'             for a black point in the path. As a way of improving the search method,
'             search has to be done in circles around the object and in two directions,
'             clockwise and counterclockwise. Also the next search starts form the
'             same angle where the perviouse point was found according to the previouse
'             postion. If no point found in the path, the radius of the circle increases
'             so that it can search the whole area.
'             You can use the Paint to make a new path and copy it and paste it to the
'             picutre box, but make sure the path is black, because that's what the
'             circle searches for ( Black points) .
'             This program uses the sine and cosine functions for making circles in the
'             searching process.
'
'Email:       yehia_sm@hotmail.com ( If you have any question, please feel free to ask. )
              
Dim StopObject As Boolean
'To remember the direction
Dim OldI, IR As Single

Private Sub CmdExit_Click()
Unload Form1
End Sub

Private Sub CmdMove_Click()

On Error GoTo Err
'Varibales
Dim XPath As Integer, YPath As Integer
Dim XPathTemp As Integer, YPathTemp As Integer, Radius As Integer
Dim XPathTemp1 As Integer, YPathTemp1 As Integer
Dim XPathTemp2 As Integer, YPathTemp2 As Integer
Dim I As Single, Stp As Single
Dim MoveObject As Boolean

'Disable some buttons
CmdMove.Enabled = False
CmdExit.Enabled = False
StopObject = False

'Action
Do
    'Objcect center cooredinates
    XPath = Obj.Left + Obj.Width / 2
    YPath = Obj.Top + Obj.Height / 2
    Radius = 0
    MoveObject = False
    
    'Search for a point in the path
    Do
        'Radius increase if failed the previouse search to find a point in the path
        'Multiplying by 1.5 makes the search for the path faster
        Radius = 1.5 * Radius + 1
        
        'As the radius of the search gets bigger lower the steps so
        'it can cover many points in the cercumference
        Stp = 1 / (10 * Radius)
        
        'OldI is very important to start from the same angle you
        'used the last time to find a point in the path
        IR = OldI
        'Search in two halves of a circle
        For I = OldI To OldI + 3.14159265359 Step Stp
            
            'Search for a point in path counter clockwise
            XPathTemp1 = XPath + Radius * Cos(I)
            YPathTemp1 = YPath + Radius * Sin(I)
            
            'Path.PSet (XPathTemp1, YPathTemp1), vbGreen
            
            'Test the point
            'If the point's color is black then move the object there
            If Path.Point(XPathTemp1, YPathTemp1) = 0 Then
                XPathTemp = XPathTemp1
                YPathTemp = YPathTemp1
                MoveObject = True
                OldI = I
                
                'Exit this interior "Do" loop to move the object
                'You don't have to search in the opposite direction
                Exit Do
            End If
            
            'Search for point in the path clockwise
            'This is very important search in order to minimize the
            'time of finding the path and in order to avoid the backward motion
            'IR deacreses where I increases but they start for the same value
            IR = IR - Stp
            XPathTemp2 = XPath + Radius * Cos(IR)
            YPathTemp2 = YPath + Radius * Sin(IR)
            
            'Path.PSet (XPathTemp2, YPathTemp2), vbRed
            
            'Test the point
            'If the point's color is black then move the object there

            If Path.Point(XPathTemp2, YPathTemp2) = vbBlack Then
                XPathTemp = XPathTemp2
                YPathTemp = YPathTemp2
                MoveObject = True
                OldI = IR
                Exit Do
            End If
            
            'When hit the stop button
            If StopObject Then GoTo ExtSub
            'If the radius is bigger than the path area then stop searching
            If Radius >= (Sqr(Path.Height ^ 2 + Path.Width ^ 2)) Then GoTo Err
            
            DoEvents
        Next I
    Loop Until MoveObject
    'Move the object
    Obj.Left = XPathTemp - Obj.Width / 2
    Obj.Top = YPathTemp - Obj.Height / 2
    DoEvents
Loop Until StopObject

GoTo ExtSub
'
Err:
MsgBox "Can't find the path", vbCritical, "Error"
'
ExtSub:
'Reenable buttons
CmdMove.Enabled = True
CmdExit.Enabled = True
CmdMove.SetFocus
End Sub

Private Sub CmdStop_Click()
StopObject = True
End Sub

