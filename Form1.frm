VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Prevent other instance (DDE example)"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4965
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   885
      TabIndex        =   0
      Text            =   "this textbox recieves/sends dde data"
      Top             =   2205
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label2 
      Height          =   600
      Left            =   105
      TabIndex        =   2
      Top             =   750
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "COMPILE THIS PROJECT.  OPEN ONE INSTANCE AND MINIMIZE IT.   TRY TO OPEN A NEW INSTANCE BY CLICKING THE EXECUTEABLE AGAIN"
      Height          =   735
      Left            =   90
      TabIndex        =   1
      Top             =   30
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PreviewsInstanceExists As Boolean
Private Sub Form_Load()
On Error GoTo ErrorHandler
     PreviewsInstanceExists = True
     Text1.LinkTopic = App.Title & "|" & Me.Name
     Text1.LinkItem = "Text1"
     Text1.LinkMode = 1 'If no Application is listening it will cause an error no DDE responce
     'If the code continues without error means ths is not the first instance
     'and we have connected with the first via DDE.
     Text1.Text = "NEW INSTANCE OPENED " & Interaction.Command
     Text1.LinkPoke
     Unload Me
ErrorHandler:
  If Err.Number = 282 Then
     'Seems no application is listening in DDE this means its the first instance.
     'Me.LinkMode = 1 'Done manualy
     Me.LinkTopic = Me.Name
     Me.Caption = "This is the first instance"
     PreviewsInstanceExists = False
  End If
End Sub


Private Sub Text1_Change()
  'NOTE: Do not have any modal forms or textboxes popping up in the _Change event
  'because after the commands in this event are executed a DDE transfer ok is sent
  If (PreviewsInstanceExists = False) And Left(Text1.Text, 19) = "NEW INSTANCE OPENED" Then
     Me.Visible = True
     Me.WindowState = vbNormal
     Me.SetFocus
     If Len(Text1.Text) > 20 Then Label2.Caption = "Command line parameters of the other instance:" & vbCrLf & Mid(Text1.Text, 20)
     Text1.Text = ""
  End If
End Sub
