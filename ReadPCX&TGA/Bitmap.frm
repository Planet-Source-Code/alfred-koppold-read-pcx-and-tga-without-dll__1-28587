VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Read all pcx and tga-files without a dll!!!!!"
   ClientHeight    =   4410
   ClientLeft      =   2010
   ClientTop       =   2370
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7275
   WindowState     =   2  'Maximiert
   Begin VB.CommandButton Command1 
      Caption         =   "Vote"
      Height          =   615
      Left            =   8400
      TabIndex        =   14
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Scale"
      Height          =   1815
      Left            =   9840
      TabIndex        =   10
      Top             =   5880
      Width           =   1455
      Begin VB.OptionButton Option7 
         Caption         =   "200%"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "100%"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "50%"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Object"
      Height          =   1095
      Left            =   8280
      TabIndex        =   7
      Top             =   5880
      Width           =   1455
      Begin VB.OptionButton Option4 
         Caption         =   "Form"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Picturebox"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pattern"
      Height          =   1692
      Left            =   7080
      TabIndex        =   4
      Top             =   5880
      Width           =   1092
      Begin VB.OptionButton Option2 
         Caption         =   "*.pcx"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   852
      End
      Begin VB.OptionButton Option1 
         Caption         =   "*.tga"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   852
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   3720
      Pattern         =   "*.tga"
      TabIndex        =   3
      Top             =   5880
      Width           =   3252
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   1920
      TabIndex        =   2
      Top             =   5880
      Width           =   1692
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   1692
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   2172
      Left            =   240
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   0
      Top             =   100
      Width           =   1932
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Dim Scaling As Single
If Option5.Value = True Then Scaling = 2 ' 50%
If Option6.Value = True Then Scaling = 1 '100%
If Option7.Value = True Then Scaling = 0.5 '200%
Form2.Hide
Pic1.Cls

If Option1.Value = True Then 'TGA-File
Dim tgaFile As New LoadTGA
If Option6.Value = False Then
tgaFile.Autoscale = False
Else
tgaFile.Autoscale = True
End If

tgaFile.LoadTGA File1.Path & "\" & File1.FileName
If tgaFile.IsTGA = True Then 'is it a TGA-File?
If Option3.Value = False Then
Form2.Caption = File1.FileName
Form2.Width = tgaFile.TGAWidth / Scaling
Form2.Height = tgaFile.TGAHeight / Scaling
Form2.Show
tgaFile.DrawTGA Form2
Else
Form1.Pic1.Width = tgaFile.TGAWidth / Scaling
Form1.Pic1.Height = tgaFile.TGAHeight / Scaling
tgaFile.DrawTGA Form1.Pic1
End If
End If
Else 'PCX-File
Dim pcxFile As New LoadPCX

If Option6.Value = False Then
pcxFile.Autoscale = False
Else
pcxFile.Autoscale = True
End If
pcxFile.LoadPCX File1.Path & "\" & File1.FileName
If pcxFile.IsPCX = True Then 'is it a PCX-File?
Form2.ScaleMode = 3
pcxFile.ScaleMode = 1
If Option3.Value = False Then
Form2.Caption = File1.FileName
Form2.Width = pcxFile.PCXWidth / Scaling
Form2.Height = pcxFile.PCXHeight / Scaling
Form2.Show
pcxFile.DrawPCX Form2
Else
Form1.Pic1.Width = pcxFile.PCXWidth / Scaling
Form1.Pic1.Height = pcxFile.PCXHeight / Scaling
pcxFile.DrawPCX Form1.Pic1
End If
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Option1_Click()
File1.Pattern = "*.tga"
End Sub

Private Sub Option2_Click()
File1.Pattern = "*.pcx"
End Sub
Private Sub Command1_Click()
    On Error GoTo ErrHandler
    Dim Temp As Integer
    Temp = MsgBox("Visit ''http://www.planet-source-code.com'' now?", vbQuestion + vbYesNo)
    
    If Temp = vbYes Then
        'This code snippet opens an .url file.
        'Please give me some credit if you use it ;)
        Shell "rundll32.exe shdocvw.dll,OpenURL " & App.Path & "\Planet Source Code.url"
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Sorry, ''Planet Source Code.url'' was not found.", vbExclamation + vbOKOnly
End Sub

