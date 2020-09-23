VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00C8AE99&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manikandan"
   ClientHeight    =   8310
   ClientLeft      =   -60
   ClientTop       =   510
   ClientWidth     =   11880
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu excelImpExp 
         Caption         =   "&Excel  File Import & Export"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub excelImpExp_Click()
    Frm_export_import.Show 1
End Sub
