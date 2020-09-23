VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_export_import 
   Caption         =   "Import and Export a Excel Sheet"
   ClientHeight    =   6240
   ClientLeft      =   1455
   ClientTop       =   1320
   ClientWidth     =   8940
   Icon            =   "Frm_export_import.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8940
   Begin VB.CommandButton Cmd_export 
      Caption         =   "&Export to a Excel Sheet"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   4800
      Width           =   3255
   End
   Begin VB.CommandButton Cmd_import 
      Caption         =   "&Import Excel Sheet"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   3255
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _Version        =   196608
      _ExtentX        =   15478
      _ExtentY        =   8070
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   20
      SpreadDesigner  =   "Frm_export_import.frx":030A
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Cmd_export_html 
      Caption         =   "Export to HTML Page"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   5280
      Width           =   3255
   End
   Begin VB.OLE OLE1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Frm_export_import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_export_Click()
    Dim a As String
    CommonDialog1.Action = 2
    tempp = CommonDialog1.FileName ' & ".xls"
    For ii = 1 To Len(CommonDialog1.FileName)
        If Mid$(tempp, Len(tempp), 1) = "\" Then
            Exit For
        Else
            tempp = Mid$(tempp, 1, Len(tempp) - 1)
        End If
    Next
    tempname = CommonDialog1.FileTitle
    For aa = 1 To Len(tempname)
        If Mid$(tempname, Len(tempname) - 3, 4) = ".xls" Or Mid$(tempname, Len(tempname) - 3, 4) = ".XLS" Then
            a = tempp & tempname
        Else
            a = tempp & tempname & ".xls"
        End If
    Next
    Call Export_excel(vaSpread1, a, "")
End Sub

Private Sub Cmd_export_html_Click()
    Dim a As String
    CommonDialog1.Action = 2
    tempp = CommonDialog1.FileName
    For ii = 1 To Len(CommonDialog1.FileName)
        If Mid$(tempp, Len(tempp), 1) = "\" Then
            Exit For
        Else
            tempp = Mid$(tempp, 1, Len(tempp) - 1)
        End If
    Next
    tempname = CommonDialog1.FileTitle
    For aa = 1 To Len(tempname)
        If Mid$(tempname, Len(tempname) - 3, 4) = ".html" Or Mid$(tempname, Len(tempname) - 3, 4) = ".HTML" Then
            a = tempp & tempname
        Else
            a = tempp & tempname & ".html"
        End If
    Next
    Call Export_html(vaSpread1, a, "")
End Sub

Private Sub Cmd_import_Click()
    Dim b As String
    CommonDialog1.Action = 1
    b = CommonDialog1.FileName
    Call Import_excel(vaSpread1, b, "")
End Sub

