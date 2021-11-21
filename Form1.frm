VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Wpusty pryzmatyczne"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12795
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   12795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Zapisz do pliku"
      Height          =   495
      Left            =   8520
      TabIndex        =   61
      Top             =   8280
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   60
      Top             =   0
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
      Min             =   0,001
      Max             =   1000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   9000
      TabIndex        =   59
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   5400
      TabIndex        =   57
      Text            =   "Text16"
      Top             =   8400
      Width           =   2055
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3600
      TabIndex        =   54
      Text            =   "B"
      Top             =   7440
      Width           =   495
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   53
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   52
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   51
      Text            =   "Text10"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   50
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   49
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5400
      TabIndex        =   48
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   7080
      TabIndex        =   47
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   3480
      TabIndex        =   43
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton koniec 
      Caption         =   "Zamknij"
      Height          =   375
      Left            =   11040
      TabIndex        =   40
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text6 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   39
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1800
      TabIndex        =   33
      Top             =   8400
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   240
      TabIndex        =   32
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   10560
      TabIndex        =   31
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   11040
      TabIndex        =   30
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Obliczenia"
      Height          =   3975
      Left            =   240
      TabIndex        =   19
      Top             =   4200
      Width           =   12375
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   58
         Text            =   "L"
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5880
         TabIndex        =   56
         Text            =   "Text15"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   405
         Left            =   4200
         TabIndex        =   55
         Text            =   "H"
         Top             =   3240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   6720
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   3105
         ScaleWidth      =   5025
         TabIndex        =   29
         Top             =   360
         Width           =   5055
      End
      Begin VB.Frame Frame4 
         Caption         =   "t1"
         Height          =   735
         Left            =   480
         TabIndex        =   23
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "t2"
         Height          =   735
         Left            =   480
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label32 
         Caption         =   "[mm]"
         Height          =   495
         Left            =   5880
         TabIndex        =   46
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label31 
         Caption         =   "[mm]"
         Height          =   375
         Left            =   5880
         TabIndex        =   45
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label30 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   44
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label28 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   9480
         TabIndex        =   38
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   37
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label25 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   36
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   35
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "Wpust pryzmatyczny"
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Dobrano wpust pryzmatyczny  typu"
         Height          =   375
         Left            =   3240
         TabIndex        =   28
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label18 
         Caption         =   "Lc="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   27
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Lo="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   26
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "D³ugoœæ ca³kowita wpustu"
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label15 
         Caption         =   "Obliczona d³ugoœæ czynna wpustu"
         Height          =   375
         Left            =   3360
         TabIndex        =   24
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label12 
         Caption         =   "G³êbokoœæ rowka na wpust w wa³ku t1"
         Height          =   615
         Left            =   360
         TabIndex        =   21
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "G³êbokoœæ rowka na wpust w piaœcie t2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   8280
      TabIndex        =   18
      Top             =   3600
      Width           =   2415
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   8280
      TabIndex        =   16
      Top             =   3120
      Width           =   3615
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   8280
      TabIndex        =   14
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form1.frx":3259E
      Left            =   8280
      List            =   "Form1.frx":325A0
      TabIndex        =   12
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8280
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Left            =   4200
      Max             =   5
      Min             =   1
      TabIndex        =   7
      Top             =   1560
      Value           =   1
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Znam moment M"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Znam Si³ê F"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   600
      Min             =   1
      TabIndex        =   1
      Top             =   3000
      Value           =   1
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Co jest dane"
      Height          =   1575
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label29 
      Caption         =   "[N]"
      Height          =   375
      Left            =   3480
      TabIndex        =   42
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "WPUSTY PRYZMATYCZNE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   41
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label10 
      Caption         =   "Okreœl rodzaj pracy"
      Height          =   255
      Left            =   5760
      TabIndex        =   17
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Okreœl warunki pracy"
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Dobierz materia³ wpustu"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Wybierz typ wpustu"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Wybierz œrednice wa³ka"
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Liczba wpustów"
      Height          =   615
      Left            =   3240
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a As Double
Public b As Double
Public c As Double
Public l As Double

        

Private Sub Combo1_Click()
 Set exTabela = CreateObject("Excel.Application")

        'Dim Zeszyt As Workbook

        exTabela.Application.Visible = False

        'Set Zeszyt = exTabela.Workbooks.Add 'Open(oFileDlg.FileName)

        Set Zeszyt = exTabela.Workbooks.Open(FileName:="C:\1.xls")

      'Dim listaEL As WorkSheet

        'Set listaEL = exTabela.Application.Workbooks.Item(1).Worksheets.Item(1)

        Set listaEL = exTabela.Application.Workbooks.Item(1).Worksheets("roboczy")

        exTabela.Application.Visible = False

        Dim pozPlik(1, 25) As Variant

        exTabela.Application.Visible = False

       

        exTabela.Cells.NumberFormat = "@"
        'Dodawanie wyszukaj.pionowo
    
'====================================================
'Szukanie danych z komórek tak samo jak funkcja INDEX, match, vlookup w excel
'Dodawanie kolumny t2

For w = 1 To 64
    If Combo1.Text = listaEL.Cells(w + 23, 7).Value Then
    Text8.Text = listaEL.Cells(w + 23, 5).Value
    Exit For
    End If
    
Next w

'===================================================
'Dodawanie t1

For w = 1 To 64
If Combo1.Text = listaEL.Cells(w + 23, 7).Value Then
Text9.Text = listaEL.Cells(w + 23, 4).Value
Exit For
End If
Next w
'====================================
'Dodawanie wartoœci b
For w = 1 To 64
If Combo1.Text = listaEL.Cells(w + 23, 7).Value Then
Text13.Text = listaEL.Cells(w + 23, 2).Value
Exit For
End If
Next w
'===================================================
'Dodawanie wartosci h
For w = 1 To 64
If Combo1.Text = listaEL.Cells(w + 23, 7).Value Then
Text14.Text = listaEL.Cells(w + 23, 3).Value
Exit For
End If
Next w

'=================================================
'Dodawanie d³ugosci l

   For w = 1 To 64
ProgressBar1 = w * 15.6
If Combo1.Text = listaEL.Cells(w + 23, 7).Value Then
Text17.Text = listaEL.Cells(w + 23, 8).Value
Text15.Text = listaEL.Cells(w + 23, 8).Value
'MsgBox Text12.Text & "=text12" & Text17.Text & "=text17"
If Text17.Text < Text12.Text Then
'MsgBox Text17
MsgBox Text12 & "=text12"
Text17.Text = "####"
Exit For
End If
End If
Next w

'===================
If Option2.Value = True Then
Dim d As Variant
Dim f, h As Variant
f = CDbl(Text1.Text)
d = CDbl(Combo1.Text)
h = CDbl(Text6.Text)
'HScroll1.Value = CDbl(Text1.Text)
h = 2 * f * 1000 / d
Text6.Text = h
End If


If Option1.Value = True Then
Text11.Text = oblL(CDbl(Text1.Text), CDbl(Text3.Text), CDbl(Text2.Text), CDbl(Text9.Text), CDbl(Text10.Text)) * 1000

Text4.Text = CDbl(Text11.Text) + CDbl(Text13.Text)
Text5.Text = CDbl(Text11.Text)
Text7.Text = CDbl(Text11.Text) + 0.5 * CDbl(Text13.Text)

Label23.Caption = Combo2.Text
If Combo2.ListIndex = 0 Then
Text12.Text = Text4.Text
Else
If Combo2.ListIndex = 1 Then
Text12.Text = Text5.Text
Else
If Combo2.ListIndex = 2 Then
Text12.Text = Text7.Text
Else
End If
End If
End If

'End If

''===================================================
''Dodawanie wartoœci b
'For w = 1 To 64
'If Combo1.Text = listaEl.Cells(w + 24, 7).Value Then
'Label24.Caption = listaEl.Cells(w + 24, 2).Value
'Exit For
'End If
'Next w
''===================================================
''Dodawanie wartosci h
'For w = 1 To 64
'If Combo1.Text = listaEl.Cells(w + 24, 7).Value Then
'Label26.Caption = listaEl.Cells(w + 24, 3).Value
'Exit For
'End If
'Next w
'
''=================================================
''Dodawanie d³ugosci
'For w = 1 To 64
'If Combo1.Text = listaEl.Cells(w + 24, 7).Value Then
'Label28.Caption = listaEl.Cells(w + 24, 9).Value
'Exit For
'End If
'Next w
End If
 Zeszyt.Close (False)
  
       End Sub
    
    Private Sub Combo2_Click()
Label23.Caption = Combo2.Text
If Combo2.ListIndex = 0 Then
Text12.Text = Text4.Text
Else
If Combo2.ListIndex = 1 Then
Text12.Text = Text5.Text
Else
If Combo2.ListIndex = 2 Then
Text12.Text = Text7.Text
Else
End If
End If
End If
End Sub

Private Sub Combo3_Click()
If Combo3.ListIndex = 0 Then
Text3.Text = 120
Else
If Combo3.ListIndex = 1 Then
Text3.Text = 145
Else
If Combo3.ListIndex = 2 Then
Text3.Text = 160
Else
If Combo3.ListIndex = 3 Then
Text3.Text = 175
Else
If Combo3.ListIndex = 4 Then
Text3.Text = 170
Else
End If
End If
End If
End If
End If
End Sub

Private Sub Combo4_Click()
       If Combo4.ListIndex = 0 And Combo5.ListIndex = 0 Then
     Text2.Text = 0.35
     Else
       If Combo4.ListIndex = 0 And Combo5.ListIndex = 1 Then
      Text2.Text = 0.15
    Else
If Combo4.ListIndex = 0 And Combo5.ListIndex = 2 Then
Text2.Text = 0.03
Else
If Combo4.ListIndex = 1 And Combo5.ListIndex = 0 Then
Text2.Text = 0.6
Else
If Combo4.ListIndex = 1 And Combo5.ListIndex = 1 Then
Text2.Text = 0.2
Else
If Combo4.ListIndex = 1 And Combo5.ListIndex = 2 Then
Text2.Text = 0.06
Else
If Combo4.ListIndex = 2 And Combo5.ListIndex = 0 Then
Text2.Text = 0.8
Else
If Combo4.ListIndex = 2 And Combo5.ListIndex = 1 Then
Text2.Text = 0.3
Else
If Combo4.ListIndex = 2 And Combo5.ListIndex = 2 Then
Text2.Text = 0.1
Else
End If
End If
End If
End If
End If
End If
End If
End If
End If

End Sub

Private Sub Combo5_click()

If Combo4.ListIndex = 0 And Combo5.ListIndex = 0 Then
Text2.Text = 0.35
Else
If Combo4.ListIndex = 1 And Combo5.ListIndex = 0 Then
Text2.Text = 0.6
Else
If Combo4.ListIndex = 2 And Combo5.ListIndex = 0 Then
Text2.Text = 0.8
Else
If Combo4.ListIndex = 0 And Combo5.ListIndex = 1 Then
Text2.Text = 0.15
Else
If Combo4.ListIndex = 1 And Combo5.ListIndex = 1 Then
Text2.Text = 0.2
Else
If Combo4.ListIndex = 2 And Combo5.ListIndex = 1 Then
Text2.Text = 0.3
Else
If Combo4.ListIndex = 0 And Combo5.ListIndex = 2 Then
Text2.Text = 0.03
Else
If Combo4.ListIndex = 1 And Combo5.ListIndex = 2 Then
Text2.Text = 0.06
Else
If Combo4.ListIndex = 2 And Combo5.ListIndex = 2 Then
Text2.Text = 0.1
Else
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Command1_Click()
'init word object
Set wordApp = New Word.Application
'show word
wordApp.Application.Visible = True
'add new document
wordApp.Documents.Add
wordApp.Documents(1).Range.InsertAfter Text:="Wprowadz dan¹" & "" & Text1.Text

Dim oDoc

  wordApp.Documents(1).Tables.Add Range:=wordApp.Documents(1).Range, NumRows:=2, NumColumns:= _
        5, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
 wordApp.Documents(1).InlineShapes.AddPicture FileName:= _
        "C:\Documents and Settings\Administrator\Moje dokumenty\Moje obrazy\1.jpg.bmp" _
        , LinkToFile:=False, SaveWithDocument:=True
        wordApp.Documents(1).Range.InsertAfter Text:="kto ja"
Close
End Sub

Private Sub Command2_Click()


MsgBox oblL(CDbl(Text1.Text), CDbl(Text3.Text), CDbl(Text2.Text), CDbl(Text9.Text), CDbl(Text10.Text))

End Sub



Private Sub Command3_Click()
For s = 1 To 1000
ProgressBar1.Value = s
Next s
End Sub

Private Sub Form_Load()
Form1.Show
a = 0.35 ' Text2.Text
 b = 1 'Text1.Text
c = 120 ' Text3.Text
l = 1 'text10.text

Text2.Text = a
Text1.Text = b
Text3.Text = c
Text10.Text = l
Set exTabela = CreateObject("Excel.Application")

        'Dim Zeszyt As Workbook

        exTabela.Application.Visible = False

        'Set Zeszyt = exTabela.Workbooks.Add 'Open(oFileDlg.FileName)

        Set Zeszyt = exTabela.Workbooks.Open(FileName:="C:\1.xls")

      'Dim listaEL As WorkSheet

        'Set listaEL = exTabela.Application.Workbooks.Item(1).Worksheets.Item(1)

        Set listaEL = exTabela.Application.Workbooks.Item(1).Worksheets("roboczy")

        exTabela.Application.Visible = False

        Dim pozPlik(1, 25) As Variant

        exTabela.Application.Visible = False

        listaEL.Name = "Zestawienie_komponentów"

        exTabela.Cells.NumberFormat = "@"
     
          '+++++++++++++++++++++++
       'Dodawanie zakresu komórek
     For Each Cell In listaEL.Range("J25:J27")
Combo2.AddItem Cell.Value
 Combo2.ListIndex = 0
 Next Cell
   'Dodawanie  wymiarów wa³ka
     For Each Cell In listaEL.Range("g24:g64")
Combo1.AddItem Cell.Value
Combo1.ListIndex = 0
 Next Cell
 'Dodaj materia³
      For Each Cell In listaEL.Range("l25:l29")
    Combo3.AddItem Cell.Value
 Combo3.ListIndex = 0
  Next Cell
 'Dodaj opisy
    Label2.Caption = "Okreœl dan¹"
 Label3.Caption = "Okreœl dan¹="
'Wartoœæ minimalna wpustów
 VScroll1.Value = 1
 'Dodawanie warunków pracy
 For Each Cell In listaEL.Range("N25:N27")
 Combo4.AddItem Cell.Value
 Combo4.ListIndex = 0
 Next
 'Dodawanie rodzaju pracy
 For Each Cell In listaEL.Range("S24:S26")
 Combo5.AddItem Cell.Value
 Combo5.ListIndex = 0
 Next
 Option1.Value = True
' Dodawanie dawnych wyjœciowych  t2
'===========================================
For w = 1 To 64
    If Combo1.Text = listaEL.Cells(w + 23, 7).Value Then
    Text8.Text = listaEL.Cells(w + 23, 5).Value
    Exit For
    End If
    
Next w

'===================================================
'Dodawanie t1

For w = 1 To 64
If Combo1.Text = listaEL.Cells(w + 23, 7).Value Then
Text9.Text = listaEL.Cells(w + 23, 4).Value
Exit For
End If
Next w
'===================================
'Dodawanie naprezen dopuszczalnych kc
For w = 1 To 64
If Combo3.Text = listaEL.Cells(w + 24, 12).Value Then
Text3.Text = listaEL.Cells(w + 24, 13).Value
Exit For
End If
Next w
'====================================
'Dodawanie wartoœci b
For w = 1 To 64
If Combo1.Text = listaEL.Cells(w + 23, 7).Value Then
Text13.Text = listaEL.Cells(w + 23, 2).Value
Exit For
End If
Next w
'===================================================
'Dodawanie wartosci h
For w = 1 To 64
If Combo1.Text = listaEL.Cells(w + 23, 7).Value Then
Text14.Text = listaEL.Cells(w + 23, 3).Value
Exit For
End If
Next w

'=================================================
'Dodawanie d³ugosci l
For w = 1 To 64
If Combo1.Text = listaEL.Cells(w + 23, 7).Value Then
Text17.Text = listaEL.Cells(w + 23, 8).Value
If Text12.Text > Text17.Text Then
Text17.Text = "$$$$"
End If
Exit For
End If
Next w
'=================================================

'$$$$$$$$$$$$$$$$$$$$$$$$$$

'$$$$$$$$$$$$$$$$$$$$$$$$$$
'Obliczenia si³y/momentu

If Option2.Value = True Then
Dim d As Variant
Dim f, h As Variant
f = Text1.Text
d = Combo1.Text
h = Text6.Text
HScroll1.Value = Text1.Text
h = 2 * f * 1000 / d
Text6.Text = h
End If



Text11.Text = oblL(CDbl(Text1.Text), CDbl(Text3.Text), CDbl(Text2.Text), CDbl(Text9.Text), CDbl(Text10.Text)) * 1000
'=======
Text4.Text = CDbl(Text11.Text) + CDbl(Text13.Text)
Text5.Text = CDbl(Text11.Text)
Text7.Text = CDbl(Text11.Text) + 0.5 * CDbl(Text13.Text)

Label23.Caption = Combo2.Text
If Combo2.ListIndex = 0 Then
Text12.Text = Text4.Text
Else
If Combo2.ListIndex = 1 Then
Text12.Text = Text5.Text
Else
If Combo2.ListIndex = 2 Then
Text12.Text = Text7.Text
Else
End If
End If
End If
Zeszyt.Close (False)

End Sub

Private Sub HScroll1_Change()
Text1.Text = HScroll1.Value
Text1_LostFocus
End Sub

Private Sub koniec_Click()
Close
End
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Label2.Caption = "Si³a F"
Label3.Caption = "Si³a F="
Label29.Caption = "[N]"
Text6.Text = Text1.Text
If HScroll1.Max > Text1.Text Then HScroll1.Value = Text1.Text
If HScroll1.Max < Text1.Text Then HScroll1.Value = HScroll1.Max
Else
Label2.Caption = "Moment M"
Label3.Caption = "Moment M="
Label29.Caption = "[Nm]"
Dim d As Variant
Dim f, h As Variant
f = Text1.Text
d = Combo1.Text
h = Text6.Text
HScroll1.Value = CDbl(Text1.Text)
h = 2 * f * 1000 / d
Text6.Text = h
End If


End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Label2.Caption = "Moment M"
Label3.Caption = "Moment M="
Label29.Caption = "[Nm]"
Dim d As Variant
Dim f, h As Variant
f = Text1.Text
d = Combo1.Text
h = Text6.Text
h = 2 * f * 1000 / d

If HScroll1.Max > Text1.Text Then HScroll1.Value = Text1.Text
'If HScroll1.Max < Text1.Text Then HScroll1.Value = Text1.Text
Text6.Text = h
Else
Label2.Caption = "Si³a F"
Label3.Caption = "Si³a F="
Text6.Text = Text1.Text
Label29.Caption = "[N]"
End If

End Sub

'==================================
'Sprawdziæ dzia³anie funkcji public przez MSGBOX wraz z innymi danymi
Public Function oblL(ByVal xF As Double, xkc As Double, xz As Double, xt As Double, xn As Double)


oblL = xF * 1000 / (xkc * 1000000# * xz * xt * xn)


End Function






Private Sub Text1_Change()
If HScroll1.Max > Text1.Text Then HScroll1.Value = Text1.Text
If HScroll1.Max < Text1.Text Then HScroll1.Value = HScroll1.Max

End Sub

Private Sub Text1_GotFocus()
Text1.Tag = Text1.Text
End Sub

Private Sub Text1_LostFocus()
Dim d As Variant

Dim f, h As Variant
f = Text1.Text
d = Combo1.Text
h = Text6.Text
'If Text1.Text > HScroll1.Max Then HScroll1.Value = HScroll1.Max

If HScroll1.Max > Text1.Text And HScroll1.Value = Text1.Text Then Text1.Text = HScroll1.Value
If HScroll1.Max < Text1.Text Then HScroll1.Value = HScroll1.Max
'HScroll1.Value = Text1.Text
'MsgBox ("Wartoœæ powy¿ej 3124 zapisz i zaakceptuj wyborem danej")
If IsNumeric(Text1.Text) = False And Text1.Text <> "" Then
Text1.Text = Text1.Tag
End If
If Option2.Value = True Then
If HScroll1.Max > Text1.Text Then Text1.Text = HScroll1.Value
If HScroll1.Max < Text1.Text Then HScroll1.Value = HScroll1.Max

h = CDbl(2 * f * 1000 / d)
If Text1.Text <> "" Then Text6.Text = h
Else
If Option1.Value = True Then
If Text1.Text <> "" Then Text6.Text = Text1.Text
Else
End If
End If
If IsNumeric(Text1.Text) = False Then
Text1.Text = Text1.Tag
End If
End Sub

Private Sub Text12_Change()
'Set exTabela = CreateObject("Excel.Application")
'
'        Dim Zeszyt As Workbook
'
'        exTabela.Application.Visible = False
'
'        Set Zeszyt = exTabela.Workbooks.Add 'Open(oFileDlg.FileName)
'
'        Set Zeszyt = exTabela.Workbooks.Open(FileName:="C:\1.xls")
'
'      Dim listaEL As Worksheet
'
'        Set listaEL = exTabela.Application.Workbooks.Item(1).Worksheets.Item(1)
'
'        Set listaEL = exTabela.Application.Workbooks.Item(1).Worksheets("roboczy")
'
'        exTabela.Application.Visible = False
'
'        Dim pozPlik(1, 25) As Variant
'
'        exTabela.Application.Visible = False
'
'
'
'        exTabela.Cells.NumberFormat = "@"
'        Dodawanie wyszukaj.pionowo
'    For w = 1 To 64
'
'If Combo1.Text = listaEL.Cells(w + 23, 7).Value Then
'Text17.Text = listaEL.Cells(w + 23, 8).Value
'Text15.Text = listaEL.Cells(w + 23, 8).Value
'MsgBox Text12.Text & "=text12" & Text17.Text & "=text17"
'If Text17.Text < Text12.Text Then
'MsgBox Text17
'MsgBox Text12 & "=text12"
'Text17.Text = "####"
'Exit For
'End If
'End If
'Next w
'Zeszyt.Close (False)
End Sub

Private Sub Text2_Change()

''$$$$$$$$$$$$$$$$$$$$$$$$$$
'Text2.Text = 0.35
'Text1.Text = 6
'Text3.Text = 120
'Text10.Text = 1
''$$$$$$$$$$$$$$$$$$$$$$$$$$
If Combo4.ListIndex >= 0 Then
Text11.Text = oblL(CDbl(Text1.Text), CDbl(Text3.Text), CDbl(Text2.Text), CDbl(Text9.Text), CDbl(Text10.Text)) * 1000


Text4.Text = CDbl(Text11.Text) + CDbl(Text13.Text)
Text5.Text = CDbl(Text11.Text)
Text7.Text = CDbl(Text11.Text) + 0.5 * CDbl(Text13.Text)

Label23.Caption = Combo2.Text
If Combo2.ListIndex = 0 Then
Text12.Text = Text4.Text
Else
If Combo2.ListIndex = 1 Then
Text12.Text = Text5.Text
Else
If Combo2.ListIndex = 2 Then
Text12.Text = Text7.Text
Else
End If
End If
End If
End If
End Sub


Private Sub Text3_Change()

'$$$$$$$$$$$$$$$$$$$$$$$$$$
'Text2.Text = a
'Text1.Text = b
'Text3.Text = c
'Text10.Text = a
'$$$$$$$$$$$$$$$$$$$$$$$$$$
If Combo3.ListIndex >= 0 And Combo3.ListIndex <= 4 Then
Text11.Text = oblL(CDbl(Text1.Text), CDbl(Text3.Text), CDbl(Text2.Text), CDbl(Text9.Text), CDbl(Text10.Text)) * 1000

Text4.Text = CDbl(Text11.Text) + CDbl(Text13.Text)
Text5.Text = CDbl(Text11.Text)
Text7.Text = CDbl(Text11.Text) + 0.5 * CDbl(Text13.Text)
Label23.Caption = Combo2.Text
If Combo2.ListIndex = 0 Then
Text12.Text = Text4.Text
Else
If Combo2.ListIndex = 1 Then
Text12.Text = Text5.Text
Else
If Combo2.ListIndex = 2 Then
Text12.Text = Text7.Text
Else

End If
End If
End If
End If
End Sub

Private Sub Text6_Change()

If IsNumeric(Text1.Text) = True And Text1.Text >= 0 Then
Text11.Text = oblL(CDbl(Text6.Text), CDbl(Text3.Text), CDbl(Text2.Text), CDbl(Text9.Text), CDbl(Text10.Text)) * 1000
'=======

Text4.Text = CDbl(Text11.Text) + CDbl(Text13.Text)
Text5.Text = CDbl(Text11.Text)
Text7.Text = CDbl(Text11.Text) + 0.5 * CDbl(Text13.Text)

Label23.Caption = Combo2.Text
If Combo2.ListIndex = 0 Then
Text12.Text = Text4.Text
Else
If Combo2.ListIndex = 1 Then
Text12.Text = Text5.Text
Else
If Combo2.ListIndex = 2 Then
Text12.Text = Text7.Text
Else
End If
End If
End If

End If
'Text16.Text = "UWAGA!!! Zalecana norma d³ugoœci wynosi"
End Sub

Private Sub VScroll1_Change()
'$$$$$$$$$$$$$$$$$$$$$$$$$$
'Text2.Text = 0.35
'Text1.Text = 6
'Text3.Text = 120
'$$$$$$$$$$$$$$$$$$$$$$$$$$
Text10.Text = VScroll1.Value

If VScroll1.Value < 6 Then
Text11.Text = oblL(CDbl(Text1.Text), CDbl(Text3.Text), CDbl(Text2.Text), CDbl(Text9.Text), CDbl(Text10.Text)) * 1000
Text4.Text = CDbl(Text11.Text) + CDbl(Text13.Text)
Text5.Text = CDbl(Text11.Text)
Text7.Text = CDbl(Text11.Text) + 0.5 * CDbl(Text13.Text)

Label23.Caption = Combo2.Text
If Combo2.ListIndex = 0 Then
Text12.Text = Text4.Text
Else
If Combo2.ListIndex = 1 Then
Text12.Text = Text5.Text
Else
If Combo2.ListIndex = 2 Then
Text12.Text = Text7.Text
Else
End If
End If
End If
Else
End If

End Sub


