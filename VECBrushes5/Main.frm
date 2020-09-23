VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Brush Vectors by Robert Rayment"
   ClientHeight    =   7635
   ClientLeft      =   165
   ClientTop       =   -420
   ClientWidth     =   11130
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   742
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPlusHairs 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&C"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   " Toggle Cross- hairs "
      Top             =   6570
      Width           =   285
   End
   Begin VB.CommandButton cmdMagnifier 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&M"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   " Toggle Magnifier "
      Top             =   6915
      Width           =   285
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Info"
      Height          =   1035
      Left            =   45
      TabIndex        =   54
      Top             =   6465
      Width           =   1965
      Begin VB.Label LabWH 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pic W,H = 10000, 10000"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   105
         TabIndex        =   57
         Top             =   750
         Width           =   1680
      End
      Begin VB.Label LabXY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "X,Y = 10000, 10000"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   105
         TabIndex        =   56
         Top             =   555
         Width           =   1305
      End
      Begin VB.Label LabNumVisStrokes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No. Strokes = 888888            No. Viisble =  888888"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   55
         Top             =   195
         Width           =   1815
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   225
      Left            =   2895
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   6750
      Width           =   4680
   End
   Begin VB.VScrollBar VS 
      Height          =   6585
      Left            =   2595
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   135
      Width           =   255
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vector operations"
      Height          =   1470
      Left            =   60
      TabIndex        =   43
      Top             =   30
      Width           =   2220
      Begin VB.CheckBox chkFastRedraw 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Fast  redraw"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   345
         TabIndex        =   66
         ToolTipText     =   " FAST REDRAW "
         Top             =   1230
         Width           =   1650
      End
      Begin VB.CommandButton cmdPicVec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Stretch"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   " STRETCH VECTORS TO FIT "
         Top             =   975
         Width           =   990
      End
      Begin VB.CommandButton cmdPicVec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fix"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   " FIX RESIZED CANVAS "
         Top             =   975
         Width           =   1050
      End
      Begin VB.CommandButton cmdPicVec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Undo"
         Height          =   270
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   210
         Width           =   525
      End
      Begin VB.CommandButton cmdPicVec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Redo"
         Height          =   270
         Index           =   1
         Left            =   570
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   210
         Width           =   525
      End
      Begin VB.CommandButton cmdPicVec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   " CLEAR PICTURE "
         Top             =   210
         Width           =   1005
      End
      Begin VB.CommandButton cmdPicVec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Redraw"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   " REDRAW PICTURE"
         Top             =   480
         Width           =   1005
      End
      Begin VB.CommandButton cmdPicVec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delete LAST"
         Height          =   255
         Index           =   3
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   " DELETE LAST VECTOR "
         Top             =   480
         Width           =   1035
      End
      Begin VB.CommandButton cmdPicVec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delete ALL"
         Height          =   240
         Index           =   5
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   " DELETE ALL VECTORS "
         Top             =   735
         Width           =   1050
      End
      Begin VB.CommandButton cmdPicVec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CLIP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   " CLIP VECTORS TO UNDO "
         Top             =   735
         Width           =   990
      End
   End
   Begin VB.Frame fraScratch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scratch"
      Height          =   3495
      Left            =   6870
      MousePointer    =   5  'Size
      TabIndex        =   29
      Top             =   405
      Width           =   2475
      Begin VB.CommandButton cmdClrScratch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear"
         Height          =   240
         Index           =   1
         Left            =   1800
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1875
         Width           =   495
      End
      Begin VB.PictureBox picScratch 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1230
         Index           =   1
         Left            =   135
         MousePointer    =   1  'Arrow
         ScaleHeight     =   78
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   143
         TabIndex        =   39
         Top             =   2160
         Width           =   2205
      End
      Begin VB.CommandButton cmdClrScratch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear"
         Height          =   240
         Index           =   0
         Left            =   1860
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   165
         Width           =   450
      End
      Begin VB.PictureBox picScratch 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1230
         Index           =   0
         Left            =   135
         MousePointer    =   1  'Arrow
         ScaleHeight     =   78
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   142
         TabIndex        =   30
         Top             =   450
         Width           =   2190
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   2445
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   15
         X2              =   2460
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Label LabScratch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Test brush"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Top             =   1890
         Width           =   885
      End
      Begin VB.Label LabScratch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LC_RC = Park_Select color"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   41
         Top             =   210
         Width           =   1740
      End
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      Height          =   6600
      Left            =   2880
      ScaleHeight     =   436
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   521
      TabIndex        =   32
      Top             =   150
      Width           =   7875
      Begin VB.PictureBox picCanvas 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4125
         Left            =   30
         ScaleHeight     =   275
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   227
         TabIndex        =   33
         Top             =   15
         Width           =   3405
         Begin VB.PictureBox picResizer 
            BorderStyle     =   0  'None
            Height          =   165
            Left            =   3195
            MousePointer    =   5  'Size
            ScaleHeight     =   11
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   53
            Top             =   3945
            Width           =   195
            Begin VB.Shape Shape2 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   45
               Left            =   60
               Top             =   60
               Width           =   75
            End
            Begin VB.Shape Shape1 
               FillStyle       =   0  'Solid
               Height          =   120
               Left            =   45
               Shape           =   3  'Circle
               Top             =   30
               Width           =   135
            End
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            X1              =   0
            X2              =   0
            Y1              =   16
            Y2              =   46
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            X1              =   25
            X2              =   57
            Y1              =   1
            Y2              =   1
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transparency %"
      Height          =   465
      Left            =   45
      TabIndex        =   22
      Top             =   3255
      Width           =   2235
      Begin VB.HScrollBar HSalpha 
         Height          =   150
         LargeChange     =   10
         Left            =   105
         Max             =   100
         TabIndex        =   34
         Top             =   225
         Width           =   1620
      End
      Begin VB.Label LabAlpha 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1770
         TabIndex        =   35
         Top             =   165
         Width           =   360
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Color"
      Height          =   2745
      Left            =   45
      TabIndex        =   14
      Top             =   3720
      Width           =   2235
      Begin VB.HScrollBar HSRGB 
         Height          =   180
         Index           =   2
         Left            =   1095
         Max             =   255
         SmallChange     =   4
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   720
         Value           =   1
         Width           =   315
      End
      Begin VB.HScrollBar HSRGB 
         Height          =   180
         Index           =   1
         Left            =   1095
         Max             =   255
         SmallChange     =   4
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   465
         Value           =   1
         Width           =   315
      End
      Begin VB.HScrollBar HSRGB 
         Height          =   180
         Index           =   0
         Left            =   1095
         Max             =   255
         SmallChange     =   4
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   210
         Value           =   1
         Width           =   315
      End
      Begin VB.PictureBox picColorBox 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C00000&
         Height          =   1020
         Left            =   75
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   16
         Top             =   225
         Width           =   1020
      End
      Begin VB.PictureBox picSepColors 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000080FF&
         Height          =   1020
         Index           =   5
         Left            =   1800
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   28
         Top             =   1605
         Width           =   330
      End
      Begin VB.PictureBox picSepColors 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000080FF&
         Height          =   1020
         Index           =   4
         Left            =   1455
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   27
         Top             =   1605
         Width           =   330
      End
      Begin VB.PictureBox picSepColors 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000080FF&
         Height          =   1020
         Index           =   3
         Left            =   1110
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   26
         Top             =   1605
         Width           =   330
      End
      Begin VB.PictureBox picSepColors 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000080FF&
         Height          =   1020
         Index           =   2
         Left            =   765
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   25
         Top             =   1605
         Width           =   330
      End
      Begin VB.PictureBox picSepColors 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000080FF&
         Height          =   1020
         Index           =   1
         Left            =   420
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   24
         Top             =   1605
         Width           =   330
      End
      Begin VB.PictureBox picSepColors 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000080FF&
         Height          =   1020
         Index           =   0
         Left            =   75
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   23
         Top             =   1605
         Width           =   330
      End
      Begin VB.TextBox txtRGB 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1425
         MaxLength       =   3
         TabIndex        =   21
         Text            =   "0"
         Top             =   690
         Width           =   405
      End
      Begin VB.TextBox txtRGB 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1425
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "0"
         Top             =   450
         Width           =   405
      End
      Begin VB.TextBox txtRGB 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1425
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "0"
         Top             =   195
         Width           =   405
      End
      Begin VB.CommandButton cmdRed 
         BackColor       =   &H00E0E0E0&
         Caption         =   "-R"
         Height          =   255
         Index           =   1
         Left            =   1485
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   " Decr red "
         Top             =   975
         Width           =   270
      End
      Begin VB.CommandButton cmdRed 
         BackColor       =   &H00E0E0E0&
         Caption         =   "+R"
         Height          =   255
         Index           =   0
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   " Incr Red "
         Top             =   975
         Width           =   285
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   1425
         Left            =   1860
         TabIndex        =   15
         ToolTipText     =   " PaintColor "
         Top             =   105
         Width           =   300
         Begin VB.Shape shpPaintColor 
            BackColor       =   &H008080FF&
            BackStyle       =   1  'Opaque
            FillStyle       =   0  'Solid
            Height          =   1200
            Left            =   60
            Top             =   150
            Width           =   165
         End
      End
      Begin VB.Shape shpShowColor 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FFC0FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   90
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label LabShowRGB 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 255"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1395
         TabIndex        =   38
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label LabShowRGB 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 255"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1020
         TabIndex        =   37
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label LabShowRGB 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 255"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   660
         TabIndex        =   36
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Brush angle"
      Height          =   675
      Left            =   60
      TabIndex        =   9
      Top             =   2580
      Width           =   2220
      Begin VB.OptionButton optBrush 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Index           =   4
         Left            =   1635
         Picture         =   "Main.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   " Fill "
         Top             =   180
         Width           =   405
      End
      Begin VB.OptionButton optBrush 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Index           =   3
         Left            =   1230
         Picture         =   "Main.frx":0B2C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   390
      End
      Begin VB.OptionButton optBrush 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Index           =   2
         Left            =   840
         Picture         =   "Main.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   390
      End
      Begin VB.OptionButton optBrush 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   1
         Left            =   450
         Picture         =   "Main.frx":0DD0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   195
         Width           =   390
      End
      Begin VB.OptionButton optBrush 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   0
         Left            =   60
         Picture         =   "Main.frx":0F22
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   195
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Brush size"
      Height          =   1065
      Left            =   45
      TabIndex        =   0
      Top             =   1515
      Width           =   2235
      Begin VB.OptionButton optBrushHead 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   7
         Left            =   1635
         Picture         =   "Main.frx":1074
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         Width           =   450
      End
      Begin VB.OptionButton optBrushHead 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   6
         Left            =   1140
         Picture         =   "Main.frx":11C6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   480
      End
      Begin VB.OptionButton optBrushHead 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   5
         Left            =   630
         Picture         =   "Main.frx":1318
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   480
      End
      Begin VB.OptionButton optBrushHead 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   4
         Left            =   120
         Picture         =   "Main.frx":146A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   480
      End
      Begin VB.OptionButton optBrushHead 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   3
         Left            =   1665
         Picture         =   "Main.frx":15BC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   435
      End
      Begin VB.OptionButton optBrushHead 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   2
         Left            =   1140
         Picture         =   "Main.frx":170E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   480
      End
      Begin VB.OptionButton optBrushHead 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   1
         Left            =   630
         Picture         =   "Main.frx":1860
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   480
      End
      Begin VB.OptionButton optBrushHead 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   0
         Left            =   120
         Picture         =   "Main.frx":19B2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   480
      End
   End
   Begin VB.Shape shpRG 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   180
      Left            =   2430
      Shape           =   3  'Circle
      Top             =   30
      Width           =   180
   End
   Begin VB.Line Line2 
      X1              =   159
      X2              =   159
      Y1              =   0
      Y2              =   703
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&FILE"
      Begin VB.Menu mnuFileOps 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "&Import vector file .vec"
         Index           =   2
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "&Export vector file .vec"
         Index           =   3
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "&Get in   binary file .veb"
         Index           =   5
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "&Put out binary file .veb"
         Index           =   6
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "I&mport parked colors .prk"
         Index           =   8
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "E&xport parked colors .prk"
         Index           =   9
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "&Save As bmp file"
         Index           =   11
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "&Quit"
         Index           =   13
      End
   End
   Begin VB.Menu mnuEdits 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "Undo"
         Index           =   0
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Redo"
         Index           =   1
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Clear picture"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Redraw picture"
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Delete last vector"
         Index           =   5
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cl&ip vectors to Undo"
         Index           =   6
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "D&elete all vectors"
         Index           =   7
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Fix new canvas size"
         Index           =   9
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Stretch vectors"
         Index           =   11
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "F&ast redraw"
         Checked         =   -1  'True
         Index           =   13
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1 Main.frm

' Brush Vectors  by Robert Rayment  Sep 2003

' 11 Oct - File New corrected when no vectors

' 12 Sep - Correction to drawing on already stretched canvas

' 11 Sep - Adjust DO_ _BRUSHES for BMP VEB files

' 9 Sep  - Added in/out binary .veb files
'        - Correction to Test brush
'        - Re-arrange cmdPicVec & mnuEdit code
'        - Add chkFastRedraw

' 8 Sep  - DIB & long array picMem() used for fast redraw

' 7 Sep  - picBack(picbox) changed to bBack(byte array)
'        - Use Bresline for BrushWidth =1 as well

' 5 Sep  - Addition - Stretch vectors

Option Explicit

Private CommonDialog1 As New OSDialog



Private Sub Form_Load()
Dim NB As Integer
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrentPath$ = PathSpec$
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   LCCount = 0    ' Left button press count
   aDRAW = False
    
   ' Set Slow redraw
   aFast = False
   mnuEdit(13).Checked = False
   
   ' Transparency scrollbar
   With HSalpha
      .Min = 0
      .Max = 100
      .SmallChange = 1
      .LargeChange = 10
   End With
   
   ' Initial values
   zalpha = 0.5
   HSalpha.Value = 50
   BrushSize = 13
   optBrushHead((BrushSize - 1) \ 2).Value = True
   BrushAngle = 1
   optBrush(BrushAngle).Value = True
   NB = 100 + 1 + 8 * BrushAngle + (BrushSize - 1) \ 2
   If NB > 133 Then NB = 133
   picCanvas.MousePointer = vbCustom
   picCanvas.MouseIcon = LoadResPicture(NB, vbResCursor)
   
   PaintColor = 0
   
   aHiLite = False
   aFileOps = False
   aStart = False
   aHairs = False
   Line3.Visible = False
   Line4.Visible = False
   aMagON = False
   aDupVECArray = False
   aStretch = False
   
   NumParkedColors = 0
   ReDim ParkedColors(1)
   ParkedColors(0) = NumParkedColors
   
   ReDim ParkedX(1), ParkedY(1)
   ReDim ParkedColor(1)
   
   ' Sizing
   With picCanvas
      .ScaleMode = vbPixels
      .Width = 256
      .Height = 256
      .Top = 0
      .Left = 0
   End With
   
   LabWH = "Pic W,H =" & Str$(picCanvas.Width) & "," & Str$(picCanvas.Height)
   picResizer.Left = picCanvas.Left + picCanvas.Width - picResizer.Width - 1
   picResizer.Top = picCanvas.Top + picCanvas.Height - picResizer.Height
   
   ' Each stroke drawn separately on to bBack()
   ' to allow X-fade check on pixels
   ReDim bBack(picCanvas.Width, picCanvas.Height)
   ReDim picMem(1, 1)
   
   'ReDim picMem(1 To picCanvas.Width, 1 To picCanvas.Height)
   
   ' Set up VEC array
   NumVectors = 0
   NumVisVectors = 0
   LabNumVisStrokes = "No. Vectors =" & Str$(NumVectors) & vbCr & "No. Visible =" & Str$(NumVisVectors)
   VECSIZE = 4
   ReDim VEC(0 To VECSIZE)
   VEC(0) = 0        ' NumVectors
   VEC(1) = picCanvas.Width
   VEC(2) = picCanvas.Height
   VEC(3) = -2       ' END
   VEC(4) = -2       ' END
   aRUNVEC = False

   ' Set up picColorBox
   For j = 0 To 63
   For i = 0 To 63
      SetPixelV picColorBox.hdc, i, j, RGB(160, 4 * i + 3, 4 * j + 3)
   Next i
   Next j
   picColorBox.Refresh
   
   FillSepColors
   
   FixScrollbars picFrame, picCanvas, HS, VS
      
   ' For space bar start stroke
   KeyPreview = True
   Xpic = -1
   Ypic = -1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aHiLite Then ClearAllHiLites
   LabXY = "X,Y =" & Str$(X) & Str$(Y)
   Xpic = -1   ' To block space bar
   Ypic = -1
End Sub

Private Sub Form_Resize()
   If WindowState = vbMinimized Then
      If aHelp Then
         Unload frmHelp
         aHelp = False
      End If
      Exit Sub
   End If

   picFrame.Width = Me.Width / STX - 222
   picFrame.Height = Me.Height / STY - 90
   FixScrollbars picFrame, picCanvas, HS, VS
   HS.Top = picFrame.Top + picFrame.Height + 2
   fraScratch.Left = Screen.Width
   fraScratch.Top = 70
   fraMOVER Form1, fraScratch, 1, Me.Width / STX, Me.Height / STX
   ' Checker picFrame
   picFrame.Cls
   picFrame.BackColor = vbWhite
   For j = 0 To picFrame.Height Step 16
   For i = 0 To picFrame.Width Step 16
      picFrame.Line (i, j)-(i + 8, j + 8), &HE0E0E0, BF
      picFrame.Line (i + 8, j + 8)-(i + 16, j + 16), &HE0E0E0, BF
   Next i
   Next j
   picFrame.Refresh
End Sub

Private Sub HSRGB_Change(Index As Integer)
txtRGB(Index).Text = LTrim$(Str$(HSRGB(Index).Value))
End Sub



'#### FILE OPS ####################################

Private Sub mnuFile_Click()
   If aHiLite Then ClearAllHiLites
   If aDRAW Then Exit Sub
   If aStretch Then Exit Sub
End Sub

Private Sub mnuFileOps_Click(Index As Integer)
   If aDRAW Then Exit Sub
   If aStretch Then Exit Sub
   
   aFileOps = True
   Set CommonDialog1 = New OSDialog
   
   Select Case Index
   Case 0       ' New
      aFileOps = False
      If NumVectors <> 0 Then
         cmdPicVec_MouseUp 5, 1, 0, 0, 0      ' Delete all vectors
      End If
      FixScrollbars picFrame, picCanvas, HS, VS
   
   Case 2   ' Import .vec
      IMPORT_VEC
      If aFileOps Then
         Screen.MousePointer = vbHourglass
         picCanvas.Picture = LoadPicture
         ReDim picMem(1 To picCanvas.Width, 1 To picCanvas.Height)
   
         DrawVectors NumVectors
         FixScrollbars picFrame, picCanvas, HS, VS
         Screen.MousePointer = vbDefault
         aDupVECArray = False
         ReDim VECORG(0)
         FixScrollbars picFrame, picCanvas, HS, VS
      End If
   Case 3   ' Export .vec
      EXPORT_VEC
   
   Case 4   ' --
   
   Case 5   ' Get in  binary .veb
      GET_VEB
      If aFileOps Then
         Screen.MousePointer = vbHourglass
         picCanvas.Picture = LoadPicture
         ReDim picMem(1 To picCanvas.Width, 1 To picCanvas.Height)
   
         DrawVectors NumVectors
         FixScrollbars picFrame, picCanvas, HS, VS
         Screen.MousePointer = vbDefault
         aDupVECArray = False
         ReDim VECORG(0)
         FixScrollbars picFrame, picCanvas, HS, VS
      End If
   
   Case 6   ' Put out binary .veb
      PUT_VEB
   
   Case 7   ' --
   
   Case 8   ' Import .prk
      IMPORT_PRK
      If aFileOps Then
         picScratch(0).Cls
         picScratch(0).DrawWidth = 20
         NumParkedColors = ParkedColors(0)
         ' Re-draw parked colors
         If NumParkedColors > 0 Then
            For i = 1 To NumParkedColors Step 3
               picScratch(0).PSet (ParkedColors(i), ParkedColors(i + 1)), ParkedColors(i + 2)
            Next i
         End If
      End If
   Case 9   ' Export .prk
      EXPORT_PRK
   
   Case 10   ' --
   
   Case 11   ' Save As bmp
      SAVE_BMP
   
   Case 12   ' --
   
   Case 13   ' Quit
      Set CommonDialog1 = Nothing
      Form_QueryUnload 1, 0
   End Select
   
   Set CommonDialog1 = Nothing
   
   If NumVisVectors > 0 Then
      picResizer.Visible = False
   Else
      picResizer.Visible = True
   End If

   aFileOps = False
End Sub

Private Sub IMPORT_VEC()
Dim FileSpec$, Title$, Filt$, InDir$
   
   Title$ = "Import VEC File"
   Filt$ = "Load vec (.vec)|*.vec"
   InDir$ = CurrentPath$ 'Pathspec$
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd

   SetCursorPos Me.Left \ STX + picFrame.Left - 4, 380
   
   If Len(FileSpec$) = 0 Then
      aFileOps = False
      Exit Sub
   Else
      READ_VEC_FILE FileSpec$
   End If
End Sub

Private Sub EXPORT_VEC()
Dim FileSpec$, Title$, Filt$, InDir$

   If VECSIZE <= 2 Then
      MsgBox "No vectors yet", vbInformation, "Export vec"
      Exit Sub
   End If
   
   Title$ = "Export VEC File"
   Filt$ = "Save vec (.vec)|*.vec"
   InDir$ = CurrentPath$ 'Pathspec$
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd

   If Len(FileSpec$) = 0 Then
      aFileOps = False
      Exit Sub
   Else
      FixExtension FileSpec$, "vec"
      CurrentPath$ = FileSpec$
      SAVE_VEC_FILE FileSpec$
   End If
End Sub

Private Sub PUT_VEB()
Dim FileSpec$, Title$, Filt$, InDir$
   
   If VECSIZE <= 2 Then
      MsgBox "No vectors yet", vbInformation, "Export vec"
      Exit Sub
   End If
   
   Title$ = "Save binary VEB File"
   Filt$ = "Save veb (.veb)|*.veb"
   InDir$ = CurrentPath$ 'Pathspec$
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd

   If Len(FileSpec$) = 0 Then
      aFileOps = False
      Exit Sub
   Else
      FixExtension FileSpec$, "veb"
      CurrentPath$ = FileSpec$
      SAVE_VEB_FILE FileSpec$
   End If
End Sub

Private Sub GET_VEB()
Dim FileSpec$, Title$, Filt$, InDir$

   Title$ = "Load binary VEB File"
   Filt$ = "Load vec (.veb)|*.veb"
   InDir$ = CurrentPath$ 'Pathspec$
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd

   SetCursorPos Me.Left \ STX + picFrame.Left - 4, 380
   
   If Len(FileSpec$) = 0 Then
      aFileOps = False
      Exit Sub
   Else
      READ_VEB_FILE FileSpec$
   End If
End Sub

Private Sub IMPORT_PRK()
Dim FileSpec$, Title$, Filt$, InDir$
   
   Title$ = "Import Parked colors File"
   Filt$ = "Load prk (.prk)|*.prk"
   InDir$ = CurrentPath$ 'Pathspec$
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd

   SetCursorPos Me.Left \ STX + picFrame.Left - 4, 380
   
   If Len(FileSpec$) = 0 Then
      aFileOps = False
      Exit Sub
   Else
      READ_PRK_FILE FileSpec$
   End If
End Sub

Private Sub EXPORT_PRK()
Dim FileSpec$, Title$, Filt$, InDir$

   If NumParkedColors = 0 Then
      MsgBox "No parked colors yet", vbInformation, "Export prk"
      Exit Sub
   End If
   
   Title$ = "Export Parked colors File"
   Filt$ = "Save prk (.prk)|*.prk"
   InDir$ = CurrentPath$ 'Pathspec$
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd

   If Len(FileSpec$) = 0 Then
      aFileOps = False
      Exit Sub
   Else
      FixExtension FileSpec$, "prk"
      CurrentPath$ = FileSpec$
      SAVE_PRK_FILE FileSpec$
   End If
End Sub

Private Sub SAVE_BMP()
Dim FileSpec$, Title$, Filt$, InDir$

   If VECSIZE <= 2 Then
      MsgBox "No vectors yet", vbInformation, "Save bmp"
      Exit Sub
   End If
   
   Title$ = "Save display as BMP File"
   Filt$ = "Save display (.bmp)|*.bmp"
   InDir$ = CurrentPath$ 'Pathspec$
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd

   If Len(FileSpec$) = 0 Then
      aFileOps = False
      Exit Sub
   Else
      FixExtension FileSpec$, "bmp"
      CurrentPath$ = FileSpec$
      SavePicture picCanvas.Image, FileSpec$
   End If

End Sub
'#### END FILE OPS ####################################

'### HELP ######################################################

Private Sub mnuHelp_Click()
Dim a$
   If aDRAW Then Exit Sub
   If aStretch Then Exit Sub

   a$ = PathSpec$ & "BrushHelp.txt"
   If Len(Dir$(a$)) = 0 Then
      MsgBox "BrushHelp.txt missing ", , "Brush vectors - Help"
      Exit Sub
   Else
      aHelp = True
      'frmHelp.Hide  ' Allows vbModal disabling other forms
      frmHelp.Show vbModeless ''vbModal
   End If
End Sub
'### END HELP ######################################################

'#### MAGNIFIER #######################################

Private Sub cmdMagnifier_Click()
'&M
   aMagON = Not aMagON
   
   If aMagON Then
      MagForm.Show
   Else
      MagForm.Hide
   End If
   picCanvas.SetFocus
End Sub

'#### TOGGLE + HAIRS ######################################

Private Sub cmdPlusHairs_Click() 'Button As Integer, Shift As Integer, X As Single, Y As Single)
   aHairs = Not aHairs

   If aHairs Then
      picCanvas.SetFocus
      Line3.Visible = True
      Line4.Visible = True
      
      Line3.X1 = picCanvas.CurrentX 'X
      Line3.X2 = picCanvas.CurrentX 'X
      Line3.Y1 = 0
      Line3.Y2 = picCanvas.Height
      Line4.X1 = 0
      Line4.X2 = picCanvas.Width
      Line4.Y1 = picCanvas.CurrentY 'Y
      Line4.Y2 = picCanvas.CurrentY 'Y
   Else
      Line3.Visible = False
      Line4.Visible = False
   End If

   picCanvas.SetFocus
End Sub

Private Sub picFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aHiLite Then ClearAllHiLites
   LabXY = "X,Y =" & Str$(X) & Str$(Y)
   Xpic = -1   ' To block space bar
   Ypic = -1
End Sub

'#### PIC RESIZER ###########################

Private Sub picResizer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   fraX = X
   fraY = Y
End Sub

Private Sub picResizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Only resizes a starting canvas
Dim HLeft As Long, HTop As Long
Dim picWidth As Long, picHeight As Long

   Xpic = -1   ' To block space bar
   Ypic = -1
   If aFileOps Then Exit Sub
   If aHiLite Then ClearAllHiLites
   
   If NumVisVectors <> 0 Then
      MsgBox " Clear picture first, resize" & vbCr & "and then Redraw", vbExclamation, "Resizing picture"
      Exit Sub
   End If
   
   On Error GoTo LabHError

   If Button = vbLeftButton Then
      ' Test  new position
      
      HLeft = picResizer.Left + (X - fraX)
      HTop = picResizer.Top + (Y - fraY)
      
      picWidth = (HLeft + picResizer.Width)
      picHeight = (HTop + picResizer.Height)
      
      ' Limit lower size to 16x16
      ' Resize picCanvas
      If picWidth < 16 Then picWidth = 16
      picCanvas.Width = picWidth
      
      If picHeight < 16 Then picHeight = 16
      picCanvas.Height = picHeight
      
      'Re-position picResizer to new picCanvas size
      
      HLeft = picCanvas.Width - picResizer.Width - 1
      HTop = picCanvas.Height - picResizer.Height - 1
      
      picResizer.Left = HLeft
      picResizer.Top = HTop
      
      FixScrollbars picFrame, picCanvas, HS, VS
      
      LabWH = "Pic W,H =" & Str$(picCanvas.Width) & "," & Str$(picCanvas.Height)

      ReDim bBack(picCanvas.Width, picCanvas.Height)
   
   
   End If

Exit Sub
'========
LabHError:
picResizer.Left = picCanvas.Left + picCanvas.Width - picResizer.Width - 1
picResizer.Top = picCanvas.Top + picCanvas.Height - picResizer.Height
End Sub
'#### END PIC RESIZER ###########################


'#### CANVAS SCROLL BARS #########################################

Private Sub HS_Change()
   Xpic = -1   ' To block space bar
   Ypic = -1
   picCanvas.Left = -HS.Value
End Sub

Private Sub HS_Scroll()
   Xpic = -1   ' To block space bar
   Ypic = -1
   picCanvas.Left = -HS.Value
End Sub

Private Sub VS_Change()
   Xpic = -1   ' To block space bar
   Ypic = -1
   picCanvas.Top = -VS.Value
End Sub

Private Sub VS_Scroll()
   Xpic = -1   ' To block space bar
   Ypic = -1
   picCanvas.Top = -VS.Value
End Sub
'#### END CANVAS SCROLL BARS #########################################

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aHiLite Then ClearAllHiLites
End Sub

Private Sub ClearAllHiLites()
   For i = 0 To 8
      cmdPicVec(i).BackColor = RGB(&HE0, &HE0, &HE0)
   Next i
   aHiLite = False
End Sub

Private Sub HSalpha_Change()
   If aFileOps Then Exit Sub
   If aDRAW Then Exit Sub
   If aStretch Then Exit Sub
   
   zalpha = HSalpha.Value / 100
   LabAlpha = Str$(100 - HSalpha.Value)
End Sub

Private Sub HSalpha_Scroll()
   If aFileOps Then Exit Sub
   If aDRAW Then Exit Sub
   If aStretch Then Exit Sub
   
   zalpha = HSalpha.Value / 100
   LabAlpha = Str$(100 * (1 - zalpha))
End Sub

'#### SELECT BRUSHES #######################################################################

Private Sub optBrushHead_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Brush Size
' 0(1), 1(3),  2(5),  3(7),
' 4(9), 5(11), 6(13), 7(15)
Dim NB As Long
   
   If aHiLite Then ClearAllHiLites
   If aFileOps Then
      optBrushHead(Index).Value = 0
      optBrushHead((BrushSize - 1) \ 2).Value = 1
      Exit Sub
   End If
   If aDRAW Then
      optBrushHead(Index).Value = 0
      optBrushHead((BrushSize - 1) \ 2).Value = 1
      Exit Sub
   End If
   
   BrushSize = 2 * Index + 1
   NB = 100 + 1 + 8 * BrushAngle + (BrushSize - 1) \ 2
   If NB > 132 Then NB = 132
   picCanvas.MousePointer = vbCustom
   picCanvas.MouseIcon = LoadResPicture(NB, vbResCursor)

   picCanvas.SetFocus
End Sub

Private Sub optBrush_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Brush Angle
' Brush Angle 0(-), 1(\), 2(/), 3(|), 4(Fill)
Dim NB As Long
   
   If aHiLite Then ClearAllHiLites
   If aFileOps Then
      optBrush(Index).Value = 0
      optBrush(BrushAngle).Value = 1
      Exit Sub
   End If
   If aDRAW Then
      optBrush(Index).Value = 0
      optBrush(BrushAngle).Value = 1
      Exit Sub
   End If
   
   BrushAngle = Index
   NB = 100 + 1 + 8 * BrushAngle + (BrushSize - 1) \ 2
   If NB > 133 Then NB = 133
   picCanvas.MousePointer = vbCustom
   picCanvas.MouseIcon = LoadResPicture(NB, vbResCursor)

   picCanvas.SetFocus
End Sub
'#### END SELECT BRUSHES #######################################################################


'#### CANVAS #########################################################

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then
      If Xpic <> -1 And Ypic <> -1 Then
         picCanvas_MouseUp 1, 0, Xpic, Ypic
      End If
   End If
End Sub

Private Sub picCanvas_LostFocus()
   Xpic = -1   ' To block space bar
   Ypic = -1
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aHiLite Then ClearAllHiLites
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aRUNVEC Then Exit Sub
   
   If X < 0 Then X = 0
   If X > picCanvas.Width - 1 Then X = picCanvas.Width - 1
   If Y < 0 Then Y = 0
   If Y > picCanvas.Height - 1 Then Y = picCanvas.Height - 1
   
   If Shift = 1 Then X = Xprev
   If Shift = 2 Then Y = Yprev
   
   If aHairs Then
      Line3.X1 = X
      Line3.X2 = X
      Line3.Y1 = 0
      Line3.Y2 = picCanvas.Height
      Line4.X1 = 0
      Line4.X2 = picCanvas.Width
      Line4.Y1 = Y
      Line4.Y2 = Y
   End If
   
   ' For Space bar start
   Xpic = X
   Ypic = Y
   
   LabXY = "X,Y =" & Str$(X) & Str$(Y)
   
   ' Show color under cursor
   Cul = picCanvas.Point(X, Y)
   If Cul >= 0 Then
      shpShowColor.FillColor = Cul
      shpShowColor.Refresh
      LabShowRGB(0) = Str$(Cul And &HFF&)
      LabShowRGB(1) = Str$((Cul And &HFF00&) / &H100&)
      LabShowRGB(2) = Str$((Cul And &HFF0000) / &H10000)
   End If
   
   If BrushAngle = 4 Then Exit Sub   ' Fill
   
   If LCCount = 1 Then
      If CInt(X) <> VEC(VECSIZE - 1) Or CInt(Y) <> VEC(VECSIZE) Then
         DO_SLOW_BRUSH picCanvas, ByVal X, ByVal Y
         VECSIZE = VECSIZE + 2
         ReDim Preserve VEC(0 To VECSIZE)
         VEC(VECSIZE - 1) = CInt(X)
         VEC(VECSIZE) = CInt(Y)
      End If
   End If
End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aRUNVEC Then Exit Sub
   
   If X < 0 Then X = 0
   If X > picCanvas.Width - 1 Then X = picCanvas.Width - 1
   If Y < 0 Then Y = 0
   If Y > picCanvas.Height - 1 Then Y = picCanvas.Height - 1
   
   If Button = vbRightButton Then
      Cul = picCanvas.Point(X, Y)
      If Cul >= 0 Then
         PaintColor = Cul
         bPaintRed = PaintColor And &HFF&
         bPaintGreen = (PaintColor And &HFF00&) / &H100&
         bPaintBlue = (PaintColor And &HFF0000) / &H10000
         shpPaintColor.FillColor = PaintColor
         shpPaintColor.Refresh
         txtRGB(0) = bPaintRed
         txtRGB(1) = bPaintGreen
         txtRGB(2) = bPaintBlue
         HSRGB(0).Value = bPaintRed
         HSRGB(1).Value = bPaintGreen
         HSRGB(2).Value = bPaintBlue
      
         If aDRAW Then
               If CInt(X) <> VEC(VECSIZE - 1) Or CInt(Y) <> VEC(VECSIZE) Then
               VECSIZE = VECSIZE + 4
               ReDim Preserve VEC(0 To VECSIZE)
               VEC(VECSIZE - 3) = CInt(X)
               VEC(VECSIZE - 2) = CInt(Y)
            Else
               VECSIZE = VECSIZE + 2
               ReDim Preserve VEC(0 To VECSIZE)
            End If
      
            VEC(VECSIZE - 1) = -2
            VEC(VECSIZE) = -2
            
            LCCount = 0 ' Ready for new stroke
            aDRAW = False
            shpRG.BackColor = vbRed
            ' Block space bar start in case cursor goes off canvas
            Xpic = -1
            Ypic = -1
         End If   ' If aDRAW Then
      
      End If   ' If Cul >= 0 Then
      Exit Sub
   End If   ' If Button = vbRightButton Then
   
   If Button = vbLeftButton Then
      
      LCCount = LCCount + 1
      
      If LCCount = 1 Then
      
         aDRAW = True
         shpRG.BackColor = vbGreen
         
         Xprev = X
         Yprev = Y
   
         Select Case BrushAngle
         Case 4   ' Fill
            Fill picCanvas, X, Y

            VEC(0) = VEC(0) + 1  ' NumVectors
            NumVectors = VEC(0)
            NumVisVectors = NumVectors
            If NumVectors = 1 Then  ' put in canvas size
               VEC(1) = picCanvas.Width
               VEC(2) = picCanvas.Height
            End If
            ' Change -2 to -1
            VEC(VECSIZE - 1) = -1
            VEC(VECSIZE) = VEC(0)
            VECSIZE = VECSIZE + 8
            ReDim Preserve VEC(0 To VECSIZE)
            VEC(VECSIZE - 7) = BrushSize
            VEC(VECSIZE - 6) = BrushAngle
            VEC(VECSIZE - 5) = CInt(bPaintRed)
            VEC(VECSIZE - 4) = CInt(bPaintGreen)
            VEC(VECSIZE - 3) = CInt(bPaintBlue)
            VEC(VECSIZE - 2) = CInt(zalpha * 100)
            VEC(VECSIZE - 1) = CInt(X)
            VEC(VECSIZE) = CInt(Y)
            VECSIZE = VECSIZE + 2
            ReDim Preserve VEC(0 To VECSIZE)
            VEC(VECSIZE - 1) = -2
            VEC(VECSIZE) = -2
            
            LCCount = 0  ' Ready for new stroke
            aDRAW = False
            shpRG.BackColor = vbRed
            ' Block space bar start in case cursor goes off canvas
            Xpic = -1
            Ypic = -1
      
         Case Else   ' Not Fill
      
            START_BRUSH picCanvas, ByVal X, ByVal Y
            
            VEC(0) = VEC(0) + 1  ' NumVectors
            NumVectors = VEC(0)
            NumVisVectors = NumVectors
            If NumVectors = 1 Then  ' put in canvas size
               VEC(1) = picCanvas.Width
               VEC(2) = picCanvas.Height
            End If
            ' Change -2 to -1
            VEC(VECSIZE - 1) = -1
            VEC(VECSIZE) = VEC(0)
            
            VECSIZE = VECSIZE + 8
            ReDim Preserve VEC(0 To VECSIZE)
            VEC(VECSIZE - 7) = BrushSize
            VEC(VECSIZE - 6) = BrushAngle
            VEC(VECSIZE - 5) = CInt(bPaintRed)
            VEC(VECSIZE - 4) = CInt(bPaintGreen)
            VEC(VECSIZE - 3) = CInt(bPaintBlue)
            VEC(VECSIZE - 2) = CInt(zalpha * 100)
            VEC(VECSIZE - 1) = CInt(X)
            VEC(VECSIZE) = CInt(Y)
         End Select
      
      Else  ' LCCount>1   END VECTOR
   
         If Shift = 1 Then X = Xprev
         If Shift = 2 Then Y = Yprev
   
         If CInt(X) <> VEC(VECSIZE - 1) Or CInt(Y) <> VEC(VECSIZE) Then
            VECSIZE = VECSIZE + 4
            ReDim Preserve VEC(0 To VECSIZE)
            VEC(VECSIZE - 3) = CInt(X)
            VEC(VECSIZE - 2) = CInt(Y)
         Else
            VECSIZE = VECSIZE + 2
            ReDim Preserve VEC(0 To VECSIZE)
         End If
      
         VEC(VECSIZE - 1) = -2
         VEC(VECSIZE) = -2
         
         LCCount = 0             ' Ready for new stroke
         aDRAW = False           ' Allow other actions
         shpRG.BackColor = vbRed ' Show ready for new stroke
         ' Block space bar start in case cursor goes off canvas
         Xpic = -1
         Ypic = -1
      End If   ' If LCCount = 1 Then
   
   End If   ' If Button = vbLeftButton Then
      
   LabNumVisStrokes = "No. Vectors =" & Str$(NumVectors) & vbCr & "No. Visible =" & Str$(NumVisVectors)
   HSalpha.SetFocus
   If NumVectors > 0 Then picResizer.Visible = False
End Sub
'#### END CANVAS #########################################################


'### COLORS ######################################################

Private Sub picColorBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   Cul = picColorBox.Point(X, Y)
   
   If Cul >= 0 Then
      shpShowColor.FillColor = Cul
      shpShowColor.Refresh
      LabShowRGB(0) = Str$(Cul And &HFF&)
      LabShowRGB(1) = Str$((Cul And &HFF00&) / &H100&)
      LabShowRGB(2) = Str$((Cul And &HFF0000) / &H10000)
   End If
End Sub

Private Sub picColorBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   Cul = picColorBox.Point(X, Y)
   If Cul >= 0 Then
      PaintColor = Cul
      bPaintRed = PaintColor And &HFF&
      bPaintGreen = (PaintColor And &HFF00&) / &H100&
      bPaintBlue = (PaintColor And &HFF0000) / &H10000
      shpPaintColor.FillColor = PaintColor
      shpPaintColor.Refresh
      txtRGB(0) = bPaintRed
      txtRGB(1) = bPaintGreen
      txtRGB(2) = bPaintBlue
      HSRGB(0).Value = bPaintRed
      HSRGB(1).Value = bPaintGreen
      HSRGB(2).Value = bPaintBlue
   End If
End Sub


Private Sub picSepColors_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   Cul = picSepColors(Index).Point(X, Y)
   If Cul >= 0 Then
      shpShowColor.FillColor = Cul
      shpShowColor.Refresh
      LabShowRGB(0) = Str$(Cul And &HFF&)
      LabShowRGB(1) = Str$((Cul And &HFF00&) / &H100&)
      LabShowRGB(2) = Str$((Cul And &HFF0000) / &H10000)
   End If
End Sub

Private Sub picSepColors_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   Cul = picSepColors(Index).Point(X, Y)
   If Cul >= 0 Then
      PaintColor = Cul
      bPaintRed = PaintColor And &HFF&
      bPaintGreen = (PaintColor And &HFF00&) / &H100&
      bPaintBlue = (PaintColor And &HFF0000) / &H10000
      shpPaintColor.FillColor = PaintColor
      shpPaintColor.Refresh
      txtRGB(0) = bPaintRed
      txtRGB(1) = bPaintGreen
      txtRGB(2) = bPaintBlue
      HSRGB(0).Value = bPaintRed
      HSRGB(1).Value = bPaintGreen
      HSRGB(2).Value = bPaintBlue
   End If
End Sub

Private Sub txtRGB_Change(Index As Integer)
Dim txtVal As Integer
   
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   'If Not IsNumeric(txtRGB(Index).Text) Then txtRGB(Index).Text = "0"
   If IsNumeric(txtRGB(Index).Text) Then
   
      If Len(txtRGB(Index).Text) = 0 Then
         txtVal = 0
      Else
         txtVal = Val(txtRGB(Index).Text)
      End If
      
      If txtVal < 0 Then
         txtVal = 0
         txtRGB(Index).Text = "0"
      End If
      If txtVal > 255 Then
         txtVal = 255
         txtRGB(Index).Text = "255"
      End If
      
      Select Case Index
      Case 0   ' R
         bPaintRed = txtVal
      Case 1   ' G
         bPaintGreen = txtVal
      Case 2   ' B
         bPaintBlue = txtVal
      End Select
      
      PaintColor = RGB(bPaintRed, bPaintGreen, bPaintBlue)
      shpPaintColor.FillColor = PaintColor
      shpPaintColor.Refresh
   
      DoEvents
   
   End If
End Sub

Private Sub cmdRed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' 0 + Red
' 1 - Red
Dim bbR As Byte
Dim bbG As Byte
Dim bbB As Byte

   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   aPMRed = False
   
   Do
      ' Vary Red in picColorBox
      
      Cul = picColorBox.Point(3, 3)
      bbR = Cul And &HFF&
      If Index = 0 Then
         If bbR <= 247 Then
            bbR = bbR + 8 'bbR = bbR + 4
         Else
            bbR = 255
         End If
         
      Else
         If bbR >= 8 Then
            bbR = bbR - 8 'bbR = bbR - 4
         Else
            bbR = 8
         End If
      End If
      LabShowRGB(0) = Str$(bbR)
      
      For j = 0 To 63
      For i = 0 To 63
         Cul = picColorBox.Point(i, j)
         bbG = (Cul And &HFF00&) / &H100&
         bbB = (Cul And &HFF0000) / &H10000
         SetPixelV picColorBox.hdc, i, j, RGB(bbR, bbG, bbB)
      Next i
      Next j
      
      DoEvents
   
   Loop Until aPMRed
   
   picColorBox.Refresh
End Sub

Private Sub cmdRed_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   aPMRed = True
End Sub

Private Sub FillSepColors()
Dim byteArray() As Byte
Dim L As Long
Dim U As Long

byteArray = LoadResData(101, "CUSTOM")
   
U = UBound(byteArray)   ' 0 to 4304
L = LBound(byteArray)

Open PathSpec$ & "~tempPAL.tmp" For Binary As #1
Put #1, , byteArray
Close
   
   'Open PathSpec$ & "AllColors.PAL" For Input As #1
   Open PathSpec$ & "~tempPAL.tmp" For Input As #1
   
   For i = 0 To 383
      Input #1, bR, bG, bB
      Select Case i
      Case Is < 64
         picSepColors(0).Line (0, i)-(picSepColors(0).Width, i), RGB(bR, bG, bB)
      Case Is < 128
         picSepColors(1).Line (0, i - 64)-(picSepColors(1).Width, i - 64), RGB(bR, bG, bB)
      Case Is < 192
         picSepColors(2).Line (0, i - 128)-(picSepColors(2).Width, i - 128), RGB(bR, bG, bB)
      Case Is < 256
         picSepColors(3).Line (0, i - 192)-(picSepColors(3).Width, i - 192), RGB(bR, bG, bB)
      Case Is < 320
         picSepColors(4).Line (0, i - 256)-(picSepColors(4).Width, i - 256), RGB(bR, bG, bB)
      Case Is < 384
         picSepColors(5).Line (0, i - 320)-(picSepColors(5).Width, i - 320), RGB(bR, bG, bB)
      End Select
   Next i
   Close
   
   Erase byteArray
   Kill PathSpec$ & "~tempPAL.tmp"
End Sub
'### END COLORS ######################################################


'#### SCRATCH #####################################################################

Private Sub picScratch_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   Select Case Index
   Case 0   ' Park & Pick-up colors
      picScratch(Index).DrawWidth = 20
      If Button = vbLeftButton Then ' Park color
         picScratch(Index).PSet (X, Y), PaintColor
         NumParkedColors = NumParkedColors + 3
         ReDim Preserve ParkedColors(0 To NumParkedColors)
         ParkedColors(0) = NumParkedColors
         ParkedColors(NumParkedColors - 2) = CLng(X)
         ParkedColors(NumParkedColors - 1) = CLng(Y)
         ParkedColors(NumParkedColors) = PaintColor

      ElseIf Button = vbRightButton Then  ' Pick-up color
         Cul = picScratch(Index).Point(X, Y)
         If Cul >= 0 Then
            PaintColor = Cul
            bPaintRed = PaintColor And &HFF&
            bPaintGreen = (PaintColor And &HFF00&) / &H100&
            bPaintBlue = (PaintColor And &HFF0000) / &H10000
            shpPaintColor.FillColor = PaintColor
            shpPaintColor.Refresh
            txtRGB(0) = bPaintRed
            txtRGB(1) = bPaintGreen
            txtRGB(2) = bPaintBlue
         End If
      End If
   Case 1   ' Test brush
      Xprev = X
      Yprev = Y
      
      If Button = vbLeftButton Then
         ReDim bBack(picCanvas.Width, picCanvas.Height)
         If BrushAngle = 4 Then
            Fill picScratch(1), X, Y
         Else
            CanvasColor = picScratch(Index).Point(X, Y)
            bCanvasRed = CanvasColor And &HFF&
            bCanvasGreen = (CanvasColor And &HFF00&) / &H100&
            bCanvasBlue = (CanvasColor And &HFF0000) / &H10000
            bR = zalpha * (1& * bPaintRed - bCanvasRed) + bCanvasRed
            bG = zalpha * (1& * bPaintGreen - bCanvasGreen) + bCanvasGreen
            bB = zalpha * (1& * bPaintBlue - bCanvasBlue) + bCanvasBlue
            picScratch(Index).PSet (X, Y), RGB(bR, bG, bB)
         End If
      End If
   End Select
End Sub

Private Sub picScratch_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   Select Case Index
   Case 0   ' Park & Pick-up colors
      Cul = picScratch(0).Point(X, Y)
      If Cul >= 0 Then
         shpShowColor.FillColor = Cul
         shpShowColor.Refresh
         LabShowRGB(0) = Str$(Cul And &HFF&)
         LabShowRGB(1) = Str$((Cul And &HFF00&) / &H100&)
         LabShowRGB(2) = Str$((Cul And &HFF0000) / &H10000)
      End If
   Case 1   ' Test brush
      If Button = vbLeftButton Then
         If BrushAngle = 4 Then Exit Sub
         DO_SLOW_BRUSH picScratch(1), X, Y
      End If
   End Select
End Sub

Private Sub cmdClrScratch_Click(Index As Integer)
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   picScratch(Index).Picture = LoadPicture
   If Index = 0 Then
      NumParkedColors = 0
      ReDim ParkedColors(1)
   End If
End Sub

Private Sub fraScratch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   fraX = X
   fraY = Y
End Sub

Private Sub fraScratch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Xpic = -1   ' To block space bar
   Ypic = -1
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   fraMOVER Form1, fraScratch, Button, X, Y
End Sub

Private Sub LabScratch_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aFileOps Then Exit Sub
   If aStretch Then Exit Sub
   If aDRAW Then Exit Sub
   
   fraX = LabScratch(Index).Left + X
   fraY = LabScratch(Index).Top + Y
End Sub

Private Sub LabScratch_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   fraMOVER Form1, fraScratch, Button, LabScratch(Index).Left + X, LabScratch(Index).Top + Y
End Sub
'#### END SCRATCH #####################################################################


'#### VECTORS ########################################################################

Private Sub chkFastRedraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aHiLite Then ClearAllHiLites
End Sub

Private Sub chkFastRedraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aHiLite Then ClearAllHiLites
   mnuEdit_Click (13)
End Sub

Private Sub mnuEdit_Click(Index As Integer)
   If aDRAW Then Exit Sub
   Select Case Index
   Case 0: cmdPicVec_MouseUp 0, 1, 0, 0, 0      ' Undo
   Case 1: cmdPicVec_MouseUp 1, 1, 0, 0, 0      ' Redo
   Case 2: cmdPicVec_MouseUp 2, 1, 0, 0, 0      ' Clear picture
   Case 3: cmdPicVec_MouseUp 4, 1, 0, 0, 0      ' Redraw
   Case 4:  ' --
   Case 5: cmdPicVec_MouseUp 3, 1, 0, 0, 0      ' Delete last vector
   Case 6: cmdPicVec_MouseUp 6, 1, 0, 0, 0      ' Clip vectors
   Case 7: cmdPicVec_MouseUp 5, 1, 0, 0, 0      ' Delete all vectors
   Case 8:   ' --
   Case 9: cmdPicVec_MouseUp 7, 1, 0, 0, 0   ' Fix new canvas size
   Case 10  ' --
   Case 11: cmdPicVec_MouseUp 8, 1, 0, 0, 0  ' Stretch vectors
   Case 12  ' --
   Case 13  ' Fast redraw
      aFast = Not aFast
      mnuEdit(13).Checked = Not mnuEdit(13).Checked
   End Select
End Sub

Private Sub cmdPicVec_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aFileOps Then Exit Sub
   If aDRAW Then Exit Sub
   
   For i = 0 To 8
      If i = Index Then
         cmdPicVec(i).BackColor = vbWhite
      Else
         cmdPicVec(i).BackColor = RGB(&HE0, &HE0, &HE0)
      End If
   Next i
   aHiLite = True
End Sub

Private Sub cmdPicVec_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim v As Long
Dim NV As Long
   
   If aFileOps Then Exit Sub
   If aDRAW Then Exit Sub
   
   If NumVectors = 0 Then
      MsgBox " No vectors yet", vbExclamation, "Vector ops"
      Exit Sub
   End If
   
   Select Case Index
   Case 0   ' Undo
      If NumVisVectors > 0 Then
         NumVisVectors = NumVisVectors - 1
          Screen.MousePointer = vbHourglass
          picCanvas.Picture = LoadPicture
          DrawVectors NumVisVectors
         Screen.MousePointer = vbDefault
      End If
      
   Case 1   ' Redo
      If NumVisVectors < NumVectors Then
         NumVisVectors = NumVisVectors + 1
          Screen.MousePointer = vbHourglass
          picCanvas.Picture = LoadPicture
          DrawVectors NumVisVectors
          Screen.MousePointer = vbDefault
      End If
   
   Case 2   ' Clear picture
      picCanvas.Picture = LoadPicture
      If aFast Then ReDim picMem(1 To picCanvas.Width, 1 To picCanvas.Height)
      NumVisVectors = 0
      LabWH = "Pic W,H =" & Str$(picCanvas.Width) & "," & Str$(picCanvas.Height)
      picResizer.Left = picCanvas.Left + picCanvas.Width - picResizer.Width - 1
      picResizer.Top = picCanvas.Top + picCanvas.Height - picResizer.Height

      
   Case 3   ' Del last vector
      If VECSIZE > 4 Then
         For v = VECSIZE To 1 Step -1
            If VEC(v) = -1 Then
               VECSIZE = v + 1
               ReDim Preserve VEC(0 To VECSIZE)
               VEC(VECSIZE - 1) = -2
               VEC(VECSIZE) = -2
               
               NumVectors = NumVectors - 1
               If NumVectors <= 32767 Then
                  VEC(0) = CInt(NumVectors)
               Else
                  If NumVectors < 65536 Then
                     VEC(0) = NumVectors - 65536
                  Else
                     VEC(0) = -32768
                  End If
               End If
               NumVisVectors = NumVectors
               
                Screen.MousePointer = vbHourglass
                picCanvas.Picture = LoadPicture
                DrawVectors NumVectors
                Screen.MousePointer = vbDefault
               Exit For
            End If
         Next v
      End If
   
   Case 4   ' Draw vectors (re-estab all vectors)
      If VECSIZE > 4 Then
         NumVisVectors = NumVectors
          Screen.MousePointer = vbHourglass
          picCanvas.Picture = LoadPicture
          DrawVectors NumVectors
          Screen.MousePointer = vbDefault
      End If
   
   Case 5   ' Del all vectors
      NumVectors = 0
      NumVisVectors = 0
      ' Set up VEC array
      VECSIZE = 4
      ReDim VEC(0 To VECSIZE)
      VEC(0) = 0        ' NumVectors
      VEC(1) = picCanvas.Width
      VEC(2) = picCanvas.Height
      VEC(3) = -2       ' END
      VEC(4) = -2       ' END
      picCanvas.Picture = LoadPicture
      If aFast Then ReDim picMem(1 To picCanvas.Width, 1 To picCanvas.Height)
      aDupVECArray = False
      ReDim VECORG(0)
      
   Case 6   ' Clip vectors down to NumVisVectors
      NV = 0
      If VECSIZE > 4 Then
         For v = 1 To VECSIZE - 1
            If VEC(v) = -1 Then ' VEC(v + 1) = vecnum
               NV = NV + 1
               If NV = NumVisVectors + 1 Then
                  Exit For ' @ end of NumVisVectors
               End If
            End If
         Next v
         If v < VECSIZE Then  ' which it should be
            VECSIZE = v + 1
            VEC(v) = -2
            VEC(v + 1) = -2
            ReDim Preserve VEC(0 To VECSIZE)
            NumVectors = NumVisVectors
            VEC(0) = NumVectors
             Screen.MousePointer = vbHourglass
             picCanvas.Picture = LoadPicture
             DrawVectors NumVectors
             Screen.MousePointer = vbDefault
         End If
      End If
   
   Case 7   ' Fix canvas
      If NumVectors = 0 Then
         MsgBox " No vectors yet", vbExclamation, "Resizing canvas"
         Exit Sub
      ElseIf NumVisVectors <> 0 Then
         MsgBox " Clear picture first, resize" & vbCr & "and then Fix canvas", vbExclamation, "Resizing canvas"
         Exit Sub
      End If
      ' Redraw
      VEC(1) = picCanvas.Width
      VEC(2) = picCanvas.Height
      If VECSIZE > 4 Then
         NumVisVectors = NumVectors
          Screen.MousePointer = vbHourglass
          picCanvas.Picture = LoadPicture
          DrawVectors NumVectors
          Screen.MousePointer = vbDefault
      End If
      FixScrollbars picFrame, picCanvas, HS, VS
   
   Case 8   ' Stretch vectors
      If NumVectors = 0 Then
         MsgBox " No vectors yet", vbExclamation, "Stretch vectors"
         Exit Sub
      ElseIf NumVisVectors <> 0 Then
         MsgBox " Clear picture first, resize" & vbCr & "and then Stretch vectors", vbExclamation, "Stretch vectors"
         Exit Sub
      End If

      If Not aDupVECArray Then  ' aDupVECArray = False @ Start, New or Del all vectors
         ReDim VECORG(0 To VECSIZE)
         CopyMemory VECORG(0), VEC(0), 2 * VECSIZE + 2
         aDupVECArray = True
      Else
         If UBound(VEC) > UBound(VECORG) Then  ' vectors have been added to VEC()
                                               ' so copy everything to VECORG()
            ReDim Preserve VECORG(0 To VECSIZE)  ' Increase to match VEC()
            CopyMemory VECORG(0), VEC(0), 2 * VECSIZE + 2
         End If
      End If

      StretchPIC
      FixScrollbars picFrame, picCanvas, HS, VS
   
   End Select

   LabNumVisStrokes = "No. Vectors =" & Str$(NumVectors) & vbCr & "No. Visible =" & Str$(NumVisVectors)

   
   If NumVisVectors > 0 Then
      picResizer.Visible = False
   Else
      picResizer.Left = picCanvas.Left + picCanvas.Width - picResizer.Width - 1
      picResizer.Top = picCanvas.Top + picCanvas.Height - picResizer.Height
      picResizer.Visible = True
   End If

   LCCount = 0
End Sub

Private Sub StretchPIC()
' Stretch vectors from origin at top-left of canvas
Dim zRW As Single
Dim zRH As Single
Dim BS As Integer
Dim v As Long

' Testing
'Dim s1 As Long
'Dim s2 As Long

   aStretch = True
   
   zRW = picCanvas.Width / VECORG(1)
   zRH = picCanvas.Height / VECORG(2)
   
   svBrushSizeAtStretch = BrushSize
   
   VEC(1) = picCanvas.Width
   VEC(2) = picCanvas.Height
   
   v = 1
   Do
      If VEC(v) = -1 Then
      
         ' Adjust Brushsize
         If zRW > zRH Then
            BS = VECORG(v + 2) * zRW
         Else
            BS = VECORG(v + 2) * zRH
         End If
         If BS < 1 Then
            BS = 1
         End If
         VEC(v + 2) = BS   ' Brushsize
         
         v = v + 8
         Do
            VEC(v) = zRW * VECORG(v)            ' newX
            VEC(v + 1) = zRH * VECORG(v + 1)    ' newY
            v = v + 2
            If VEC(v) = -1 Then Exit Do
            If VEC(v) = -2 Then Exit Do
         Loop
      Else
         v = v + 1
      End If
      If VEC(v) = -2 Then Exit Do
   Loop
   's1 = UBound(VEC)
   NumVisVectors = NumVectors
   Screen.MousePointer = vbHourglass
   picCanvas.Picture = LoadPicture
   ReDim picMem(1 To picCanvas.Width, 1 To picCanvas.Height)
   DrawVectors NumVectors
   Screen.MousePointer = vbDefault
   
   's2 = UBound(VEC)
   aStretch = False
   BrushSize = svBrushSizeAtStretch
   LCCount = 0
End Sub

Private Sub DrawVectors(ByVal NStrokes As Long)
Dim v As Long
Dim NS As Long

   aRUNVEC = True
   
   ' Save current values
   svzalpha = zalpha
   svBrushSize = BrushSize
   svBrushAngle = BrushAngle
   svPaintColor = PaintColor
   
   If VEC(0) >= 0 Then
      NumVectors = CInt(VEC(0))
   Else
      NumVectors = 65536 + VEC(0)
   End If
   
   picCanvas.Width = VEC(1)
   picCanvas.Height = VEC(2)
   picCanvas.Refresh
   
   If aFast Then
      ReDim picMem(1 To picCanvas.Width, 1 To picCanvas.Height)
      Cul = RGB(255, 255, 255)
      For j = 1 To picCanvas.Height
      For i = 1 To picCanvas.Width
         picMem(i, j) = Cul
      Next i
      Next j
   End If
   
   ReDim bBack(picCanvas.Width, picCanvas.Height)
   
   LabWH = "Pic W,H =" & Str$(picCanvas.Width) & "," & Str$(picCanvas.Height)
   picResizer.Left = picCanvas.Left + picCanvas.Width - picResizer.Width - 1
   picResizer.Top = picCanvas.Top + picCanvas.Height - picResizer.Height
   picResizer.Visible = False
   
   DoEvents
   
   NS = 0
   v = 3
   
   Do
      If VEC(v) = -2 Then Exit Do   ' No strokes
      
      If VEC(v) = -1 Then  ' Start of stroke
         NS = NS + 1
         'If NStrokes < 32766 Then
            If NS > NStrokes Then Exit Do
         'End If
         v = v + 2: BrushSize = VEC(v)
         v = v + 1: BrushAngle = VEC(v)
         v = v + 1: bPaintRed = CByte(VEC(v))
         v = v + 1: bPaintGreen = CByte(VEC(v))
         v = v + 1: bPaintBlue = CByte(VEC(v))
         PaintColor = RGB(bPaintRed, bPaintGreen, bPaintBlue)
         v = v + 1: zalpha = CSng(VEC(v) / 100)
         
         If zalpha > 1 Then
            zalpha = 1
            VEC(v) = 100
         End If
         v = v + 1
         
         If BrushAngle = 4 Then   ' Fill
            If aFast Then
               ShowWholePicture
               Fill picCanvas, CSng(VEC(v)), CSng(VEC(v + 1))
               GETDIB picCanvas.Image
            Else
               Fill picCanvas, CSng(VEC(v)), CSng(VEC(v + 1))
            End If
         Else
            START_BRUSH picCanvas, CSng(VEC(v)), CSng(VEC(v + 1))
         End If
         
         v = v + 2
         If VEC(v) = -2 Then Exit Do   ' Single click
         Do
            
            If VEC(v) = -1 Then Exit Do   ' Next stroke
            
            If Not aFast Then
               DO_SLOW_BRUSH picCanvas, CSng(VEC(v)), CSng(VEC(v + 1))
               picCanvas.Refresh ' Show progress  (a bit slower)
            Else
               DO_FAST_BRUSH CSng(VEC(v)), CSng(VEC(v + 1))
            End If
            
            v = v + 2
            If VEC(v) = -1 Then Exit Do   ' Next stroke
            If VEC(v) = -2 Then Exit Do   ' END
         Loop
         If VEC(v) = -2 Then Exit Do   ' END
      End If
   Loop
   
   If aFast Then
      ShowWholePicture
   End If
   
   ' Restore current values
   zalpha = svzalpha
   BrushSize = svBrushSize
   BrushAngle = svBrushAngle
   PaintColor = svPaintColor
   bPaintRed = PaintColor And &HFF&
   bPaintGreen = (PaintColor And &HFF00&) / &H100&
   bPaintBlue = (PaintColor And &HFF0000) / &H10000
   
   
   HSalpha.Value = 100 * zalpha
   If aStretch Then
      optBrushHead((svBrushSizeAtStretch - 1) \ 2).Value = True
   Else
      optBrushHead((BrushSize - 1) \ 2).Value = True
   End If
   
   optBrush(BrushAngle).Value = True
   
   LabNumVisStrokes = "No. Vectors =" & Str$(NumVectors) & vbCr & "No. Visible =" & Str$(NumVisVectors)

   aRUNVEC = False
End Sub
'#### END VECTORS ########################################################################


Private Sub ShowWholePicture()
Dim W As Long, H As Long

   FillBMPStruc
   
   W = picCanvas.Width
   H = picCanvas.Height
   
   If StretchDIBits(picCanvas.hdc, _
      0, 0, W, H, _
      0, 0, W, H, _
      picMem(1, 1), bm, _
      DIB_RGB_COLORS, vbSrcCopy) = 0 Then
         MsgBox "Blit Error", vbCritical, "Brush Vectors"
         Form_Unload 0
   End If
   picCanvas.Refresh
End Sub


'#### QUITTING STUFF #########################################

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Form As Form

   If aHelp Then
      aHelp = False
      Unload frmHelp
   End If
   
   If UnloadMode = 0 Then    'Close on Form1 pressed
         
      response = MsgBox("", vbQuestion + vbYesNo, "Quit Application ?")
      If response = vbNo Then
         Cancel = True
      Else  'response= Yes
         Cancel = False
         
         Screen.MousePointer = vbDefault
         
         ' Make sure all forms cleared
         For Each Form In Forms
            Unload Form
            Set Form = Nothing
         Next Form
         End
      
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Form As Form

   Screen.MousePointer = vbDefault

   ' Make sure all forms cleared
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
   End
End Sub

'   EXAMPLE:-
'   VECSIZE = 30
'   ReDim VEC(0 To VECSIZE)
'   VEC(0) = 2        ' NumVectors
'   VEC(1) = 256      ' picCanvas.Width
'   VEC(2) = 256      ' picCanvas.Height
'   VEC(3) = -1       ' New stroke
'   VEC(4) = 1        ' Vector num
'   VEC(5) = 13       ' BrushSize
'   VEC(6) = 1        ' BrushAngle
'   VEC(7) = R        ' R PaintColor
'   VEC(8) = G        ' G PaintColor
'   VEC(9) = B        ' B PaintColor
'   VEC(10) = 50      ' zalpha*100 (ie zalpha=0.5)
'   VEC(11) = 30      ' X1
'   VEC(12) = 30      ' Y1
'   VEC(13) = 100     ' X2
'   VEC(14) = 150     ' Y2
'   VEC(15) = 80      ' X3
'   VEC(16) = 180     ' Y3
'
'   VEC(17) = -1      ' New stroke
'   VEC(18) = 2       ' Vector num
'   VEC(19) = 13      ' BrushSize
'   VEC(20) = 2       ' BrushAngle
'   VEC(21) = R       ' R PaintColor
'   VEC(22) = G       ' G PaintColor
'   VEC(23) = B       ' B PaintColor
'   VEC(24) = 20      ' zalpha*100 (ie zalpha=0.2)
'   VEC(25) = 150     ' X1
'   VEC(26) = 30      ' Y1
'   VEC(27) = 10      ' X2
'   VEC(28) = 30      ' Y2
'   VEC(29) = 200     ' X3
'   VEC(30) = 180     ' Y3
'   VEC(31) = -2      ' END
'   VEC(32) = -2      ' END

