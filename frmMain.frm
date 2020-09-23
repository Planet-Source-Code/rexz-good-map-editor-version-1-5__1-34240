VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map Editor"
   ClientHeight    =   6015
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox EditMap 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Edit Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   21
      Top             =   360
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Frame SelFrame 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   15
      Left            =   220
      TabIndex        =   19
      Top             =   580
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   7935
      TabIndex        =   9
      Top             =   0
      Width           =   7960
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Event"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   6840
         TabIndex        =   22
         Top             =   0
         Width           =   690
         Visible         =   0   'False
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Event: None"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5880
         TabIndex        =   20
         Top             =   0
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Layer: 1/9"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5040
         TabIndex        =   16
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Walkable: No"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3960
         TabIndex        =   15
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FXType: 5"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1920
         TabIndex        =   14
         Top             =   0
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name: Noname"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2760
         TabIndex        =   13
         Top             =   0
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Tile: 0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   720
         TabIndex        =   12
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tile: 0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   345
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Tiles"
      Height          =   2535
      Left            =   6600
      TabIndex        =   1
      Top             =   600
      Width           =   1335
      Begin VB.TextBox txtLayer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         MaxLength       =   1
         TabIndex        =   18
         Text            =   "1"
         Top             =   2120
         Width           =   495
      End
      Begin VB.CheckBox chkWalk 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Walkable"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdFlood 
         Caption         =   "Flood"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.PictureBox Current 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   7
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox SelTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   720
         Picture         =   "frmMain.frx":0000
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox SelTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   240
         Picture         =   "frmMain.frx":06BA
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox SelTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   720
         Picture         =   "frmMain.frx":0D74
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox SelTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":142E
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Layer:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   135
         TabIndex        =   6
         Top             =   1080
         Width           =   480
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   6510
      TabIndex        =   0
      Top             =   240
      Width           =   6540
      Begin VB.Image Tile 
         Height          =   375
         Index           =   287
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   286
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   285
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   284
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   283
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   282
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   281
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   280
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   279
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   278
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   277
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   276
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   275
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   274
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   273
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   272
         Left            =   720
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   271
         Left            =   360
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   270
         Left            =   0
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   269
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   268
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   267
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   266
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   265
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   264
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   263
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   262
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   261
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   260
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   259
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   258
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   257
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   256
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   255
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   254
         Left            =   720
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   253
         Left            =   360
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   252
         Left            =   0
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   251
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   250
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   249
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   248
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   247
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   246
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   245
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   244
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   243
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   242
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   241
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   240
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   239
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   238
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   237
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   236
         Left            =   720
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   235
         Left            =   360
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   234
         Left            =   0
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   233
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   232
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   231
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   230
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   229
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   228
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   227
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   226
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   225
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   224
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   223
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   222
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   221
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   220
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   219
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   218
         Left            =   720
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   217
         Left            =   360
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   216
         Left            =   0
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   215
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   214
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   213
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   212
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   211
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   210
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   209
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   208
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   207
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   206
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   205
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   204
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   203
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   202
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   201
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   200
         Left            =   720
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   199
         Left            =   360
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   198
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   197
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   196
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   195
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   194
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   193
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   192
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   191
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   190
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   189
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   188
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   187
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   186
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   185
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   184
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   183
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   182
         Left            =   720
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   181
         Left            =   360
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   180
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   179
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   178
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   177
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   176
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   175
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   174
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   173
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   172
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   171
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   170
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   169
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   168
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   167
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   166
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   165
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   164
         Left            =   720
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   163
         Left            =   360
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   162
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   161
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   160
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   159
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   158
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   157
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   156
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   155
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   154
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   153
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   152
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   151
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   150
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   149
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   148
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   147
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   146
         Left            =   720
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   145
         Left            =   360
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   144
         Left            =   0
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   143
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   142
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   141
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   140
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   139
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   138
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   137
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   136
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   135
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   134
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   133
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   132
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   131
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   130
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   129
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   128
         Left            =   720
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   127
         Left            =   360
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   126
         Left            =   0
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   125
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   124
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   123
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   122
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   121
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   120
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   119
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   118
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   117
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   116
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   115
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   114
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   113
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   112
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   111
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   110
         Left            =   720
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   109
         Left            =   360
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   108
         Left            =   0
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   107
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   106
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   105
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   104
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   103
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   102
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   101
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   100
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   99
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   98
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   97
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   96
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   95
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   94
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   93
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   92
         Left            =   720
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   91
         Left            =   360
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   90
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   89
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   88
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   87
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   86
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   85
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   84
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   83
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   82
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   81
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   80
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   79
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   78
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   77
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   76
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   75
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   74
         Left            =   720
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   73
         Left            =   360
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   72
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   71
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   70
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   69
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   68
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   67
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   66
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   65
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   64
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   63
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   62
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   61
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   60
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   59
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   58
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   57
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   56
         Left            =   720
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   55
         Left            =   360
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   54
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   53
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   52
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   51
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   50
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   49
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   48
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   47
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   46
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   45
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   44
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   43
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   42
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   41
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   40
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   39
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   38
         Left            =   720
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   37
         Left            =   360
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   36
         Left            =   0
         Stretch         =   -1  'True
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   35
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   34
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   33
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   32
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   31
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   30
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   29
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   28
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   27
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   26
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   25
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   24
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   23
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   22
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   21
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   20
         Left            =   720
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   19
         Left            =   360
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   18
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   17
         Left            =   6120
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   16
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   15
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   14
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   13
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   12
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   11
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   10
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   9
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   8
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   7
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   6
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   5
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   4
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   3
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   2
         Left            =   720
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   1
         Left            =   360
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Tile 
         Height          =   375
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog cm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuNew 
         Caption         =   "New Map"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Map"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Map"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuRndPattern 
         Caption         =   "Random Pattern"
      End
      Begin VB.Menu mnuAdvProp 
         Caption         =   "Advanced Properties"
      End
      Begin VB.Menu mnuSetName 
         Caption         =   "Set Name"
      End
      Begin VB.Menu mnuEvent 
         Caption         =   "Add Event"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHlp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cur_Tile As Integer, SelectedTile As Integer

Private Sub chkWalk_Click()
ThisMap.Tiles(SelectedTile).Walkable = chkWalk.Value
End Sub

Private Sub cmdFlood_Click()
Flood
End Sub

Private Sub Current_Click()
Cur_Tile = 5
Current.Picture = LoadPicture()
End Sub

Private Sub EditMap_Click()
If EditMap.Value = 1 Then
Frame.Enabled = True
Else
Frame.Enabled = False
End If
End Sub

Private Sub Form_Load()
SelectedTile = 0
Current.Picture = SelTile(Cur_Tile).Picture
Dim strString As String
    Dim lngDword As Long
    Dim Record As String
    
        If Command$ <> "%1" And Command$ <> "" Then
        'Command$ is the file you need To open!
        'Load the file
        LoadMap Command$
        Else
        NewMap
        End If
End Sub

Private Sub Label10_Click()
Select Case Left(ThisMap.Tiles(SelectedTile).Event, 3)
Case "MSG"
MsgBox "Tile Index: " & SelectedTile & vbCrLf & "Event: Message" & vbCrLf & "Message: " & Mid(ThisMap.Tiles(SelectedTile).Event, InStr(ThisMap.Tiles(SelectedTile).Event, "=") + 1), vbInformation + vbOKOnly, "Event Info"
Case "WAR"
MsgBox "Tile Index: " & SelectedTile & vbCrLf & "Event: WARP" & vbCrLf & "Script: " & Mid(ThisMap.Tiles(SelectedTile).Event, InStr(ThisMap.Tiles(SelectedTile).Event, "=") + 1), vbInformation + vbOKOnly, "Event Info"
Case "DAM"
MsgBox "Tile Index: " & SelectedTile & vbCrLf & "Event: Damage Point" & vbCrLf & "Amount: " & Mid(ThisMap.Tiles(SelectedTile).Event, InStr(ThisMap.Tiles(SelectedTile).Event, "=") + 1), vbInformation + vbOKOnly, "Event Info"
End Select
End Sub

Private Sub Label4_Click()
Dim sname As String
sname = InputBox("Enter a name for the map", "Name the map")
If sname <> "" Then
ThisMap.sname = sname
Label4.Caption = "Name: " & sname
Else
Exit Sub
End If
End Sub

Private Sub mnuAbout_Click()
MsgBox "Tile Map Editor by Hans Bjerndell" & vbCrLf & "For Tilebased games" & vbCrLf & "Copyright © 2002 Hans Bjerndell", vbInformation + vbOKOnly, "About"
End Sub

Private Sub mnuAdvProp_Click()
frmAdvanced.Show
End Sub

Private Sub mnuEvent_Click()
frmEvent.iIndex = SelectedTile
frmEvent.lblTile.Caption = "Tile: " & SelectedTile
frmEvent.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuHlp_Click()
frmHelp.Show
End Sub

Private Sub mnuNew_Click()
NewMap
End Sub

Private Sub mnuOpen_Click()
cm1.Filter = "Supported types |*.cms|"
cm1.ShowOpen
If cm1.Filename <> "" Then
LoadMap cm1.Filename
Else
GoTo error:
End If
error:
End Sub

Private Sub mnuRndPattern_Click()
frmRandom.Show
End Sub

Private Sub mnuSave_Click()
On Error GoTo error:
cm1.Filter = "Supported types |*.cms|"
cm1.ShowSave
If cm1.Filename <> "" Then
SaveMap cm1.Filename
Else
GoTo error:
End If
error:
End Sub

Private Sub mnuSetName_Click()
Label4_Click
End Sub

Private Sub SelFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SelFrame.Visible = False
End Sub

Private Sub SelTile_Click(Index As Integer)
Cur_Tile = Index
Current.Picture = SelTile(Index).Picture
End Sub

Function Flood()
Dim i
For i = 0 To 287
If Cur_Tile = 5 Then
Tile(i).Picture = LoadPicture()
Else
Tile(i).Picture = SelTile(Cur_Tile).Picture
End If
ThisMap.Tiles(i).FXType = Cur_Tile
ThisMap.Tiles(i).Walkable = chkWalk.Value
ThisMap.Tiles(i).Layer = CInt(txtLayer.Text)
Next i
End Function

Private Sub Tile_DblClick(Index As Integer)
If EditMap.Value = 0 Then
frmEvent.iIndex = Index
frmEvent.lblTile.Caption = "Tile: " & Index
frmEvent.Show
Else
End If
End Sub

Private Sub Tile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If EditMap.Value = 1 Then
If Cur_Tile = 5 Then
Tile(Index).Picture = LoadPicture()
ThisMap.Tiles(Index).FXType = Cur_Tile
ThisMap.Tiles(Index).Layer = CInt(txtLayer.Text)
If chkWalk.Value = 1 Then
ThisMap.Tiles(Index).Walkable = 1
Else
ThisMap.Tiles(Index).Walkable = 0
End If
Else
Tile(Index).Picture = SelTile(Cur_Tile)
ThisMap.Tiles(Index).FXType = Cur_Tile
ThisMap.Tiles(Index).Layer = CInt(txtLayer.Text)
If chkWalk.Value = 1 Then
ThisMap.Tiles(Index).Walkable = 1
Else
ThisMap.Tiles(Index).Walkable = 0
End If
End If
Else
If ThisMap.Tiles(Index).Event <> "" Then
Label10.Visible = True
Else
Label10.Visible = False
End If
End If
Else
If EditMap.Value = 1 Then
Tile(Index).Picture = LoadPicture()
ThisMap.Tiles(Index).Layer = 1
ThisMap.Tiles(Index).FXType = 5
Else
End If
End If

Label3.Caption = "Selected Tile: " & Index
SelectedTile = Index
If ThisMap.Tiles(Index).Walkable = 1 Then
chkWalk.Value = 1
Else
chkWalk.Value = 0
End If
End Sub

Private Sub Tile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "Tile: " & Index
Label5.Caption = "FXType: " & ThisMap.Tiles(Index).FXType
Label7.Caption = "Layer: " & ThisMap.Tiles(Index).Layer & "/9"
If ThisMap.Tiles(Index).Event <> "" Then
Select Case Left(ThisMap.Tiles(Index).Event, 3)
Case "MSG"
Label9.Caption = "Event: MSG"
Case "WAR"
Label9.Caption = "Event: WARP"
Case "DAM"
Label9.Caption = "Event: Damage Point"
End Select
Else
Label9.Caption = "Event: None"
End If
SelFrame.Top = Tile(Index).Top + 580
SelFrame.Left = Tile(Index).Left + 220
SelFrame.Visible = True
If ThisMap.Tiles(Index).Walkable = 1 Then
Label6.Caption = "Walkable: Yes"
Else
Label6.Caption = "Walkable: No"
End If
End Sub

Function SaveMap(Filename As String)
Dim i
Open Filename For Output As #1
Print #1, ThisMap.sname
For i = 0 To 287
If ThisMap.Tiles(i).Event <> "" Then
Print #1, ThisMap.Tiles(i).FXType & ":" & ThisMap.Tiles(i).Walkable & ":" & ThisMap.Tiles(i).Layer & "," & ThisMap.Tiles(i).Event
Else
Print #1, ThisMap.Tiles(i).FXType & ":" & ThisMap.Tiles(i).Walkable & ":" & ThisMap.Tiles(i).Layer
End If
Next i
Close #1
End Function

Function LoadMap(Filename As String)
Dim i, temp As String, arr() As String, arr2() As String
Open Filename For Input As #1
Input #1, ThisMap.sname
For i = 0 To 287
Line Input #1, temp
arr = Split(temp, ":")
If UBound(arr()) < 2 Then
MsgBox "This seemes to be an old version, or an unsupported filetype. Unable to open file.", vbCritical + vbOKOnly, "Error"
Close #1
Exit Function
Else
End If
ThisMap.Tiles(i).FXType = CInt(arr(0))
ThisMap.Tiles(i).Walkable = arr(1)
If FindPart(arr(2), ",") = 1 Then
ThisMap.Tiles(i).Layer = Mid(arr(2), 1, 1)
ThisMap.Tiles(i).Event = Mid(arr(2), InStr(arr(2), ",") + 1)
Else
ThisMap.Tiles(i).Layer = arr(2)
End If
If arr(0) = "5" Then
Else
Tile(i).Picture = SelTile(arr(0))
End If
Next i
Close #1
Me.Caption = "Map Editor - " & ThisMap.sname
Label4.Caption = "Name: " & ThisMap.sname
Exit Function
End Function

Private Sub Tile_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Tile(Index).Picture = 0 Then
Label5.Caption = "FXType: 5"
Else
Label5.Caption = "FXType: " & ThisMap.Tiles(Index).FXType
End If
End Sub

Function NewMap()
Dim i
ThisMap.sname = "Noname"
Label4.Caption = "Name: Noname"
For i = 0 To 287
Tile(i).Picture = LoadPicture()
ThisMap.Tiles(i).FXType = 5
ThisMap.Tiles(i).Layer = 1
ThisMap.Tiles(i).Walkable = 0
Next i
End Function
Private Sub txtLayer_KeyPress(KeyAscii As Integer)
KeyAscii = IIf(Not KeyAscii = 8 And Not Val((Chr(KeyAscii))) > 0, 0, KeyAscii)
'If IsNumeric(txtLayer.Text) = True Then
'txtLayer.Text = "0"
'Else
'If txtLayer.Text > bla Then
'txtLayer.Text = "5"
'ThisMap.Tiles(SelectedTile).Layer = CInt(txtLayer.Text)
'Else
'ThisMap.Tiles(SelectedTile).Layer = CInt(txtLayer.Text)
'End If
'End If
End Sub
