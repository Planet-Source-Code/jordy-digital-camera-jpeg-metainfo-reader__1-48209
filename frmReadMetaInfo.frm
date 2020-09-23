VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMetaInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXIF Metatag Parser"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btClose 
      Caption         =   "Close"
      Height          =   420
      Left            =   5310
      TabIndex        =   3
      Top             =   945
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CDL 
      Left            =   6435
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "JPG Files|*.jpg;*.jpeg;*.jpe;*.jfif|All files|*.*"
   End
   Begin VB.CommandButton btRead 
      Caption         =   "Open file..."
      Height          =   420
      Left            =   5310
      TabIndex        =   1
      Top             =   495
      Width           =   2175
   End
   Begin VB.TextBox txInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7305
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   495
      Width           =   5235
   End
   Begin VB.Image imPic 
      Height          =   2625
      Left            =   5310
      Top             =   1485
      Width           =   2175
   End
   Begin VB.Label lbFile 
      Caption         =   "<select an image file>"
      Height          =   465
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   5190
   End
End
Attribute VB_Name = "frmMetaInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EXF As New clsEXIF
Dim m_ImageWidth As Single
Dim m_ImageHeight As Single

Private Sub btClose_Click()
    Unload Me
End Sub

'--- browse for file and extract its metatag information ---
Private Sub btRead_Click()
    On Error GoTo 100
    CDL.CancelError = True
    CDL.ShowOpen
    On Error GoTo 200
    'clear controls
    txInfo.Text = ""
    Set imPic.Picture = LoadPicture("")
    'retrieve picture information
    EXF.ImageFile = CDL.FileName 'set the image file property
    txInfo.Text = EXF.ListInfo 'list all tags into the text box
    lbFile.Caption = CDL.FileName
    'show the picture on the form
    imPic.Stretch = False
    Set imPic.Picture = LoadPicture(CDL.FileName)
    SizeImage imPic
100
    Exit Sub
200
    MsgBox "Error parsing the JPEG header.", vbCritical
End Sub

Private Sub Form_Load()
    m_ImageWidth = imPic.Width
    m_ImageHeight = imPic.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EXF = Nothing
End Sub

'--- resizes the image to fit the image control ---
Private Sub SizeImage(ImageHolder As Image)
    Dim Pw As Single, Ph As Single, Aspect As Single
    With ImageHolder
        Pw = .Width
        Ph = .Height
        Aspect = Pw / Ph
        If (Pw > m_ImageWidth And Ph > m_ImageHeight) Then
            If Ph > m_ImageHeight Then
                Ph = m_ImageHeight
                Pw = Aspect * Ph
            End If
            If Pw > m_ImageWidth Then
                Pw = m_ImageWidth
                Ph = Pw / Aspect
            End If
            .Width = Pw  'Take the new values to the pic
            .Height = Ph
        End If
        If Aspect > 1 Then
            If .Width < m_ImageWidth Then
                .Width = m_ImageWidth
                .Height = m_ImageWidth / Aspect
            End If
        Else
            If .Height < m_ImageHeight Then
                .Height = m_ImageHeight
                .Width = m_ImageHeight * Aspect
            End If
        End If
        .Stretch = True
    End With
End Sub
