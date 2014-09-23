VERSION 5.00
Object = "{94A0E92D-43C0-494E-AC29-FD45948A5221}#1.0#0"; "WIAAut.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "ImageFile UnitTest"
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnThumb 
      Caption         =   "ShowThumbnail"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton btnCrop 
      Caption         =   "Crop"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   7560
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   3600
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   6975
      Left            =   1560
      ScaleHeight     =   461
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      Begin VB.Shape Selection 
         BorderStyle     =   3  'Dot
         Height          =   2655
         Left            =   720
         Top             =   960
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Carga Foto"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7560
      Width           =   1335
   End
   Begin WIACtl.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   7680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim img As ImageFile
Dim prcs As New ImageProcess
Dim thmb As ImageFile
Dim x1 As Integer
Dim y1 As Integer
Dim x2 As Integer
Dim y2 As Integer
Dim bSelecting As Boolean

Private Sub UpdatePicture()
    If Not img Is Nothing Then
        Dim w, h As Integer
       
        w = img.Width
        h = img.Height
        
        Set Picture1.Picture = img.ARGBData.Picture(w, h)
        
        While (prcs.Filters.Count > 0)
            prcs.Filters.Remove 1
        Wend
        
        prcs.Filters.Add prcs.FilterInfos("Scale").FilterID
        
        
        Set thmb = prcs.Apply(img)
        
        If Not thmb Is Nothing Then
            If thmb.FrameCount = img.FrameCount Then
                thmb.ActiveFrame = img.ActiveFrame
            End If
            w = thmb.Width
            h = thmb.Height
            Set Picture2.Picture = thmb.ARGBData.Picture(w, h)
        End If
    End If
    
    If img.FrameCount > 1 Then
        HScroll1.Min = 1
        HScroll1.Max = img.FrameCount
        HScroll1.Value = img.ActiveFrame
        HScroll1.Visible = True
    Else
        HScroll1.Visible = False
    End If
End Sub

Private Sub btnStamp_Click()
    If Not img Is Nothing Then
        Dim stamp As ImageFile
        
        While (prcs.Filters.Count > 0)
            prcs.Filters.Remove 1
        Wend
        
        prcs.Filters.Add prcs.FilterInfos("Scale").FilterID
        
        prcs.Filters(1).Properties(1).Value = img.Width * 3 / 4
        prcs.Filters(1).Properties(2).Value = img.Height * 3 / 4
        
        Set stamp = prcs.Apply(img)
        
        If stamp.FrameCount = img.FrameCount Then
            stamp.ActiveFrame = img.ActiveFrame
        End If
        
        If Not stamp Is Nothing Then
            Dim stamped As ImageFile
            
            While (prcs.Filters.Count > 0)
                prcs.Filters.Remove 1
            Wend
        
            prcs.Filters.Add prcs.FilterInfos("Stamp").FilterID
        
            prcs.Filters(1).Properties(1).Value = stamp
            prcs.Filters(1).Properties(2).Value = img.Width - stamp.Width
            prcs.Filters(1).Properties(3).Value = img.Height - stamp.Height
        
            Set stamped = prcs.Apply(img)
            
            If stamped.FrameCount = img.FrameCount Then
                stamped.ActiveFrame = img.ActiveFrame
            End If
            
            Set img = stamped
            UpdatePicture
        End If
    End If
End Sub

Private Sub btnThumb_Click()
    Dim v As Vector
    
    On Error Resume Next
    Set v = img.Properties("ThumbnailData").Value
    If Err.Number <> 0 Then Exit Sub
    
    Set img = v.ImageFile
    UpdatePicture
End Sub

Private Sub HScroll1_Change()
    img.ActiveFrame = HScroll1.Value
    UpdatePicture
End Sub

Private Sub btnLoad_Click()
    CommonDialog2.Filter = "JPEG File (*.jpg)|*.jpg|GIF File (*.gif)|*.gif|PNG File (*.png)|*.png|TIFF File (*.tif)|*.tif|BMP File (*.bmp)|*.bmp|All Files (*.*)|*.*"
    CommonDialog2.ShowOpen
    
    If CommonDialog2.FileName = "" Then Exit Sub
    Set img = New ImageFile
    img.LoadFile CommonDialog2.FileName
    
    UpdatePicture
End Sub

Private Sub btnFlipRotate_Click()
    If Not img Is Nothing Then
        Dim pic As ImageFile
        
        While (prcs.Filters.Count > 0)
            prcs.Filters.Remove 1
        Wend
        
        prcs.Filters.Add prcs.FilterInfos("RotateFlip").FilterID
        
        
        If chkHorizontal.Value = 1 Then prcs.Filters(1).Properties(2).Value = True
        If chkVertical.Value = 1 Then prcs.Filters(1).Properties(3).Value = True
        
        Set pic = prcs.Apply(img)
        If pic Is Nothing Then
            MsgBox "Failed to FlipRotate"
            Exit Sub
        End If
        
        If pic.FrameCount = img.FrameCount Then
            pic.ActiveFrame = img.ActiveFrame
        End If
        Set img = pic
        UpdatePicture
    End If
End Sub

Private Sub btnCrop_Click()
    If Not img Is Nothing Then
        If Selection.Visible = True Then
            Dim pic As ImageFile
            
            Selection.Visible = False
            
            While (prcs.Filters.Count > 0)
                prcs.Filters.Remove 1
            Wend
        
            prcs.Filters.Add prcs.FilterInfos("Crop").FilterID
        
            prcs.Filters(1).Properties(1).Value = x1
            prcs.Filters(1).Properties(2).Value = y1
            prcs.Filters(1).Properties(3).Value = img.Width - x2
            prcs.Filters(1).Properties(4).Value = img.Height - y2
        
            Set pic = prcs.Apply(img)
            If pic Is Nothing Then
                MsgBox "Failed to Crop"
                Exit Sub
            End If
        
            If pic.FrameCount = img.FrameCount Then
                pic.ActiveFrame = img.ActiveFrame
            End If
            Set img = pic
            UpdatePicture
            
        End If
    End If
End Sub


Private Sub btnSave_Click()
    Dim sType As String
    Dim pic As ImageFile
    
    CommonDialog2.Filter = "BMP File (*.bmp)|*.bmp|JPEG File (*.jpg)|*.jpg|GIF File (*.gif)|*.gif|PNG File (*.png)|*.png|TIFF File (*.tif)|*.tif"
    
    CommonDialog2.ShowSave
    If CommonDialog2.FileName = "" Then Exit Sub
    
    If CommonDialog2.FilterIndex = 1 Then
        sType = wiaFormatBMP
    ElseIf CommonDialog2.FilterIndex = 2 Then
        sType = wiaFormatJPEG
    ElseIf CommonDialog2.FilterIndex = 3 Then
        sType = wiaFormatGIF
    ElseIf CommonDialog2.FilterIndex = 4 Then
        sType = wiaFormatPNG
    Else
        sType = wiaFormatTIFF
    End If
    
    While (prcs.Filters.Count > 0)
        prcs.Filters.Remove 1
    Wend
        
    prcs.Filters.Add prcs.FilterInfos("Convert").FilterID
    prcs.Filters(1).Properties(1).Value = sType
    Set pic = prcs.Apply(img)
    If pic Is Nothing Then
        MsgBox "Failed to Convert"
        Exit Sub
    End If
    
    pic.SaveFile CommonDialog2.FileName
End Sub

Private Sub btnProperties_Click()
    Load Form2
    Set Form2.img = img
    Form2.Show
End Sub



Private Sub Form_Resize()
    Dim NewWidth As Integer
    Dim NewHeith As Integer
    
    NewWidth = Width - Picture1.Left - 200
    NewHeight = Height - Picture1.Top - 600
    
    If NewWidth < 1000 Then NewWidth = 1000
    If NewHeight < 1000 Then NewHeight = 1000
    Picture1.Width = NewWidth
    Picture1.Height = NewHeight
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not img Is Nothing Then
        x1 = X
        x2 = X
        y1 = Y
        y2 = Y
        
        Selection.Left = x1
        Selection.Top = y1
        Selection.Width = 10
        Selection.Height = 10
        Selection.Visible = True
        bSelecting = True
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim newX As Integer
    Dim newY As Integer
        
    newX = X
    newY = Y
    
    If bSelecting Then
        If newX > img.Width Then newX = img.Width
        If newY > img.Height Then newY = img.Height
        If newX > x1 + 10 Then Selection.Width = newX - x1
        If newY > y1 + 10 Then Selection.Height = newY - y1
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim newX As Integer
    Dim newY As Integer
    
    If bSelecting Then
        newX = X
        newY = Y
        If newX > img.Width Then newX = img.Width
        If newY > img.Height Then newY = img.Height
    
        x2 = newX
        y2 = newY
    
        If x2 > x1 + 10 And y2 > y1 + 10 Then
            Selection.Width = x2 - x1
            Selection.Height = y2 - y1
        Else
            Selection.Visible = False
        End If
        bSelecting = False
    End If
End Sub
