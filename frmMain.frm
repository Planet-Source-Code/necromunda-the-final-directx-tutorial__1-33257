VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectX"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmdiFile 
      Left            =   5400
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Picture Files |*.bmp .jpeg|"
   End
   Begin VB.ListBox lstReport 
      Height          =   1230
      Left            =   6240
      TabIndex        =   7
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Frame fmeExtra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Extra Options"
      Height          =   2895
      Left            =   6120
      TabIndex        =   4
      Top             =   2520
      Width           =   3015
      Begin VB.CommandButton cmdMoveCam 
         Caption         =   "Move Camera"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdModel 
         Caption         =   "Change Model"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdTexture 
         Caption         =   "Change Texture"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox chkDrag 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Drag Mode"
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox chkZ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Z"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   2040
         Width           =   495
      End
      Begin VB.CheckBox chkY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Y"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   2040
         Width           =   495
      End
      Begin VB.CheckBox chkX 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&X"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   2040
         Value           =   1  'Checked
         Width           =   495
      End
      Begin VB.CheckBox chkMove 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Translate"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox chkScale 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Scale"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   735
      End
      Begin VB.CheckBox chkRotate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Rotate in:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   3000
         Y1              =   1920
         Y2              =   1920
      End
   End
   Begin VB.Frame fmeSetup 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Setup Options"
      Height          =   1215
      Left            =   6120
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
      Begin VB.ComboBox cmbMode 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   1320
         List            =   "frmMain.frx":0010
         TabIndex        =   15
         Top             =   660
         Width           =   1335
      End
      Begin VB.OptionButton optWindow 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Window"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optFull 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fullscreen"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblMode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Mode:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Unload Objects and Exit"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "St&art Scene"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox picDirectX 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   120
      Picture         =   "frmMain.frx":003C
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   389
      TabIndex        =   2
      ToolTipText     =   "Click and hold to drag!"
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'Function to retrieve the time in millisecondds the
'computer has been on, so that we can calculate the
'fps of the scene
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim FPS_LastCheck As Long
Dim FPS_Count As Long
Dim FPS_Current As Integer


Private Sub Form_Load()
        
    'Initialise the random
    'number generator
    Call Randomize
    
    'Set the list to normal mode
    cmbMode.ListIndex = 0
    
    'Set the model and textures to their defaults
    currentModel = "default"
    currentTexture = "default"
    


    
End Sub
Private Sub Form_Unload(Cancel As Integer)
            
    'Unload all the objects that have been loaded
    Call unloadObjects

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
        'Check for keypresses
    
        Select Case KeyCode
            Case Is = vbKeyEscape
                Call unloadObjects
            Case Is = vbKeyR        'Rotate
                chkRotate.Value = IIf(chkRotate.Value = 1, 0, 1)
            Case Is = vbKeyS        'Scale
                chkScale.Value = IIf(chkScale.Value = 1, 0, 1)
            Case Is = vbKeyT        'Transform
                chkMove.Value = IIf(chkMove.Value = 1, 0, 1)
            Case Is = vbKeyX        'X-Axis Rotate
                chkX.Value = IIf(chkX.Value = 1, 0, 1)
            Case Is = vbKeyY        'Y-Axis Rotate
                chkY.Value = IIf(chkY.Value = 1, 0, 1)
            Case Is = vbKeyZ        'Z-Axis Rotate
                chkZ.Value = IIf(chkZ.Value = 1, 0, 1)
        End Select

End Sub


Private Sub cmdModel_Click()

    'Set the commondialog to display only .X  model files
    cmdiFile.FileName = ""
    cmdiFile.Filter = "Microsoft DirectX .X Files (*.X) |*.x"
    cmdiFile.ShowOpen
    
    'Check if a model has been selected and load the model if it has
    If cmdiFile.FileName <> "" Then
        Call loadModel(cmdiFile.FileName) '[mModels]
        D3DDevice.SetRenderState D3DRS_SPECULARENABLE, 0
    
    End If
    
End Sub
Private Sub cmdTexture_Click()
    
    'Set filter to display only bitmaps and jpegs
    cmdiFile.FileName = ""
    cmdiFile.Filter = "Picture Files (*.bmp, *.jpeg, *.jpg) | *.bmp; *.jpeg; *.jpg"
    cmdiFile.ShowOpen
    
    'Check for file and load
    If cmdiFile.FileName <> "" Then Call setTexture(cmdiFile.FileName) '[mTextures]

End Sub
Private Sub cmdMoveCam_Click()
    Dim moveTo As D3DVECTOR
    'Vector storing the user inputs for the camera
    
    On Error GoTo quitsub
    
    'Get the user input for the camera location and
    'reset the camera
    moveTo.X = InputBox("Set X location. Note: Viewing range is from 0.1 to 100", "Move Camera")
    moveTo.Y = InputBox("Set Y location. Note: Viewing range is from 0.1 to 100", "Move Camera")
    moveTo.Z = InputBox("Set Z location. Note: Viewing range is from 0.1 to 100", "Move Camera")
    
    
    'Keep the camera in the viewing range of the
    'projection matrix (0.1 to 100 metres)
    With moveTo
        If .X > 90 Then .X = 90
        If .Y > 90 Then .Y = 90
        If .Z > 90 Then .Z = 90
    End With
    
    Call moveCamera(moveTo) '[mGeometry]
    
quitsub:
End Sub
Private Sub cmdStart_Click()
    'Every time the start button is pressed,
    'the scene is reloaded again.
            
    'Everything is cleared and set to defaults
    lstReport.Clear
    cmdTexture.Enabled = False
    cmdModel.Enabled = False
    currentApp = notRunning
       
    'This function calls the directX scene initialisation
    'coding, passing through a boolean value for windowed
    'or fullscreen, and the handle of where you want to
    'display to. If there is an error, the function returns
    'false, and a message s displayed to the user.
    
    'initD3D see[mInit]
    If InitD3D(optWindow.Value, IIf(optWindow.Value, picDirectX.hWnd, frmFullDisp.hWnd)) = True Then
                    
        'Clear away the logo - preventing possible
        'glitches as the picture may show through
        Set picDirectX = Nothing
        addReport ("Window DX Scene Loaded!")
        addReport ("--------")
              
                            
        cmdMoveCam.Enabled = True
                        
        'This select case structure determines which
        'display mode we shall be using, out of:
        'Normal: simple vertex types, no lighting/textures/models
        'Lighting: lit vertex types, no textures/models
        'Texturing: textured vertex types: no lighting/models
        'Modelling: Meshes
        'Once the display mode has been found the
        'appropriate function calls are made
        'to setup the scene
                                
        Select Case cmbMode.text
            Case Is = "Normal"
                addReport ("Normal mode!")
                addReport ("Loading vertices...")
   
                currentApp = Normal
                Call initGeometry 'see[mGeometry]
                
            'Summary:
            'Initialise the normal scene by creating
            'an octahedron out of vectors and loading
            'it into vertexbuffer.
                        
            Case Is = "Lighting"
                addReport ("Lighting mode!")
                addReport ("Loading lighting vertices...")
                addReport ("Setting up lights...")
                
                currentApp = Lighting
                Call initLightingGeo 'see[mLights]
                Call SetupLights 'see[mLights]
            
            'Summary:
            'Create the same octahedron shape as above,
            'but also adds vector normals for each triangle
            'then creates a randomly coloured light
            
            Case Is = "Texturing"
                addReport ("Texturing mode!")
                addReport ("Loading texture vertices...")
                addReport ("Loading texture...")
                
                currentApp = Texturing
                cmdTexture.Enabled = True
                Call initTexCube 'see[mTextures]
                Call setTexture(currentTexture)  'see[mTextures]
            
            'Summary:
            'Create a cube with texture coordinates included
            'and load up the default texture
                        
            Case Is = "Modelling"
                addReport ("Modelling mode!")
                addReport ("Loading model...")
                addReport ("Setting up lights...")
            
                currentApp = Modelling
                cmdModel.Enabled = True
                Call loadModel(currentModel)  '[mModels]
                Call SetupModelLights '[mModels]
                               
                
            'Summary:
            'Load the default model into a mesh and
            'setup the lights for modelling
                        
            Case Else
                currentApp = Normal
                Call initGeometry '[mInit]
            'In case anything else happens set it
            'to normal mode.
            
        End Select
        
        addReport ("Initialising matrices...")
        Call initMatrix 'see[mGeometry]
        
        addReport ("Successfully loaded!")
        Do While Not currentApp = notRunning
            Call Render '[mRender]
        
            If GetTickCount() - FPS_LastCheck >= 100 Then
                FPS_Current = FPS_Count * 10
                FPS_Count = 0
                FPS_LastCheck = GetTickCount()
            End If
        
            FPS_Count = FPS_Count + 1
            
            If optFull.Value = False Then Me.Caption = FPS_Current & "fps"
            
                 
            'Let Windows take a breath
            DoEvents
        
        Loop
    Else
        MsgBox "Initialisation failed!", vbCritical + vbOKOnly
    End If
    
    
End Sub
Private Sub cmdExit_Click()

    Call unloadObjects '[mObjects]

End Sub


Private Sub chkDrag_Click()
    'This is just some cosmetic code to disable the other checkboxes
    'when you put drag mode on.
        
    If chkDrag.Value = 1 Then
        
        'Turn off timer and enable picbox
        'Disable other checkboxes
        picDirectX.Enabled = True
    
        chkRotate.Enabled = False
        chkScale.Enabled = False
        chkMove.Enabled = False
        
        chkX.Enabled = False
        chkY.Enabled = False
        chkZ.Enabled = False
    Else
        'Opposite of above
        picDirectX.Enabled = False
    
        chkRotate.Enabled = True
        chkScale.Enabled = True
        chkMove.Enabled = True
        
        chkX.Enabled = True
        chkY.Enabled = True
        chkZ.Enabled = True
    End If
        
End Sub
Private Sub chkRotate_Click()
    
    'Disable the X,Y and Z checkboxes when the rotation checkbox is disabled
    If chkRotate.Value = 1 Then
        chkX.Visible = True: chkY.Visible = True: chkZ.Visible = True
    Else
        chkX.Visible = False: chkY.Visible = False: chkZ.Visible = False
    End If

End Sub

Private Sub optFull_Click()
    
    'Prevent user from running dragmode at the same
    'time as running fullscreen
    chkDrag.Enabled = False

    If chkDrag.Value = 1 Then
        chkDrag.Value = 0
        Call chkDrag_Click
    End If

End Sub
Private Sub optWindow_Click()
    
    'Enabled dragmode for windowed mode
    chkDrag.Enabled = True

End Sub

'Code to stop the user editting the values in the runningmode combobox
Private Sub cmbMode_KeyDown(KeyCode As Integer, Shift As Integer)
        
    KeyCode = 0

End Sub
Private Sub cmbMode_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub
Private Sub cmbMode_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub

Private Sub picDirectX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim matPic As D3DMATRIX
    'DRAG MODE coding
    
    'This code allows you to drag the 3D shapes/models around.
    'It takes the x and y of the mouse location and scales it down
    ' eg.(X / picDirectX.ScaleWidth) * 2.5
    'This is then loaded into a rotation matrix and applied as a
    'world transform.
        
    If Button <> 0 And currentApp <> notRunning Then
        'Rotate around the X and Z axis
        D3DXMatrixRotationYawPitchRoll matPic, _
            (X / picDirectX.ScaleWidth) * 2.5, _
            0, _
            (Y / picDirectX.ScaleHeight) * 2.5
        D3DDevice.SetTransform D3DTS_WORLD, matPic
    End If

End Sub



