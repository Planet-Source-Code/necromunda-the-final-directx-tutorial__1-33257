Attribute VB_Name = "mObjects"
Option Explicit

'Before creating these objects below you will need:
'   1. The DirectX8 runtime files (www.microsoft.com/directx)
'   2. The DirectX8 Type Libraries for VB
'   3. To have loaded the Type Libraries
'   Project > References > DirectX 8 for VB type libraries


'Main Control Object for DirectX 8
'Required for setting up all other main DirectX objects
Public DX8 As DirectX8

'D3D Object
'Contains all that is necessary to make 3d thingys
Public D3D As Direct3D8

'D3DX object
'Contains helper functions for doing things like loading textures.
Public D3DX As New D3DX8

'D3D Device
'Represents the hardware that renders the scene
Public D3DDevice As Direct3DDevice8
'Vertex Buffer
'An area of memory set aside for storing a list of verticies
Public vertexBuffer As Direct3DVertexBuffer8

'Enumeration for whichever state we are running in
Enum runningMode
    notRunning
    Normal
    Lighting
    Texturing
    Modelling
End Enum

Public Const PI = 3.14159265358979

'RunningMOde value
Public currentApp As runningMode

Sub addReport(text As String)

    frmMain.lstReport.AddItem text

End Sub

Sub unloadObjects()
    'Stop the render loop if its running
    'then unload all the objects in reverse order for some reason.
    'I think it's to do with how windows handles memory.
        
    currentApp = notRunning
    
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set DX8 = Nothing
    Set D3DX = Nothing
    
    Unload frmFullDisp
    
    End

End Sub

