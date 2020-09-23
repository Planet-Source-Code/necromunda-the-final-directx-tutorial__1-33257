Attribute VB_Name = "mInit"
Option Explicit

'-----------
'#InitD3D
'# Creates a fullscreen or windowed 3d scene
'-----------

Function InitD3D(windowed As Boolean, hWnd As Long) As Boolean

On Error GoTo initFail

'Stores the display mode data
Dim DispMode As D3DDISPLAYMODE
'Stores the way in which directx is rendering a scene
Dim D3DWindow As D3DPRESENT_PARAMETERS


'Setup the DirectX 8 Objects and check that they have been loaded
Set DX8 = New DirectX8
If DX8 Is Nothing Then addReport ("DX8 Object Error!"): GoTo initFail
addReport ("DX8 Object loaded...")

Set D3D = DX8.Direct3DCreate
If D3D Is Nothing Then addReport ("D3D Object Error!"): GoTo initFail
addReport ("D3D Object loaded...")

'Why all these objects? you may ask.  Well, firstly, for any DirectX 8
'application you are gonna need the DirectX8 object.  It's the main
'thing that lets you create all the other things - D3D, DXInput etc.
'As we are using Direct3D, we need a Direct3D8 ojbect, which is
'setup using the DX8 object. See? Now that the DX8 object has created the
'D3D8 object, we can use it to create the Direct3DDevice8 object.

'This line gets the current display mode settings, making it easier
'to setup the renderer
D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode


If windowed = True Then
    addReport ("WINDOWED mode!")
       

    With D3DWindow
        'Make this a windowed scene, (0 for fullscreen - default)
        .windowed = 1
        
        'Set the swap effect (the way that DirectX swaps
        'between the screen and the backbuffer) to be in
        'sync with the refresh rate
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        
        'Set and enable the Z-Buffer format, so that the scene
        'now has a 3D effect.
        .AutoDepthStencilFormat = D3DFMT_D16
        .EnableAutoDepthStencil = 1
                    
        'Set the format of the back buffer to the
        'current displays settings
        .BackBufferFormat = DispMode.Format
        
        'And this is where the window will actually be.
        'It can be pretty much anywhere (on the form,
        'in a picbox etc.) , all you need is the hWnd
        .hDeviceWindow = hWnd
    End With
    
    'd3d.CreateDevice(Adapter as Long, _
                DeviceType as CONST_D3DDEVTYPE, _
                hFocusWindow as Long, _
                BehaviourFlags as Long, _
                PresentationParameters as D3DPRESENT_PARAMETERS
    
    'Adapter    - Most graphics cards will use the default setting,
    '             so you can say D3DADAPTER_DEFAULT in most cases
    
    'DeviceType - Tells DirectX what type of device we wish to use.
    '             The main ones that you will use are:
    '             Hardware renderer = D3DDEVTYPE_HAL
    '             Sofware renderer  = D3DDEVTYPE_REF
    'The hardware renderer is your 3D Card doing
    'all the rendering.
    'The software renderer is a software emulator that
    'pretends to be a graphics card. It has a very slow frame
    'rate but supports all DirectX 8 features
    'There is a 3rd option but you will never be likely to use
    'it.  It allows the use of other software rendering devices.
    
    'hFocusWindow - The handle (hWnd) of where we want the scene
    '               to be displayed. To display straight to the
    '               form you would use frmMain.hWnd.  In this case
    '               we are using a picture box called picDirectX,
    '               so we use frmMain.picDirectX.hWnd
    
    'BehaviourFlags - This is how the created device will behave.
    '                 For most graphics cards (ones that do not
    '                 support hardware transform and lighting)
    '                 you would use D3DCREATE_SOFTWARE_VERTEXPROCESSING,
    '                 otherwise you would use D3DCREATE_HARDWARE_VERTEXPROCESSING
    
    'PresentationParameters - Just use the D3DWINDOW object setup above
    
    
        
    If D3D.CheckDeviceType(D3DADAPTER_DEFAULT, _
                           D3DDEVTYPE_HAL, _
                           D3DWindow.BackBufferFormat, _
                           D3DWindow.BackBufferFormat, _
                           True) = D3D_OK Then
        addReport ("Hardware rendering supported!")
                                  
    'This line above checks that the computer actually has a graphics card.
    'It doesn't check to see what the card can do, but that it's there. There
    'are ways of checking the graphics cards capbilities but i won't go into
    'that here. It is know as device enumeration if you wish to look into it
    'more.
                                  
        'If the computer has a graphics card, then use it!
        Set D3DDevice = _
        D3D.CreateDevice(D3DADAPTER_DEFAULT, _
                         D3DDEVTYPE_HAL, _
                         hWnd, _
                         D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                         D3DWindow)
        addReport ("D3DDevice object setup...")
        
    Else
        
        'Oh dear, no graphics card. In this case we inform the user
        'and use the software rasteriser.  This has full support for all
        'the DirectX features, but is far too slow to be used for anything
        'other than debugging.
        
        MsgBox "Hardware rendering not supported!  Switching to Software...", vbExclamation + vbOKOnly, "Render Mode Error!"
        addReport ("No 3D hardware!!!")
        addReport ("Software rendering used!")
        
        Set D3DDevice = _
        D3D.CreateDevice(D3DADAPTER_DEFAULT, _
                         D3DDEVTYPE_REF, _
                         hWnd, _
                         D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                         D3DWindow)
        addReport ("D3D Object Setup...")
        
    End If
         
    
Else    'Fullscreen mode
    
    'Now we come to full screen mode. The only difference
    'here is that you have to set the dimensions of the
    'backbuffer.  You can only use certain dimensions (640*480,
    '800*600,etc) here.  in this case,  i have used 800*600 with
    'a 16bit colour depth (D3DFMT_R5G6B5 = 5 Red,6 Green,5 Blue)
    
    With D3DWindow
        .windowed = False
        
        '800*600*16
        .BackBufferHeight = 600
        .BackBufferWidth = 800
                               
        'Set the backbuffer count to 1
        .BackBufferCount = 1
        .BackBufferFormat = DispMode.Format
        
        'Same as in windowed (refresh in sync with screen)
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
                
        '16bit Z-buffer
        .AutoDepthStencilFormat = D3DFMT_D16
        .EnableAutoDepthStencil = 1
                                
        'And use the hWnd of frmFullDisp
        .hDeviceWindow = hWnd
    End With
        
    'Check again to see which renderer to use
    If D3D.CheckDeviceType(D3DADAPTER_DEFAULT, _
                           D3DDEVTYPE_HAL, _
                           D3DWindow.BackBufferFormat, _
                           D3DWindow.BackBufferFormat, _
                           False) = D3D_OK Then
    
        'Hardware Rendering
        Set D3DDevice = _
        D3D.CreateDevice(D3DADAPTER_DEFAULT, _
                         D3DDEVTYPE_HAL, _
                         hWnd, _
                         D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                         D3DWindow)

    Else
        
        MsgBox "Hardware rendering not supported!  Switching to Software...", vbExclamation + vbOKOnly, "Render Mode Error!"
        
        'Not hardware rendering
        Set D3DDevice = _
        D3D.CreateDevice(D3DADAPTER_DEFAULT, _
                         D3DDEVTYPE_REF, _
                         hWnd, _
                         D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                         D3DWindow)
    
    End If
    
End If
    
    'We've finished, with no errors! Return true
    '(this is a function) so that the code can continue.
    InitD3D = True


'This initialisation function will only create a blank scene.
'You may adapt and add to this if you want to use it in another
'project of your own - Look! It's even in a nice separate module
'for you!

Exit Function
initFail:

addReport ("Initialisation FAILED!")
InitD3D = False

End Function


