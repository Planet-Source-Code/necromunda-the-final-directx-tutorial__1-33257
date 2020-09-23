Attribute VB_Name = "mTextures"
Option Explicit

'The object that holds our texture
Public texture As Direct3DTexture8

'Constant for textured vertices
Public Const FVF_TEX = (D3DFVF_XYZ Or D3DFVF_TEX1)

'Type for textured vertices
Type TEXVERTEX
  X As Single
  Y As Single
  Z As Single
  tu As Single
  tv As Single
End Type

'Path of the current texture
Public currentTexture As String

Sub initTexCube()
    Dim verts(35) As TEXVERTEX
    
    'In this section, instead of an octahedron,
    'i'm going to be using a cube, as it is easier
    'when dealing with texture co-ordinates.
        
    'When dealing with textures, you have a new
    'thing to add to the vertex type - texture coordinates.
    'These are used to line up the texture with your
    'shape. Imagine your texture as square with sides
    '1 unit long.  To work out the texture coordinates for a shape
    'you can imagine laying you shape on top of the texture.
    'from this you can work out roughly what the coordinates should
    'be.  If you are using a square then its easy.  More complex shapes
    'will require some working out.

    'Eg.
    '       Texture
    '(0,0)----------------(1,0)
    '  |        / \         |
    '  |       /   \        |
    '  |      /     \       |
    '  |     /       \      |
    '  |    /         \     |
    '  |   /           \    |
    '  |  /   Shape     \   |
    '  | /               \  |
    '(1,0)----------------(1,1)
    'For this example shape, I would set the bottom coordinates
    'to (1,0) and (1,1) and the top to (0.5,0).
    
    'If you still dont get it this formula might help:
    '((1/Width)*X, (1/Height)*Y)
    'where width, height, x and y are all pixel measurements.
    
    verts(0) = createTexVert(-1, 1, -1, 0, 0)
    verts(1) = createTexVert(1, 1, -1, 1, 0)
    verts(2) = createTexVert(-1, 1, 1, 0, 1)
        
    verts(3) = createTexVert(1, 1, -1, 1, 0)
    verts(4) = createTexVert(1, 1, 1, 1, 1)
    verts(5) = createTexVert(-1, 1, 1, 0, 1)
    
    verts(6) = createTexVert(-1, -1, -1, 0, 0)
    verts(7) = createTexVert(1, -1, -1, 1, 0)
    verts(8) = createTexVert(-1, -1, 1, 0, 1)
        
    verts(9) = createTexVert(1, -1, -1, 1, 0)
    verts(10) = createTexVert(1, -1, 1, 1, 1)
    verts(11) = createTexVert(-1, -1, 1, 0, 1)
    
    verts(12) = createTexVert(-1, 1, -1, 0, 0)
    verts(13) = createTexVert(-1, 1, 1, 1, 0)
    verts(14) = createTexVert(-1, -1, -1, 0, 1)
        
    verts(15) = createTexVert(-1, 1, 1, 1, 0)
    verts(16) = createTexVert(-1, -1, 1, 1, 1)
    verts(17) = createTexVert(-1, -1, -1, 0, 1)
    
    verts(18) = createTexVert(1, 1, -1, 0, 0)
    verts(19) = createTexVert(1, 1, 1, 1, 0)
    verts(20) = createTexVert(1, -1, -1, 0, 1)
        
    verts(21) = createTexVert(1, 1, 1, 1, 0)
    verts(22) = createTexVert(1, -1, 1, 1, 1)
    verts(23) = createTexVert(1, -1, -1, 0, 1)
        
    verts(24) = createTexVert(-1, 1, 1, 0, 0)
    verts(25) = createTexVert(1, 1, 1, 1, 0)
    verts(26) = createTexVert(-1, -1, 1, 0, 1)
    
    verts(27) = createTexVert(1, 1, 1, 1, 0)
    verts(28) = createTexVert(1, -1, 1, 1, 1)
    verts(29) = createTexVert(-1, -1, 1, 0, 1)
    
    verts(30) = createTexVert(-1, 1, -1, 0, 0)
    verts(31) = createTexVert(1, 1, -1, 1, 0)
    verts(32) = createTexVert(-1, -1, -1, 0, 1)
        
    verts(33) = createTexVert(1, 1, -1, 1, 0)
    verts(34) = createTexVert(1, -1, -1, 1, 1)
    verts(35) = createTexVert(-1, -1, -1, 0, 1)
    
    'Set the vertex shader to FVF_TEX
    D3DDevice.SetVertexShader FVF_TEX
    
    'Disable culling and lighting
    'Enable Z-Buffer
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    
    'Load the cube into the vertxbuffer
    Set vertexBuffer = D3DDevice.CreateVertexBuffer(Len(verts(0)) * 36, _
                                                    0, _
                                                    FVF_TEX, _
                                                    D3DPOOL_DEFAULT)
                
    'Set the vertexbuffer
    D3DVertexBuffer8SetData vertexBuffer, 0, Len(verts(0)) * 36, 0, verts(0)
     
End Sub

Sub setTexture(ByVal textureName As String)
    On Error GoTo report
    'If there is an error loading the texture then
    'display a messagebox
    
    'Set the default texture
    If textureName = "default" Then textureName = App.Path & "\green.jpeg"
    
    'Retreive the current display mode
    Dim DispMode As D3DDISPLAYMODE
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
'   CreateTextureFromFile( _
'           Device As Direct3DDevice8, _
'           SrcFile As String) As Direct3DTexture8
    
'    CreateTextureFromFileEx( _
'           Device As Direct3DDevice8, _
'           SrcFile As String, _
'           Width As Long, Height As Long, _
'           MipLevels As Long, Usage As Long, _
'           Format As CONST_D3DFORMAT, Pool As CONST_D3DPOOL, _
'           Filter As Long, MipFilter As Long, _
'           ColorKey As Long, SrcInfo As Any, Palette As Any) _
'           As Direct3DTexture8
    
    'The first function above is the simple version of the
    'second, it simply takes the string path of the texture -
    'D3D does the rest.
    
    'The second funtion give more control for doing more advanced things
    'The device is simple enough (D3DDevice), and the srcFile.
    'The width and height are the dimensions you wish to
    'store the texture as, D3D will resize it to fit.
    'The miplevels and usage parameters are to do with a technique
    'known as mip-mapping, which we aren't using.
    'Set them to 1 and 0 normally.
    'The format is best kept as the same as the current format,
    'and the pool is how (or where) the texture is stored in memory
    'terms.  Using D3DPOOL_MANAGED will let the drivers decide how
    'it should be stored.
    'The filter and mip filter are used for various texture techniques
    'again, default is usually a safe bet.
    'the final 3 parameters are rarely used, so set them all to 0.
    
    'Set texture = D3DX.CreateTextureFromFile(D3DDevice, textureName)
            
    Set texture = D3DX.CreateTextureFromFileEx(D3DDevice, _
                            textureName, _
                            256, 256, _
                            1, 0, _
                            DispMode.Format, _
                            D3DPOOL_MANAGED, _
                            D3DX_DEFAULT, D3DX_DEFAULT, _
                            0, ByVal 0, ByVal 0)

    currentTexture = textureName
    
    'After loading the texture, it's simple -
    'just set the texture.
    
    D3DDevice.setTexture 0, texture
    
Exit Sub

report:

    If Err.Number <> D3D_OK Then
        MsgBox "Error loading texture! Ensure that all texture files are of the correct format (*.bmp/jpeg)", vbOKOnly + vbExclamation
    
    End If


End Sub


Private Function createTexVert(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal tv As Single, ByVal tu As Single) As TEXVERTEX

'This function makes it easier and quicker to create a  textured type vertex.
'Inputs:    X,Y,Z vector
'           U,V texture coordinates
'Outputs:   Vertex of the TEXVERTEX type

    With createTexVert
        .X = X
        .Y = Y
        .Z = Z
        .tu = tu
        .tv = tv
    End With

End Function

