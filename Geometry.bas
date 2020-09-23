Attribute VB_Name = "mGeometry"
Option Explicit

'#Requires
'# Direct3DVertexBuffer8
'# That initD3D has been run

'# Constant for untransformed and lit vertices - XYZ and DIFFUSE
Public Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

'# Type for lit vertex
Type VERTEX
  X As Single
  Y As Single
  Z As Single
  Colour As Long
End Type

'-------------------------
'#initGeometry
'# Creates a coloured octahedral shape.

'#initMatrix
'# Initialises the view and projection matrices.

'#createVertex
'# Creates a lit vertex

'#D3DVec
'# Creates a D3DVECTOR
'-------------------------

'Now, here's where we make up our luvly little shapes, using equally nice
'vectors.  Vectors are great wee things that tell us where we are in
'3D space, using an X, Y and Z value. A point in 3D space (like the corner
'of a cube) is known as a vertex.

'Eg.
'Using the normal type co-ordinate system (X across the screen,
'Y up the screen, Z into the screen):

'verts(0) = createVertex(-1, 0, -1, vbBlue) means go -1 in the X-axis,
'0 in the Y-axis, and -1 in the Z-axis, then make the vertex there.
'And make it blue.

'As we are making an octahedron (think of 2 pyramids
'stuck base to base) here we need to make 24 verticies. This works out as
'8 triangles - 3 verticies per triangle.

'For a cube, we would need 36 verticies
'(cube has 6 faces, but each face is made up of 2 triangles.
'so that makes 12 triangles = 36 verticies). It really helps if you work out
'the points that you need beforehand, using pen and paper...


'Right. That's the basic theory explained...now the VB stuff.
'In the declarations section you'll notice a type (FVF_VERTEX) and
'a constant (VERTEX).  These describe what type of vertex that you're
'gonna be using, using the Flexible Vertex Format system in DirectX.
'This lovely system means that you can mix and match your vertex types.

'In this module we only need a basic vertex type,
'the X,Y and Z of each vertex, and to make it look nice, the colour.
'For the X,Y and Z we need to add D3DFVF_XYZ to the constant, and to
'the type:

'Type VERTEX
'  X As Single \
'  Y As Single  = D3DFVF_XYZ
'  Z As Single /
'End Type

'For the colour we add D3DFVF_DIFFUSE to the constant, making the type:

'Type VERTEX
'  X As Single \
'  Y As Single  = D3DFVF_XYZ
'  Z As Single /
'  Colour As Long } D3DFVF_DIFFUSE
'End Type

'Looking at this way can make it a bit easier to understand. You tell
'DirectX what qualities you wish your vertex to have, and build up the
'UDT from that.
'Below is a list of some of the FVF types you are liely to use, and
'a wee bit about what they do..

'Type EXAMPLEVERTEX
'   X as Single   \
'   Y as Single    | D3DFVF_XYZ - the x,y,z of each vertex. Fairly essential.
'   Z as Single   /
'   rhw as Single/ D3DFVF_XYZRHW - for doing 2D graphics. You cant do any matrix stuff with this one.
'   nX as Single \
'   nY as Single  = D3DFVF_NORMAL - for lighting use, specifies the normal of each vertex
'   nZ as Single /
'   tu as Single \ D3DFVF_TEX1 - for basic texturing (texture coordinates use u,v,w instead of x,y,z)
'   tv as single /
'   Colour as Long = D3DFVF_DIFFUSE - not too tough
'   Specular as Long = D3DFVF_SPECULAR - for making things look shiny...
'End Type

'NOTE: You cannot use the D3DFVF_XYZRHW flag with the XYZ and NORMAL flags

'These should allow you to do most of the basic things.
'When making your vertex type, cut out the bits you dont really need
'(you won't need normals if you are having making unlit textured cube etc.)
'For a full list of the FVF things look up the DirectX 8 SDK...

Sub initGeometry()
Dim verts(23) As VERTEX
Dim rt2 As Single

    rt2 = Sqr(2)

    'Making an octahedron with coloured corners.
    'I've separated them into each triangle (8 in all)

    'Top - Front face
    verts(0) = createVertex(-1, 0, -1, vbBlue)
    verts(1) = createVertex(0, rt2, 0, vbCyan)
    verts(2) = createVertex(1, 0, -1, vbRed)
    
    'Top - Back face
    verts(3) = createVertex(-1, 0, 1, vbRed)
    verts(4) = createVertex(0, rt2, 0, vbCyan)
    verts(5) = createVertex(1, 0, 1, vbBlue)

    'Top - Right Face
    verts(6) = createVertex(1, 0, 1, vbBlue)
    verts(7) = createVertex(0, rt2, 0, vbCyan)
    verts(8) = createVertex(1, 0, -1, vbRed)
    
    'Top - Left Face
    verts(9) = createVertex(-1, 0, 1, vbRed)
    verts(10) = createVertex(0, rt2, 0, vbCyan)
    verts(11) = createVertex(-1, 0, -1, vbBlue)

    'Bottom - Front Face
    verts(12) = createVertex(-1, 0, -1, vbBlue)
    verts(13) = createVertex(0, -rt2, 0, vbCyan)
    verts(14) = createVertex(1, 0, -1, vbRed)
    
    'Bottom - Back Face
    verts(15) = createVertex(-1, 0, 1, vbRed)
    verts(16) = createVertex(0, -rt2, 0, vbCyan)
    verts(17) = createVertex(1, 0, 1, vbBlue)

    'Bottom - Right Face
    verts(18) = createVertex(1, 0, 1, vbBlue)
    verts(19) = createVertex(0, -rt2, 0, vbCyan)
    verts(20) = createVertex(1, 0, -1, vbRed)

    'Bottom - Left Face
    verts(21) = createVertex(-1, 0, 1, vbRed)
    verts(22) = createVertex(0, -rt2, 0, vbCyan)
    verts(23) = createVertex(-1, 0, -1, vbBlue)
        
    
    'Set DirectX to use out vertex type that we've made
    D3DDevice.SetVertexShader FVF_VERTEX
    
    'Enable lighting, the z-buffer and set the cullmode (see lighting) to none.
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    
    'Now we come to the vertex buffer.
    'This is an area set aside especially for storing verticies, which
    'makes it generally easier to deal with all those verticies.
    'The verticies are first loaded into an array of the corresponding vertex type.
    'You then simply load them into the vertex buffer (declared in mObjects)
    'by stating the size of them (the size of one multiplied by the number there are)
    'and setting the FVF type. The usage and pool should be set to 0 and D3DPOOL_DEFAULT
    'most of the time, you will rarely need to change them.  Once they have been
    'loaded into the vertexbuffer you inform DirectX that it exists and fill in the
    'parameters. Simple.
    'During the render phase you have to set the vertexbuffer using the statement:
    'D3DDevice.SetStreamSource N, vertexBuffer, vertexSize
    'Where N is the number of the vertexbuffer,
    'vertexBuffer is the direct3D vertexbuffer and
    'vertexsize is the size of the vertices used in the buffer.
        
    Set vertexBuffer = D3DDevice.CreateVertexBuffer(Len(verts(0)) * 24, _
                                                    0, _
                                                    FVF_VERTEX, _
                                                    D3DPOOL_DEFAULT)
        
    D3DVertexBuffer8SetData vertexBuffer, 0, Len(verts(0)) * 24, 0, verts(0)
   

End Sub

Sub initMatrix()
    Dim matView As D3DMATRIX
    Dim matProj As D3DMATRIX
    
    'Right, now the matrix stuff.
    'Matrices are used to set up the scene so that it actually looks
    '3D, and to do stuff to our nice shapes, like rotating and
    'scaling them.
            
    'The three main matricies that you will be dealing with are the
    'view, projection and the world matrix.
    
    'The view matrix is best thought of as a camera.  It has 3 vector
    'values, the location of it, where it's looking at and which direction up
    'is.  This is the simplest matrix to use of the 3.
    
    'The projection matrix is used least often out of the 3 types, you should only
    'need to set it once at the start of your program.  It dictates how the camera
    'views our scene.
    'The fovy value specifies the viewing angle in radians, PI/4 is equal to 45 degrees
    'to convert an angle from degrees to radians you just multiply by PI/180
    'The aspect ratio is to do how the vertical aspect of the scene is related to
    'the horizontal.  1 is usually a good bet for this value, higher or lower
    'will stretch the scene in wierd ways.
    'The zf and zn are simple enough, the zf vaule is the far distance to render to
    'and the zn is the near distance.  Anything beyond the far distance is culled as
    'is anything before the near distance. The distances in D3D are usually measured
    'in metres, though it doesn't really matter much.
            
    'The world matrix is the main matrix that you will be using, as it does all
    'the rotating, scaling and translating (moving) of vertices.
    
            
    D3DXMatrixLookAtLH matView, _
        D3DVec(1#, 2#, -5.5), _
        D3DVec(0#, 0#, 0#), _
        D3DVec(0#, 1#, 0#)
    D3DDevice.SetTransform D3DTS_VIEW, matView
    
    
    D3DXMatrixPerspectiveFovLH matProj, _
        PI / 4, _
        1, _
        0.1, 100
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj


End Sub

Sub moveCamera(location As D3DVECTOR)
    Dim matMove As D3DMATRIX

    'Takes the user input from frmMain and moves the camera accordingly
    D3DXMatrixLookAtLH matMove, _
        D3DVec(location.X, location.Y, location.Z), _
        D3DVec(0#, 0#, 0#), _
        D3DVec(0#, 1#, 0#)
    D3DDevice.SetTransform D3DTS_VIEW, matMove
    
End Sub

Sub matrixTransforms()
    Dim matRotate As D3DMATRIX
    Dim matTrans As D3DMATRIX
    Dim matScale As D3DMATRIX
    Dim matWorld As D3DMATRIX
    
    With frmMain
    
        'ROTATING
        If .chkRotate.Value = 1 Then
            'This function combines the X,Y and Z
            'rotations into one.  you have to pass
            'in a matrix to recieve the result, and
            'values for X,Y and Z.
            '(The iif statements just check which options
            'the user has selected)
            'I have set this to swing around the X-Axis
            '(hence the sin(timer))  and just rotate and Y & Z
            
            D3DXMatrixRotationYawPitchRoll matRotate, _
                IIf(.chkX.Value = 1, 1.2 * Sin(Timer), 0), _
                IIf(.chkY.Value = 1, 0.8 * Timer, 0), _
                IIf(.chkZ.Value = 1, Timer, 0)
        Else
            
            'If no rotations, set the rotation matrix
            'to the equivelant of 1
            D3DXMatrixIdentity matRotate
        End If
            
        'TRANSLATION
        If .chkMove.Value = 1 Then
            'Like the rotation matrix function, except
            'this is for translating around the X,Y & Z.
            'I have set it to move around in a simple circle
            'in X and Y
            D3DXMatrixTranslation matTrans, _
                Cos(Timer), _
                Sin(Timer), _
                0
        Else
            
            'set to 1
            D3DXMatrixIdentity matTrans
        End If
            
        
        'SCALING
        If .chkScale.Value = 1 Then
            
            'This is the function for making the scene
            'shrink or grow.
            'Note: if you set the vaules to different things,
            'then the shape will warp and stretch in wierd ways
            'Also, setting a negative number will turn the shape
            'inside out, ruining any lighting effect.
            
            D3DXMatrixScaling matScale, _
                 Abs(Sin(Timer)), _
                 Abs(Sin(Timer)), _
                 Abs(Sin(Timer))
                      
        Else
            
            'Set to one
            D3DXMatrixIdentity matScale
        End If
        
    End With
        
    'This last section before setting the transformations combines
    'all the matrix effects into one matrix, using the matrix
    'multiply function.  The order in which you multiply the matrices
    'matters here, X * Y is not the same as Y * X.
    'The order I have here rotates the shape around 0, moves it to
    'a new location, then performs any scaling.
    'If I translated first, then I would be moving the shape, then
    'performing the rotation at this new position.
    
    D3DXMatrixMultiply matWorld, matRotate, matTrans
    D3DXMatrixMultiply matWorld, matScale, matWorld
                
    'After combining all the effects, set them as the world matrix.
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
   
    
End Sub

Private Function createVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal Colour As Long) As VERTEX
    
    'A function that makes it easier to create all the vertices.
    'It loads all the given values into the vertex
    'Inputs:  XYZ
    '         Colour
    'Outputs: Vertex of type VERTEX
    
    With createVertex
        .X = X
        .Y = Y
        .Z = Z
        .Colour = Colour
    End With

End Function

Function D3DVec(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
    
    'Create a D3DVECTOR
    'Inputs:  XYZ
    'Outputs: D3DVector
        
    With D3DVec
        .X = X
        .Y = Y
        .Z = Z
    End With

End Function
