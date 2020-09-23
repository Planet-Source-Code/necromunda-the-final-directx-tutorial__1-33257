Attribute VB_Name = "mLights"
Option Explicit

'#Requires
'# Direct3DVertexBuffer8
'# That initD3D has been run

'# Constant for untransformed and unlit verticies
Public Const NORMAL_FVF = (D3DFVF_XYZ Or D3DFVF_NORMAL)

'# Type for untransformed and unlit verticies.
Type NORMALVERTEX
    X As Single
    Y As Single
    Z As Single
    nX As Single
    nY As Single
    nZ As Single
End Type

'--------------------------

'#initLightingGeo
'# Creates a lit octahedron

'#GenerateTriangleNormals
'# Calculates the normals for each triangle

'#SetupLights
'# Sets the material of the scene and
'# creates a coloured point light.

'--------------------------

Sub initLightingGeo()
Dim normVec As D3DVECTOR
Dim rt2 As Single
Dim verts(23) As NORMALVERTEX
    
    rt2 = Sqr(2)
    
    'These are the same vertices that are used to make the
    'octahedron in normal mode, except that in this module,
    'they are rendered in clockwise order.
    
    '1 - 2  or  0 - 1
    '|            /
    '0          2
    
    'You must do this so that the faces are lit properly in respect to
    'the position of the light, which means you could end up
    'with faces that are lit up when facing away from the light.
    'Also, with the verticies in clockwise order, you can
    'enable culling. Culling is when only the vertices that
    'can be seen are rendered, so that the back faces of the octahedron
    'wont show up. You can cull in clockwise or anticlockwise direction
    'by setting the d3dDevice.RenderState to D3DCULL_CW or D3DCULL_CCW
    'The default vaule for this is D3DCULL_NONE
    
        
    'Top - Front
    verts(0) = createNormVert(0, rt2, 0, 0, 0, 0)
    verts(1) = createNormVert(1, 0, -1, 0, 0, 0)
    verts(2) = createNormVert(-1, 0, -1, 0, 0, 0)
        'Work out triangle nornals using the 3 vertices of the triangle
        normVec = GenerateTriangleNormals(verts(0), verts(1), verts(2))
        'Fill in the normal values
        verts(0).nX = normVec.X: verts(0).nY = normVec.Y: verts(0).nZ = normVec.Z
        verts(1).nX = normVec.X: verts(1).nY = normVec.Y: verts(1).nZ = normVec.Z
        verts(2).nX = normVec.X: verts(2).nY = normVec.Y: verts(2).nZ = normVec.Z
        
    'Top - Back
    verts(3) = createNormVert(0, rt2, 0, 0, 0, 0)
    verts(4) = createNormVert(-1, 0, 1, 0, 0, 0)
    verts(5) = createNormVert(1, 0, 1, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(3), verts(4), verts(5))
        verts(3).nX = normVec.X: verts(3).nY = normVec.Y: verts(3).nZ = normVec.Z
        verts(4).nX = normVec.X: verts(4).nY = normVec.Y: verts(4).nZ = normVec.Z
        verts(5).nX = normVec.X: verts(5).nY = normVec.Y: verts(5).nZ = normVec.Z
    
    'Top - Right
    verts(6) = createNormVert(0, rt2, 0, 0, 0, 0)
    verts(7) = createNormVert(1, 0, 1, 0, 0, 0)
    verts(8) = createNormVert(1, 0, -1, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(6), verts(7), verts(8))
        verts(6).nX = normVec.X: verts(6).nY = normVec.Y: verts(6).nZ = normVec.Z
        verts(7).nX = normVec.X: verts(7).nY = normVec.Y: verts(7).nZ = normVec.Z
        verts(8).nX = normVec.X: verts(8).nY = normVec.Y: verts(8).nZ = normVec.Z
    
    'Top - Left
    verts(9) = createNormVert(0, rt2, 0, 0, 0, 0)
    verts(10) = createNormVert(-1, 0, -1, 0, 0, 0)
    verts(11) = createNormVert(-1, 0, 1, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(9), verts(10), verts(11))
        verts(9).nX = normVec.X: verts(9).nY = normVec.Y: verts(9).nZ = normVec.Z
        verts(10).nX = normVec.X: verts(10).nY = normVec.Y: verts(10).nZ = normVec.Z
        verts(11).nX = normVec.X: verts(11).nY = normVec.Y: verts(11).nZ = normVec.Z
    
    'Bottom - Front
    verts(12) = createNormVert(-1, 0, -1, 0, 0, 0)
    verts(13) = createNormVert(1, 0, -1, 0, 0, 0)
    verts(14) = createNormVert(0, -rt2, 0, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(12), verts(13), verts(14))
        verts(12).nX = normVec.X: verts(12).nY = normVec.Y: verts(12).nZ = normVec.Z
        verts(13).nX = normVec.X: verts(13).nY = normVec.Y: verts(13).nZ = normVec.Z
        verts(14).nX = normVec.X: verts(14).nY = normVec.Y: verts(14).nZ = normVec.Z
    
    'Bottom - Back
    verts(15) = createNormVert(1, 0, 1, 0, 0, 0)
    verts(16) = createNormVert(-1, 0, 1, 0, 0, 0)
    verts(17) = createNormVert(0, -rt2, 0, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(15), verts(16), verts(17))
        verts(15).nX = normVec.X: verts(15).nY = normVec.Y: verts(15).nZ = normVec.Z
        verts(16).nX = normVec.X: verts(16).nY = normVec.Y: verts(16).nZ = normVec.Z
        verts(17).nX = normVec.X: verts(17).nY = normVec.Y: verts(17).nZ = normVec.Z
    
    'Bottom - Right
    verts(18) = createNormVert(1, 0, -1, 0, 0, 0)
    verts(19) = createNormVert(1, 0, 1, 0, 0, 0)
    verts(20) = createNormVert(0, -rt2, 0, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(18), verts(19), verts(20))
        verts(18).nX = normVec.X: verts(18).nY = normVec.Y: verts(18).nZ = normVec.Z
        verts(19).nX = normVec.X: verts(19).nY = normVec.Y: verts(19).nZ = normVec.Z
        verts(20).nX = normVec.X: verts(20).nY = normVec.Y: verts(20).nZ = normVec.Z
        
    'Bottom - Left
    verts(21) = createNormVert(-1, 0, 1, 0, 0, 0)
    verts(22) = createNormVert(-1, 0, -1, 0, 0, 0)
    verts(23) = createNormVert(0, -rt2, 0, 0, 0, 0)
        normVec = GenerateTriangleNormals(verts(21), verts(22), verts(23))
        verts(21).nX = normVec.X: verts(21).nY = normVec.Y: verts(21).nZ = normVec.Z
        verts(22).nX = normVec.X: verts(22).nY = normVec.Y: verts(22).nZ = normVec.Z
        verts(23).nX = normVec.X: verts(23).nY = normVec.Y: verts(23).nZ = normVec.Z
    
    'Set the device to use the lit vertices
    D3DDevice.SetVertexShader NORMAL_FVF
    
    'Enable lighting
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1
    
    'Set an ambient light that illuminates the scene evenly
    D3DDevice.SetRenderState D3DRS_AMBIENT, &H101010
    
    'Set the cullmode to cull anticlockwise, as our shape's
    'vertices are in clockwise direction.
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    
    Set vertexBuffer = D3DDevice.CreateVertexBuffer(Len(verts(0)) * 24, _
                                                    0, _
                                                    NORMAL_FVF, _
                                                    D3DPOOL_DEFAULT)

    D3DVertexBuffer8SetData vertexBuffer, 0, Len(verts(0)) * 24, 0, verts(0)


End Sub

Private Function GenerateTriangleNormals(p0 As NORMALVERTEX, p1 As NORMALVERTEX, p2 As NORMALVERTEX) As D3DVECTOR
    Dim vNorm As D3DVECTOR
    Dim temp1 As D3DVECTOR
    Dim temp2 As D3DVECTOR
    
    'This function generates the normals required to use lighting.
    'The normals are worked out using a bit of vector maths, so it
    'might be helpful if you looked at some textbook or website that
    'explains it.
    
    'The first step is to create 2 vectors from the 1st vertex of
    'the triangle to the second and third vertcies. Vector subtraction
    'works by subtracting each X,Y and Z value from each other
    '(x1,y1,z1) - (x2,y2,z2) = (x1 - x2), (y1 - y2), (z1 - z2)
    
    'The next step is to calculate the cross product of these 2 new
    'vectors. The cross product is:
    '(x1,y1,z1) * (x2,y2,z2) = (y1*z2 - z1*y2, z1*x2 - x1*z2, x1*y2 - y1*x2)
    
    ' 0-------1     >      0-------1    >       0------------1
    '        /      >      |            >       |\
    '       /       >      |            >       | \
    '      /        >      |            >       |  \
    '     /         >      |            >       |   \
    '    /          >      |            >       |    \
    '   /           >      |            >       |     \
    '  /            >      |            >       |      \
    ' 2             >      2            >       2       normal
    
    'The resulting vector is returned by the function and applied to all the
    'verticies in the triangle
        
    'Subtract vector 1 from vector 0
    temp1.X = p1.X - p0.X
    temp1.Y = p1.Y - p0.Y
    temp1.Z = p1.Z - p0.Z
    
    'Subtract vectr 2 from vector 0
    temp2.X = p2.X - p0.X
    temp2.Y = p2.Y - p0.Y
    temp2.Z = p2.Z - p0.Z
       
    'Work out the cross product of the 2
    D3DXVec3Cross vNorm, temp1, temp2
    
    'Normalise the vectors
    D3DXVec3Normalize vNorm, vNorm

    GenerateTriangleNormals.X = vNorm.X
    GenerateTriangleNormals.Y = vNorm.Y
    GenerateTriangleNormals.Z = vNorm.Z

End Function

Sub SetupLights()
    Dim light As D3DLIGHT8 'Direct3D Light
    Dim Material As D3DMATERIAL8 'Direct3D material
    
    'The material of the scene means the way that
    'the shapes that are being rendered reflect light.
    'A material that has equal colour vaules will
    'reflect colours evenly, so the object will reflect
    'the colour of the light.  An object with red colour
    'vaules will only reflect red light.
    
    'Material.Ambient  = how the material reflects ambient
    '                    light.  Ambient light is set using
    '                    D3ddevice.SetRenderState D3DRS_AMBIENT
    '                    and a hexidecimal value for the colour.
    '                    Ambient light lights the whole scene
    '                    evenly, there is no source or direction
    '                    to it.
    'Material.Diffuse  = how the material reflects diffuse light.
    '                    Diffuse light is light that comes from
    '                    a D3D light source.
    'Material.Emissive = This gives the impression of the material
    '                    emitting light.  Note that it doesn't
    '                    actually emit any light, so nearby objects
    '                    won't be affected.
    'Material.Specular = The shinyness of the material, and how it
    '                    reflects this shinyness.  This only works if
    '                    D3DDevice.SetRenderState D3DRS_SPECULARENABLE, 1
    '                    is set and if the light source has a specular
    '                    component.  It works in conjunction with the
    '                    material.power value
    'Material.power    = The degree of shinyness of the material.  Smaller
    '                    values make the material more shiny, with values
    '                    between 20 and 50 looking best.
    
    'Reflect all ambient and diffuse light
    Material.Ambient = loadColours(1, 1, 1, 1)
    Material.diffuse = loadColours(1, 1, 1, 1)
    
    'Set the material to the scene
    D3DDevice.SetMaterial Material
                                                
    'There are three types of light, and they are are quite simple to
    'use and understand.
    'The easiest type (excluding ambient light) is the directional light.
    'The only things you need to specify with the directional light
    'are direction and colour.  A directional light can be thought of as
    'a very bright light very far away, such as the sun.
    
    'The next light type is the one demonstrated below, the point light.
    'The point light has the direction and colour as the directional light,
    'but also position, range and attenuation. Attenuation is how much
    'darker the light becomes over the range that is set.  There are 3 different
    'attenuation variables, 0, 1 and 2 relating to Constant, Linear and Quadratic
    'attentuation.  Messing about with these can create strange light effects,
    'like lights the get brighter the further away you get, or dark lights that
    'cast darkness.
    
    'The final light type is the spotlight. These are relatively difficult to
    'setup compared to the others as you need to specify two extra values.
    'Spot lights have two light cones, an inner cone and an outer cone.
    'You have to specify the angle of each cone (in radians again) Theta being
    'the inner cone and Phi being the outer cone.  This should create a light
    'with a bright inner light fading out towards the edge.
                                            
    With light
        .Type = D3DLIGHT_POINT
        
        .Position = D3DVec(0, 1, -5)
        .Direction = D3DVec(0, 0, 0)
        
        .diffuse = loadColours(1, Rnd() * 2, Rnd() * 2, Rnd() * 2)
        
                        
        .Range = 10
        .Attenuation1 = 0.3
    End With

    D3DDevice.SetLight 0, light
    D3DDevice.LightEnable 0, 1

End Sub

Private Function createNormVert(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal nX As Single, ByVal nY As Single, ByVal nZ As Single) As NORMALVERTEX

'This function makes it easier and quicker to create a normal type vertex.
'Inputs:    X,Y,Z vector
'           Normal vector
'Outputs:   Vertex of the NORMALVERTEX type

    With createNormVert
        .X = X
        .Y = Y
        .Z = Z
        .nX = nX
        .nY = nY
        .nZ = nZ
    End With

End Function

Private Function loadColours(ByVal a As Single, ByVal r As Single, ByVal g As Single, ByVal b As Single) As D3DCOLORVALUE
    
'This function makes it easier to make a colour value
'Inputs:    Alpha, Red, Green, Blue values
'Outputs:   D3DCOLOURVALUE
    
    With loadColours
        .a = a
        .r = r
        .g = g
        .b = b
    End With

End Function

