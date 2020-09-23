Attribute VB_Name = "mRender"
Option Explicit

'This is where the actual rendering gets done!

Public Sub Render()
    Dim vertexSize As Single    'Stores the size of the current vertex type
    Dim loopval As Single   'Looping variable
    
    Dim sizeVert As VERTEX           'Needed to work out the size of VERTEX
    Dim sizeNormVert As NORMALVERTEX 'Needed to work out the size of NORMALVERTEX
    Dim sizeTexVert As TEXVERTEX     'Needed to work out the size of TEXVERTEX
        
    'As we are using vertexbuffers, we need to know the
    'size of the vertex steps we shall be taking.
    'This section below determines the vertex type being
    'used and sets the vertexSize variable to the len()
    'of that.
    
    If currentApp = Normal Then
        vertexSize = Len(sizeVert)  'Normal vertex
    ElseIf currentApp = Lighting Then
        vertexSize = Len(sizeNormVert)  'Lighting vertex
    ElseIf currentApp = Texturing Then
        vertexSize = Len(sizeTexVert)   'Texture vertex
    End If
    
    'Before we begin rendering we have to clear out the
    'backbuffer to leave us with a clean sheet to render
    'to.
    'The clear function contains a lot of parameters for various
    'unimportant things.  At this basic level, you wont be changing
    'any except the colour value, which sets the background colour
    'of the scene. the rest of the parameters can be kept the same
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, RGB(225, 100, 100), 1#, 0
    D3DDevice.BeginScene
    'Tell D3D we're going to start rendering
        
    'This checks if dragmode is enabled.
    'If it isnt the transform (rotation etc.) sub is run
    If frmMain.chkDrag.Value = False Then matrixTransforms
    
    'The current running mode is texturing
    If currentApp = Texturing Then
        
        'We are using texture vertices here, so the above
        'code should have set the vertexsize to the correct value.
        'The vertexsize is required to tell directX the size of
        'each value in the vertexbuffer, which is essentially an
        'array.
        
        'First thing to do is to set which vertex buffer is
        'being used. This line below does this.
                
        'The main things you have to set here are the stride
        '(vertex size) and the StreamData (the vertexbuffer)
        'These should stay the same for most things.
        'If you wish to change to a different buffer, you can
        'just call this again with the new buffers name and
        'vertex size
        D3DDevice.SetStreamSource 0, vertexBuffer, vertexSize
        
        'This line gets the drawing to the backbuffer done.
        'DrawPrimitive is the render statement you use for
        'vertex buffers.  If you just wanted to render
        'from an array of vertices (eg. verts(0 to 3) as FVF)
        'you would use the DrawPrimitiveUP statement.
        'The D3DPT_TRIANGLELIST statement means that we are
        'rendering separate triangles in the buffer. ie. each
        'triangle is self contained (123-456-789).
        'this is not a very memory efficient way, but it is
        'the simplest to deal with.
        'Other ways the display the triangle are
        'D3DPT_TRIANGLEFAN and D3DPT_TRIANGLESTRIP
        'Triangle fan is where all the triangles share one
        'central point. (eg. 1-23-34-45-56) This is useful
        'for making things like circles easily.
        'Triangle strip is where each triangle forms the side
        'of another (eg. 123-234-345)  This is more efficient,
        'but is more confusing when working out the shape.
        
        'The final 2 numbers say where to start from, and
        'how many triangles there are (12 as this is a cube)
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 12
        
        
            
            
    ElseIf currentApp = Modelling Then
        
        'When rendering in modelling mode, things are
        'done slightly differently.  Now you dont have
        'to muck about with vertex buffers and things.
        'All you need is a loop for the number of materials.
        'Each model is broken down into several subsets -
        'eg. one for texture A, another for texture B etc.
        'Running through the loop will set the texture
        '(or material) for that stage, then render the
        'sub section  (cubeModel.modelMesh.DrawSubset loopval)
                
        If cubeModel.nMaterials = 0 Then
            'If there are no materials, then there
            'must only be one section
            cubeModel.modelMesh.DrawSubset 0
            
        Else
            'Loop for the model
            For loopval = 0 To cubeModel.nMaterials
                
                'Set current texture
                D3DDevice.setTexture 0, cubeModel.MeshTextures(loopval)
                'Set current material
                D3DDevice.SetMaterial cubeModel.MeshMaterials(loopval)
                'Draw section
                cubeModel.modelMesh.DrawSubset loopval
         
            Next loopval
        
        End If
        
    Else 'currentApp = lighting or normal
    
        'this is as above, but the vertex of different, and there
        'are only 8 triangles this time.
        '(lighting and normal mode use the same shape so they
        'can both use this render section)
        D3DDevice.SetStreamSource 0, vertexBuffer, vertexSize
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 8
        
    End If
    
    'Tell D3D we've finished rendering
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    'And finally present the scene!
    'This just swaps the screen buffer and the back buffer about.
    'All the parameters can be set to zero (ByVal 0 for Any types)

End Sub

