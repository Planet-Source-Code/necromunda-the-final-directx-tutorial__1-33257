Attribute VB_Name = "mModels"
Option Explicit

'#Type containing all variables needed to load a model
Type modelType
    modelMesh As D3DXBaseMesh           'Holds the model
    nMaterials As Long                  'Stores the number of materials/textures
    MeshMaterials() As D3DMATERIAL8     'Array of the materials of the model
    MeshTextures() As Direct3DTexture8  'Array of the textures of the model
End Type

'#Path of the current model
Public currentModel As String


Public cubeModel As modelType

'----------------------------
'#loadModel
'# Loads a .X file and all associated textures and materials

'#setupModelLghts
'# Creates 2 white point lights

'#LoadColours
'# Creates a D3DCOLORVALUE
'-----------------------------

'Modelling can sound like quite a hard topic, but after
'you know how it works you'll wonder how you ever did without it.
'The type of files that DirectX uses are .X files.  You can create these
'using 3d modelling software such as Milkshape 3D (30 day-shareware) or
'by downloading them from 3d model sites.  There is a program within the
'SDK that will convert 3D Studio Max files (*.3ds) to .X files for you,
'most files on the web are of the .3ds format.

'Once you have a .X model file (and textures that go with it, if any)
'it's quite a simple matter to load it.
'There are 4 main things associated with loading a model:
'First you have the D3DXMesh, which holds the vertex information about the
'model.
'You can't have a decent model without some sort of textures and/or materials on it,
'so you have to have something to store them in. You'll need 2 open-ended arrays
'for the textures and the materials, open-ended as you wont know how many
'textures/materials until runtime.
'The last thing you need is a simple long variable to store the number of
'textures and materials.

'The first step is to load the model into the mesh using the D3DX library
'function D3DX.LoadMeshFromX.
'eg set model = D3DX.LoadMeshFromX PathOfModel, D3DXMESH_MANAGED, _
                        D3DDevice, Nothing, D3DXBuffer, NumberMaterials)

'You need to pass in a D3DXBuffer to store the materials and a long variable
'for the number of materials, and the path of the model obviously.
'The rest of the values can be left the same for each time you use it.
'Once the model is loaded you need to redimension the 2 arrays using
'the long value that is now storing the number of materials/textures.
'The next step is to actualy load the textures and materials into the arrays.
'A loop is used to run through through the number of materials and textures.
'Each material and texture are loaded into the array, if there is no texture then
'it skips the loading line.

'When using models, use render them using the cubeModel.modelMesh.DrawSubset.
'Again here you have to use a loop for the number of materials:
'eg.

'For loop = 1 to nMaterials
'   cubeModel.modelMesh.DrawSubset loop
'
'next loop


Sub loadModel(modelName As String)
    Dim mtrlBuffer As D3DXBuffer        'Materials buffer
    Dim textureLoc As String            'Holds the dir of the textures
    Dim textureName As String           'holds the name of the texture
    Dim loopval As Long                 'Looping variable
    
    On Local Error GoTo report
    
    
    
    If modelName = "default" Then
          'The D3DX library contains many useful functions for creating
          'models, such as this one below for creating a teapot.
          'There are several other useful one for creating spheres and
          'boxes.
                
        'Set cubeModel.modelMesh = D3DX.CreateTeapot(D3DDevice, mtrlBuffer)
        Set cubeModel.modelMesh = D3DX.CreateSphere(D3DDevice, 2, 80, 80, mtrlBuffer)
        'Exit the sub as the teapot has no textures or materials.
        Exit Sub
    End If
    
    'Retrieves the directory of the model file
    'This is assuming that the texture files will be in the same
    'directory as the model.
    textureLoc = Left(modelName, InStrRev(modelName, "\"))

    'Load the model into cubeModel
    Set cubeModel.modelMesh = D3DX.LoadMeshFromX(modelName, D3DXMESH_MANAGED, _
                            D3DDevice, Nothing, mtrlBuffer, cubeModel.nMaterials)
    'Store the name of the current model
    currentModel = modelName

    'Redimension the texture and material arrays
    ReDim cubeModel.MeshMaterials(cubeModel.nMaterials)
    ReDim cubeModel.MeshTextures(cubeModel.nMaterials)
    
    'Loop to load the materials and textures
    For loopval = 0 To cubeModel.nMaterials - 1
        
        'Retrieve the material from the mtrlBuffer
        D3DX.BufferGetMaterial mtrlBuffer, loopval, cubeModel.MeshMaterials(loopval)
        
        'Retrieve the name of the texture
        textureName = D3DX.BufferGetTextureName(mtrlBuffer, loopval)
            
        'Fill in the gaps of the material
        cubeModel.MeshMaterials(loopval).Ambient = cubeModel.MeshMaterials(loopval).diffuse
    
        'Check if there is a texture to load
        If textureName <> "" Then
            
            'Load the model texture
            Set cubeModel.MeshTextures(loopval) = _
            D3DX.CreateTextureFromFileEx(D3DDevice, textureLoc & textureName, _
                                         256, 256, D3DX_DEFAULT, 0, _
                                         D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                         D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, _
                                         0, ByVal 0, ByVal 0)
        End If
    
    Next loopval
    
    'Disable culling
    'Some model files do not have the vertices in the correct order.
    'This will prevent any holes.
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

Exit Sub
report:

    If Err.Number <> D3D_OK Then
        MsgBox "Error loading model file! Check that all files are present", vbExclamation + vbOKOnly, "Models"
    
    End If
    

End Sub

Sub SetupModelLights()
    Dim Material As D3DMATERIAL8
    Dim light As D3DLIGHT8
        
    'If the model has no materials then we have to set some for it
    If cubeModel.nMaterials = 0 Then
        
        'White ambient
        Material.Ambient = loadColours(1, 1, 1, 1)
        'Light blue
        Material.diffuse = loadColours(1, 0.3, 0.3, 1)
        
        'Set white specular, which is how shiny the
        'material will look
        Material.specular = loadColours(1, 1, 1, 1)
        
        'Set the specular power
        Material.power = 25
    
        'Enable specular lighting
        D3DDevice.SetRenderState D3DRS_SPECULARENABLE, 1
        D3DDevice.SetMaterial Material
    
    End If
    
    With light
        .Type = D3DLIGHT_POINT
                
        .position = D3DVec(0, 8, -3)
        .Direction = D3DVec(0, 0, 0)
                    
        .diffuse = loadColours(1, 1, 1, 1)
        .specular = loadColours(1, 1, 1, 1)
                
                
        .Range = 20
        .Attenuation1 = 0.12
    End With

    D3DDevice.SetLight 0, light
    
    With light
        .Type = D3DLIGHT_POINT
        
        .position = D3DVec(0, -2, -3)
        .Direction = D3DVec(1, 1, 1)
        
        .diffuse = loadColours(1, 1, 1, 1)
        .specular = loadColours(1, 1, 1, 1)

        
        .Range = 20
        .Attenuation1 = 0.5
    End With
    
    D3DDevice.SetLight 1, light
    
    D3DDevice.LightEnable 0, 1
    D3DDevice.LightEnable 1, 1
    


End Sub

Private Function loadColours(ByVal a As Single, ByVal r As Single, ByVal g As Single, ByVal b As Single) As D3DCOLORVALUE
    
    With loadColours
        .a = a
        .r = r
        .g = g
        .b = b
    End With

End Function

