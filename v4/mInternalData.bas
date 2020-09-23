Attribute VB_Name = "mInternalData"
Option Explicit

' Define the name of this class/module for error-trap reporting.
Private Const m_strModuleName As String = "mInternalData"

Public g_strFeedback() As String

Public Function CreateCube() As mdr3DObject

    With CreateCube
        .Caption = "Cube Object"
        .Description = "This cube object has 1 Polyhedron (the cube), 8 Vertices and 6 Sides made from 12 Polygons (ie. Triangles)."
        
        .Pitch = 0
        .Roll = 0
        .Yaw = 0
        
        .PointingAt.x = 0
        .PointingAt.y = 0
        .PointingAt.z = 0
        .PointingAt.w = 1
        
        .UniformScale = 1
        
        .WorldPosition.x = 0
        .WorldPosition.y = 0
        .WorldPosition.z = 0
        .WorldPosition.w = 1
        
        ReDim .Polyhedra(0)
        With .Polyhedra(0)
            .Caption = "Cube Polyhedron"
            .Description = "Description of the Cube Polyhedron"
            
            .IdentityMatrix = MatrixIdentity
            .PointingAt.x = 0
            .PointingAt.y = 0
            .PointingAt.z = 0
            .PointingAt.w = 1
            
            .Pitch = 0
            .Roll = 0
            .Yaw = 0
            
        End With
        
    End With
    
End Function

Public Sub InitPSCFeedback()
    
    ReDim g_strFeedback(9)
    
    g_strFeedback(0) = """This code is so beautifully simple, really cool, I loved it!!! 5 globes, would give you more..."", Jonathan D (PSC Voter)"
    g_strFeedback(1) = """Very good tutorial for beginners, 5 Globes"", Josh Nixon (PSC Voter)"
    g_strFeedback(2) = """Excellent, 5 globes from me."", Eugene Wolff (PSC Voter)"
    g_strFeedback(3) = """5 globes from me, have you think to release a more advanced tutorial?"", Carlos Bomtempo (PSC Voter)"
    g_strFeedback(4) = """gold 5 stars"", RPG MAKER (PSC Voter)"
    g_strFeedback(5) = """Simply excellent..."", Carles P.V. (PSC Voter)"
    g_strFeedback(6) = """what can i say? just excellent"", Carlos Bomtempo (PSC Voter)"
    g_strFeedback(7) = """5 Stars from me. Now i 'm visiting your website"", jomblokeren (PSC Voter)"
    g_strFeedback(8) = """...Im sure this will help someone, as most of your work, Peter...."", Ole Chrisitian Spro (PSC Voter)"
    g_strFeedback(9) = """I like it very much... 5 From Me"", Jerous (PSC Voter)"

End Sub

