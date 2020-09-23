Attribute VB_Name = "mDataStructures"
Option Explicit


' =========================================================================================
' For this application we only need low precision values of PI; ie. The "Single" data type.
' =========================================================================================
Public Const g_sngPI As Single = 3.141594!
Public Const g_sngPIDivideBy180 As Single = 0.0174533!
Public Const g_sng180DivideByPI As Single = 57.29578!


' =========================================================================
' This is a 4 dimensional vector, because it holds 4 values (X, Y, Z & W).
' Now you understand multi-dimensional vectors. See?, Vectors are not hard!
' =========================================================================
Public Type mdrVector4
    x As Single
    y As Single
    z As Single
    w As Single         ' Named 'w' because we ran out of letters! w is not often used (so you can optimise lots of code because of this.)
End Type


' =======================================
' A 4x4 Matrix - RC stands for RowColumn.
' =======================================
Public Type mdrMATRIX4
    rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
    rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
    rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
    rc41 As Single: rc42 As Single: rc43 As Single: rc44 As Single
End Type


' ============================================================
' Vertices are the simplest building blocks in 3D.
' ie.
'    A 2D triangle is made up from only 3 Vertices.
'    A 3D pyramid is made up from 5 Vertices.
'    A 3D cube is made up from 8 Vertices.
' Remember, a vertex is just a single point (or dot) in space.
' ============================================================
Public Type mdrVertex
    Vertex As mdrVector4
    
    RGB_Red As Single                   '   0 to 1 (I suppose this could be 0->255, I just prefer 0->1)
    RGB_Green As Single                 '   0 to 1
    RGB_Blue As Single                  '   0 to 1
End Type


' ===========================================================================
' A Polyhedron is a solid object contained by many faces. Typically these are
' the sub-parts of your 3D object. They can usually be rotated separately.
' ===========================================================================
Public Type mdrPolyhedron
    Caption As String                   '   Helicopter Blades, Landing Gear, Gun Turret, Leg, Head, Arm, etc. (Optional)
    Description As String               '   A Caption should always have a Description. (Optional)
    
    RGB_Red As Single                   '   0 to 1
    RGB_Green As Single                 '   0 to 1
    RGB_Blue As Single                  '   0 to 1
    
    Vertices() As mdrVertex             '   The original vertices that make up the object (these never changed once defined)
    VerticesT() As mdrVertex            '   The transformed vertices; a temporary working area.
    Faces() As Variant                  '   Connect the dots [Vertices] together to form shapes.
    
    PointingAt As mdrVector4            '   Defines the direction the object is pointing. (Alternative to Pitch, Roll & Yaw.)
    Pitch As Single                     '   Angle in degrees
    Roll As Single                      '   Angle in degrees
    Yaw As Single                       '   Angle in degrees
    
    IdentityMatrix As mdrMATRIX4        '   This holds the initial or default starting position for the polyhedron (rotation, size & position). (Optional)
End Type


' ======================================================================
' A 3D object is usually a collection of smaller objects (ie. Polyhedra)
' ======================================================================
Public Type mdr3DObject
    Caption As String                   '   Helicopter, Tank, Space Ship, Monster, etc. (Optional)
    Description As String               '   A Caption should always have a Description. (Optional)
        
    WorldPosition As mdrVector4         '   Position of the Object in World Coordinates.
    PointingAt As mdrVector4            '   Defines the direction the object is pointing. (Alternative to Pitch, Roll & Yaw.)
    Pitch As Single                     '   Angle in degrees
    Roll As Single                      '   Angle in degrees
    Yaw As Single                       '   Angle in degrees
    UniformScale As Single              '   Uniform scale on all axes. Typically equals 1.0
    
    Vector As mdrVector4                '   Direction and Magnitude of the 3D Object. ie. Which way is the object moving, and how fast?
    
    Polyhedra() As mdrPolyhedron        '   This object is made up from Polyhedra.
End Type


' =====================================
' This is our Virtual 3D Camera object.
' =====================================
Public Type mdr3DCamera
    Caption As String                   '   Camera1, Director's Chair, Birds-eye View, etc. (Optional)
    Description As String               '   A Caption should always have a Description.     (Optional)
    
    WorldPosition As mdrVector4         '   Position of the Camera in World Coordinates.
    LookAtPoint As mdrVector4           '   This is where the Camera is looking at in World Coordinates.
    VUP As mdrVector4                   '   Which way is UP?
    
    FOV As Single                       '   Field Of View (FOV). "90 degree FOV" = "1x Zoom". If you update FOV, don't forget to update Zoom.
    Zoom As Single                      '   (The Zoom value is calculated from the FOV. Normally you define one of them, then calculate the other.)
    
    ClipFar As Single                   '   Don't draw Vertices further away than this value. Any value higher than 0.
    ClipNear As Single                  '   Don't draw Vertices that are this close to us (or behind us). Typically 0, but can be higher.
End Type

