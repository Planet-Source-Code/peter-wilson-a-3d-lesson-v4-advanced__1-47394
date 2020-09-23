VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "MIDAR's Simple 3D Lesson"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MCI.MMControl MMControl1 
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   979
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer TimerMain 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   60
      Top             =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Define the name of this class/module for error-trap reporting.
Private Const m_strModuleName As String = "Form1"


' API used for reading the keyboard.
' ==================================
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


' =========================================================================================
' Virtual Camera: World Position, LookAt point, Tilt, FOV, Zoom, Near & Far Clipping plane.
' This is pretty comprehensive stuff!
' =========================================================================================
Private m_VirtualCamera As mdr3DCamera
Private m_VUP As mdrVector4


' m_Dots & m_Temp will hold an Array of Vectors (as defined in 'mMaths' module.)
' =================================================================================
Private m_Dots() As mdrVector4  ' << We define our Dots only once, and store them here.
Private m_Temp() As mdrVector4  ' << The original dots are transformed (by the camera code) into this temporary storage area.


' Misc. Settings / Temp Variables
' ===============================
Private m_lngCounter As Long

Public Sub DrawDots()

    ' ===============================
    ' Draws the Dots onto the screen.
    ' ===============================
    
    On Error GoTo errTrap
    
    Dim lngIndex As Long
    Dim PixelX As Single
    Dim PixelY As Single
    Dim sngDistance As Single
    
    Dim intBrightness As Integer
    Dim sngDeltaVisible As Single   ' Distance between the near and far clip distances.
    sngDeltaVisible = m_VirtualCamera.ClipFar - m_VirtualCamera.ClipNear
    
    
    ' Set the drawing style and width, etc.
    ' =====================================
    Me.DrawWidth = 1                    '   << Set the Width of the Pen. Any value higher than 1 will slow down animation.
    
    
    ' Loop through from the "Lower Boundry" of the Array, to the "Upper Boundry" of the Array.
    For lngIndex = LBound(m_Temp) To UBound(m_Temp)
                    
        PixelX = m_Temp(lngIndex).x
        PixelY = -m_Temp(lngIndex).y        ' Negated to make positive Y go up (and not down like Microsoft wants us to do... this is an aesthetics issue.)
        sngDistance = m_Temp(lngIndex).z    ' The Z coordinate, now represents the distance between the current Dot and the Camera.
        
        ' Only draw dots in front of the camera (and not behind us),
        ' but no further away than the Far clipping distance.
        ' ==========================================================
        If (sngDistance > m_VirtualCamera.ClipNear) And (sngDistance < m_VirtualCamera.ClipFar) Then
        
            ' Ignore Pixels that extend outside of the viewing window. Although the OS will pretty much do this
            ' for us 99% of the time, it fails with 'OverFlow Errors' the rest of the time when the OS tries to
            ' plot extreamly large pixel values... I consider this a Microsoft bug. Good on ya MS! :-p
            If (Abs(PixelX) < (1 / m_VirtualCamera.Zoom)) And (Abs(PixelY) < (1 / m_VirtualCamera.Zoom)) Then
            
            
                ' Shading dots is an important depth-cue (two methods)
                'intBrightness = CInt((sngDistance / sngDeltaVisible) * 360)
                intBrightness = 255 - CInt((sngDistance / sngDeltaVisible) * 255)
                
                
                ' Plot the point (two methods)
                'Me.PSet (PixelX, PixelY), HSV(intBrightness, 1,1)
                Me.PSet (PixelX, PixelY), RGB(intBrightness, intBrightness, intBrightness)
                
                
            End If ' Is Pixel within the window?
        End If ' Is the Pixel within the near & far clip values?
        
     Next lngIndex
    
    Exit Sub
errTrap:
    ' Just ignore any errors.
    
End Sub


Private Sub CreateTestData()

    ' =====================================================
    ' Create a nice big test grid (and some ground clutter)
    ' =====================================================
    
    Screen.MousePointer = vbHourglass
    
    Dim lngIndex As Long
    
    Dim intX As Integer
    Dim intY As Integer
    Dim intZ As Integer
    Dim tempVector As mdrVector4
    
    lngIndex = -1 ' Reset to -1, because we'll soon be increasing this value to 0 (the start of our array)
    
    ' ===========================================================================
    ' Create a random star field.
    ' Create a random vector, then normalize it (to give appearance of a sphere),
    ' then scale it to expand the sphere to be the distant stars.
    ' ===========================================================================
    For intY = 100 To 200 Step 100
        For intX = 0 To 300
            lngIndex = lngIndex + 1
            ReDim Preserve m_Dots(lngIndex)
            m_Dots(lngIndex).x = Rnd - 0.5
            m_Dots(lngIndex).y = Rnd - 0.5
            m_Dots(lngIndex).z = Rnd - 0.5
            m_Dots(lngIndex).w = 1
            
            ' Force dots into a sphere (Remark this out to see the difference, as this is a *excellent* way to visualize what normalizing does!)
            m_Dots(lngIndex) = VectorNormalize(m_Dots(lngIndex))
            
            ' Make the sphere bigger.
            m_Dots(lngIndex) = VectorMultiplyByScalar(m_Dots(lngIndex), CSng(intY))
            
        Next intX
    Next intY
    
''    ' ===============================================================
''    ' Create some random ground clutter (ie. grass blades, whatever?)
''    ' (If you are feeling adventurous, you might like to introduce
''    ' colour into this application to make the grass green.
''    ' ===============================================================
''    For intX = 0 To 100                             '   << Try increase the number of ground clutter dots
''        lngIndex = lngIndex + 1
''        ReDim Preserve m_Dots(lngIndex)
''        m_Dots(lngIndex).x = (Rnd * 100) - 50       '   << ie. Random number between -50 and +50
''        m_Dots(lngIndex).y = 0                      '   << Because this is the ground, the elevation is zero.
''        m_Dots(lngIndex).z = (Rnd * 100) - 50
''        m_Dots(lngIndex).w = 1
''    Next intX
    
    
    ' ====================================================================
    ' Create 3 large lines out of dots, representing the 3 axes (x, y & z)
    ' ====================================================================
    For intX = -100 To 100 Step 5                  '   << Positive X points to the Right
        For intZ = -100 To 100 Step 5              '   << Positive Z points *into* the monitor - away from you.
            For intY = -100 To 100 Step 5          '   << Positive Y goes Up

                If (intX = 0 And intY = 0) Or (intX = 0 And intZ = 0) Or (intY = 0 And intZ = 0) Then

                    lngIndex = lngIndex + 1
                    ReDim Preserve m_Dots(lngIndex)
                    m_Dots(lngIndex).x = intX
                    m_Dots(lngIndex).y = intY
                    m_Dots(lngIndex).z = intZ
                    m_Dots(lngIndex).w = 1

                End If
            Next intY
        Next intZ
    Next intX
    
    
    ' ===================
    ' Create a Test Grid.
    ' ===================
    For intX = -100 To 100 Step 20                  '   << Positive X points to the Right
        For intZ = -100 To 100 Step 20              '   << Positive Z points *into* the monitor - away from you.
            For intY = -100 To 100 Step 20          '   << Positive Y goes Up

                If (Abs(intX) = 100 Or Abs(intZ) = 100 Or intX = 0 Or intZ = 0) And (intY = 0 Or Abs(intY) = 100) Then

                    ' Create the basement (below ground level), floor and roof (of the test grid)
                    lngIndex = lngIndex + 1
                    ReDim Preserve m_Dots(lngIndex)
                    m_Dots(lngIndex).x = intX / 100
                    m_Dots(lngIndex).y = intY / 100
                    m_Dots(lngIndex).z = intZ / 100
                    m_Dots(lngIndex).w = 1

                ElseIf Abs(intX) = 100 And Abs(intZ) = 100 Then

                    ' Put some corners on it (ie. 4 support beams)
                    lngIndex = lngIndex + 1
                    ReDim Preserve m_Dots(lngIndex)
                    m_Dots(lngIndex).x = intX / 100
                    m_Dots(lngIndex).y = intY / 100
                    m_Dots(lngIndex).z = intZ / 100
                    m_Dots(lngIndex).w = 1

                End If
            Next intY
        Next intZ
    Next intX
    
'''    For intX = -100 To 100 Step 10       '   << Try adjusting the Step value to 1, 2, 5, 10 or 25.
'''        For intZ = -100 To 100 Step 10   '   << Try adjusting the Step value to 1, 2, 5, 10 or 25.
'''
'''            lngIndex = lngIndex + 1
'''            ReDim Preserve m_Dots(lngIndex)
'''            m_Dots(lngIndex).x = intX / 100                                 '   << Positive X points to the Right
'''            m_Dots(lngIndex).y = Cos(Sqr(intX * intX + intZ * intZ) / 30)   '   << Positive Y goes Up
'''            m_Dots(lngIndex).Z = intZ / 100                                 '   << Positive Z points *into* the monitor - away from you.
'''            m_Dots(lngIndex).W = 1
'''
'''        Next intZ
'''    Next intX
    
           
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub DrawCrossHairs()

    ' Draws cross-hairs going through the origin of the 2D window.
    ' ============================================================
    Me.DrawWidth = 1
    
    ' Draw Horizontal line (slightly darker to compensate for CRT monitors)
    Me.ForeColor = RGB(0, 32, 32)
    Me.Line (Me.ScaleLeft, 0)-(Me.ScaleWidth, 0)
    
    ' Draw Vertical line
    Me.ForeColor = RGB(0, 48, 48)
    Me.Line (0, Me.ScaleTop)-(0, Me.ScaleHeight)
    
End Sub

Private Sub DrawParameters(InDemo As Boolean)

    ' ==========================================================================================
    ' This routine slows down the program, because printing text is very slow.
    ' Speed has been sacrificed for instructional clarity for beginners to 3D Computer Graphics.
    ' Remember that by-and-large I am programming things the slow way, in an effort to be clear.
    ' You can always speed up my code by making your own clever adjustments.
    ' ==========================================================================================
    
    Dim sngX As Single
    Me.FontSize = 8
    Me.FontBold = False
    
    ' Set our start printing position
    ' Remember, The origin of our screen has been moved into the center of the window, but we want text top-left.
    Me.ForeColor = RGB(255, 255, 192)
    sngX = Me.ScaleLeft
    Me.CurrentY = Me.ScaleTop
    
    
    ' Show product name.
    Me.CurrentX = sngX
    Me.Print App.ProductName & " - " & App.LegalCopyright
    
    
    ' Show helpful reminders.
    If InDemo = False Then
        Me.CurrentX = sngX
        Me.Print "Keys:  ESC, Left/Right/Up/Down, Shift-Up/Down, Page-Up/Down, Spacebar. Modify Camera LookAt point in code: 'm_VirtualCamera.LookAtPoint'"
        
        ' Show helpful reminders.
        Me.CurrentX = sngX
        Me.Print "Mouse:  Move mouse over dots to display original defined coordinates." & vbNewLine
    End If
    
    ' Show current Camera position.
    Me.CurrentX = sngX
    Me.Print "Camera:  x: " & Format(m_VirtualCamera.WorldPosition.x, "Fixed") & "   y: " & Format(m_VirtualCamera.WorldPosition.y, "Fixed") & "   z: " & Format(m_VirtualCamera.WorldPosition.z, "Fixed")
    
    ' Show current LookAt point.
    Me.CurrentX = sngX
    Me.Print "LookAt:  x: " & Format(m_VirtualCamera.LookAtPoint.x, "Fixed") & "   y: " & Format(m_VirtualCamera.LookAtPoint.y, "Fixed") & "   z: " & Format(m_VirtualCamera.LookAtPoint.z, "Fixed")
    
    ' Show current VUP vector.
    Me.CurrentX = sngX
    Me.Print "VUP:  x: " & Format(m_VirtualCamera.VUP.x, "Fixed") & "   y: " & Format(m_VirtualCamera.VUP.y, "Fixed") & "   z: " & Format(m_VirtualCamera.VUP.z, "Fixed")
    
    ' Show current Camera Zoom value.
    Me.CurrentX = sngX
    Me.Print "Camera Zoom: " & Format(m_VirtualCamera.Zoom, "Fixed")
    
    ' Show current Field Of View value.
    Me.CurrentX = sngX
    Me.Print "Field Of View: " & Format(m_VirtualCamera.FOV, "Fixed") & "°" & vbNewLine
    
    
End Sub

Private Sub Init_MIDISoundtrack()

    ' ==========================================
    ' Init. Multimedia Control & Open MIDI File.
    ' ==========================================
    If Me.MMControl1.Enabled = True Then        ' Allow user to easily disable this item.
        Me.MMControl1.Visible = False
        Me.MMControl1.DeviceType = "Sequencer"
        Me.MMControl1.FileName = App.Path & "\2001 - A Space Odyssey.mid"
        Me.MMControl1.UpdateInterval = 1
        Me.MMControl1.Command = "OPEN"
'        Me.MMControl1.From = 45
        Me.MMControl1.Command = "PLAY"
    End If
    
End Sub

Public Sub Init_VirtualCamera()
    
    ' =========================================================
    ' Don't forget that...
    '   * Positive X points to the Right
    '   * Positive Z points *into* the monitor - away from you.
    '   * Positive Y goes Up
    ' =========================================================

    With m_VirtualCamera
    
        ' Fill in some comments (optional)
        .Caption = "Director's Chair"
        .Description = "This is the view as seen from the Director's chair."
        
        ' Reset the clipping distances.
        .ClipNear = 0
        .ClipFar = 500
        
        ' Reset the Zoom and FOV. Don't forget to call the form's Resize code after changing the Zoom value!
        .Zoom = 1
        .FOV = ConvertZoomtoFOV(.Zoom)
        
        With .WorldPosition
            .x = 0
            .y = 0
            .z = 2
            .w = 1
        End With
        
        With .LookAtPoint
            .x = 0
            .y = 0
            .z = 3
            .w = 1
        End With
        
        With .VUP
            .x = 0
            .y = 1
            .z = 0
            .w = 1
        End With
        
    End With
    
End Sub

Private Sub Form_Load()
    
    ' ========================================
    ' Initializes the random-number generator.
    ' ========================================
    Randomize
    
    
    ' ===============================
    ' Set some basic form properties.
    ' ===============================
    Me.AutoRedraw = True
    Me.ForeColor = RGB(255, 255, 255)
    Me.BackColor = RGB(0, 0, 0)
    Me.Font = "Arial"
    
    
    ' ===========================================================
    ' Init Feedback Array (Planet Source Code comments/feedback).
    ' ===========================================================
    Call InitPSCFeedback
    m_lngCounter = Int(Rnd * UBound(g_strFeedback))
    
    
    ' =======================================================================
    ' Create some test data. ie. Star Field, Ground Clutter, Major Axes, etc.
    ' =======================================================================
    Call CreateTestData
    
    
    ' =============================
    ' Initilize the Virtual Camera.
    ' =============================
    Call Init_VirtualCamera
    
    
    ' ======================================================================================
    ' Initilize the MIDI Soundtrack.
    ' This will also initilize the 'Sequencer' timer that will double-up as a Timer control.
    ' The Multimedia Control helps us Synchronize the Music to the action... very cool!
    ' ======================================================================================
    Call Init_MIDISoundtrack
    
    
    ' ================================================================
    ' Hide Mouse (by moving it to the far right)
    ' This method causes less problems than actually hiding the mouse!
    ' ================================================================
    Call SetCursorPos(Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' This routine slows down the program when the mouse is moved.
    ' ============================================================
    
    On Error GoTo errTrap
    
    Dim lngIndex As Long
    Dim PixelX As Single
    Dim PixelY As Single
    Dim sngAutoTolerance As Single
    
    If Me.TimerMain.Enabled = False Then Exit Sub
    
    Me.DrawWidth = 4
    Me.ForeColor = RGB(255, 255, 0)
    
    sngAutoTolerance = m_VirtualCamera.Zoom / 50
    
    ' Loop through from the "Lower Boundry" of the Array, to the "Upper Boundry" of the Array.
    For lngIndex = LBound(m_Temp) To UBound(m_Temp)
        
        ' Only consider dots in front of us.
        If m_Temp(lngIndex).z > 0 Then
            
            PixelX = m_Temp(lngIndex).x
            PixelY = -m_Temp(lngIndex).y
            
            ' Is the mouse close to the X coordinate?
            If Abs(PixelX - x) < sngAutoTolerance Then

                ' Is the mouse close to the Y coordinate?
                If Abs(PixelY - y) < sngAutoTolerance Then

                    ' Plot the pixel
                    Me.PSet (PixelX, PixelY)
                    Me.Print "x:" & Format(m_Dots(lngIndex).x, "Fixed") & " y:" & Format(m_Dots(lngIndex).y, "Fixed") & " z:" & Format(m_Dots(lngIndex).z, "Fixed")

                End If
            End If
        End If
     Next lngIndex
     
    Exit Sub
errTrap:
    
End Sub

Private Sub Form_Resize()

    ' Reset the width and height of our form, and also move the origin (0,0) into
    ' the centre of the form. This makes our life much easier.
    Dim sngAspectRatio As Single
    sngAspectRatio = Me.Width / Me.Height
    
    Me.ScaleWidth = 2 * (1 / m_VirtualCamera.Zoom)
    Me.ScaleLeft = -ScaleWidth / 2
    
    Me.ScaleHeight = 2 * (1 / m_VirtualCamera.Zoom) / sngAspectRatio
    Me.ScaleTop = -Me.ScaleHeight / 2
    
End Sub
Public Sub CalculateNewDotPositions()

    On Error GoTo errTrap
    
    Dim lngIndex As Long
    
    ReDim m_Temp(UBound(m_Dots))
    
    Dim matTilt As mdrMATRIX4
    Dim matView As mdrMATRIX4
    Dim vectVPN As mdrVector4 ' View Plane Normal (VPN) - The direction that the Virtual Camera is pointing.
    Dim vectVUP As mdrVector4 ' View UP direction (VUP) - Which way is up? This is used for tilting (or not tilting) the Camera.
    Dim vectVRP As mdrVector4 ' View Reference Point (VRP) - The World Position of the Virtual Camera.
    
    
    ' Subtract the Camera's world position from the 'LookAt' point to give us the View Plane Normal (VPN).
    vectVPN = VectorSubtract(m_VirtualCamera.LookAtPoint, m_VirtualCamera.WorldPosition)
    
    vectVRP = m_VirtualCamera.WorldPosition
    
    vectVUP = m_VirtualCamera.VUP
''    ' Calculate the Camera's Tilt value (if any)
''    Dim vectTemp As mdrVector4
''    matTilt = MatrixRotationZ(ConvertDeg2Rad(15))
''    vectVUP = MatrixMultiplyVector(matTilt, vectVUP)
    m_VUP = vectVUP
    
    
    ' Calculate the View Orientation Matrix.
    matView = MatrixViewOrientation(vectVPN, vectVUP, vectVRP)
    
    
    ' Loop through from the "Lower Boundry" of the Array, to the "Upper Boundry" of the Array.
    For lngIndex = LBound(m_Dots) To UBound(m_Dots)
    
        ' Apply the 'View Orientation Matrix' to the dots.
        m_Temp(lngIndex) = MatrixMultiplyVector(matView, m_Dots(lngIndex))
        
        
        ' ========================================================
        ' Transform the 3D vector down to a 2D vector by simply
        ' dividing the X and Y coordinates by their Z counterpart.
        ' ========================================================
        If m_Temp(lngIndex).z <> 0 Then
            m_Temp(lngIndex).x = m_Temp(lngIndex).x / m_Temp(lngIndex).z
            m_Temp(lngIndex).y = m_Temp(lngIndex).y / m_Temp(lngIndex).z
        End If
        
    Next lngIndex
    
    Exit Sub
errTrap:
    
End Sub


Public Sub UpdateCameraParameters()

    ' ===================================================================
    ' This routine looks at the keyboard, and adjusts the camera position
    ' and Zoom values depending on which keys are held down.
    ' ===================================================================
    
    Dim lngKeyState As Long
    Dim sngCameraStep As Single
    
    sngCameraStep = 0.5 ' << Adjust this to move the camera faster or slower (any value not zero)
    
    lngKeyState = GetKeyState(vbKeyControl)
    If (lngKeyState And &H8000) Then
    
        ' ======================================
        ' Move Camera's LookAt Point Left/Right.
        ' ======================================
        lngKeyState = GetKeyState(vbKeyLeft)
        If (lngKeyState And &H8000) Then m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x - sngCameraStep
        lngKeyState = GetKeyState(vbKeyRight)
        If (lngKeyState And &H8000) Then m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + sngCameraStep
    
    Else
    
        ' =======================
        ' Move Camera Left/Right.
        ' =======================
        lngKeyState = GetKeyState(vbKeyLeft)
        If (lngKeyState And &H8000) Then m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x - sngCameraStep
        lngKeyState = GetKeyState(vbKeyRight)
        If (lngKeyState And &H8000) Then m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + sngCameraStep
    
        lngKeyState = GetKeyState(vbKeyShift)
        If (lngKeyState And &H8000) Then
            
            ' ======================================================================
            ' Shift Key is down, the user must want to move closer, or further away.
            ' ======================================================================
            lngKeyState = GetKeyState(vbKeyUp)
            If (lngKeyState And &H8000) Then m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + sngCameraStep
            lngKeyState = GetKeyState(vbKeyDown)
            If (lngKeyState And &H8000) Then m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z - sngCameraStep
        
        Else
        
            ' =============================================
            ' Shift Key is *not* down. Move camera up/down.
            ' =============================================
            lngKeyState = GetKeyState(vbKeyUp)
            If (lngKeyState And &H8000) Then m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + sngCameraStep
            lngKeyState = GetKeyState(vbKeyDown)
            If (lngKeyState And &H8000) Then m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y - sngCameraStep
            
        End If ' Is Shift Key held down?
    End If ' Is Control Key held down?
    
    ' ==============================================================================================
    ' Modify the following:
    '   * Field Of View (FOV)
    '   * Camera's Zoom value
    '
    '   Note: These two values are pretty much the same thing, it depends on how you think about it.
    '         You could also think of this as the "Perspective Distortion" as well.
    '
    ' All of this is achieved simply by adjusting the height/width of the window.
    ' It might sound simple, but in reality this is pretty much what the complex 3D engines do.
    ' ==============================================================================================
    lngKeyState = GetKeyState(vbKeyPageUp)
    If (lngKeyState And &H8000) Then
        If m_VirtualCamera.Zoom > 0.05 Then
            m_VirtualCamera.Zoom = m_VirtualCamera.Zoom - 0.05
            m_VirtualCamera.FOV = ConvertZoomtoFOV(m_VirtualCamera.Zoom)
            Call Form_Resize                                                '   Redefine the Height/Width of our drawing window.
        End If
    End If
    lngKeyState = GetKeyState(vbKeyPageDown)
    If (lngKeyState And &H8000) Then
        m_VirtualCamera.Zoom = m_VirtualCamera.Zoom + 0.05
        m_VirtualCamera.FOV = ConvertZoomtoFOV(m_VirtualCamera.Zoom)
        Call Form_Resize                                                '   Redefine the Height/Width of our drawing window.
    End If
    
    
    ' ====================================
    ' Reset Camera to a starting position.
    ' ====================================
    lngKeyState = GetKeyState(vbKeySpace)
    If (lngKeyState And &H8000) Then
        ' Reset Camera
        m_VirtualCamera.WorldPosition.x = 0
        m_VirtualCamera.WorldPosition.y = 3
        m_VirtualCamera.WorldPosition.z = -15

        ' Reset the Camera's LookAt point.
        ' ================================
        m_VirtualCamera.LookAtPoint.x = 0
        m_VirtualCamera.LookAtPoint.y = 0
        m_VirtualCamera.LookAtPoint.z = 0
        
        ' Reset Zoom/FOV
        m_VirtualCamera.Zoom = 1
        m_VirtualCamera.FOV = ConvertZoomtoFOV(m_VirtualCamera.Zoom)
        Call Form_Resize
        
    End If
    
    
    ' ========================================
    ' Check for ESC / Quit / Exit Application.
    ' ========================================
    lngKeyState = GetKeyState(vbKeyEscape)
    If (lngKeyState And &H8000) Then
        ' Quit Application
        Me.TimerMain.Enabled = False
        Unload Me
    End If
    
    
End Sub

Private Sub MMControl1_StatusUpdate()
    
    ' Sorry, this routine is a little messy, and it might be hard for you to understand some of it.
    ' Basically I'm changing the Virtual Camera at important musical times.
    ' I usually tell the Camera where I would like it to be, then I calculate the difference between
    ' where the camera currently is, and where it should be, then move the camera accordingly.
    ' This method, is not exact and may produce slightly different results  depending on how fast your computer is.
    ' I really just wanted to experient with Syncronizing Music to Animation using the MMControl, that's all.
    '
    ' Cheers,
    ' peter@midar.com
    
    
    ' ===============================================
    ' "2001 - A Space Odyssey"  :   Major music hits.
    ' ===============================================
    '16
    '24
    '32
    '47
    '66
    '81
    '88
    '96
    '111
    '113
    '132
    '144
    '152
    '160
    '175
    '177
    '192
    '201
    '209
    '214
    '222
    '232
    '237
    '241
    '273
    
    Static blnToggle As Boolean
    Dim strMsg As String
    
    Dim lngPosition As Long
    lngPosition = Me.MMControl1.Position
    
    Static tempA As Single
    Static tempB As Single
    Dim matTilt As mdrMATRIX4
    Dim vectTemp As mdrVector4
    Dim vectTemp1 As mdrVector4
    Dim vectTemp2 As mdrVector4
        
    Me.Cls
    Me.ForeColor = RGB(255, 255, 255)
    
    If lngPosition < 16 Then
        Me.FontSize = 9
        Me.FontBold = True
        Me.FontItalic = True
        
        strMsg = g_strFeedback(m_lngCounter)
        
        Me.CurrentX = -Me.TextWidth(strMsg) / 2
        Me.CurrentY = -Me.TextHeight(strMsg) / 2
        Me.Print strMsg
        Me.FontItalic = False
    
        ' Calculate the Camera's Tilt value (if any)
        tempA = tempA + 0.1: tempB = tempB + 1
        vectTemp.x = 0: vectTemp.y = 1: vectTemp.z = 0: vectTemp.w = 1
        matTilt = MatrixRotationZ(ConvertDeg2Rad(tempA))
        m_VirtualCamera.VUP = MatrixMultiplyVector(matTilt, vectTemp)
        Call CalculateNewDotPositions
        Call DrawDots

    ElseIf lngPosition < 24 Then
        Me.FontSize = 72
        Me.FontBold = True
        
        strMsg = "Peter Wilson"
        Me.CurrentX = -Me.TextWidth(strMsg) / 2
        Me.CurrentY = -Me.TextHeight(strMsg) / 2
        Me.Print strMsg
        
        ' Calculate the Camera's Tilt value (if any)
        tempA = tempA + 0.1: tempB = tempB + 1
        vectTemp.x = 0: vectTemp.y = 1: vectTemp.z = 0: vectTemp.w = 1
        matTilt = MatrixRotationZ(ConvertDeg2Rad(tempA))
        m_VirtualCamera.VUP = MatrixMultiplyVector(matTilt, vectTemp)
        Call CalculateNewDotPositions
        Call DrawDots
        
    ElseIf lngPosition < 32 Then
        Me.FontSize = 26
        
        strMsg = "in association with"
        Me.CurrentX = -Me.TextWidth(strMsg) / 2
        Me.CurrentY = -Me.TextHeight(strMsg)
        Me.Print strMsg
        
        strMsg = "www.PlanetSourceCode.com"
        Me.CurrentX = -Me.TextWidth(strMsg) / 2
        Me.CurrentY = 0 'Me.TextHeight(strMsg)
        Me.Print strMsg
    
        ' Calculate the Camera's Tilt value (if any)
        tempA = tempA + 0.1: tempB = tempB + 1
        vectTemp.x = 0: vectTemp.y = 1: vectTemp.z = 0: vectTemp.w = 1
        matTilt = MatrixRotationZ(ConvertDeg2Rad(tempA))
        m_VirtualCamera.VUP = MatrixMultiplyVector(matTilt, vectTemp)
        Call CalculateNewDotPositions
        Call DrawDots
        
    ElseIf lngPosition < 47 Then
        Me.FontSize = 26
        strMsg = "presents"
        Me.CurrentX = -Me.TextWidth(strMsg) / 2
        Me.CurrentY = -Me.TextHeight(strMsg) / 2
        Me.Print strMsg
    
        ' Calculate the Camera's Tilt value (if any)
        tempA = tempA + 0.1: tempB = tempB + 1
        vectTemp.x = 0: vectTemp.y = 1: vectTemp.z = 0: vectTemp.w = 1
        matTilt = MatrixRotationZ(ConvertDeg2Rad(tempA))
        m_VirtualCamera.VUP = MatrixMultiplyVector(matTilt, vectTemp)
        Call CalculateNewDotPositions
        Call DrawDots
            
    ElseIf lngPosition < 48 Then
        Me.FontBold = True
        Me.FontSize = 36
        
        strMsg = "3D Computer Graphics"
        Me.CurrentX = -Me.TextWidth(strMsg) / 2
        Me.CurrentY = -Me.TextHeight(strMsg) * 2
        Me.Print strMsg
        
        strMsg = "for Visual Basic Programmers:"
        Me.CurrentX = -Me.TextWidth(strMsg) / 2
        Me.CurrentY = -Me.TextHeight(strMsg)
        Me.Print strMsg
        
        Me.FontBold = False
        Me.FontSize = 26
        
        strMsg = "Theory, Practice and Source Code."
        Me.CurrentX = -Me.TextWidth(strMsg) / 2
        Me.CurrentY = 0
        Me.Print strMsg
        
        ' Calculate the Camera's Tilt value (if any)
        tempA = tempA + 0.1: tempB = tempB + 1
        vectTemp.x = 0: vectTemp.y = 1: vectTemp.z = 0: vectTemp.w = 1
        matTilt = MatrixRotationZ(ConvertDeg2Rad(tempA))
        m_VirtualCamera.VUP = MatrixMultiplyVector(matTilt, vectTemp)
        Call CalculateNewDotPositions
        Call DrawDots
            
    ElseIf lngPosition < 66 Then
        Me.FontBold = True
        Me.FontSize = 36
        
        strMsg = "3D Computer Graphics"
        Me.CurrentX = -Me.TextWidth(strMsg) / 2
        Me.CurrentY = -Me.TextHeight(strMsg) * 2
        Me.Print strMsg
        
        strMsg = "for Visual Basic Programmers:"
        Me.CurrentX = -Me.TextWidth(strMsg) / 2
        Me.CurrentY = -Me.TextHeight(strMsg)
        Me.Print strMsg
        
        Me.FontBold = False
        Me.FontSize = 26
        
        strMsg = "Theory, Practice and Source Code."
        Me.CurrentX = -Me.TextWidth(strMsg) / 2
        Me.CurrentY = 0
        Me.Print strMsg
        
        Me.ForeColor = RGB(56, 56, 56)
        Me.FontSize = 36
        Me.FontBold = True
        strMsg = "v4.0"
        Me.CurrentX = Me.TextWidth(strMsg) * 3
        Me.CurrentY = Me.TextHeight(strMsg) / 2
        Me.Print strMsg
        
        ' Calculate the Camera's Tilt value (if any)
        tempA = tempA + 0.1: tempB = tempB + 1
        vectTemp.x = 0: vectTemp.y = 1: vectTemp.z = 0: vectTemp.w = 1
        matTilt = MatrixRotationZ(ConvertDeg2Rad(tempA))
        m_VirtualCamera.VUP = MatrixMultiplyVector(matTilt, vectTemp)
        Call CalculateNewDotPositions
        Call DrawDots
            
    ElseIf lngPosition < 80 Then
        
        ' Pan Down (so stars seem to fall)
        If blnToggle = False Then
            ' Reset tilt.
            m_VirtualCamera.VUP.x = 0
            m_VirtualCamera.VUP.y = 1
            m_VirtualCamera.VUP.z = 0
            
            m_VirtualCamera.WorldPosition.x = 0
            m_VirtualCamera.WorldPosition.y = 250
            m_VirtualCamera.WorldPosition.z = -50
            
            m_VirtualCamera.LookAtPoint.x = 0
            m_VirtualCamera.LookAtPoint.y = 250
            m_VirtualCamera.LookAtPoint.z = 0
            
            blnToggle = Not blnToggle
        Else
        
            vectTemp.x = 0
            vectTemp.y = 0
            vectTemp.z = -50
            vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
            m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.02)
            m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.02)
            m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.02)
            
            vectTemp2.x = 0
            vectTemp2.y = 0
            vectTemp2.z = 0
            vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
            m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.02)
            m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.02)
            m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.02)
        End If
        
        Call CalculateNewDotPositions
        Call DrawDots
        
    ElseIf lngPosition < 96 Then
    
        vectTemp.x = 0
        vectTemp.y = 0
        vectTemp.z = -150
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.02)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.02)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.02)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.02)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.02)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.02)
        
        Call CalculateNewDotPositions
        Call DrawDots
        
    ElseIf lngPosition < 111 Then
        
        vectTemp.x = -200
        vectTemp.y = 300
        vectTemp.z = -200
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.02)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.02)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.02)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.02)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.02)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.02)
    
        Call CalculateNewDotPositions
        Call DrawDots
        
    ElseIf lngPosition < 144 Then
    
        vectTemp.x = 200
        vectTemp.y = 300
        vectTemp.z = 200
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.01)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.02)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.02)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.02)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.02)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.02)
        
        Call CalculateNewDotPositions
        Call DrawDots
    
    ElseIf lngPosition < 152 Then

        vectTemp.x = 1
        vectTemp.y = 350
        vectTemp.z = -15
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.1)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.1)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.1)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.1)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.1)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.1)
        
        Call CalculateNewDotPositions
        Call DrawDots
        
    ElseIf lngPosition < 160 Then
    
        vectTemp.x = 1
        vectTemp.y = 250
        vectTemp.z = -15
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.07)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.07)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.07)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.1)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.1)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.1)
        
        Call CalculateNewDotPositions
        Call DrawDots
    
    ElseIf lngPosition < 175 Then
    
        vectTemp.x = 1
        vectTemp.y = 150
        vectTemp.z = -15
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.05)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.05)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.05)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.1)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.1)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.1)
        
        Call CalculateNewDotPositions
        Call DrawDots
    
    ElseIf lngPosition < 209 Then
    
        vectTemp.x = 10
        vectTemp.y = 5
        vectTemp.z = -15
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.02)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.01)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.01)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.1)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.1)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.1)
        
        Call CalculateNewDotPositions
        Call DrawDots
    
    ElseIf lngPosition < 222 Then
    
        vectTemp.x = 3
        vectTemp.y = 3
        vectTemp.z = -3
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.02)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.01)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.01)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.1)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.1)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.1)
        
        Call CalculateNewDotPositions
        Call DrawDots
        
    ElseIf lngPosition < 232 Then
    
        vectTemp.x = 5
        vectTemp.y = 0
        vectTemp.z = 0
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.02)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.01)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.01)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.1)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.1)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.1)
        
        Call CalculateNewDotPositions
        Call DrawDots
        
    ElseIf lngPosition < 237 Then
    
        vectTemp.x = 0
        vectTemp.y = -5
        vectTemp.z = -1
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.02)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.02)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.02)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.1)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.1)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.1)
        
        Call CalculateNewDotPositions
        Call DrawDots
        
    ElseIf lngPosition < 241 Then
    
        vectTemp.x = -5
        vectTemp.y = 0
        vectTemp.z = -1
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.02)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.02)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.02)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.1)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.1)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.1)
        
        Call CalculateNewDotPositions
        Call DrawDots
        
    ElseIf lngPosition < 273 Then
            
        vectTemp.x = 0
        vectTemp.y = 3
        vectTemp.z = -15
        vectTemp = VectorSubtract(vectTemp, m_VirtualCamera.WorldPosition)
        m_VirtualCamera.WorldPosition.x = m_VirtualCamera.WorldPosition.x + (vectTemp.x * 0.008)
        m_VirtualCamera.WorldPosition.y = m_VirtualCamera.WorldPosition.y + (vectTemp.y * 0.008)
        m_VirtualCamera.WorldPosition.z = m_VirtualCamera.WorldPosition.z + (vectTemp.z * 0.008)
        
        vectTemp2.x = 0
        vectTemp2.y = 0
        vectTemp2.z = 0
        vectTemp2 = VectorSubtract(vectTemp2, m_VirtualCamera.LookAtPoint)
        m_VirtualCamera.LookAtPoint.x = m_VirtualCamera.LookAtPoint.x + (vectTemp2.x * 0.05)
        m_VirtualCamera.LookAtPoint.y = m_VirtualCamera.LookAtPoint.y + (vectTemp2.y * 0.05)
        m_VirtualCamera.LookAtPoint.z = m_VirtualCamera.LookAtPoint.z + (vectTemp2.z * 0.05)
        
        
        Call CalculateNewDotPositions
        Call DrawDots
    Else
        
        m_VirtualCamera.VUP.x = 0
        m_VirtualCamera.VUP.y = 1
        m_VirtualCamera.VUP.z = 0
        
        m_VirtualCamera.WorldPosition.x = 0
        m_VirtualCamera.WorldPosition.y = 3
        m_VirtualCamera.WorldPosition.z = -15
        
        m_VirtualCamera.LookAtPoint.x = 0
        m_VirtualCamera.LookAtPoint.y = 0
        m_VirtualCamera.LookAtPoint.z = 0
        
        Me.MMControl1.UpdateInterval = 0
        Me.MMControl1.Enabled = False
        Me.TimerMain.Enabled = True
    End If
    
    
    
ExitSub:
    ' ======================
    ' Check for ESC of Demo.
    ' ======================
    Dim lngKeyState As Long
    lngKeyState = GetKeyState(vbKeyEscape)
    If (lngKeyState And &H8000) Then
        Me.MMControl1.Command = "Next"
    End If

End Sub


Private Sub TimerMain_Timer()

    ' Apply Virtual Camera to original 'Dots'
    Call CalculateNewDotPositions
    
    ' Draw Stuff
    ' ==========
    Me.Cls
    Call DrawCrossHairs
    Call DrawDots
    Call DrawParameters(False)
    
    ' Process keyboard commands
    Call UpdateCameraParameters
    
End Sub

