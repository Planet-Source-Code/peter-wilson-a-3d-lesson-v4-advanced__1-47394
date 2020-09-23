=======
Forward
=======
At the time of writing, I would like to thank the following users for their votes and their feedback:
	Jonathan D 
	Carles P.V.
	Carlos Bomtempo
	Eugene Wolff
	Ole Chrisitian Spro
	dafhi
	shadow8883
	Josh Nixon
	jveracrus
	Christopher Brim
	RPG MAKER
Without people like this, I never would have created this project or improved it. In my humble opinion, it's people such as these that keep PlanetSourceCode alive! Well done to you!


=======
Preface
=======
This is version three of my 3D lessons. If this version is too complicated for you, I strongly suggest you download my earlier PlanetSourceCode submission listed below:

	"A 3D Lesson v2, Simple"
	http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=47100&lngWId=1

	^^^ This is where you should start if you a beginner to 3D Computer Graphics ^^^


================================
Overview: 3D Lesson v4, Advanced
================================
*** Turn up your speakers! ***
The major update in version 4 is the synchronization of MIDI music and 3D animation. I also wanted to experiment with some opening credits at the beginning. You can press the ESC key to skip the animation.

The other major change is the various Camera variables have been cleaned up and placed into their own data type: 'mdr3DCamera'.  You might also notice that I've created some additional data types (ie. mdr3DObject, mdrPolyhedron & mdrVertex). This is in anticipation of getting rid of the dots, and creating some lines drawings... should be exciting!

There hasn't been too many changes to the core 3D maths; most of it is fairly stable and probably won't change much in future versions. If you see other programmers placing their Sines and Cosines in different positions (like the columns and rows have been swapped over - ie. Transposed) then this probably means that they have coded all of their algorithims to a different "notation standard". This application follows the conventions used in the ledgendary bible "Computer Graphics Principles and Practice", Foley·vanDam·Feiner·Hughes which illustrates mathematical formulas using Column-Vector notation. Other books like "3D Math Primer for Graphics and Game Development", Fletcher Dunn·Ian Parberry, use Row-Vector notation. Both are correct, however it's important to know which standard you code to, because it affects the way in which you build your matrices and the order in which you should multiply them to obtain the correct result. OpenGL uses Column Vectors (like this application). DirectX uses Row Vectors. This is often a source of confusion for newbie 3D programmers.

Scroll down to the history section below.


Keyboard & Mouse Controls
-------------------------
Move Camera Left/Right	=	Left and Right arrow keys.
Move Camera Up/Down	=	Up and Down arrow keys.
Move Camera In/Out	=	SHIFT-Up / SHIFT-Down

Zoom Camera In/Out	=	Page-Up / Page-Down
(Also known as 'Perspective Distortion' and/or 'Field Of View')

Mouse Move		=	Hi-lights Dots that are close by	(optional routine - slows down program)

Reset Camera		=	Space Bar
Quit Application	=	Esc Key


===============
Version History
===============
------------------------------------------------
Version 1.0 initial release to PSC on 22-July-03
------------------------------------------------
	* Displays 3D Dots. Camera can be moved left/right, up/down, in/out.
	* Simple application of division to produce 3D:
		PixelX = X/Z
		PixelY = Y/Z


---------------------------------------------
Version 2.0 minor update to PSC on 24-July-03
---------------------------------------------
	* In version 1.0, it wasn't immediately clear which way X, Y and Z pointed. Version 2.0 cleared this up considerably.
	* Positive X points to the Right
	* Positive Z points *into* the monitor - away from You.
	* Positive Y goes Up

           +y
            |   +z (away from you - into the monitor)
            |  / 
            | /
	    |/
-x  --------+--------  +x
           /|
          / |
         /  |
       -z   |
	   -y

	* The MS Operating Systems has Y=0 at the top of the monitor, with increasing values of Y going down towards the bottom. However, in this application I wanted the Origin (0,0,0) to be in the centre of the screen, and for Y to go up so I flipped the sign from + to - (see subroutine: DrawDots)
	* Version 2.0 allows you to move the mouse, and have the Dots hi-lighted that a close to the mouse. This is just for fun, and it also slows down the program... but I just wanted to show you how to do it.


---------------------------------------------
Version 3.0 major update to PSC on 26-July-03
---------------------------------------------
	The major update in version 3, is the Virtual Camera code. You can move the camera anywhere (using the keyboard), and make the camera 'look at' a certain point. To add more realism, Near and Far Clipping values have been incorporated. Only Dots between the Near and Far Clipping distances are visible. The benefit of this, is that the Dots can be shaded, depending on their distance from the Camera. This 3D lesson includes some very advanced 'mathematical' topics, however I have listed the code as 'Moderate' or 'Intermediate' because you don't need to understand the advanced parts to have fun with this project. Besides, I have a much more complicated 3D project up my sleeve that will be listed as 'Advanced'. Drawing polygons adds extra code, and I wanted this lesson to be a simple as possible. For this reason there are no line-drawings, triangles, polyhedra or polygons - just Dots. Besides, pixels are a lot faster to plot than lines - so we can have move of them!

	* New Feature:	Major update to include "Virtual Camera" code.
	* New Feature:	Matrix Multiplication is now used to move, rotate and tilt the camera.
	* New Feature:	Demo animation at start (can press ESC to cancel it)
	* New Feature:	Near and Far clipping distances can now be set. Can also shade dots depending on their distance from Camera.


---------------------------------------------
Version 3.1 minor update to PSC on 26-July-03
---------------------------------------------
	* Bug Fix: 	Fixed 'Overflow Errors' when attempting to plot outside of the viewing window.
	* Improvement:	Cleaned up code, added more comments.
	* Improvement:	Made animation faster (mainly by reducing the number of dots), and also made it more interesting.
	* Improvement:	Sped up pixel drawing routine.
	* Improvement:	Added 'Field Of View' code to supplement the Zoom values used in earlier versions (See diagram below)
	
	(top view - of FOV diagram)

        +z
	 |
 a	 |       b
  \	 |      /
   \--c--|--d--/
    \	 |    /
     \   |   /
      \  |  /
       \ | /
-x _____\|/_______ +x
	 |
	 |
        -z

	* c & d represent the window where the dots get drawn. c is the left side of the screen, d is the right.
	* The angle between a and b, is the Field Of View. In the above ASCII-art diagram, this angle is about 60°. (I measured it on screen with a protractor.)
		60 degrees gives us a Zoom value of 1.73.
		
		Q. How did I work this out?

		Zoom = 1 / Tan(FOV/2)	Note: FOV should be specified in Radians not Degrees.
			or
		FOV = 2 * ATan(1/Zoom)	Note: Convert the result from Radians to Degrees.

	* Zoom is directly related to FOV. You specific one of them, then work out the other. 'Conversion' functions have been provided to do this for you.


---------------------------------------------
Version 4 Major update to PSC on XX-August-03
---------------------------------------------
	* Improvement:  Added Multimedia Control and used this to syncronize MIDI music to the animation and credits.
	* Improvement:	Moved separate camera variables into a single Camera object. (see: mDataStructures.mdr3DCamera)
	* Improvement:	Added several new functions to help rotate objects on the X, Y and Z axes.
	* Improvement:	Form Resize code now correctly takes into account the Aspect Ratio of the window.
	* Note:		Camera Tilting is *not* intuitive, although it is in there and does work.



Cheers,

Peter Wilson
peter@midar.com
http://dev.midar.com/
