MIDARs HSV to RGB Converter
===========================
Instead of using VB's internal RGB function, use my HSV function. It is used in exactly the same way.


RGB Function (the old way)
==========================
Picture1.ForeColor = RGB(255,0,0)


HSV Function (the better way)
=============================
Picture1.ForeColor = HSV(0,1,1)


	LongReturnValue = HSV(Hue, Saturation, Lightness)

	Hue		0 to 360
	Saturation	0 to 1
	Lightness	0 to 1


Once you get used to it... you'll never look back!

- Peter
http://dev.midar.com/
