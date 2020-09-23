<div align="center">

## Mouse manipulation


</div>

### Description

At some point, you may find it useful to manipulate the location of the mouse cursor. Perhaps you are

designing an interactive tutorial, a walkthrough, or maybe you plan on controlling another application

through mouse events. Regardless, you will quickly find a number of hurdles to overcome and it is the

goal of this example to help you over, under, or around these hurdle
 
### More Info
 
First, a little bit of background about the mouse in general. Since the days of DOS, mouse drivers have

reported their location in graphic applications by returning X/Y coordinates based up on a resolution

independent coordinate system. This coordinate system neatly breaks the screen down into 65535

units on each axis. A unit of measurement in this system is known as a Mickey. This system was

devised to insure that the specification for mouse drivers would be a lasting one, and that screen

resolution would never overtake the resolution of the mouse driver.

Why mention this? Well, the Win32 API function call which allows you to specify the location for the

mouse wants the location provided in mickeys. And the first hurdle to overcome is converting screen

coordinates to mouse coordinates.

In order to make the coversion, we first need to get the screen's height and width with

GetSystemMetrics. The GetScreenRes subroutine illustrates how this is done.

Once the resolution of the display is known, we can convert the pixels returned by GetScreenRes into

mickeys. There are four conversion routines included with this example, two to handle pixel

conversions to mickeys (PixelXToMickey, PixelYToMickey), and two to handle mickey to pixel

conversions (MickeyXToPixel, MickeyYToPixel).

Now that we have conversion routines, we can actually do some work. Included with this example is

CenterMouseOn, a function that will center the mouse cursor on anything that has an hWnd. An

example of using this function to put the cursor over a commandbutton appears as:

CenterMouseOn (command1.hWnd)

If you need to move the mouse but don't have an hWnd to reference, the MouseMove function will

allow you to specify an X/Y coordinate for the mouse cursor. And once it is moved, you can use the

MouseFullClick function to simulate a mouseclick.

There are a series of mouse coordinate to screen coordinate routines included in this example. Due to

rounding problems, it is quite likely that the calculations may be off by a pixel. If your application

requires extremely precise pointer placement, you may want to develop or look for a more precise

method.

One of the uglier portions of this code are the mickey to pixel routines. They use a series of

temporary singles to store values prior to being converted. This was done to improve the accuracy of

the conversion, but even so, rounding errors continue to creep in. If you know of a better, more

accurate way to accomplish the same task, I would appreciate hearing about it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tim's VB 5 tips and tricks](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tim-s-vb-5-tips-and-tricks.md)
**Level**          |Unknown
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tim-s-vb-5-tips-and-tricks-mouse-manipulation__1-450/archive/master.zip)

### API Declarations

```
' ----------------------------------------------
     ' *    MouseEvent Related Declares     *
     ' ----------------------------------------------
     Private Const MOUSEEVENTF_ABSOLUTE = &H8000
     Private Const MOUSEEVENTF_LEFTDOWN = &H2
     Private Const MOUSEEVENTF_LEFTUP = &H4
     Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
     Private Const MOUSEEVENTF_MIDDLEUP = &H40
     Private Const MOUSEEVENTF_MOVE = &H1
     Private Const MOUSEEVENTF_RIGHTDOWN = &H8
     Private Const MOUSEEVENTF_RIGHTUP = &H10
     Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
       ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, _
       ByVal dwExtraInfo As Long)
     ' ----------------------------------------------
     ' *   GetSystemMetrics Related Declares   *
     ' ----------------------------------------------
     Private Const SM_CXSCREEN = 0
     Private Const SM_CYSCREEN = 1
     Private Const TWIPS_PER_INCH = 1440
     Private Const POINTS_PER_INCH = 72
     Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex _
       As Long) As Long
     ' ----------------------------------------------
     ' *    GetWindowRect Related Declares    *
     ' ----------------------------------------------
     Private Type RECT
         Left As Long
         Top As Long
         Right As Long
         Bottom As Long
     End Type
     Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
       lpRect As RECT) As Long
     ' ----------------------------------------------
     ' *    Internal Constants and Types     *
     ' ----------------------------------------------
     Private Const MOUSE_MICKEYS = 65535
     Public Enum enReportStyle
       rsPixels
       rsTwips
       rsInches
       rsPoints
     End Enum
     Public Enum enButtonToClick
       btcLeft
       btcRight
       btcMiddle
     End Enum
```


### Source Code

```

     ' Returns the screen size in pixels or, optionally,
     ' in others scalemode styles
     Public Sub GetScreenRes(ByRef X As Long, ByRef Y As Long, Optional ByVal _
       ReportStyle As enReportStyle)
       X = GetSystemMetrics(SM_CXSCREEN)
       Y = GetSystemMetrics(SM_CYSCREEN)
       If Not IsMissing(ReportStyle) Then
         If ReportStyle <> rsPixels Then
           X = X * Screen.TwipsPerPixelX
           Y = Y * Screen.TwipsPerPixelY
           If ReportStyle = rsInches Or ReportStyle = rsPoints Then
             X = X \ TWIPS_PER_INCH
             Y = Y \ TWIPS_PER_INCH
             If ReportStyle = rsPoints Then
               X = X * POINTS_PER_INCH
               Y = Y * POINTS_PER_INCH
             End If
           End If
         End If
       End If
     End Sub
     ' Convert's the mouses coordinate system to
     ' a pixel position.
     Public Function MickeyXToPixel(ByVal mouseX As Long) As Long
       Dim X As Long
       Dim Y As Long
       Dim tX As Single
       Dim tmouseX As Single
       Dim tMickeys As Single
       GetScreenRes X, Y
       tX = X
       tMickeys = MOUSE_MICKEYS
       tmouseX = mouseX
       MickeyXToPixel = CLng(tmouseX / (tMickeys / tX))
     End Function
     ' Converts mouse Y coordinates to pixels
     Public Function MickeyYToPixel(ByVal mouseY As Long) As Long
       Dim X As Long
       Dim Y As Long
       Dim tY As Single
       Dim tmouseY As Single
       Dim tMickeys As Single
       GetScreenRes X, Y
       tY = Y
       tMickeys = MOUSE_MICKEYS
       tmouseY = mouseY
       MickeyYToPixel = CLng(tmouseY / (tMickeys / tY))
     End Function
     ' Converts pixel X coordinates to mickeys
     Public Function PixelXToMickey(ByVal pixX As Long) As Long
       Dim X As Long
       Dim Y As Long
       Dim tX As Single
       Dim tpixX As Single
       Dim tMickeys As Single
       GetScreenRes X, Y
       tMickeys = MOUSE_MICKEYS
       tX = X
       tpixX = pixX
       PixelXToMickey = CLng((tMickeys / tX) * tpixX)
     End Function
     ' Converts pixel Y coordinates to mickeys
     Public Function PixelYToMickey(ByVal pixY As Long) As Long
       Dim X As Long
       Dim Y As Long
       Dim tY As Single
       Dim tpixY As Single
       Dim tMickeys As Single
       GetScreenRes X, Y
       tMickeys = MOUSE_MICKEYS
       tY = Y
       tpixY = pixY
       PixelYToMickey = CLng((tMickeys / tY) * tpixY)
     End Function
     ' The function will center the mouse on a window
     ' or control with an hWnd property. No checking
     ' is done to ensure that the window is not obscured
     ' or not minimized, however it does make sure that
     ' the target is within the boundaries of the
     ' screen.
     Public Function CenterMouseOn(ByVal hwnd As Long) As Boolean
       Dim X As Long
       Dim Y As Long
       Dim maxX As Long
       Dim maxY As Long
       Dim crect As RECT
       Dim rc As Long
       GetScreenRes maxX, maxY
       rc = GetWindowRect(hwnd, crect)
       If rc Then
         X = crect.Left + ((crect.Right - crect.Left) / 2)
         Y = crect.Top + ((crect.Bottom - crect.Top) / 2)
         If (X >= 0 And X <= maxX) And (Y >= 0 And Y <= maxY) Then
           MouseMove X, Y
           CenterMouseOn = True
         Else
           CenterMouseOn = False
         End If
       Else
         CenterMouseOn = False
       End If
     End Function
     ' Simulates a mouse click
     Public Function MouseFullClick(ByVal MBClick As enButtonToClick) As Boolean
       Dim cbuttons As Long
       Dim dwExtraInfo As Long
       Dim mevent As Long
       Select Case MBClick
         Case btcLeft
           mevent = MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP
         Case btcRight
           mevent = MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP
         Case btcMiddle
           mevent = MOUSEEVENTF_MIDDLEDOWN Or MOUSEEVENTF_MIDDLEUP
         Case Else
           MouseFullClick = False
           Exit Function
       End Select
       mouse_event mevent, 0&, 0&, cbuttons, dwExtraInfo
       MouseFullClick = True
     End Function
     Public Sub MouseMove(ByRef xPixel As Long, ByRef yPixel As Long)
       Dim cbuttons As Long
       Dim dwExtraInfo As Long
       mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, _
         PixelXToMickey(xPixel), PixelYToMickey(yPixel), cbuttons, dwExtraInfo
     End Sub
```

