'***************************************************************************************
Option Explicit
#Const blnDeveloperMode = False
Private Const strModuleName As String = "ROWANIMATION"
'**** Author : Robert M Kreegier, Internet
'**** Purpose: Easing Functions for animations
'**** Source : The easing functions themselves have been copied and retooled from jQuery:
'****           https://github.com/danro/jquery-easing/blob/master/jquery.easing.js
'**** Notes  : t: current time, b: beginning value, c: change In value, d: duration
'***************************************************************************************

'***************************************************************************************************
' Author : Robert Kreegier
' Purpose: Set the height of the rows to a new height.
' Params :
'   rngRows         A Range of rows to set the height of.
'   dblHeight       The height to set the rows to.
'   lngTime         Gives the amount of time that the animation should take.
'                   If lngTime is zero, then there is no animation and the change happens instantly
'   strEasing       Optionally an easing function can be named to define how the animation looks.
'                   See the module "EASING_FUNCTIONS" as a reference.
'***************************************************************************************************
Sub SetRowHeight(ByRef rngRows As Range, ByVal dblHeight As Double, ByVal lngTime As Long, Optional ByVal strEasing As String = "easeInOutSine")
    #If Not blnDeveloperMode Then
        On Error GoTo Whoa
    #End If
    '***********************************************************************************************
    
    If lngTime > 0 And GetVar("V_SHOW_ANIMATIONS").Value Then
        Dim blnUpdating As Boolean: blnUpdating = Application.ScreenUpdating
        Dim lngCalc As Long: lngCalc = Application.Calculation
        Dim blnEvent As Boolean: blnEvent = Application.EnableEvents
        
        Dim lngStartTick As Long: lngStartTick = GetTickCount
        Dim lngEndTick As Long: lngEndTick = GetTickCount + lngTime
        Dim dblDistance As Double: dblDistance = dblHeight - rngRows.Cells(1).RowHeight
        Dim dblStartHeight As Double: dblStartHeight = rngRows.Cells(1).RowHeight
        
        ' When this function is called, things are bound to be suppressed, but we need ScreenUpdating to be active
        ' so we can see the animation. So we save the old value of ScreenUpdating, change it, then revert it
        ' back to its old value at the end of the procedure.
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        'Application.DisplayStatusBar = False
        
        ' In Excel 2010, we could do an animation without using DoEvents. Later versions require the screen to be updated each iteration of the animation, so they require DoEvents...however, it makes the animation slow
        If Application.Version > 14 Then
            While GetTickCount <= lngEndTick
                rngRows.RowHeight = Application.Run(strEasing, GetTickCount - lngStartTick, dblStartHeight, dblDistance, lngTime)
                DoEvents
            Wend
        Else
            While GetTickCount <= lngEndTick
                rngRows.RowHeight = Application.Run(strEasing, GetTickCount - lngStartTick, dblStartHeight, dblDistance, lngTime)
            Wend
        End If
        
        ' final step
        If dblHeight >= 0 Then
            rngRows.RowHeight = dblHeight
        End If
        
        If dblHeight > 0 Then
            rngRows.EntireRow.AutoFit
        End If
        
        ' return everything to the way it was
        Application.ScreenUpdating = blnUpdating
        Application.Calculation = lngCalc
        Application.EnableEvents = blnEvent
        'Application.DisplayStatusBar = True
    Else
        If dblHeight >= 0 Then
            rngRows.RowHeight = dblHeight
        End If
    End If
    
    '***********************************************************************************************
Letscontinue:
    Exit Sub
Whoa:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": SetRowHeight", True
        Resume Letscontinue
    #End If
End Sub

Function easeInQuad(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    easeInQuad = C * (t / d) ^ 2 + B
End Function

Function easeOutQuad(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    easeOutQuad = -C * (t / d) * ((t / d) - 2) + B
End Function

'    easeInOutQuad: function (x, t, b, c, d) {
'        if ((t/=d/2) < 1) return c/2*t*t + b;
'        return -c/2 * ((--t)*(t-2) - 1) + b;
'    },
'' doesn't work completely
'Function easeInOutQuad(ByVal t, ByVal b, ByVal c, ByVal d) As Double
'    If ((t / d) / 2) < 1 Then
'        easeInOutQuad = c / 2 * (t / d) ^ 2 + b
'    End If
'
'    easeInOutQuad = -c / 2 * (((t / d) - 1) * (((t / d) - 1) - 2) - 1) + b
'End Function

Function easeInCubic(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    easeInCubic = C * (t / d) ^ 3 + B
End Function

Function easeOutCubic(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    easeOutCubic = C * ((t / d - 1) ^ 3 + 1) + B
End Function

Function easeInQuart(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    easeInQuart = C * (t / d) ^ 4 + B
End Function

Function easeOutQuart(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    easeOutQuart = -C * ((t / d - 1) ^ 4 - 1) + B
End Function

Function easeInQuint(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    easeInQuint = C * (t / d) ^ 5 + B
End Function

Function easeOutQuint(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    easeOutQuint = -C * ((t / d - 1) ^ 5 - 1) + B
End Function

Function easeInSine(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    easeInSine = -C * Math.Cos(t / d * (WorksheetFunction.PI / 2)) + C + B
End Function

Function easeOutSine(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    easeOutSine = C * Math.Sin(t / d * (WorksheetFunction.PI / 2)) + B
End Function

Function easeInOutSine(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    easeInOutSine = -C / 2 * (Math.Cos(WorksheetFunction.PI * t / d) - 1) + B
End Function

Function easeInExpo(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    If t = 0 Then
        easeInExpo = B
    Else
        easeInExpo = C * (2 ^ (10 * (t / d - 1))) + B
    End If
End Function

Function easeOutExpo(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    If t = d Then
        easeOutExpo = B + C
    Else
        easeOutExpo = C * (-(2 ^ (-10 * t / d)) + 1) + B
    End If
End Function

'    easeInOutExpo: function (x, t, b, c, d) {
'        if (t==0) return b;
'        if (t==d) return b+c;
'        if ((t/=d/2) < 1) return c/2 * Math.pow(2, 10 * (t - 1)) + b;
'        return c/2 * (-Math.pow(2, -10 * --t) + 2) + b;
'    },
'Function easeInOutExpo(ByVal t, ByVal b, ByVal c, ByVal d) As Double
'    If t = 0 Then
'        easeInOutExpo = b
'    End If
'
'    If t = d Then
'        easeInOutExpo = b + c
'    End If
'
'    If ((t / d) / 2) < 1 Then
'        easeInOutExpo = c / 2 * (2 ^ (10 * ((t / d) - 1))) + b
'    End If
'
'    easeInOutExpo = c / 2 * (-(2 ^ (-10 * ((t / d) - 1))) + 2) + b
'End Function

'    easeInOutCubic: function (x, t, b, c, d) {
'        if ((t/=d/2) < 1) return c/2*t*t*t + b;
'        return c/2*((t-=2)*t*t + 2) + b;
'    },
'    easeInOutQuart: function (x, t, b, c, d) {
'        if ((t/=d/2) < 1) return c/2*t*t*t*t + b;
'        return -c/2 * ((t-=2)*t*t*t - 2) + b;
'    },
'    easeInOutQuint: function (x, t, b, c, d) {
'        if ((t/=d/2) < 1) return c/2*t*t*t*t*t + b;
'        return c/2*((t-=2)*t*t*t*t + 2) + b;
'    },
'    easeInCirc: function (x, t, b, c, d) {
'        return -c * (Math.sqrt(1 - (t/=d)*t) - 1) + b;
'    },
'    easeOutCirc: function (x, t, b, c, d) {
'        return c * Math.sqrt(1 - (t=t/d-1)*t) + b;
'    },
'    easeInOutCirc: function (x, t, b, c, d) {
'        if ((t/=d/2) < 1) return -c/2 * (Math.sqrt(1 - t*t) - 1) + b;
'        return c/2 * (Math.sqrt(1 - (t-=2)*t) + 1) + b;
'    },
'    easeInElastic: function (x, t, b, c, d) {
'        var s=1.70158;var p=0;var a=c;
'        if (t==0) return b;  if ((t/=d)==1) return b+c;  if (!p) p=d*.3;
'        if (a < Math.abs(c)) { a=c; var s=p/4; }
'        else var s = p/(2*Math.PI) * Math.asin (c/a);
'        return -(a*Math.pow(2,10*(t-=1)) * Math.sin( (t*d-s)*(2*Math.PI)/p )) + b;
'    },
'    easeOutElastic: function (x, t, b, c, d) {
'        var s=1.70158;var p=0;var a=c;
'        if (t==0) return b;  if ((t/=d)==1) return b+c;  if (!p) p=d*.3;
'        if (a < Math.abs(c)) { a=c; var s=p/4; }
'        else var s = p/(2*Math.PI) * Math.asin (c/a);
'        return a*Math.pow(2,-10*t) * Math.sin( (t*d-s)*(2*Math.PI)/p ) + c + b;
'    },

Function easeOutElastic(ByVal t, ByVal B, ByVal C, ByVal d) As Double
    Dim s As Double: s = 1.70158
    Dim p As Double: p = 0
    Dim a As Double: a = C
    Dim PI As Double: PI = 3.14159265359

    If t = 0 Then easeOutElastic = B

    t = t / d
    If (t = 1) Then easeOutElastic = B + C

    If Not p Then p = d * 0.3
    
    If a < Math.Abs(C) Then
        a = C
        s = p / 4
    Else
        s = p / (2 * PI) * Application.WorksheetFunction.Asin(C / a)
    End If

    easeOutElastic = a * (2 ^ (-10 * t)) * Math.Sin((t * d - s) * (2 * PI) / p) + C + B
End Function

'    easeInOutElastic: function (x, t, b, c, d) {
'        var s=1.70158;var p=0;var a=c;
'        if (t==0) return b;  if ((t/=d/2)==2) return b+c;  if (!p) p=d*(.3*1.5);
'        if (a < Math.abs(c)) { a=c; var s=p/4; }
'        else var s = p/(2*Math.PI) * Math.asin (c/a);
'        if (t < 1) return -.5*(a*Math.pow(2,10*(t-=1)) * Math.sin( (t*d-s)*(2*Math.PI)/p )) + b;
'        return a*Math.pow(2,-10*(t-=1)) * Math.sin( (t*d-s)*(2*Math.PI)/p )*.5 + c + b;
'    },
'    easeInBack: function (x, t, b, c, d, s) {
'        if (s == undefined) s = 1.70158;
'        return c*(t/=d)*t*((s+1)*t - s) + b;
'    },
'    easeOutBack: function (x, t, b, c, d, s) {
'        if (s == undefined) s = 1.70158;
'        return c*((t=t/d-1)*t*((s+1)*t + s) + 1) + b;
'    },
'    easeInOutBack: function (x, t, b, c, d, s) {
'        if (s == undefined) s = 1.70158;
'        if ((t/=d/2) < 1) return c/2*(t*t*(((s*=(1.525))+1)*t - s)) + b;
'        return c/2*((t-=2)*t*(((s*=(1.525))+1)*t + s) + 2) + b;
'    },
'    easeInBounce: function (x, t, b, c, d) {
'        return c - jQuery.easing.easeOutBounce (x, d-t, 0, c, d) + b;
'    },
'    easeOutBounce: function (x, t, b, c, d) {
'        if ((t/=d) < (1/2.75)) {
'            return c*(7.5625*t*t) + b;
'        } else if (t < (2/2.75)) {
'            return c*(7.5625*(t-=(1.5/2.75))*t + .75) + b;
'        } else if (t < (2.5/2.75)) {
'            return c*(7.5625*(t-=(2.25/2.75))*t + .9375) + b;
'        } else {
'            return c*(7.5625*(t-=(2.625/2.75))*t + .984375) + b;
'        }
'    },
'    easeInOutBounce: function (x, t, b, c, d) {
'        if (t < d/2) return jQuery.easing.easeInBounce (x, t*2, 0, c, d) * .5 + b;
'        return jQuery.easing.easeOutBounce (x, t*2-d, 0, c, d) * .5 + c*.5 + b;
'    }
