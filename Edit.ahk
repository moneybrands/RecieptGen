#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


^n::
Send, ^c
ClipWait, 1 
n:= Clipboard
Send, % Spell(n) " Only"
return
 
Spell(n) { ; recursive function to spell out the name of a max 36 digit integer, after leading 0s removed 
    Static p1=" Thousand",p2=" Million",p3=" Billion" ,p4=" Trillion",p5=" Quadrillion",p6=" Quintillion " 
         , p7=" Sextillion ",p8=" Septillion ",p9=" Octillion ",p10=" Nonillion ",p11=" Decillion " 
         , t2="Twenty",t3="Thirty",t4="Forty",t5="Fifty",t6="Sixty",t7="Seventy",t8="Eighty",t9="Ninety" 
         , o0="Zero",o1="One",o2="Two",o3="Three",o4="Four",o5="Five",o6="Six",o7="Seven",o8="Eight" 
         , o9="Nine",o10="Ten",o11="Eleven",o12="Twelve",o13="Thirteen",o14="Fourteen",o15="Fifteen" 
         , o16="Sixteen",o17="Seventeen",o18="Eighteen",o19="Nineteen" 
 
    n :=RegExReplace(n,"^0+(\d)","$1") ; remove leading 0s from n 
 
    If  (11 < d := (StrLen(n)-1)//3)   ; #of digit groups of 3 
        Return "Number too big" 
 
    If (d)                             ; more than 3 digits 
        Return Spell(SubStr(n,1,-3*d)) p%d% ((s:=SubStr(n,1-3*d)) ? ", " Spell(s) : "") 
 
    i := SubStr(n,1,1) 
    If (n > 99)                        ; 3 digits 
        Return o%i% " Hundred " ((s:=SubStr(n,2)) ? Spell(s) : "") 
 
    If (n > 19)                        ; n = 20..99 
        Return t%i% ((o:=SubStr(n,2)) ? "-" o%o% : "") 
 
    Return o%n%                        ; n = 0..19 
} 

^d::
FormatTime, CurrentDateTime,, yyyy-MM-dd HH:mm:ss
Send %CurrentDateTime%
return
