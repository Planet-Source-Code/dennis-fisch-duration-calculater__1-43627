<div align="center">

## Duration calculater


</div>

### Description

This code takes a value of seconds as an input value and then calculates the duration in seconds, minutes, hours and days without using any VB date function. This code runs super fast. I wrote this for my IRC Server and thought it might be any useful for you... if you like it, you may want to vote, if not, well the not, i guess ;)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dennis Fisch](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dennis-fisch.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, VBA MS Access, VBA MS Excel
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dennis-fisch-duration-calculater__1-43627/archive/master.zip)





### Source Code

<p>Public Function Duration(ByVal InSeconds As Long) As String<br>
Dim Seconds As Long, mins As Long, Hours As Long, Days As Long<br>
Seconds = InSeconds Mod 60<br>
mins = (InSeconds \ 60) Mod 60<br>
Hours = ((InSeconds \ 60) \ 60) Mod 24<br>
Days = ((InSeconds \ 60) \ 60) \ 24<br>
Duration = Days &amp; &quot; days &quot; &amp; Format$(Hours, &quot;00&quot;) &amp; &quot;:&quot; &amp; Format$(mins, &quot;00&quot;) &amp;
&quot;:&quot; &amp; Format$(Seconds, &quot;00&quot;)<br>
End Function</p>

