<div align="center">

## IsExpired


</div>

### Description

This code checks the difference between today and any expiration date. Using VB6 functions DateDiff and TimeValue it will evaluate the dates and tell you if we are past the expiration date or not. [Highly commented.]
 
### More Info
 
All you need to input is the expiration date and the expiration time. These can be in any format.

This code can be used for many things, including shareware locks. You can set an expiration date (e.g. "30 days from today.") and check each day if it has expired. The simple boolean return will tell you if the shareware lock has expired or not.

This code returns True if we are past the expiration date. It returns False if it has not yet expired.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ignis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ignis.md)
**Level**          |Beginner
**User Rating**    |3.3 (10 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ignis-isexpired__1-24055/archive/master.zip)





### Source Code

```
Function IsExpired(ExpireDate As Date, ExpireTime As Date) As Boolean
 Dim lngDayDiff As Long
 Dim lngTimeDiff As Long
 ' Using DateDiff, a function unique to VB6, we check the
 ' difference between the current date (extracted from Now)
 ' and the expiration date.
 lngDayDiff = DateDiff("d", Now, ExpireDate)
 ' If the difference is a negative that means that we are
 ' past the expired date so of course it is expired.
 If lngDayDiff < 0 Then
  GoTo YesExpired
 ' If the difference is a zero that means we are ON the
 ' date of expiration. We check the time for a difference
 ' to determine if the time has expired.
 ElseIf lngDayDiff = 0 Then
  ' Get the time difference. Note that we use TimeValue(Now)
  ' instead of just Now because it will return the exact
  ' time, not the date/time.
  lngTimeDiff = DateDiff("n", TimeValue(Now), ExpireTime)
  ' If the time difference is a negative, we are past it so
  ' the date is expired.
  If lngTimeDiff <= 0 Then
   GoTo YesExpired
  ' Otherwise (if we are on the time, or before it) then
  ' we are not yet expired.
  Else
   GoTo NoExpired
  End If
 ' Otherwise (if we are on the date, or before it) then
 ' we are not yet expired.
 Else
  GoTo NoExpired
 End If
YesExpired:
 IsExpired = True
 Exit Function
NoExpired:
 IsExpired = False
 Exit Function
End Function
```

