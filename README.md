<div align="center">

## Count spaces or occurances of words or anything with one simple line of code


</div>

### Description

Count spaces or occurances of words or anything with one simple line of code.
 
### More Info
 
text1.text would be the string you want to count the occurances in. this particular code would count how amny times "hi" appears in text1.text

a common use of this to count both lower and uppercase would to first use the line

text1.text = Lcase(text1.text)    to make text1 all lower case. you coudl aslo use this

searchtextbox.text = lcase(searchtextbox.text)

occuranceslabel.caption= ubound(split(document.text, searchtextbox.text))

that would put in occuranceslabel the number of times that the text in searchtextbox appears in the document textbox.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\_\+Seanp2k\+\_](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/seanp2k.md)
**Level**          |Beginner
**User Rating**    |3.3 (13 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/seanp2k-count-spaces-or-occurances-of-words-or-anything-with-one-simple-line-of-code__1-40950/archive/master.zip)





### Source Code

```
Ubound(Split(text1.text,"hi"))
this can easily be transformed into a function
Public Function fCount(ByVal sString As String, ByVal sDelim As String) As Long
  fCount = 0
  On Error Resume Next
  fCount = UBound(Split(sString, sDelim))
End Function
```

