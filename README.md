<div align="center">

## GetTag, GetTagText, CutTag


</div>

### Description

This is a set of three functions that pull tagged data, such as that from HTML or XML, based on the tag name. I've used this in a number of applications where I've need to store multiple bits of variable-length data in a single string or file.
 
### More Info
 
A string containing tags which in turn contain data you are searching for AND the name of the tag.

For instance, if you're pulling from HTML for anything between <H1> and </H1> then you would enter "H1" as the tag name.

GetTag() returns the data between the specified tags along with the tags.

GetTagText() returns the data between the specified tags, without the tags.

CutTag() returns the original text minus the specified tags and the data between them.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jamie Richard Wilson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jamie-richard-wilson.md)
**Level**          |Beginner
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jamie-richard-wilson-gettag-gettagtext-cuttag__1-9927/archive/master.zip)





### Source Code

```
Public Function GetTag(SourceString As String, Tag As String) As String
  'Gets the tag and text between it
  If InStr(SourceString, "<" & Tag & ">") = 0 Then
    GetTag = ""
    Exit Function
  End If
  GetTag = Mid$(SourceString, InStr(SourceString, "<" & Tag & ">"), InStr(SourceString, "</" & Tag & ">") + Len("</" & Tag & ">") - 1)
End Function
Public Function GetTagText(SourceString As String, Tag As String) As String
  'Grabs the text between tags
  If InStr(SourceString, "<" & Tag & ">") = 0 Then
    GetTagText = ""
    Exit Function
  End If
  GetTagText = Mid$(SourceString, InStr(SourceString, "<" & Tag & ">") + Len("<" & Tag & ">"), (InStr(SourceString, "</" & Tag & ">")) - (InStr(SourceString, "<" & Tag & ">") + Len("<" & Tag & ">")))
End Function
 Public Function CutTag(SourceString As String, Tag As String) As String
  'Cuts the entire tag out of the text
  If InStr(SourceString, "<" & Tag & ">") = 0 Then
    CutTag = ""
    Exit Function
  End If
  CutTag = Left$(SourceString, InStr(SourceString, "<" & Tag & ">") - 1) & Mid$(SourceString, InStrRev(SourceString, "</" & Tag & ">") + Len("</" & Tag & ">"))
End Function
```

