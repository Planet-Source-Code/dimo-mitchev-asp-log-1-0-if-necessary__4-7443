<div align="center">

## Asp log 1\.0 \(if necessary\)


</div>

### Description

with this code can be written log file from asp page
 
### More Info
 
string to be writen

can be adobted for many files, some formating...

'(may be in next versions)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dimo Mitchev](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dimo-mitchev.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dimo-mitchev-asp-log-1-0-if-necessary__4-7443/archive/master.zip)

### API Declarations

free


### Source Code

```
sub writelog(lstr1)
	dim fs,fname
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	If not (fs.FileExists("logasp.txt"))=true Then
		set fname=fs.CreateTextFile("logasp.txt",true)
		fname.WriteLine(lstr1)
		fname.Close
	else
		dim f
		set f=fs.OpenTextFile("logasp.txt",8,true)
		f.WriteLine(lstr1)
		f.Close
		set f=Nothing
	end if
	set fname=nothing
	set fs=nothing
end sub
```

