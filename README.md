<div align="center">

## Process Parsing/Detecting Tutorial


</div>

### Description

FULLY COMMENTED TUTORIAL ON PROCESS DETECTION USING THE PROCESSFIRST, PROCESSNEXT, AND CREATETOOLHELP API'S.

Ok here is the deal, I spent alot of time trying to figure out how to find a process, parse out the exe file name and it's process id to use for process memory manipulation. However, there were no eazy to read tutorials that could make me understand how the api's worked. After figuring everything out, I thought I should make a fully commented source for others who may end up stuck in my shoes. I hope you enjoy. The sub is the easier to understand than all on psc thus far.
 
### More Info
 
input the case sensitive exe file name.

It is extreemely commented, to the point where someone who don't know vb at all could prolly understand it. To view the Processentry32 type declaration and related api's refer to the module.

determines if the process is running and if so it will msgbox the processes parsed szexefile name and it's process id number.

I have not tried this code on any other environment than vb6.


<span>             |<span>
---                |---
**Submitted On**   |2003-09-09 13:01:36
**By**             |[Steven Sherrod](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steven-sherrod.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Process\_Pa164358992003\.zip](https://github.com/Planet-Source-Code/steven-sherrod-process-parsing-detecting-tutorial__1-48380/archive/master.zip)

### API Declarations

```
processfirst
processnext
createtoolhelpsnapshot
type processentry32
```





