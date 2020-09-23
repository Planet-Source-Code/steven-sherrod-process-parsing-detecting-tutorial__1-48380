VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the case-sensitive process exe file name--i.e. IEXPLORER.EXE"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I created this example because I spent alot of time searching through
'psc and other sources looking for an example for a real NEWBIE to the
'process detection/manipulation api calls and usage, and never found
'one that I could understand instantly.  It took two weeks comparing
'four different sources before I could grasp the concept.
'Now that I have it, I wish to share it with others in hopes that
'they may find it much easier to understand.  Feel free to reuse
'this source in any project, and enjoy.  I hope it helps!


Private Sub myGameProcess()
Dim gPid As Long
Dim myGameProcess As PROCESSENTRY32 '<--Refer to the type listing in the module
Dim myGameSnapshot As Long '<--SnapShot of the process list
mygameid = Text1.Text     '<--Storage for the process the user named
Dim success As Long       '<--Used to run through the processes
Dim Name As String        '<--Used to parse the .szexefile(process exe file name) to compare to the users process request
Dim ProcThere As Boolean  '<--To toggle weather or not the process is running
ProcThere = False         '<--Defaulting the toggle switch
Dim IsThere As String     '<--Just to store the msgbox info string
myGameSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0) '<--(say cheese process list)Snap!  Took a picture of the process list

myGameProcess.dwSize = Len(myGameProcess) '<--I'm not entirely sure why it's done.. but it is a MUST
success = ProcessFirst(myGameSnapshot, myGameProcess) '<--Gets the first process in the list
    
    If success = -1 Then '**********************  Incase there are
        MsgBox "No processes available"       '*  no processes
        Exit Sub                              '*  running
    End If '************************************
  
While success <> 0  '<--Loop through the processes
    Name = Left(myGameProcess.szexeFile, InStr(myGameProcess.szexeFile, vbNullChar) - 1) '<--parse out the exe fileName string
    If Name = mygameid Then '******************** Compare the name
        gPid = myGameProcess.th32ProcessID     '* of the exe file
        ProcThere = True                       '* to the file the user
    End If  '************************************ named in text1.text

myGameProcess.dwSize = Len(myGameProcess) '<--get ready for the next process

success = ProcessNext(myGameSnapshot, myGameProcess) '<--obtain the next process in the list

Wend

If ProcThere = True Then '************************************************* Check the toggle
   IsThere = mygameid + "(ProcessId =" & gPid & ") is currently running" '* switch to see if
    MsgBox IsThere                                                       '* the process is running
Else                                                                     '* and advise the user
                                                                         '*
    MsgBox "The process is not currently running"                        '*
End If '*******************************************************************

End Sub

Private Sub Command1_Click()
Call myGameProcess
End Sub
