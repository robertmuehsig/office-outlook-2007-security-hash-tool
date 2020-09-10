'''''''''''''''''''''''''''''''''''''''''''''
'
'  Outlook 12 Add-in Hash Generator Utility
'
'''''''''''''''''''''''''''''''''''''''''''''

'  Description:
'  ------------
'  This script allows an administrator to determine the correct hash value
'  for an add-in DLL to be added to the trusted add-in policy keys.
'
'  To Run:  
'  -------
'  This is the format for this script:
'  
'      cscript addinhash.vbs <path to dll file>
'  
'  NOTE:  If you want to execute this script directly from Windows, use 
'  'wscript' instead of 'cscript'. 
'
'''''''''''''''''''''''''''''''''''''''''''''

' Initalize error checking
On Error Resume Next

' initalize variables
Dim ArgCount, TargetFile, HashCtrl, HashOutput

' Default values
ArgCount = 0
TargetFile = ""
HashOutput = ""

' Parse command line arguments
While ArgCount < Wscript.Arguments.Count
	' Determine switches used
	Select Case WScript.Arguments(ArgCount)
		Case "-h", "-?", "/?":
			Call UsageMsg
		Case Else:
			TargetFile = WScript.Arguments(ArgCount)
	End Select
	
	' Move our counter to the next argument
	ArgCount = ArgCount + 1
Wend

If Len(TargetFile) = 0 Then
	Call UsageMsg
End If


' Initalize the hash control
Set HashCtrl = CreateObject("Hash.HashCtl")
If Err.Number <> 0 Then
	Call HashCtrlError(1)
End If

' Hash the input file and get the result
WScript.Echo "Generating trusted add-in settings..."
WScript.Echo "Please copy and paste the following information into the group policy editor:"
WScript.Echo ""
WScript.Echo "Value Name: " & TargetFile

HashOutput = HashCtrl.HashFile(TargetFile)
If Err.Number <> 0 Then
	Call HashCtrlError(2)
End If

WScript.Echo "Value     : " & HashOutput

Sub UsageMsg()
	Wscript.Echo "Usage: cscript addinhash.vbs <add-in dll file>"
	WScript.Quit
End Sub

Sub HashCtrlError(errorNumber)
	Select Case errorNumber
		Case 1:
			Wscript.Echo "Unable to load hashctl.dll.  Please make sure this file has been registered."
			Wscript.Echo " ex. regsvr32 hashctl.dll"
	
		Case 2:
			WScript.Echo "An error occured generating a hash of the file.  Check the file name and try again."
		Case Else:
			WScript.Echo "An unknown error occured with the hash control."
	End Select
	
	Wscript.Quit(errorNumber)
End Sub