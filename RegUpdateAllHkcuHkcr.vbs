Option Explicit

'-----------------------------------------------------------------------------
' Class: RegUpdateAllHkcuHkcr
' 
' Modify HKCU and/or HKCR registry key(s) for ALL users on a system.
' 
' Run with cscript to suppress dialogs:
' 
' ====== Text ======
' cscript.exe RegUpdateAllHkcuHkcr.vbs
' ==================
'
' Version:
'
' 	2.0.0
'
' Author:
'
' 	* <Mick Grove: http://micksmix.wordpress.com> (2012-2013)
'	* Eduardo Mozart de Oliveira (2014-2017)
'
Class RegUpdateAllUsers
	Private DAT_NTUSER, DAT_USRCLASS
	Private RegRoot	
	Private WshShell, WshShellApp, objFSO, objRegistry
	
	Private AHKCUKeys, AHKCRKeys
	
	Private debugEnabled
	
	' Enable or disable debug logging. If enabled, debug messages are 
	' logged to the enabled facilities. Otherwise debug messages are 
	' silently discarded. This property is disabled by default.
	Public Property Get Debug
		Debug = debugEnabled
	End Property

	Public Property Let Debug(ByVal enable)
		debugEnabled = CBool(enable)
		objRegistry.Debug = debugEnabled
	End Property
	
	' -------------------------------------------------------------------------------
	
	'-----------------------------------------------------------------------------
	' Function: Class_Initialize
	'
	' - Declare RegRoot variable: This is where our HKCU is temporarily loaded, and where we need to write to it. You don't really need to change this, but you can if you want.
	' - Declare AHKCUKeys and AHKCRKeys variables: Arrays that contain Registry Keys/Values informed by the user that will be modified by the class. See <Validate>.
	' - Set WshShell, WshShellApp, objFSO and objRegistry objects. 
	'
	Private Sub Class_Initialize()
	  DAT_NTUSER 	= &H70000000
	  DAT_USRCLASS	= &H70000001
	  
	  RegRoot = "HKLM\TEMPHIVE"
	  
	  AHKCUKeys = Array()
	  AHKCRKeys = Array()
	  
	  Set WshShell = CreateObject("WScript.shell")
	  Set WshShellApp = CreateObject("Shell.Application")
	  Set objFSO = CreateObject("Scripting.FileSystemObject")
	  Set objRegistry = New CWMIReg
	  
	  debugEnabled = False
	End Sub 
	
	'-----------------------------------------------------------------------------
	' Function: Class_Terminate
	'
	' - Set WshShell, WshShellApp, objFSO and objRegistry objects to Nothing. 
	'
	Private Sub Class_Terminate()
		Set WshShell = Nothing
		Set WshShellApp = Nothing
		Set objFSO = Nothing
		Set objRegistry = Nothing
	End Sub
	
	' -------------------------------------------------------------------------------
	
	'-----------------------------------------------------------------------------
	' Function: SetValue
	'
	' Validate SetValue function arguments sent by the user.
	'
	' Parameters:
	'
	'	strRegistryKey - The Registry Key to be modified (Example: "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce").
	'	strValue       - The value to be set (Example: "Chrome").
	'	TypeIn_        - (Optional) A data type (Example: "REG_SZ", "REG_EXPAND_SZ", "REG_BINARY", "REG_DWORD", "REG_MULTI_SZ", "REG_QWORD"). Use "" to objRegistry.SetValue detect data type to use.
	'
	' See Also:
	'
	'	<Validate>
	'
	Public Sub SetValue(strRegistryKey, strValue, TypeIn_) 
	    Call Validate("SetValue", strRegistryKey, strValue, TypeIn_)
	End Sub
	
	'-----------------------------------------------------------------------------
	' Function: Delete
	'
	' Validate Delete function arguments sent by the user.
	'
	' Parameters:
	'
	'	strRegistryKey - The Registry Key/Value to be deleted. Add "\" at end for keys (Example: "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce\Chrome\").
	'
	' See Also:
	'
	'	<Validate>
	'
	Public Sub Delete(strRegistryKey)
		Call Validate("Delete", strRegistryKey, "", "")
	End Sub
	
	'-----------------------------------------------------------------------------
	' Function: Validate
	'
	' Parse strRegistryKey argument sent by <SetValue> and <Delete> functions.
	'
	' Parameters:
	'
	'	strMethod      - Declare main function call (<SetValue> or <Delete>). 
	'	strRegistryKey - The Registry Key to be modified (Example: "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce"). Note that for <Delete> function it may represents the Registry Key AND Value to be deleted. 
	'	strValue       - (SetValue) The value to be set (Example: "Chrome").
	'	TypeIn_        - (SetValue) (Optional) A data type (Example: "REG_SZ", "REG_EXPAND_SZ", "REG_BINARY", "REG_DWORD", "REG_MULTI_SZ", "REG_QWORD"). Use "" to objRegistry.SetValue detect data type to use.
	'
	' Returns:
	'
	'	This functions populates AHKCUKeys and AHKCRKeys Arrays accordantly with Registry Keys/Values sent if strRegistryKey is valid (HKCU or HKCR).
	'
	Private Sub Validate(strMethod, strRegistryKey, strValue, TypeIn_)
		Dim sRoot, sPartialPath, AList
	    Dim iResult, strErr, KeyLevel
	   
	    ' HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\DisablePasswordCaching
	    sRoot = Left(strRegistryKey, InStr(strRegistryKey, "\")-1) ' HKCU
	    sPartialPath = Mid(strRegistryKey, InStr(strRegistryKey, "\")+1, Len(strRegistryKey)) ' Software\Microsoft\Windows\CurrentVersion\Internet Settings\DisablePasswordCaching
		
		If UCase(sRoot) = "HKCU" Or UCase(sRoot) = "HKEY_CURRENT_USER" Then
			If UCase(Left(sPartialPath, Len("SOFTWARE\CLASSES\"))) = "SOFTWARE\CLASSES\" Then
		    	' HKCU\Software\Classes = HKCR
		    	
		    	If Len(sRoot) > 4 Then
		    		sRoot = "HKEY_CLASSES_ROOT"
		    	Else
		    		sRoot = "HKCR"
		    	End If
		    	
		    	sPartialPath = Mid(sPartialPath, Len("Software\Classes\")+1, Len(sPartialPath))
			End If
		End If
		
	    

	    If strMethod = "SetValue" Then 
	    	AList = Array("SetValue", sPartialPath, strValue, TypeIn_)

	    ElseIf strMethod = "Delete" Then			
			AList = Array("Delete", sPartialPath)
			
	    End If
	    
	    If (UCase(sRoot) = "HKCR" Or UCase(sRoot) = "HKEY_CLASSES_ROOT") Then
	    	ReDim Preserve AHKCRKeys(UBound(AHKCRKeys) + 1)
	    	AHKCRKeys(UBound(AHKCRKeys)) = AList
	    	
	    ElseIf UCase(sRoot) = "HKCU" Or UCase(sRoot) = "HKEY_CURRENT_USER" Then
    		ReDim Preserve AHKCUKeys(UBound(AHKCUKeys) + 1)
    		AHKCUKeys(UBound(AHKCUKeys)) = AList  	
	    	
	    Else
    		strErr = "*** Error: " & strErr & Quotes(sRoot) & " is not supported" & vbCrLf
    		WScript.Echo strErr
			
	    End If
	End Sub
	
	'-----------------------------------------------------------------------------
	' Function: SetValue_
	'
	' Private function called by KeysToModify function to set values on Registry.
	'
	' Parameters:
	'
	'	strRegistryKey - The Registry Key to be modified (Example: "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce").
	'	strValue       - The value to be set (Example: "Chrome").
	'	TypeIn_        - (Optional) A data type (Example: "REG_SZ", "REG_EXPAND_SZ", "REG_BINARY", "REG_DWORD", "REG_MULTI_SZ", "REG_QWORD"). Use "" to objRegistry.SetValue detect data type to use.
	'
	' Returns:
	'
	'	Echo objRegistry.SetValue error code on error.
	'
	' See Also:
	'
	'	<KeysToModify>
	'	<CheckError>
	'
	Private Sub SetValue_(strRegistryKey, strValue, TypeIn_)
		Dim iResult, strErr, x
		
	    iResult = objRegistry.SetValue(strRegistryKey, strValue, TypeIn_)
		
		If (iResult = 0) Then
	       ' WScript.Echo strRegistryKey & " (" & TypeIn_ & ") value added successfully"
	    Else
	        strErr = "*** Error adding " & strValue & " (" & TypeIn_ & ") value at " & strRegistryKey
	        
	        strErr = strErr & ": " & CheckError(iResult) & " (" & iResult & ")"
			
	        WScript.Echo strErr
	    End If 
	End Sub
	
	'-----------------------------------------------------------------------------
	' Function: Delete_
	'
	' Private function called by KeysToModify function to delete keys/values on Registry.
	'
	' Parameters:
	'
	'	strRegistryKey - The Registry Key/Value to be deleted (Example: "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce\Chrome").
	'
	' Returns:
	'
	'	Echo objRegistry.Delete error code on error.
	'
	' See Also:
	'
	'	<KeysToModify>
	'	<CheckError>
	'
	Private Sub Delete_(strRegistryKey)
		Dim iResult, strErr
		
		iResult = objRegistry.Delete(strRegistryKey)
		If (iResult = 0) Then
			' WScript.Echo strRegistryKey & " deleted successfully"
		Else
			strErr = "*** Error deleting " & strRegistryKey
					
			strErr = strErr & ": " & CheckError(iResult) & " (" & iResult & ")"
				
			WScript.Echo strErr
		End If
	End Sub
	
	'-----------------------------------------------------------------------------
	' Function: KeysToModify
	'
	' For Each AHKCUKeys and AHKCRKeys Arrays to <SetValue_> or <Delete_> functions.
	'
	' Parameters:
	'
	'	sRegistryRootToUse - Declare User SID Key to modify (Example: "HKEY_USERS\S-1-5-18").
	'	DAT_FILE           - Declare Root Key to modify (DAT_NTUSER for HKCU, DAT_USRCLASS for HKCR). 
	'
	' See Also:
	'
	'	<LoadProfileHive>
	'	<LoadRegistry>
	'
	Private Sub KeysToModify(sRegistryRootToUse, DAT_FILE) 
		Dim strMethod, strRegistryKey, strValue, TypeIn_
		Dim AHKCUKey, AHKCRKey
	
		If DAT_FILE = DAT_NTUSER Then ' This is for updating HKCU keys
			For Each AHKCUKey In AHKCUKeys ' AHKCUKeys is a Global variable from Validate()
				strMethod = AHKCUKey(0)
				strRegistryKey = AHKCUKey(1)
				If strMethod = "SetValue" Then
					strValue = AHKCUKey(2)
					TypeIn_ = AHKCUKey(3)
				End If
				
				If strMethod = "SetValue" Then
					Call SetValue_(sRegistryRootToUse & "\" & strRegistryKey, strValue, TypeIn_)
				ElseIf strMethod = "Delete" Then
					Delete_(sRegistryRootToUse & "\" & strRegistryKey)
				End If
			Next
			
		ElseIf DAT_FILE = DAT_USRCLASS Then ' This is for updating HKCR keys per-user
			For Each AHKCRKey In AHKCRKeys 
				strMethod = AHKCRKey(0)
				strRegistryKey = AHKCRKey(1)
				If strMethod = "SetValue" Then
					strValue = AHKCRKey(2)
					TypeIn_ = AHKCRKey(3)
				End If
				
				If strMethod = "SetValue" Then
					Call SetValue_(sRegistryRootToUse & "\" & strRegistryKey, strValue, TypeIn_)
				ElseIf strMethod = "Delete" Then
					WScript.Echo "Keys to modify: Deleting " & sRegistryRootToUse & "\" & strRegistryKey
					Delete_(sRegistryRootToUse & "\" & strRegistryKey)
				End If
			Next
		End If	
	End Sub
	
	' -------------------------------------------------------------------------------
	
	'-----------------------------------------------------------------------------
	' Function: GetDefaultUserPath
	'
	' Get Default User Path (Example: "C:\Users\Default" on Vista or newer).
	'
	' Returns:
	'
	'	Default User Path from Registry.
	'
	' See Also:
	'
	'	<LoadRegistry>
	'
	Private Function GetDefaultUserPath	    
	    Dim strKeyPath
	    Dim strDefaultUser
	    Dim strDefaultPath
	    Dim strResult
	  
	    strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	  
	    strDefaultUser = objRegistry.GetValue("HKLM\" & strKeyPath & "\DefaultUserProfile")
	    
	    strDefaultPath = objRegistry.GetValue("HKLM\" & strKeyPath & "\ProfilesDirectory")
	          
	    If Len(strDefaultUser) < 1 or IsEmpty(strDefaultUser) or IsNull(strDefaultUser) Then
	        ' Must be on Vista or newer.
	        strResult = objRegistry.GetValue("HKLM\" & strKeyPath & "\Default")
	    Else
	        ' Must be on XP.
	        strResult = strDefaultPath & "\" & strDefaultUser
	    End If
	      
	    GetDefaultUserPath = strResult
	End Function
	
	'-----------------------------------------------------------------------------
	' Function: LoadProfileHive
	'
	' Load specified Registry Hive from User Profile on RegRoot. See <Class_Initialize>.
	'
	' Parameters:
	'
	'	sProfileDataFilePath - Declare absolute path of DAT file to modify (Example: "C:\Users\Default\NTUSER.DAT").
	'	sCurrentUser         - Relative User Profile Path (Example: "Default")
	'	DAT_FILE             - Declare Root Key to modify (DAT_NTUSER for HKCU, DAT_USRCLASS for HKCR).
	'
	' Returns:
	'
	'	Load User Hive from sProfileDataFilePath to RegRoot.
	'
	' See Also:
	'
	'	<GetPathToDatFileToUpdate>
	'	<LoadRegistry>
	'	<KeysToModify>
	'
	Private Sub LoadProfileHive(sProfileDatFilePath, sCurrentUser, DAT_FILE)
	    Dim intResultLoad, intResultUnload, sUserSID
	 
	    'Load user's HKCU into temp area under HKLM 
	    intResultLoad = WshShell.Run("reg.exe load " & RegRoot & " " & Quotes(sProfileDatFilePath), 0, True) 
	    If intResultLoad <> 0 Then
	        ' This profile appears to already be loaded... Let's update it under the HKEY_USERS hive. 
	        Dim strSubKey2 
	        Dim strKeyPath2, strValueName2, strValue2 
	        Dim strSubPath2, arrSubKeys2 
	   
	        strKeyPath2 = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	        objRegistry.EnumKeys "HKLM\" & strKeyPath2, arrSubkeys2 
	        sUserSID = ""
	  
	        For Each strSubkey2 In arrSubkeys2 
	            strValueName2 = "ProfileImagePath"
	            strSubPath2 = strKeyPath2 & "\" & strSubkey2 
	            strValue2 = objRegistry.GetValue("HKLM\" & strSubPath2 & "\" & strValueName2)
	            If Right(UCase(strValue2),Len(sCurrentUser)+1) = "\" & UCase(sCurrentUser) Then
	                ' This is the one we want. 
	                sUserSID = objSubkey2 
	            End If
	        Next
	  
	        If Len(sUserSID) > 1 Then
	            WScript.Echo "  Updating another logged-on user: " & sCurrentUser & vbCrLf 
				
				If DAT_FILE = DAT_NTUSER Then
					Call KeysToModify("HKEY_USERS\" & sUserSID, DAT_FILE) 
				ElseIf DAT_FILE = DAT_USRCLASS Then
					Call KeysToModify("HKEY_USERS\" & sUserSID & "_Classes", DAT_FILE) 
				End If		
	        Else
	            WScript.Echo("  *** An error occurred while loading HKCU for this user: " & sCurrentUser) 
	        End If
	    Else
	        WScript.Echo("  HKCU loaded for this user: " & sCurrentUser) 
	    End If
	  
	    '' 
	    If sUserSID = "" then ' Check to see if we just updated this user b/c they are already logged on.
	        Call KeysToModify(RegRoot, DAT_FILE) ' Update registry settings for this selected user.
	    End If
	    '' 
	  
	    If sUserSID = "" then ' Check to see if we just updated this user b/c they are already logged on.
	        intResultUnload = WshShell.Run("reg.exe unload " & RegRoot,0, True) ' Unload HKCU from HKLM.
	        If intResultUnload <> 0 Then
	            WScript.Echo("  *** An error occurred while unloading HKCU for this user: " & sCurrentUser & vbCrLf) 
	        Else
	            WScript.Echo("  HKCU UN-loaded for this user: " & sCurrentUser & vbCrLf) 
	        End If
	    End If
	End Sub
	
	'-----------------------------------------------------------------------------
	' Function: GetUserRunningScript
	'
	' Get User running the Script.
	'
	' Returns:
	'
	'	User running the Script from %USERNAME% or %USERPROFILE% Environment Variable.
	'
	' See Also:
	'
	'	<LoadRegistry>
	'
	Private Function GetUserRunningScript()
		Dim sUserRunningScript, sComputerName 
	    sUserRunningScript = WshShell.ExpandEnvironmentStrings("%USERNAME%") ' Holds name of current logged on user running this script.
	    sComputerName = UCase(WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")) 
	     
	    If sUserRunningScript = "%USERNAME%" or sUserRunningScript = sComputerName & "$"  Then
	        ' This script might be run by the SYSTEM account or a service account.
	        Dim sTheProfilePath 
	        sTheProfilePath = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") 'Holds name of current logged on user running this script.
	    
	        sUserRunningScript = Mid(sTheProfilePath, (InStrRev(sTheProfilePath, "\") + 1), Len(sTheProfilePath))
	    End If
		
		GetUserRunningScript = sUserRunningScript
	End Function
	
	'-----------------------------------------------------------------------------
	' Function: GetPathToDatFileToUpdate
	'
	' Get Absolute Path to DAT File from the User Profile Path specified.
	'
	' Parameters:
	'
	'	sProfilePath - Absolute Path to User Profile Path (Example: "C:\Users\Default").
	'	DAT_FILE     - Declare DAT File to modify (DAT_NTUSER for HKCU, DAT_USRCLASS for HKCR).
	'
	' Returns:
	'
	'	Returns Absolute Path to "NTUSER.DAT" (DAT_NTUSER) or "USRCLASS.DAT" (DAT_USRCLASS) for sProfilePath, or "" if file could not be found.
	'
	' See Also:
	'
	'	<LoadRegistry>
	'
	Private Function GetPathToDatFileToUpdate(sProfilePath, DAT_FILE)	
		Dim sDatFile, sPathToDat
		Dim sAppData, sCurrentUserProfilePath

		sAppData = WshShellApp.NameSpace(28).Self.Path
		sAppData = Mid(sAppData, Len(WshShellApp.NameSpace(40).Self.Path & "\")+1, Len(sAppData))
		
		sPathToDat = "" 'default
		
		If Right(sProfilePath,1) = "\" Then sProfilePath = Left(sProfilePath, Len(sProfilePath)-1)
		
		If DAT_FILE = DAT_NTUSER Then
			sDatFile = "NTUSER.DAT"
			
			If objFSO.FileExists(sProfilePath & "\" & sDatFile) Then
				sPathToDat = sProfilePath & "\" & sDatFile		
			End If
		ElseIf DAT_FILE = DAT_USRCLASS Then
			sDatFile = "USRCLASS.DAT"

			If objFSO.FileExists(sProfilePath & "\" & sAppData & "\Microsoft\Windows\" & sDatFile) Then
				sPathToDat = sProfilePath & "\" & sAppData & "\Microsoft\Windows\" & sDatFile
			End If
		End If
		
		GetPathToDatFileToUpdate = sPathToDat
	End Function
	
	'-----------------------------------------------------------------------------
	' Function: LoadRegistry
	'
	' Call <LoadProfileHive> function for all users, and <KeysToModify> for Default User.
	'
	' See Also:
	'
	'	<LoadProfileHive>
	'	<KeysToModify>
	'
	Public Sub LoadRegistry()
	    Dim sUserRunningScript 
	    Dim objSubkey 
	    Dim strKeyPath, strValueName, strSubPath, arrSubKeys 
	    Dim sCurrentUser, sProfilePath, sNewUserProfile
		Dim sPathToDatFile
		Dim ADATList, DAT_FILE
	    
		ADATList = Array()
		If Not UBound(AHKCUKeys) = -1 Then ' DAT_NTUSER (HKCU)
			ReDim Preserve ADATList(UBound(ADATList) + 1)
			ADATList(UBound(ADATList)) = DAT_NTUSER
		End If
		
		If Not UBound(AHKCRKeys) = -1 Then ' DAT_USRCLASS (HKCR)
			ReDim Preserve ADATList(UBound(ADATList) + 1)
			ADATList(UBound(ADATList)) = DAT_USRCLASS
		End If
	    
	    sUserRunningScript = GetUserRunningScript        
	    
	    sNewUserProfile = GetDefaultUserPath 
	    
	    For Each DAT_FILE In ADATList
		    WScript.Echo "Updating the logged-on user: " & sUserRunningScript & vbCrLf 
			''
			
	    	If DAT_FILE = DAT_NTUSER Then
	    			Call KeysToModify("HKCU", DAT_FILE) ' Update registry settings for the user running the script.
	        ElseIf DAT_FILE = DAT_USRCLASS Then
	                Call KeysToModify("HKCR", DAT_FILE) ' Update registry settings which affects newly created profiles.
	        End If
	    	''
			
			' Default user does not have HKCR (DAT_USRCLASS).
			If DAT_FILE = DAT_NTUSER Then 
					sPathToDatFile = GetPathToDatFileToUpdate(sNewUserProfile, DAT_FILE)
				
			    If Len(sPathToDatFile) > 0 Then
			        WScript.Echo "Updating the DEFAULT user profile which affects newly created profiles." & vbCrLf
			        Call LoadProfileHive(sPathToDatFile, "Default User Profile", DAT_FILE) 
			    Else
			        WScript.Echo "Unable to update the DEFAULT user profile, because it could not be found at: " _
			            & vbCrLf & sPathToDatFile & vbCrLf
			    End If
		    End If
		    
		    strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
		    objRegistry.EnumKeys "HKLM\" & strKeyPath, arrSubkeys 
		     
		    For Each objSubkey In arrSubkeys 
		        strValueName = "ProfileImagePath"
		        strSubPath = strKeyPath & "\" & objSubkey 
		        sProfilePath = objRegistry.GetValue("HKLM\" & strSubPath & "\" & strValueName)
		        sCurrentUser = Mid(sProfilePath, (InStrRev(sProfilePath, "\") + 1), Len(sProfilePath))
		     
		        If ((UCase(sCurrentUser) <> "ALL USERS") and _ 
		            (UCase(sCurrentUser) <> UCase(sUserRunningScript)) and _ 
		            (UCase(sCurrentUser) <> "LOCALSERVICE") and _ 
		            (UCase(sCurrentUser) <> "SYSTEMPROFILE") and _ 
		            (UCase(sCurrentUser) <> "NETWORKSERVICE")) then 
		             
					sPathToDatFile = GetPathToDatFileToUpdate(sProfilePath, DAT_FILE)
					
					If Len(sPathToDatFile) > 0 Then
						WScript.Echo "Preparing to update the user: " & sCurrentUser
						Call LoadProfileHive(sPathToDatFile, sCurrentUser, DAT_FILE)
					End If
		        End If
		    Next
		Next
	End Sub
	
	'-----------------------------------------------------------------------------
	' Function: Quotes
	'
	' Wraps String in Quotes ("").
	'
	' Parameters:
	'
	'	strString - String to format.
	'
	' Returns:
	'
	'	Quoted strString.
	'
	' See Also:
	'
	'	<Validate>
	'	<LoadProfileHive>
	'
	Private Function Quotes(strString)
		Quotes = Chr(34) & strString & Chr(34)
	End Function
	
	' Private Function RemoveItemFromArray(arr, removeIndex)
		' From: http://stackoverflow.com/questions/17753462/delete-an-element-from-an-array-in-classic-asp
	'	For x=removalIndex To UBound(arr)-1
	'	    arr(x) = arr(x + 1)
	'	Next
	'	ReDim Preserve arr(UBound(arr) - 1)
	' End Function

	'-----------------------------------------------------------------------------
	' Function: CheckError
	'
	' Parse iResult to output a generic error message.
	'
	' Parameters:
	'
	'	iResult - Number that represents an error code from objRegistry.
	'
	' Returns:
	'
	'	Returns a generic error message string if error is known.
	'
	' See Also:
	'
	'	<Validate>
	'	<LoadProfileHive>
	'
	Private Function CheckError(iResult)
		Select Case iResult
	        Case -1
	        	CheckError = "Invalid Path"
	        Case -2
	        	CheckError = "Invalid HKey"
	        Case -3
	        	CheckError = "Invalid Key Path. Add " & Quotes("\") & " at end for keys"
			Case -4
				CheckError = "Permission denied"
			Case -5
				CheckError = "OS arch mismatch"
			Case -6
				CheckError = "Incoming value not valid"
			Case -7
				CheckError = "Invalid data type value sent"
	    End Select
	End Function
End Class