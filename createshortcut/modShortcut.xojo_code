#tag Module
Protected Module modShortcut
	#tag Method, Flags = &h0, CompatibilityFlags = (TargetDesktop and (Target32Bit or Target64Bit))
		Function CreateShortcut(Extends poOrigin As FolderItem, poShortcutFile As FolderItem, psLinuxDisplayname As String = "", poLinuxIconFile As FolderItem = nil) As Boolean
		  #If TargetMacOS Or TargetWindows Then
		    #Pragma unused psLinuxDisplayname
		  #EndIf
		  
		  Try
		    'Check Origin
		    If (poOrigin = Nil) Then Return False
		    If (poOrigin.Exists = False) Then Return False
		    
		    'Check Shortcut File
		    If (poShortcutFile = Nil) Or (poShortcutFile.Parent = Nil) Then Return False
		    Var oShortcutInFolder As FolderItem = poShortcutFile.Parent
		    If (Not oShortcutInFolder.Exists) Or (Not oShortcutInFolder.IsFolder) Then Return False
		    
		    'Check Shortcut Filename
		    Var sShortcutFilename As String = poShortcutFile.Name
		    If (sShortcutFilename = "") Then sShortcutFilename = poOrigin.DisplayName
		    If (sShortcutFilename = "") Then Return False
		    
		    #If TargetWindows Then
		      If (sShortcutFilename.NthField(".", sShortcutFilename.CountFields(".")) <> "lnk") Then
		        Break 'Expected a File with Extension .lnk
		        Return False
		      End If
		    #ElseIf TargetLinux Then
		      If (sShortcutFilename.NthField(".", sShortcutFilename.CountFields(".")) <> "desktop") Then
		        Break 'Expected a File with Extension .desktop
		        Return False
		      End If
		    #EndIf
		    
		    'Note: The Shortcut File gets overwritten if it already exist
		    
		    #If TargetWindows Then
		      #Pragma unused poLinuxIconFile
		      
		      Try
		        //https://docs.microsoft.com/en-us/troubleshoot/windows-client/admin-development/create-desktop-shortcut-with-wsh
		        //Windows Script Host
		        Var oOLEObject As New OLEObject("{F935DC22-1CF0-11D0-ADB9-00C04FD58A0B}")
		        If (oOLEObject = Nil) Then Return False
		        
		        Var oOLEShortcutObject As OLEObject = oOLEObject.CreateShortcut(poShortcutFile.NativePath)
		        If (oOLEShortcutObject = Nil) Then Return False
		        
		        oOLEShortcutObject.Description = sShortcutFilename
		        oOLEShortcutObject.TargetPath = poOrigin.NativePath
		        If poOrigin.IsFolder Then
		          oOLEShortcutObject.WorkingDirectory = poOrigin.NativePath
		        Else
		          If (poOrigin.Parent <> Nil) Then
		            oOLEShortcutObject.WorkingDirectory = poOrigin.Parent.NativePath
		          Else
		            oOLEShortcutObject.WorkingDirectory = ""
		          End If
		        End If
		        oOLEShortcutObject.Save
		        
		        Return ((poShortcutFile <> Nil) And poShortcutFile.Exists)
		        
		      Catch errOLE As OLEException
		        Return False
		      End Try
		    #EndIf
		    
		    #If TargetMacOS Then
		      #Pragma unused poLinuxIconFile
		      
		      //https://developer.apple.com/documentation/foundation/nsurl/1408532-writebookmarkdata?language=objc
		      Declare Function NSClassFromString Lib "Cocoa" (className As CFStringRef) As Ptr
		      Declare Function fileURLWithPath Lib "Foundation" selector "fileURLWithPath:" (ptrNSURLClass As Ptr, path As CFStringRef) As Ptr
		      Declare Function bookmarkDataWithOptions Lib "Foundation" selector "bookmarkDataWithOptions:includingResourceValuesForKeys:relativeToURL:error:" (ptrNSURL As Ptr, options As Integer, keys As Ptr, relativeURL As Ptr, error As Ptr) As Ptr
		      Declare Function writeBookmarkData Lib "Foundation" selector "writeBookmarkData:toURL:options:error:" (ptrNSURLClass As Ptr, bookmarkData As Ptr, bookmarkFileURL As Ptr, options As Integer, error As Ptr) As Boolean
		      
		      Const kNSURLBookmarkCreationSuitableForBookmarkFile = 1024
		      
		      Var ptrNSURLClass As Ptr = NSClassFromString("NSURL")
		      If (ptrNSURLClass = Nil) Then Return False
		      
		      Var ptrAppURL As Ptr = fileURLWithPath(ptrNSURLClass, poOrigin.NativePath)
		      If (ptrAppURL = Nil) Then Return False
		      
		      Var ptrAliasURL As Ptr = fileURLWithPath(ptrNSURLClass, poShortcutFile.NativePath)
		      If (ptrAliasURL = Nil) Then Return False
		      
		      Var ptrBookmarkData As Ptr = bookmarkDataWithOptions(ptrAppURL, kNSURLBookmarkCreationSuitableForBookmarkFile, Nil, Nil, Nil)
		      If (ptrBookmarkData = Nil) Then Return False
		      
		      Return writeBookmarkData(ptrNSURLClass, ptrBookmarkData, ptrAliasURL, kNSURLBookmarkCreationSuitableForBookmarkFile, Nil)
		    #EndIf
		    
		    #If TargetLinux Then
		      Var sExeFile As String = poOrigin.NativePath 'unescaped
		      Var sIconFile As String = ""
		      If (poLinuxIconFile <> Nil) And (Not poLinuxIconFile.IsFolder) And poLinuxIconFile.Exists Then
		        sIconFile = poLinuxIconFile.NativePath 'unescaped
		      End If
		      
		      Var sName As String = If(psLinuxDisplayname <> "", psLinuxDisplayname, poShortcutFile.Name)
		      If (sName.RightBytes(8) = ".desktop") Then sName = sName.LeftBytes(sName.Bytes - 8) 'remove .desktop
		      
		      Var sContent As String = "[Desktop Entry]" + EndOfLine.UNIX + _
		      "Encoding=UTF-8" + EndOfLine.UNIX + _
		      "Name=" + sName + EndOfLine.UNIX + _
		      "Exec=" + sExeFile + EndOfLine.UNIX + _
		      "Icon=" + sIconFile + EndOfLine.UNIX + _
		      "Terminal=false" + EndOfLine.UNIX + _
		      "Type=Application" + EndOfLine.UNIX + _
		      "Categories=Office"
		      sContent = ConvertEncoding(sContent, Encodings.UTF8)
		      
		      Var oStream As TextOutputStream
		      Try
		        oStream = TextOutputStream.Create(poShortcutFile)
		        oStream.Encoding = Encodings.UTF8
		      Catch errIO As IOException
		        Return False
		      End Try
		      
		      If (oStream = Nil) Then Return False
		      oStream.Write(sContent)
		      oStream.Close
		      
		      Var shlHost As New Shell
		      shlHost.Execute "chmod 755 " + poShortcutFile.ShellPath
		      shlHost.Close
		      
		      'Note: Depending on the Linux Distribution you need to manually
		      '      right click the created file and choose 'Allow Launching'
		      Return ((poShortcutFile <> Nil) And poShortcutFile.Exists)
		    #EndIf
		    
		    
		  Catch errCreateShortcut As RuntimeException
		    'Shortcut could not be created
		    'ignore...
		    Return False
		    
		  Finally
		    'if we get here, the Shortcut has not been created
		    Return False
		    
		  End Try
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetLinuxUserLocalSharedAppsFolder() As FolderItem
		  Var oFolder As FolderItem = SpecialFolder.UserHome
		  If (oFolder = Nil) Or (Not oFolder.IsFolder) Then Return Nil
		  
		  oFolder = oFolder.Child(".local")
		  If (oFolder = Nil) Or (Not oFolder.IsFolder) Then Return Nil
		  
		  oFolder = oFolder.Child("share")
		  If (oFolder = Nil) Or (Not oFolder.IsFolder) Then Return Nil
		  
		  oFolder = oFolder.Child("applications")
		  If (oFolder = Nil) Or (Not oFolder.IsFolder) Then Return Nil
		  
		  Return oFolder
		  
		End Function
	#tag EndMethod


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
