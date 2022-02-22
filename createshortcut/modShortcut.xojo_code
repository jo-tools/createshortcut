#tag Module
Protected Module modShortcut
	#tag Method, Flags = &h0, CompatibilityFlags = (TargetDesktop and (Target32Bit or Target64Bit))
		Function CreateShortcut(Extends poOrigin As FolderItem, poShortcutFile As FolderItem, poLinuxIconFile As FolderItem = nil) As Boolean
		  Try
		    'Check Origin
		    If (poOrigin = Nil) Then Return False
		    If (poOrigin.Exists = False) Then Return False
		    
		    'Check Shortcut File
		    If (poShortcutFile = Nil) Or (poShortcutFile.Parent = Nil) Then Return False
		    Dim oShortcutInFolder As FolderItem = poShortcutFile.Parent
		    If (Not oShortcutInFolder.Exists) Or (Not oShortcutInFolder.Directory) Then Return False
		    
		    'Check Shortcut Filename
		    Dim sShortcutFilename As String = poShortcutFile.Name
		    If (sShortcutFilename = "") Then sShortcutFilename = poOrigin.DisplayName
		    If (sShortcutFilename = "") Then Return False
		    
		    #If TargetWindows Then
		      If (NthField(sShortcutFilename, ".", CountFields(sShortcutFilename, ".")) <> "lnk") Then
		        Break 'Expected a File with Extension .lnk
		        Return False
		      End If
		    #ElseIf TargetLinux Then
		      If (NthField(sShortcutFilename, ".", CountFields(sShortcutFilename, ".")) <> "desktop") Then
		        Break 'Expected a File with Extension .desktop
		        Return False
		      End If
		    #EndIf
		    
		    'Note: The Shortcut File gets overwritten if it already exist
		    
		    #If TargetWindows Then
		      #Pragma unused poLinuxIconFile
		      
		      Try
		        //https://docs.microsoft.com/en-us/troubleshoot/windows-client/admin-development/create-desktop-shortcut-with-wsh
		        //Windows Script Host Shell Object
		        Dim oOLEObject As New OLEObject("{F935DC22-1CF0-11D0-ADB9-00C04FD58A0B}")
		        If (oOLEObject = Nil) Then Return False
		        
		        Dim oOLEShortcutObject As OLEObject = oOLEObject.CreateShortcut(poShortcutFile.NativePath)
		        If (oOLEShortcutObject = Nil) Then Return False
		        
		        oOLEShortcutObject.Description = sShortcutFilename
		        oOLEShortcutObject.TargetPath = poOrigin.NativePath
		        If poOrigin.Directory Then
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
		      
		      Declare Function NSClassFromString Lib "Cocoa" (className As CFStringRef) As Ptr
		      Declare Function fileURLWithPath Lib "Foundation" selector "fileURLWithPath:" (ptrNSURLClass As Ptr, path As CFStringRef) As Ptr
		      Declare Function bookmarkDataWithOptions Lib "Foundation" selector "bookmarkDataWithOptions:includingResourceValuesForKeys:relativeToURL:error:" (ptrNSURL As Ptr, options As Integer, keys As Ptr, relativeURL As Ptr, error As Ptr) As Ptr
		      Declare Function writeBookmarkData Lib "Foundation" selector "writeBookmarkData:toURL:options:error:" (ptrNSURLClass As Ptr, bookmarkData As Ptr, bookmarkFileURL As Ptr, options As Integer, error As Ptr) As Boolean
		      
		      Const kNSURLBookmarkCreationSuitableForBookmarkFile = 1024
		      
		      Dim ptrNSURLClass As Ptr = NSClassFromString("NSURL")
		      If (ptrNSURLClass = Nil) Then Return False
		      
		      Dim ptrAppURL As Ptr = fileURLWithPath(ptrNSURLClass, poOrigin.NativePath)
		      If (ptrAppURL = Nil) Then Return False
		      
		      Dim ptrAliasURL As Ptr = fileURLWithPath(ptrNSURLClass, poShortcutFile.NativePath)
		      If (ptrAliasURL = Nil) Then Return False
		      
		      Dim ptrBookmarkData As Ptr = bookmarkDataWithOptions(ptrAppURL, kNSURLBookmarkCreationSuitableForBookmarkFile, Nil, Nil, Nil)
		      If (ptrBookmarkData = Nil) Then Return False
		      
		      Return writeBookmarkData(ptrNSURLClass, ptrBookmarkData, ptrAliasURL, kNSURLBookmarkCreationSuitableForBookmarkFile, Nil)
		    #EndIf
		    
		    #If TargetLinux Then
		      Dim sExeFile As String = poOrigin.NativePath 'unescaped
		      Dim sIconFile As String = ""
		      If (poLinuxIconFile <> Nil) And (Not poLinuxIconFile.Directory) And poLinuxIconFile.Exists Then
		        sIconFile = poLinuxIconFile.NativePath 'unescaped
		      End If
		      
		      Dim sName As String = poShortcutFile.Name
		      If (RightB(sName, 8) = ".desktop") Then sName = LeftB(sName, LenB(sName)-8) 'remove .desktop
		      
		      Dim sContent As String = "[Desktop Entry]" + EndOfLine.UNIX + _
		      "Encoding=UTF-8" + EndOfLine.UNIX + _
		      "Name=" + sName + EndOfLine.UNIX + _
		      "Exec=" + sExeFile + EndOfLine.UNIX + _
		      "Icon=" + sIconFile + EndOfLine.UNIX + _
		      "Terminal=false" + EndOfLine.UNIX + _
		      "Type=Application" + EndOfLine.UNIX + _
		      "Categories=Office"
		      sContent = ConvertEncoding(sContent, Encodings.UTF8)
		      
		      Dim oStream As TextOutputStream
		      Try
		        oStream = TextOutputStream.Create(poShortcutFile)
		        #If (XojoVersion >= 2019.02) Then
		          oStream.Encoding = Encodings.UTF8
		        #EndIf
		      Catch errIO As IOException
		        Return False
		      End Try
		      
		      If (oStream = Nil) Then Return False
		      oStream.Write(sContent)
		      oStream.Close
		      
		      Dim shlHost As New Shell
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


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
