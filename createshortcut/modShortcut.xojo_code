#tag Module
Protected Module modShortcut
	#tag Method, Flags = &h0
		Function CreateShortcut(Extends poOrigin As FolderItem, poShortcutInFolder As FolderItem, psShortcutCaption As String, poLinuxIconFile As FolderItem) As Boolean
		  Try
		    'Check Origin
		    If (poOrigin = Nil) Then Return False
		    If (poOrigin.Exists = False) Then Return False
		    
		    'Check Destination Folder
		    If (poShortcutInFolder = Nil) Then Return False
		    If (poShortcutInFolder.Exists = False) Then Return False
		    If (poShortcutInFolder.Directory = False) Then Return False
		    
		    'Check Shortcut Caption
		    If (psShortcutCaption = "") Then psShortcutCaption = poOrigin.DisplayName
		    If (psShortcutCaption = "") Then Return False
		    
		    Dim sShortcutFilename As String = psShortcutCaption
		    #If TargetWindows Then
		      sShortcutFilename = sShortcutFilename + ".lnk"
		    #elseif TargetLinux then
		      sShortcutFilename = sShortcutFilename + ".desktop"
		    #EndIf
		    sShortcutFilename = ReplaceAll(sShortcutFilename, "\", "")
		    sShortcutFilename = ReplaceAll(sShortcutFilename, "/", "")
		    sShortcutFilename = ReplaceAll(sShortcutFilename, ":", "")
		    
		    'FolderItem for Shortcut (.TrueChild to get the Alias/Shortcut itself!)
		    Dim oShortcutFile As FolderItem = poShortcutInFolder.TrueChild(sShortcutFilename)
		    If (oShortcutFile = Nil) Then Return False
		    If oShortcutFile.Exists Then oShortcutFile.Delete
		    
		    
		    #If TargetWindows Then
		      #pragma unused poLinuxIconFile
		      
		      Try
		        Dim oOLEObject As New OLEObject("{F935DC22-1CF0-11D0-ADB9-00C04FD58A0B}")
		        If (oOLEObject = Nil) Then Return False
		        
		        Dim oOLEShortcutObject As OLEObject = oOLEObject.CreateShortcut(oShortcutFile.NativePath)
		        If (oOLEShortcutObject = Nil) Then Return False
		        
		        oOLEShortcutObject.Description = psShortcutCaption
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
		        
		        Return ((oShortcutFile <> Nil) And oShortcutFile.Exists)
		        
		      Catch errOLE As OLEException
		        Return False
		      End Try
		    #EndIf
		    
		    #If TargetMacOS Then
		      #pragma unused poLinuxIconFile
		      
		      Declare Function NSClassFromString Lib "Cocoa" (className As CFStringRef) As Ptr
		      Declare Function fileURLWithPath Lib "Foundation" selector "fileURLWithPath:" (ptrNSURLClass As Ptr, path As CFStringRef) As Ptr
		      Declare Function bookmarkDataWithOptions Lib "Foundation" selector "bookmarkDataWithOptions:includingResourceValuesForKeys:relativeToURL:error:" (ptrNSURL As Ptr, options As Integer, keys As Ptr, relativeURL As Ptr, error As Ptr) As Ptr
		      Declare Function writeBookmarkData Lib "Foundation" selector "writeBookmarkData:toURL:options:error:" (ptrNSURLClass As Ptr, bookmarkData As Ptr, bookmarkFileURL As Ptr, options As Integer, error As Ptr) As Boolean
		      
		      Const kNSURLBookmarkCreationSuitableForBookmarkFile = 1024
		      
		      Dim ptrNSURLClass As Ptr = NSClassFromString("NSURL")
		      If (ptrNSURLClass = Nil) Then Return False
		      
		      Dim ptrAppURL As Ptr = fileURLWithPath(ptrNSURLClass, poOrigin.NativePath)
		      If (ptrAppURL = Nil) Then Return False
		      
		      Dim ptrAliasURL As Ptr = fileURLWithPath(ptrNSURLClass, oShortcutFile.NativePath)
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
		      
		      Dim sContent As String = "[Desktop Entry]" + EndOfLine.UNIX + _
		      "Encoding=UTF-8" + EndOfLine.UNIX + _
		      "Name=" + psShortcutCaption + EndOfLine.UNIX + _
		      "Exec=" + sExeFile + EndOfLine.UNIX + _
		      "Icon=" + sIconFile + EndOfLine.UNIX + _
		      "Terminal=false" + EndOfLine.UNIX + _
		      "Type=Application" + EndOfLine.UNIX + _
		      "Categories=Office"
		      sContent = ConvertEncoding(sContent, Encodings.UTF8)
		      
		      Dim oStream As TextOutputStream
		      Try
		        oStream = TextOutputStream.Create(oShortcutFile)
		      Catch errIO As IOException
		        Return False
		      End Try
		      
		      If (oStream = Nil) Then Return False
		      oStream.Write(sContent)
		      oStream.Close
		      
		      Dim shlHost As New Shell
		      shlHost.Execute "chmod 755 " + oShortcutFile.ShellPath
		      shlHost.Close
		      
		      Return ((oShortcutFile <> Nil) And oShortcutFile.Exists)
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
