Option Explicit

''' <summary>Contains information about run-time errors. Accepts the Raise and Clear methods for generating and clearing run-time errors.</summary>
Class Err

	Sub Clear()
	End Sub

	Property Get Description
	End Property

	Property Get HelpContext
	End Property

	Property Get HelpFile
	End Property

	Property Get Number
	End Property

	Sub Raise(number)
	End Sub
	
	Sub Raise(number, source, description, helpfile, helpcontext)
	End Sub

	Property Get Source
	End Property
	
End Class

Class FileSystemObject

	Function BuildPath(path, name) ' As String
	End Function

	Sub CopyFile(source, destination)
	End Sub
	Sub CopyFile(source, destination, overwrite)
	End Sub

	Sub CopyFolder(source, destination)
	End Sub
	Sub CopyFolder(source, destination, overwrite)
	End Sub

	Function CreateFolder(foldername) ' As Folder
	End Function

	Function CreateTextFile(filename) ' As TextStream
	End Function
	Function CreateTextFile(filename, overwrite) ' As TextStream
	End Function
	Function CreateTextFile(filename, overwrite, unicode) ' As TextStream
	End Function

	Sub DeleteFile(filename)
	End Sub
	Sub DeleteFile(filename, force)
	End Sub

	Sub DeleteFolder(filename)
	End Sub
	Sub DeleteFolder(filename, force)
	End Sub

	Property Get Drives ' As DriveCollection
	End Property

	Function DriveExists(drive) ' As Boolean
	End Function

	Function FileExists(filename) ' As Boolean
	End Function

	Function FolderExists(foldername) ' As Boolean
	End Function

	Function GetAbsolutePathName(path) ' As String
	End Function

	Function GetBaseName(path) ' As String
	End Function

	Function GetDrive(drive) ' As Drive
	End Function

	Function GetDriveName(drive) ' As String
	End Function

	Function GetExtensionName(path) ' As String
	End Function

	Function GetFile(filename) ' As File
	End Function

	Function GetFileName(filename) ' As String
	End Function

	Function GetFileVersion(filename) ' As String
	End Function

	Function GetFolder(foldername) ' As Folder
	End Function

	Function GetParentFolderName(foldername) ' As String
	End Function

	Function GetSpecialFolder(folderspec) ' As Folder
	End Function
	
	Function GetStandardStream(StandardStreamType) ' As TextStream
	End Function
	Function GetStandardStream(StandardStreamType, Unicode) ' As TextStream
	End Function

	Function GetTempName() ' As String
	End Function

	Sub MoveFile(source, destination)
	End Sub

	Sub MoveFolder(source, destination)
	End Sub

	Function OpenTextFile(filename) ' As TextStream
	End Function
	Function OpenTextFile(filename, iomode) ' As TextStream
	End Function
	Function OpenTextFile(filename, iomode, create) ' As TextStream
	End Function
	Function OpenTextFile(filename, iomode, create, format) ' As TextStream
	End Function
	
End Class
