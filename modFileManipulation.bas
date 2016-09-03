Attribute VB_Name = "modFileManipulation"
'This module is used to gather the contents of a file quickly and to grab the MD5 of a file quickly by using API functions. Use this
'code in any projects you wish, no need to give credit. Please vote though.

'marcin@malwarebytes.org if you have any questions.

'Special thanks to Hossein Moradi for the optimizations with CryptHashData()

Option Explicit

Public Const OPEN_EXISTING As Long = 3
Public Const GENERIC_READ As Long = &H80000000
Public Const FILE_SHARE_READ As Long = &H1
Public Const PROV_RSA_FULL As Long = 1
Public Const CRYPT_VERIFYCONTEXT = &HF0000000
Public Const HP_HASHVAL As Long = 2
Public Const CALG_MD5 As Long = 32771
Public Const lMD5Length As Long = 16

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Public Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Public Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal pCryptHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pcbData As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Public Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long

Public Sub hasher()
    Dim lStart&, lEnd&, sMD5$, i&
    
        'Get the start time
        lStart = GetTickCount
        
        'Use the function, this is just an example, let us assume you have C:\Windows\Explorer.exe
        For i = 1 To 10
            sMD5 = GetMD5("C:\Windows\Explorer.exe")
        Next i
        
        'Get the end time
        lEnd = GetTickCount
        
        'Display results
        MsgBox "The hash of Explorer.exe is " & sMD5 & ". It was acquired 10 times in " & (lEnd - lStart) / 1000 & _
               " seconds. Please vote for my code!" & vbNewLine & vbNewLine & "If you have any ideas on optimizing this code, please contact me.", vbInformation
End Sub

Public Function GetMD5(sFile$) As String
    Dim hFile&, uBuffer() As Byte, lFileSize&, lBytesRead&, uMD5(lMD5Length) As Byte
    Dim i&, hCrypt&, hHash&, sMD5$

    'Get a handle to the file
    hFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    
    'Check if file opened successfully
    If hFile > 0 Then
        'Get the file size
        lFileSize = GetFileSize(hFile, ByVal 0&)
    
        'File size must be greater than 0
        If lFileSize > 0 Then
            'Prepare the buffer
            ReDim uBuffer(lFileSize - 1)
    
            'Read the file
            If ReadFile(hFile, uBuffer(0), lFileSize, lBytesRead, ByVal 0&) <> 0 Then
                If lBytesRead <> lFileSize Then
                    ReDim Preserve uBuffer(lBytesRead - 1)
                End If
                
                'Acquire the context, create the hash, and hash the data
                If CryptAcquireContext(hCrypt, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
                    If CryptCreateHash(hCrypt, CALG_MD5, 0&, 0&, hHash) <> 0 Then
                        If CryptHashData(hHash, uBuffer(0), lBytesRead, ByVal 0&) <> 0 Then
                            If CryptGetHashParam(hHash, HP_HASHVAL, uMD5(0), lMD5Length, 0) <> 0 Then
                                'Build the MD5 string
                                For i = 0 To lMD5Length - 1
                                    sMD5 = sMD5 & (Right$("0" & Hex$(uMD5(i)), 2))
                                Next i
                            End If
                        End If

                        'Destroy the hash
                        CryptDestroyHash hHash
                    End If
                    
                    'Release the context
                    CryptReleaseContext hCrypt, 0
                End If
            End If
        End If
        
        'Close the handle to the file
        CloseHandle hFile
    End If
    
    'Convert to lower case
    GetMD5 = LCase$(sMD5)
End Function
