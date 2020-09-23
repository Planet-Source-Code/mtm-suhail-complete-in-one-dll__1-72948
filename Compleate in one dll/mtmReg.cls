VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mtmRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' =====================
' = Type Declarations =
' =====================

Private Type FILETIME

lLowDateTime    As Long
lHighDateTime   As Long

End Type

' =========================
' = Variable Declarations =
' =========================

Private lDataSize As Long
Private lBufferSize As Long
Private ByteArray() As Byte
Private DisplayErrorMsg As Boolean
Private hKey As Long, MainKeyHandle As Long
Private rtn As Long, lBuffer As Long, sBuffer As String
 
' =============
' = Constants =
' =============
 
Private Const REG_SZ = 1&
Private Const REG_NONE = 0&
Private Const REG_LINK = 6&
Private Const REG_DWORD = 4&
Private Const REG_BINARY = 3&
Private Const REG_MULTI_SZ = 7&
Private Const REG_EXPAND_SZ = 2&
Private Const REG_RESOURCE_LIST = 8&

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1009&
Private Const ERROR_BADKEY = 1010&
Private Const ERROR_MORE_DATA = 234&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const REG_DWORD_BIG_ENDIAN = 5&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const REG_DWORD_LITTLE_ENDIAN = 4&
Private Const ERROR_INVALID_PARAMETER = 87&
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Private Const KEY_NOTIFY = &H10&
Private Const WRITE_DAC = &H40000
Private Const KEY_SET_VALUE = &H2&
Private Const WRITE_OWNER = &H80000
Private Const KEY_QUERY_VALUE = &H1&
'Private Const KEY_EXECUTE = KEY_READ
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const KEY_CREATE_LINK = &H20&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY

' =====================
' = API Declarations  =
' =====================

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lplDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long

Private Sub Class_Initialize()

'** This should be turned on if you want error message to show

'// sets DisplayErrorMsg to false

DisplayErrorMsg = True
    
End Sub

Public Property Let SetDisplayErrorMsg(vNewValue As Variant)

DisplayErrorMsg = vNewValue

End Property

Public Function SetDWordValue(ByVal sKey As String, ByVal sKeyName As String, ByVal KeyValue As Long)

'// **** Stores Dword Value Value in Regsitry ***** //
CreateKey sKey
SetDWordValue = False    '// sets SetDWordValue to false

'// calls the private function ParseKey

Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then    '// if main key handle

'// calls the api to open the registry key

rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_WRITE, hKey)
   
   If rtn = ERROR_SUCCESS Then    '//  if registry key opened successfully
   
      '// calls the api to store the Dword Value in registry
      
      rtn = RegSetValueExA(hKey, sKeyName, 0, REG_DWORD, KeyValue, 4)
      
      If Not rtn = ERROR_SUCCESS Then    '// if Dword Value saving fails
      
         If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
         
         '// calls the private function GetErrorMsg to get the correct error info
         
            MsgBox GetErrorMsg(rtn), vbExclamation
            
         End If
         
      Else    '// Dword Value value saving succeeded
      
         SetDWordValue = True    '// SetDWordValue to true
         
      End If
      
      '// calls the api to close the registry key
      
      rtn = RegCloseKey(hKey)
      
Else    '// opening the regstry key has failed
   
      If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
      
      '// calls the private function GetErrorMsg to get the correct error info
      
         MsgBox GetErrorMsg(rtn), vbExclamation
         
      End If
      
   End If
   
End If

End Function

Public Function GetDWordValue(ByVal sKey As String, ByVal sKeyName As String)

'// **** Gets Dword Value from Regsitry ***** //

'// calls the private function ParseKey

Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then     '// if main key handle

'// calls the api to open the registry key

rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)

   If rtn = ERROR_SUCCESS Then     '//  if registry key opened successfully
   
      '// call the api to query the Dword Value from registry
      
      rtn = RegQueryValueExA(hKey, sKeyName, 0, REG_DWORD, lBuffer, 4)
      
      If rtn = ERROR_SUCCESS Then    '// if quering the Dword Value succeeded
         
         '// calls the api to close the registry key
         
         rtn = RegCloseKey(hKey)
         
         GetDWordValue = lBuffer    '// GetDWordValue = queried Dword Value
         
      Else    '// quering Dword Value failed
      
         GetDWordValue = "Error"    '// GetDWordValue = error
         
         If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
         
         '// calls the private function GetErrorMsg to get the correct error info
         
            MsgBox GetErrorMsg(rtn), vbExclamation
            
         End If
         
      End If
      
Else    '// opening registry key has failed
   
      GetDWordValue = "Error"    '// GetDWordValue = error
      
      If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
      
      '// calls the private function GetErrorMsg to get the correct error info
      
         MsgBox GetErrorMsg(rtn), vbExclamation
         
      End If
      
   End If
   
End If

End Function

Public Function SetBinaryValue(ByVal sKey As String, ByVal sKeyName As String, KeyValue As String)

'// **** Stores Binary Value in Regsitry ***** //



Dim i As Long    '// declares local variable


CreateKey sKey

SetBinaryValue = False    '// sets SetBinaryValue to false

'// calls the private function ParseKey

Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then    '// if main key handle

'// calls the api to open the registry key

rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_WRITE, hKey)
   
   If rtn = ERROR_SUCCESS Then    '// if registry key opened successfully
   
      lDataSize = Len(KeyValue)    '// gets the length of the KeyValue
      
      ReDim ByteArray(lDataSize)    '// redims the bytearray to fit the lDataSize
      
      '// puts the KeyValue in bytearray
      
      For i = 1 To lDataSize
      ByteArray(i) = Asc(Mid$(KeyValue, i, 1))
      Next
      
      '// calls the api to stores the BianryValue in registry
      
      rtn = RegSetValueExB(hKey, sKeyName, 0, REG_BINARY, ByteArray(1), lDataSize)
      
      If Not rtn = ERROR_SUCCESS Then    '// if binary value saving fails
      
         If DisplayErrorMsg = True Then     '// if DisplayErrorMsg = true
         
         '// calls the private function GetErrorMsg to get the correct error info
         
            MsgBox GetErrorMsg(rtn), vbExclamation
            
         End If
         
      Else    '// BinaryValue saving succeeded
      
         SetBinaryValue = True    '// SetBinaryValue = true
         
      End If
      
      '// calls the api to close the registry key
      
      rtn = RegCloseKey(hKey)
      
Else    '// opening registry key has failed
   
      If DisplayErrorMsg = True Then     '// if DisplayErrorMsg = true
       
      '// calls the private function GetErrorMsg to get the correct error info
      
         MsgBox GetErrorMsg(rtn), vbExclamation
         
      End If
      
   End If
   
End If

End Function

Public Function GetBinaryValue(ByVal sKey As String, ByVal sKeyName As String)

'// **** Gets Binary Value from Regsitry ***** //

'// calls the private function ParseKey

Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then    '// if main key handle

'// calls the api to open the registry key
    
rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)

   If rtn = ERROR_SUCCESS Then     '// if registry key opened successfully
   
      lBufferSize = 1    '// lBufferSize = 1
      
      '// calls the api to open the registry key for binary reading
       
      rtn = RegQueryValueEx(hKey, sKeyName, 0, REG_BINARY, 0, lBufferSize)
      
      sBuffer = Space(lBufferSize)    '// determiens how much is the buffer size to hold the binary value
      
      '// calls the api to open the registry key for binary reading
      
      rtn = RegQueryValueEx(hKey, sKeyName, 0, REG_BINARY, sBuffer, lBufferSize)
      
      If rtn = ERROR_SUCCESS Then    '// if quering the binary value succeeded
         
         '// calls the api to close the registry key
         
         rtn = RegCloseKey(hKey)
         
         GetBinaryValue = sBuffer    '// GetBinaryValue = queried bianry value
         
      Else    '// quering binary value failed
      
         GetBinaryValue = "Error"    '// GetBinaryValue = error
         
         If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
            
         '// calls the private function GetErrorMsg to get the correct error info
            
            MsgBox GetErrorMsg(rtn), vbExclamation
            
         End If
         
      End If
      
Else    '// opening registry key has failed

      GetBinaryValue = "Error"    '// GetBinaryValue = error
      
      If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
          
      '// calls the private function GetErrorMsg to get the correct error info
      
         MsgBox GetErrorMsg(rtn), vbExclamation
         
      End If
      
   End If
   
End If

End Function

Public Function SetStringValue(ByVal sKey As String, ByVal sKeyName As String, ByVal KeyValue As String)

'// **** Stores String Value in Regsitry ***** //
CreateKey sKey
SetStringValue = False    '// sets SetStringValue to false

'// calls the private function ParseKey

Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then     '// if main key handle

'// calls the api to open the registry key

rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_WRITE, hKey)
   
   If rtn = ERROR_SUCCESS Then    '// if registry key opened successfully
      
      '// calls the api to store the String value in registry
      
      rtn = RegSetValueEx(hKey, sKeyName, 0, REG_SZ, ByVal KeyValue, Len(KeyValue))
      
      If Not rtn = ERROR_SUCCESS Then    '// if StringValue saving fails
      
         If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
         
            '// calls the private function GetErrorMsg to get the correct error info
         
            MsgBox GetErrorMsg(rtn), vbExclamation
            
         End If
         
      Else    '// StringValue saving succeeded
      
         SetStringValue = True    '// SetStringValue = true
         
      End If
      
      '// calls the api to close the registry key
      
      rtn = RegCloseKey(hKey)
      
Else    '// opening registry key has failed

      If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
       
      '// calls the private function GetErrorMsg to get the correct error info
      
         MsgBox GetErrorMsg(rtn), vbExclamation
         
      End If
      
   End If
   
End If

End Function

Public Function GetStringValue(ByVal sKey As String, ByVal sKeyName As String)

'// **** Gets String Value from Regsitry ***** //

lBufferSize = 0    '// sets lBufferSize = 0
sBuffer = ""    '// empties the sBuffer

'// calls the private function ParseKey

Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then    '// if main key handle

'// calls the api to open the registry key

rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)

   If rtn = ERROR_SUCCESS Then    '// if registry key opened successfully
   
      sBuffer = Space(255)    '// limit sBuffer to 255
      lBufferSize = Len(sBuffer)    '// lBufferSize = lenght of sBuffer
      
      '// calls the api to query the StringValue from registry
      
      rtn = RegQueryValueEx(hKey, sKeyName, 0, REG_SZ, sBuffer, lBufferSize)
      
      If rtn = ERROR_SUCCESS Then    '// if querying the StringValue Succeeded
      
      '// calls the api to close the registry key
      
         rtn = RegCloseKey(hKey)
         
         sBuffer = Trim(sBuffer)    '// clear the leading and ending sapces if any in sBuffer
         
         GetStringValue = Left(sBuffer, lBufferSize - 1)    '// sBuffer now holds only the StringValue
         
      Else    '// if querying the StringValue failed
      
         GetStringValue = "Error"    '// GetStringValue = error
         
         If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
         
         '// calls the private function GetErrorMsg to get the correct error info
         
            MsgBox GetErrorMsg(rtn), vbExclamation
            
         End If
         
      End If
      
Else    '// opening registry key has failed
   
      GetStringValue = "Error"    '// GetStringValue = error
      
      If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
      
      '// calls the private function GetErrorMsg to get the correct error info
      
         MsgBox GetErrorMsg(rtn), vbExclamation
         
      End If
      
   End If
   
End If

End Function

Public Function CreateKey(ByVal sKey As String)

'// **** Creates a Key in Regsitry ***** //

CreateKey = False    '// sets CreateKey to false

'// calls the private function ParseKey

Call ParseKey(sKey, MainKeyHandle)
    
If MainKeyHandle Then    '// if main key handle

'// calls the api to create the registry key

rtn = RegCreateKey(MainKeyHandle, sKey, hKey)
   
  If rtn = ERROR_SUCCESS Then    '// if registry key successfully created
  
  '// calls the api to close the registry key
  
     rtn = RegCloseKey(hKey)
     
     CreateKey = True    '// sets CreateKey to true
     
  End If
  
End If

End Function

Public Function DeleteKey(ByVal Keyname As String)

'// **** Deletes a Key from Regsitry ***** //

DeleteKey = False    '// sets DeleteKey to false

'// calls the private function ParseKey

Call ParseKey(Keyname, MainKeyHandle)
    
If MainKeyHandle Then    '// if main key handle

'// calls the api to delete the registry key

rtn = RegDeleteKey(MainKeyHandle, Keyname)

   If (rtn <> ERROR_SUCCESS) Then    '// if delete registry key failed
    
       If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
        
       '// calls the private function GetErrorMsg to get the correct error info
       
           MsgBox GetErrorMsg(rtn), vbExclamation, App.Title
            
       End If
        
   Else    '// if registry key successfully deleted
    
       DeleteKey = True    '// sets DeleteKey to true
        
   End If
    
End If
    
End Function

Public Function DeleteKeyValue(ByVal sKeyName As String, ByVal sValueName As String)

'// **** Deletes a KeyValue from Regsitry ***** //

DeleteKeyValue = False    '// sets DeleteKeyValue to false

Dim hKey As Long    '// declares local variable

'// calls the private function ParseKey

Call ParseKey(sKeyName, MainKeyHandle)

If MainKeyHandle Then    '// if main key handle

'// calls the api to open the registry key

rtn = RegOpenKeyEx(MainKeyHandle, sKeyName, 0, KEY_WRITE, hKey)

   If (rtn = ERROR_SUCCESS) Then    '// if registry key opened successfully
       
   '// calls the api to delete a KeyValue in registry
       
       rtn = RegDeleteValue(hKey, sValueName)
       
       If (rtn <> ERROR_SUCCESS) Then    '// if deleting the KeyValue from the registry failed
       
           If DisplayErrorMsg = True Then    '// if DisplayErrorMsg = true
           
           '// calls the private function GetErrorMsg to get the correct error info
           
               MsgBox GetErrorMsg(rtn), vbExclamation, App.Title
               
           End If
           
       Else    '// deleting the KeyValue from registry succeeded
       
           DeleteKeyValue = True    '// sets DeleteKeyValue to true
           
       End If
       
       '// calls the api to close the registry key
       
       rtn = RegCloseKey(hKey)

   End If

End If

End Function

Public Function KeyExist(ByVal sKey As String)

'// **** Checks if a Key exist in Regsitry ***** //

Dim hKey As Long    '// declares local variable

'// calls the private function ParseKey

Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then    '// if main key handle

'// calls the api to open the registry key

rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)

   If rtn = ERROR_SUCCESS Then    '// if registry key opened successfully
   
       KeyExist = True    '// sets KeyExist to true
       
   Else    '// opening registry key has failed
   
       KeyExist = False    '// sets KeyExist to false
       
   End If
   
End If
    
End Function

Public Function KeyValueExist(ByVal sKey As String, ByVal sKeyName As String)

'// **** Checks if a KeyValue exist in Regsitry ***** //

'// dacalres local variables

Dim hKey As Long
Dim lSize As Long
Dim sTmp As String
Dim lActualType As Long

'// calls the private function ParseKey

Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then    '// if main key handle
    
'// calls the api to open the registry key

rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)

   If (rtn = ERROR_SUCCESS) Then    '// if registry key opened successfully
        
   '// calls the api to query the registry KeyValue
   
       rtn = RegQueryValueEx(hKey, ByVal sKeyName, 0&, lActualType, sTmp, lSize)
       
       If (rtn = ERROR_SUCCESS) Then    '// if registry KeyValue query successfull
       
           KeyValueExist = True    '// sets KeyValueExist to true
           
       Else
       
           KeyValueExist = False    '// sets KeyValueExist to false
           
       End If
    
   End If
    
End If

End Function

Private Sub ParseKey(Keyname As String, Keyhandle As Long)
    
'// **** Private function ParseKey ***** //
    
rtn = InStr(Keyname, "\")      '// checks if passed Keyname has "\"

'// checks if passed Keyname has "HKEY_" and at the end "\"

If Left(Keyname, 5) <> "HKEY_" Or Right(Keyname, 1) = "\" Then

'// shows msgbox saying the passed Keyname is in incorrect format

MsgBox "Incorrect Format: " + Chr(10) + Chr(10) + Keyname
   
Exit Sub    '// exits the sub

ElseIf rtn = 0 Then    '// if Keyname has no "\"

'// calls the private function GetMainKeyHandle

   Keyhandle = GetMainKeyHandle(Keyname)
   
   Keyname = ""    '// empties Keyname
   
Else   '// if Keyname has "\"

'// calls the private function GetMainKeyHandle

   Keyhandle = GetMainKeyHandle(Left(Keyname, rtn - 1))
   
   Keyname = Right(Keyname, Len(Keyname) - rtn)
   
End If

End Sub

Private Function GetMainKeyHandle(MainKeyName As String) As Long

'// **** Private function GetMainKeyHandle ***** //

'// declares constants

Const HKEY_USERS = &H80000003
Const HKEY_DYN_DATA = &H80000006
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_PERFORMANCE_DATA = &H80000004
   
'// selects MainKeyName and returns GetMainKeyHandle according to the constants

Select Case MainKeyName

Case "HKEY_CLASSES_ROOT"

    GetMainKeyHandle = HKEY_CLASSES_ROOT
    
Case "HKEY_CURRENT_USER"

    GetMainKeyHandle = HKEY_CURRENT_USER
    
Case "HKEY_LOCAL_MACHINE"

    GetMainKeyHandle = HKEY_LOCAL_MACHINE
    
Case "HKEY_USERS"

    GetMainKeyHandle = HKEY_USERS
    
Case "HKEY_PERFORMANCE_DATA"

    GetMainKeyHandle = HKEY_PERFORMANCE_DATA
    
Case "HKEY_CURRENT_CONFIG"

    GetMainKeyHandle = HKEY_CURRENT_CONFIG
    
Case "HKEY_DYN_DATA"

    GetMainKeyHandle = HKEY_DYN_DATA
     
End Select    '// ends selection of MainKeyName

End Function

Private Function GetErrorMsg(lErrorCode As Long) As String
    
'// **** Private function GetErrorMsg ***** //
        
'// selects the lErrorCode and return the correct error information
        
Select Case lErrorCode

Case 1009, 1015

     GetErrorMsg = "The Registry Database Is Corrupt!"
     
Case 2, 1010

     GetErrorMsg = "Bad Key Name"
     
Case 1011

     GetErrorMsg = "Can't Open Key"
     
Case 4, 1012

     GetErrorMsg = "Can't Read Key"
     
Case 5

     GetErrorMsg = "Access To this key Is Denied"
     
Case 1013

     GetErrorMsg = "Can't Write Key"
     
Case 8, 14

     GetErrorMsg = "Out Of memory"
     
Case 87

     GetErrorMsg = "Invalid Parameter"
     
Case 234

     GetErrorMsg = "There Is More Data Than The Buffer Has Been Allocated To Hold"
     
Case Else

     GetErrorMsg = "Undefined Error Code : " & Str$(lErrorCode)
            
End Select

End Function
