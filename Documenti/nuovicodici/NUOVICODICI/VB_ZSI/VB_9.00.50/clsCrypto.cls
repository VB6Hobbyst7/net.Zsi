VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Const CRYPT_NEWKEYSET = &H8
Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Private Const PROV_RSA_FULL = 1
Private Const ALG_CLASS_DATA_ENCRYPT = 24576
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_TYPE_ANY = 0
Private Const ALG_TYPE_BLOCK = 1536
Private Const ALG_TYPE_STREAM = 2048
Private Const ALG_SID_RC2 = 2
Private Const ALG_SID_RC4 = 1
Private Const ALG_SID_MD5 = 3
Private Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC2 = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_RC2)
Private Const CALG_RC4 = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
Private Const ENCRYPT_ALGORITHM = CALG_RC4
Private Const ENCRYPT_BLOCK_SIZE = 1
Private Const CRYPT_EXPORTABLE = 1

Public Function Encrypt(ByVal InputBuffer As String, ByVal EncryptionKey As String, Optional ByVal UseUUEncoding As Boolean = False) As String
  Dim lHHash As Long
  Dim lHkey As Long
  Dim lResult As Long
  Dim lHExchgKey As Long
  Dim lHCryptprov As Long
  Dim sContainer As String
  Dim lCryptLength As Long
  Dim lCryptBufLen As Long
  Dim sCryptBuffer As String
  Dim sProvider As String
  Dim lngRet As Long
  If InputBuffer = "" Then Exit Function
  On Error Resume Next
  sProvider = MS_DEF_PROV & vbNullChar
  lngRet = CryptAcquireContext(lHCryptprov, 0&, sProvider, PROV_RSA_FULL, 0)
  If lngRet = 0 Then
    lngRet = CryptAcquireContext(lHCryptprov, 0&, sProvider, PROV_RSA_FULL, CRYPT_NEWKEYSET)
  End If
  If lngRet Then
    If CryptCreateHash(lHCryptprov, CALG_MD5, 0, 0, lHHash) Then
      If CryptHashData(lHHash, EncryptionKey, Len(EncryptionKey), 0) Then
        If CryptDeriveKey(lHCryptprov, ENCRYPT_ALGORITHM, lHHash, 0, lHkey) Then
          CryptDestroyHash (lHHash)
          lHHash = 0
          lCryptLength = Len(InputBuffer)
          lCryptBufLen = lCryptLength * 2
          sCryptBuffer = String(lCryptBufLen, vbNullChar)
          LSet sCryptBuffer = InputBuffer
          If CryptEncrypt(lHkey, 0, 1, 0, sCryptBuffer, lCryptLength, lCryptBufLen) Then
            sCryptBuffer = Mid$(sCryptBuffer, 1, lCryptLength)
            If UseUUEncoding Then
              sCryptBuffer = UUEncode(sCryptBuffer)
            End If
            Encrypt = sCryptBuffer
          End If
        End If
      End If
    End If
  End If
  If Err.LastDllError Then Err = Err.LastDllError
  If lHkey Then lResult = CryptDestroyKey(lHkey)
  If lHExchgKey Then CryptDestroyKey (lHExchgKey)
  If lHHash Then CryptDestroyHash (lHHash)
  If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)
  If Err Then Err.Raise Err, "CryptoAPI.Encrypt"
End Function

Public Function Decrypt(ByVal InputBuffer As String, ByVal EncryptionKey As String, Optional ByVal UUEncoded As Boolean = False) As String
  Dim lHExchgKey As Long
  Dim lHCryptprov As Long
  Dim lHHash As Long
  Dim lHkey As Long
  Dim lResult As Long
  Dim sProvider As String
  Dim sCryptBuffer As String
  Dim lCryptBufLen As Long
  Dim lCryptPoint As Long
  Dim lPasswordPoint As Long
  Dim lPasswordCount As Long
  If InputBuffer = "" Then Exit Function
  On Error Resume Next
  If UUEncoded Then
    InputBuffer = UUDecode(InputBuffer)
  End If
  sProvider = vbNullChar
  sProvider = MS_DEF_PROV & vbNullChar
  If CryptAcquireContext(lHCryptprov, 0&, sProvider, PROV_RSA_FULL, 0) Then
    If CryptCreateHash(lHCryptprov, CALG_MD5, 0, 0, lHHash) Then
      If CryptHashData(lHHash, EncryptionKey, Len(EncryptionKey), 0) Then
        If CryptDeriveKey(lHCryptprov, ENCRYPT_ALGORITHM, lHHash, 0, lHkey) Then
          CryptDestroyHash (lHHash)
          lHHash = 0
          lCryptBufLen = Len(InputBuffer) * 2
          sCryptBuffer = String(lCryptBufLen, vbNullChar)
          LSet sCryptBuffer = InputBuffer
          If CryptDecrypt(lHkey, 0, 1, 0, sCryptBuffer, lCryptBufLen) Then
            Decrypt = Mid$(sCryptBuffer, 1, lCryptBufLen / 2)
          End If
        End If
      End If
    End If
  End If
  If Err.LastDllError Then Err = Err.LastDllError
  If (lHkey) Then lResult = CryptDestroyKey(lHkey)
  If lHExchgKey Then CryptDestroyKey (lHExchgKey)
  If lHHash Then CryptDestroyHash (lHHash)
  If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)
  If Err Then Err.Raise Err, "CryptoAPI.Decrypt"
End Function

Public Function UUEncode(ByVal InputBuffer As String) As String
  Dim bytChar() As Byte
  Dim bytBuffer() As Byte
  Dim intStrPos As Integer
  Dim strOutBuffer As String
  On Error Resume Next
  For intStrPos = 1 To Len(InputBuffer) Step 3
    bytBuffer = StrConv(Mid(InputBuffer, intStrPos, 3), vbFromUnicode)
    ReDim Preserve bytChar(IIf(UBound(bytBuffer) = 3, 3, UBound(bytBuffer) + 1))
    If UBound(bytBuffer) < 3 Then ReDim Preserve bytBuffer(UBound(bytBuffer) + 1)
    bytChar(0) = Int(bytBuffer(0) / 4)
    bytChar(0) = IIf(bytChar(0), (bytChar(0) And 63) + 32, 96)
    bytChar(1) = ((bytBuffer(0) * 16) And 48) Or (Int(bytBuffer(1) / 16) And 15)
    bytChar(1) = IIf(bytChar(1), (bytChar(1) And 63) + 32, 96)
    bytChar(2) = ((bytBuffer(1) * 4) And 60) Or (Int(bytBuffer(2) / 64) And 3)
    bytChar(2) = IIf(bytChar(2), (bytChar(2) And 63) + 32, 96)
    bytChar(3) = (bytBuffer(2) And 63)
    bytChar(3) = IIf(bytChar(3), (bytChar(3) And 63) + 32, 96)
    strOutBuffer = strOutBuffer & StrConv(bytChar, vbUnicode)
  Next intStrPos
  UUEncode = strOutBuffer
  Err.Clear
End Function

Public Function UUDecode(ByVal InputBuffer As String) As String
  Dim bytChar() As Byte
  Dim intLen As Integer
  Dim strOutBuffer As String
  Dim intFill As Integer
  On Error Resume Next
  intLen = 1
  intFill = Len(InputBuffer) Mod 4
  InputBuffer = InputBuffer & String(intFill, "`")
  Do While intLen <= Len(InputBuffer)
    bytChar = StrConv(Mid(InputBuffer, intLen, 4), vbFromUnicode)
    ReDim Preserve bytChar(3)
    strOutBuffer = strOutBuffer & Chr(((((bytChar(0) - 32) And 63) * 4) Or Int(((bytChar(1) - 32) And 63) / 16)) And &HFF)
    strOutBuffer = strOutBuffer & Chr(((((bytChar(1) - 32) And 63) * 16) Or Int(((bytChar(2) - 32) And 63) / 4)) And &HFF)
    strOutBuffer = strOutBuffer & Chr(((((bytChar(2) - 32) And 63) * 64) Or ((bytChar(3) - 32) And 63)) And &HFF)
    intLen = intLen + 4
  Loop
  Select Case intFill
    Case 1: UUDecode = Left(strOutBuffer, Len(strOutBuffer) - intFill)
    Case 2: UUDecode = Left(strOutBuffer, Len(strOutBuffer) - intFill)
    Case 3: UUDecode = Left(strOutBuffer, Len(strOutBuffer) - intFill - 1)
    Case Else: UUDecode = strOutBuffer
  End Select
  Err.Clear
End Function



