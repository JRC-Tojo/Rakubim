VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EncryptDes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

'***************************************************************************************************
' Windows API 宣言（32bit / 64bit VBA 両対応）
'***************************************************************************************************
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32.dll" (ByRef phProv As LongPtr, ByVal pszContainer As LongPtr, ByVal pszProvider As LongPtr, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal pbData As LongPtr, ByVal dwDataLen As Long, ByVal hPubKey As LongPtr, ByVal dwFlags As Long, ByRef phKey As LongPtr) As Long
Private Declare PtrSafe Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
Private Declare PtrSafe Function CryptSetKeyParam Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal dwParam As Long, ByRef pbData As Any, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal hHash As LongPtr, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As LongPtr, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare PtrSafe Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal hHash As LongPtr, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As LongPtr, ByRef pdwDataLen As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long

'***************************************************************************************************
' API定数宣言
'***************************************************************************************************
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_3DES As Long = &H6603
Private Const KP_MODE As Long = 4
Private Const KP_MODE_CBC As Long = 1
Private Const KP_IV As Long = 1
Private Const CP_UTF8 As Long = 65001

'***************************************************************************************************
'定数宣言
'***************************************************************************************************
'初期ベクトル
Private Const INITIALIZATION_VECTOR = "12345678" '必ず8文字分
'暗号用共通鍵
Private Const TRIPLE_DES_KEY = "motorakudekirunj" '必ず16文字分

'***************************************************************************************************
'機 能：TripleDESによる暗号化（TripleDES暗号化⇒BASE64符号化）
'引 数：暗号化対象平文
'戻り値：暗号文（正常終了） or Null（異常終了)
'***************************************************************************************************
Public Function EncryptStringTripleDES(plain_string As String, Optional key As String = "") As Variant
    Dim hProv As LongPtr
    Dim hKey As LongPtr
    Dim plain_byte_data() As Byte
    Dim keyBytes() As Byte
    Dim ivBytes() As Byte
    Dim blob() As Byte
    Dim dataLen As Long
    Dim bufLen As Long
    Dim buffer() As Byte
    Dim mode As Long
    
    EncryptStringTripleDES = Null
    
    On Error GoTo FunctionError
    
    ' 文字列をUTF-8バイト配列に変換
    plain_byte_data = StringToUTF8(plain_string)
    keyBytes = StringToUTF8(GetEncryptKey(key))
    ivBytes = StringToUTF8(INITIALIZATION_VECTOR)
    
    ' 暗号化プロバイダの取得
    If CryptAcquireContextW(hProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then GoTo Cleanup
    
    ' キーBLOB作成と鍵のインポート
    blob = CreateKeyBlob(keyBytes)
    If CryptImportKey(hProv, VarPtr(blob(0)), UBound(blob) + 1, 0, 0, hKey) = 0 Then GoTo Cleanup
    
    ' CBCモードとIVの設定
    mode = KP_MODE_CBC
    Call CryptSetKeyParam(hKey, KP_MODE, mode, 0)
    Call CryptSetKeyParam(hKey, KP_IV, ivBytes(0), 0)
    
    ' バッファの準備（PKCS7パディング用に8バイトの倍数サイズを保持）
    On Error Resume Next
    dataLen = UBound(plain_byte_data) - LBound(plain_byte_data) + 1
    If Err.Number <> 0 Then dataLen = 0
    On Error GoTo FunctionError
    
    bufLen = ((dataLen \ 8) + 1) * 8
    ReDim buffer(0 To bufLen - 1)
    
    If dataLen > 0 Then
        CopyMemory buffer(0), plain_byte_data(0), dataLen
    End If
    
    ' 暗号化実行
    If CryptEncrypt(hKey, 0, 1, 0, VarPtr(buffer(0)), dataLen, bufLen) <> 0 Then
        ReDim Preserve buffer(0 To dataLen - 1)
        EncryptStringTripleDES = BytesToBase64(buffer)
    End If

Cleanup:
    If hKey <> 0 Then CryptDestroyKey hKey
    If hProv <> 0 Then CryptReleaseContext hProv, 0
    Exit Function

FunctionError:
    GoTo Cleanup
End Function

'***************************************************************************************************
'機 能：TripleDESによる復号化（BASE64復号化⇒DES復号化）
'引 数：暗号文
'戻り値：平文（正常終了） or Null（異常終了)
'***************************************************************************************************
Public Function DecryptStringTripleDES(encrypted_string As String, Optional key As String = "") As Variant
    Dim hProv As LongPtr
    Dim hKey As LongPtr
    Dim encrypted_byte_data() As Byte
    Dim plain_byte_data() As Byte
    Dim keyBytes() As Byte
    Dim ivBytes() As Byte
    Dim blob() As Byte
    Dim dataLen As Long
    Dim mode As Long
    
    DecryptStringTripleDES = Null
    
    On Error GoTo FunctionError
    
    If encrypted_string = "" Then Exit Function
    
    ' Base64デコード
    encrypted_byte_data = Base64toBytes(encrypted_string)
    keyBytes = StringToUTF8(GetEncryptKey(key))
    ivBytes = StringToUTF8(INITIALIZATION_VECTOR)
    
    ' 暗号化プロバイダの取得
    If CryptAcquireContextW(hProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then GoTo Cleanup
    
    ' キーBLOB作成と鍵のインポート
    blob = CreateKeyBlob(keyBytes)
    If CryptImportKey(hProv, VarPtr(blob(0)), UBound(blob) + 1, 0, 0, hKey) = 0 Then GoTo Cleanup
    
    ' CBCモードとIVの設定
    mode = KP_MODE_CBC
    Call CryptSetKeyParam(hKey, KP_MODE, mode, 0)
    Call CryptSetKeyParam(hKey, KP_IV, ivBytes(0), 0)
    
    On Error Resume Next
    dataLen = UBound(encrypted_byte_data) - LBound(encrypted_byte_data) + 1
    If Err.Number <> 0 Then dataLen = 0
    On Error GoTo FunctionError
    
    If dataLen = 0 Then GoTo Cleanup
    
    ' 復号化実行
    If CryptDecrypt(hKey, 0, 1, 0, VarPtr(encrypted_byte_data(0)), dataLen) <> 0 Then
        If dataLen > 0 Then
            ReDim plain_byte_data(0 To dataLen - 1)
            CopyMemory plain_byte_data(0), encrypted_byte_data(0), dataLen
            DecryptStringTripleDES = UTF8ToString(plain_byte_data)
        Else
            DecryptStringTripleDES = ""
        End If
    End If

Cleanup:
    If hKey <> 0 Then CryptDestroyKey hKey
    If hProv <> 0 Then CryptReleaseContext hProv, 0
    Exit Function

FunctionError:
    GoTo Cleanup
End Function

'***************************************************************************************************
' 内部補助関数: PLAINTEXTKEYBLOB (24バイト 3DES 鍵構造) の生成
'***************************************************************************************************
Private Function CreateKeyBlob(ByRef keyBytes() As Byte) As Byte()
    Dim blob(0 To 35) As Byte
    Dim i As Long
    
    ' BLOBHEADER
    blob(0) = &H8 ' PLAINTEXTKEYBLOB
    blob(1) = &H2 ' CUR_BLOB_VERSION
    blob(2) = &H0
    blob(3) = &H0
    blob(4) = &H3 ' ALG_ID = CALG_3DES (&H6603)
    blob(5) = &H66
    blob(6) = &H0
    blob(7) = &H0
    
    ' Key Length (24 bytes)
    blob(8) = 24
    blob(9) = 0
    blob(10) = 0
    blob(11) = 0
    
    ' K1 (8バイト)
    For i = 0 To 7
        blob(12 + i) = keyBytes(i)
    Next i
    ' K2 (8バイト)
    For i = 0 To 7
        blob(20 + i) = keyBytes(8 + i)
    Next i
    ' K1 (8バイト)
    For i = 0 To 7
        blob(28 + i) = keyBytes(i)
    Next i
    
    CreateKeyBlob = blob
End Function

'***************************************************************************************************
' 内部補助関数: 文字列(UTF-16) -> UTF-8バイト配列
'***************************************************************************************************
Private Function StringToUTF8(ByVal txt As String) As Byte()
    Dim length As Long
    Dim buffer() As Byte
    
    If Len(txt) > 0 Then
        length = WideCharToMultiByte(CP_UTF8, 0, StrPtr(txt), Len(txt), 0, 0, 0, 0)
        If length > 0 Then
            ReDim buffer(0 To length - 1)
            ' 修正: Callキーワードを追加
            Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(txt), Len(txt), VarPtr(buffer(0)), length, 0, 0)
        End If
    End If
    StringToUTF8 = buffer
End Function

'***************************************************************************************************
' 内部補助関数: UTF-8バイト配列 -> 文字列(UTF-16)
'***************************************************************************************************
Private Function UTF8ToString(ByRef utf8Bytes() As Byte) As String
    Dim length As Long
    Dim bufLen As Long
    
    On Error Resume Next
    bufLen = UBound(utf8Bytes) - LBound(utf8Bytes) + 1
    If Err.Number <> 0 Or bufLen <= 0 Then Exit Function
    On Error GoTo 0
    
    length = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8Bytes(0)), bufLen, 0, 0)
    If length > 0 Then
        UTF8ToString = String$(length, 0)
        ' 修正: Callキーワードを追加
        Call MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8Bytes(0)), bufLen, StrPtr(UTF8ToString), length)
    End If
End Function

'***************************************************************************************************
'関数名：BytesToBase64
'機 能：Byte配列→base64文字列への変換（Windows標準MSXML使用）
'***************************************************************************************************
Public Function BytesToBase64(varBytes() As Byte) As String
    With CreateObject("MSXML2.DomDocument").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = Replace(.Text, vbLf, "")
    End With
End Function

'***************************************************************************************************
'関数名：Base64toBytes
'機 能：base64文字列→Byte配列への変換（Windows標準MSXML使用）
'***************************************************************************************************
Public Function Base64toBytes(varStr As String) As Byte()
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
        .DataType = "bin.base64"
        .Text = varStr
        Base64toBytes = .nodeTypedValue
    End With
End Function

'***************************************************************************************************
'関数名：GetEncryptKey
'機 能：暗号化に用いる共通鍵を取得
'***************************************************************************************************
Private Function GetEncryptKey(key As String) As String
    GetEncryptKey = TRIPLE_DES_KEY
    If (key = "") Then Exit Function
    
    Const targetLength = 16
    Const padChar = "_"

    GetEncryptKey = Right(String(targetLength, padChar) & key, targetLength)
End Function