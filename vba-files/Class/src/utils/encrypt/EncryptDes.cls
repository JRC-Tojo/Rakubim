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
Private Declare PtrSafe Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwLen As Long, ByRef pbBuffer As Any) As Long
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
Private Const IV_LENGTH As Long = 8 '3DESのブロック長（バイト）

'***************************************************************************************************
' エラーコード宣言（vbObjectErrorを基点にした固有のエラー番号）
' 呼び出し側（RepoSecret等）がErr.Number / Err.Descriptionから失敗原因を判別できるようにするため、
' 失敗箇所ごとに個別のコードを割り当てる
'***************************************************************************************************
Private Const ERR_BASE As Long = vbObjectError + 5100
Private Const ERR_ACQUIRE_CONTEXT As Long = ERR_BASE + 1 '暗号化プロバイダの取得に失敗
Private Const ERR_GEN_RANDOM As Long = ERR_BASE + 2 '初期化ベクトルの生成に失敗
Private Const ERR_IMPORT_KEY As Long = ERR_BASE + 3 '暗号鍵のインポートに失敗
Private Const ERR_SET_MODE As Long = ERR_BASE + 4 '暗号モードの設定に失敗
Private Const ERR_SET_IV As Long = ERR_BASE + 5 '初期化ベクトルの設定に失敗
Private Const ERR_ENCRYPT As Long = ERR_BASE + 6 '暗号化処理に失敗
Private Const ERR_DECRYPT As Long = ERR_BASE + 7 '復号化処理に失敗
Private Const ERR_INVALID_DATA As Long = ERR_BASE + 8 '復号対象のデータ形式が不正

'***************************************************************************************************
'定数宣言
'***************************************************************************************************
'暗号用共通鍵
Private Const TRIPLE_DES_KEY = "motorakudekirunj" '必ず16文字分

'***************************************************************************************************
'機 能：TripleDESによる暗号化（TripleDES暗号化⇒BASE64符号化）
'引 数：plain_string (String) 暗号化対象平文
'       key (String, 省略可) 暗号化に用いる鍵の元となる文字列
'戻り値：暗号文（Base64文字列、先頭IV_LENGTHバイトに毎回ランダムな初期化ベクトルを含む）
'        失敗時はErr.Raiseで呼び出し元にエラーを伝播する
'***************************************************************************************************
Public Function EncryptStringTripleDES(plain_string As String, Optional key As String = "") As String
    Dim hProv As LongPtr
    Dim hKey As LongPtr
    Dim plain_byte_data() As Byte
    Dim keyBytes() As Byte
    Dim ivBytes(0 To IV_LENGTH - 1) As Byte
    Dim blob() As Byte
    Dim dataLen As Long
    Dim bufLen As Long
    Dim buffer() As Byte
    Dim combined() As Byte
    Dim mode As Long
    Dim errNum As Long, errSource As String, errDesc As String

    On Error GoTo ErrHandler

    ' 文字列をUTF-8バイト配列に変換
    plain_byte_data = StringToUTF8(plain_string)
    keyBytes = StringToUTF8(GetEncryptKey(key))

    ' 暗号化プロバイダの取得
    If CryptAcquireContextW(hProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        Err.Raise ERR_ACQUIRE_CONTEXT, "EncryptDes.EncryptStringTripleDES", "暗号化プロバイダの取得に失敗しました。"
    End If

    ' 暗号化ごとに異なるランダムな初期化ベクトルを生成
    ' （固定IVの使い回しを避け、同一平文から同一暗号文が生成されるパターン漏洩を防止する）
    If CryptGenRandom(hProv, IV_LENGTH, ivBytes(0)) = 0 Then
        Err.Raise ERR_GEN_RANDOM, "EncryptDes.EncryptStringTripleDES", "初期化ベクトルの生成に失敗しました。"
    End If

    ' キーBLOB作成と鍵のインポート
    blob = CreateKeyBlob(keyBytes)
    If CryptImportKey(hProv, VarPtr(blob(0)), UBound(blob) + 1, 0, 0, hKey) = 0 Then
        Err.Raise ERR_IMPORT_KEY, "EncryptDes.EncryptStringTripleDES", "暗号鍵のインポートに失敗しました。"
    End If

    ' CBCモードとIVの設定
    mode = KP_MODE_CBC
    If CryptSetKeyParam(hKey, KP_MODE, mode, 0) = 0 Then
        Err.Raise ERR_SET_MODE, "EncryptDes.EncryptStringTripleDES", "暗号モードの設定に失敗しました。"
    End If
    If CryptSetKeyParam(hKey, KP_IV, ivBytes(0), 0) = 0 Then
        Err.Raise ERR_SET_IV, "EncryptDes.EncryptStringTripleDES", "初期化ベクトルの設定に失敗しました。"
    End If

    ' バッファの準備（PKCS7パディング用に8バイトの倍数サイズを保持）
    dataLen = SafeArrayLength(plain_byte_data)
    bufLen = ((dataLen \ 8) + 1) * 8
    ReDim buffer(0 To bufLen - 1)
    If dataLen > 0 Then CopyMemory buffer(0), plain_byte_data(0), dataLen

    ' 暗号化実行
    If CryptEncrypt(hKey, 0, 1, 0, VarPtr(buffer(0)), dataLen, bufLen) = 0 Then
        Err.Raise ERR_ENCRYPT, "EncryptDes.EncryptStringTripleDES", "暗号化処理に失敗しました。"
    End If
    ReDim Preserve buffer(0 To dataLen - 1)

    ' 初期化ベクトルを暗号文の先頭に付与してBASE64化
    ReDim combined(0 To IV_LENGTH + dataLen - 1)
    CopyMemory combined(0), ivBytes(0), IV_LENGTH
    CopyMemory combined(IV_LENGTH), buffer(0), dataLen
    EncryptStringTripleDES = BytesToBase64(combined)

Cleanup:
    ReleaseCryptoHandles hKey, hProv
    Exit Function

ErrHandler:
    errNum = Err.Number
    errSource = Err.Source
    errDesc = Err.Description
    ReleaseCryptoHandles hKey, hProv
    Err.Raise errNum, errSource, errDesc
End Function

'***************************************************************************************************
'機 能：TripleDESによる復号化（BASE64復号化⇒DES復号化）
'引 数：encrypted_string (String) 暗号文（先頭IV_LENGTHバイトに初期化ベクトルを含むBase64文字列）
'       key (String, 省略可) 暗号化に用いた鍵の元となる文字列
'戻り値：平文（空文字列を渡した場合は空文字列を返す）
'        失敗時はErr.Raiseで呼び出し元にエラーを伝播する
'***************************************************************************************************
Public Function DecryptStringTripleDES(encrypted_string As String, Optional key As String = "") As String
    Dim hProv As LongPtr
    Dim hKey As LongPtr
    Dim combined() As Byte
    Dim ivBytes(0 To IV_LENGTH - 1) As Byte
    Dim encrypted_byte_data() As Byte
    Dim plain_byte_data() As Byte
    Dim keyBytes() As Byte
    Dim blob() As Byte
    Dim totalLen As Long
    Dim dataLen As Long
    Dim mode As Long
    Dim errNum As Long, errSource As String, errDesc As String

    If encrypted_string = "" Then
        DecryptStringTripleDES = ""
        Exit Function
    End If

    On Error GoTo ErrHandler

    ' Base64デコード（先頭IV_LENGTHバイトが初期化ベクトル、残りが暗号文）
    combined = Base64toBytes(encrypted_string)
    totalLen = SafeArrayLength(combined)
    dataLen = totalLen - IV_LENGTH
    If dataLen <= 0 Or (dataLen Mod 8) <> 0 Then
        Err.Raise ERR_INVALID_DATA, "EncryptDes.DecryptStringTripleDES", "復号対象のデータ形式が不正です。"
    End If

    CopyMemory ivBytes(0), combined(0), IV_LENGTH
    ReDim encrypted_byte_data(0 To dataLen - 1)
    CopyMemory encrypted_byte_data(0), combined(IV_LENGTH), dataLen

    keyBytes = StringToUTF8(GetEncryptKey(key))

    ' 暗号化プロバイダの取得
    If CryptAcquireContextW(hProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        Err.Raise ERR_ACQUIRE_CONTEXT, "EncryptDes.DecryptStringTripleDES", "暗号化プロバイダの取得に失敗しました。"
    End If

    ' キーBLOB作成と鍵のインポート
    blob = CreateKeyBlob(keyBytes)
    If CryptImportKey(hProv, VarPtr(blob(0)), UBound(blob) + 1, 0, 0, hKey) = 0 Then
        Err.Raise ERR_IMPORT_KEY, "EncryptDes.DecryptStringTripleDES", "暗号鍵のインポートに失敗しました。"
    End If

    ' CBCモードとIVの設定
    mode = KP_MODE_CBC
    If CryptSetKeyParam(hKey, KP_MODE, mode, 0) = 0 Then
        Err.Raise ERR_SET_MODE, "EncryptDes.DecryptStringTripleDES", "暗号モードの設定に失敗しました。"
    End If
    If CryptSetKeyParam(hKey, KP_IV, ivBytes(0), 0) = 0 Then
        Err.Raise ERR_SET_IV, "EncryptDes.DecryptStringTripleDES", "初期化ベクトルの設定に失敗しました。"
    End If

    ' 復号化実行
    If CryptDecrypt(hKey, 0, 1, 0, VarPtr(encrypted_byte_data(0)), dataLen) = 0 Then
        Err.Raise ERR_DECRYPT, "EncryptDes.DecryptStringTripleDES", "復号化処理に失敗しました。"
    End If

    If dataLen > 0 Then
        ReDim plain_byte_data(0 To dataLen - 1)
        CopyMemory plain_byte_data(0), encrypted_byte_data(0), dataLen
        DecryptStringTripleDES = UTF8ToString(plain_byte_data)
    Else
        DecryptStringTripleDES = ""
    End If

Cleanup:
    ReleaseCryptoHandles hKey, hProv
    Exit Function

ErrHandler:
    errNum = Err.Number
    errSource = Err.Source
    errDesc = Err.Description
    ReleaseCryptoHandles hKey, hProv
    Err.Raise errNum, errSource, errDesc
End Function

'***************************************************************************************************
'関数名：ReleaseCryptoHandles
'機 能：内部補助関数：CryptoAPIのキー・プロバイダハンドルを解放する
'引 数：hKey (LongPtr) CryptImportKeyで取得したキーハンドル
'       hProv (LongPtr) CryptAcquireContextWで取得したプロバイダハンドル
'戻り値：なし（解放後、引数のハンドルは0にクリアされる）
'***************************************************************************************************
Private Sub ReleaseCryptoHandles(ByRef hKey As LongPtr, ByRef hProv As LongPtr)
    If hKey <> 0 Then CryptDestroyKey hKey
    If hProv <> 0 Then CryptReleaseContext hProv, 0
    hKey = 0
    hProv = 0
End Sub

'***************************************************************************************************
'関数名：SafeArrayLength
'機 能：内部補助関数：未初期化の可能性があるByte配列の要素数を安全に取得する
'引 数：arr (Byte配列) 対象の配列
'戻り値：配列の要素数（未初期化の場合は0）
'***************************************************************************************************
Private Function SafeArrayLength(ByRef arr() As Byte) As Long
    On Error Resume Next
    SafeArrayLength = UBound(arr) - LBound(arr) + 1
    If Err.Number <> 0 Then SafeArrayLength = 0
    On Error GoTo 0
End Function

'***************************************************************************************************
'関数名：CreateKeyBlob
'機 能：内部補助関数：PLAINTEXTKEYBLOB (24バイト 3DES 鍵構造) の生成
'引 数：keyBytes (Byte配列) 16バイトの鍵データ（3DESの2キー運用のためK1をK3として複製する）
'戻り値：CryptImportKeyに渡せる36バイトのBLOBデータ（Byte配列）
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
'関数名：StringToUTF8
'機 能：内部補助関数：文字列(UTF-16) -> UTF-8バイト配列への変換
'引 数：txt (String) 変換対象の文字列
'戻り値：UTF-8でエンコードされたByte配列（空文字列を渡した場合は未初期化の配列）
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
'関数名：UTF8ToString
'機 能：内部補助関数：UTF-8バイト配列 -> 文字列(UTF-16)への変換
'引 数：utf8Bytes (Byte配列) UTF-8でエンコードされたバイト配列
'戻り値：変換後の文字列（配列が未初期化・空の場合は空文字列）
'***************************************************************************************************
Private Function UTF8ToString(ByRef utf8Bytes() As Byte) As String
    Dim length As Long
    Dim bufLen As Long

    bufLen = SafeArrayLength(utf8Bytes)
    If bufLen <= 0 Then Exit Function

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
'引 数：varBytes (Byte配列) 変換対象のバイト配列
'戻り値：Base64エンコードされた文字列（改行コードを含まない一行の文字列）
'***************************************************************************************************
Private Function BytesToBase64(varBytes() As Byte) As String
    With CreateObject("MSXML2.DomDocument").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = Replace(Replace(.Text, vbCr, ""), vbLf, "")
    End With
End Function

'***************************************************************************************************
'関数名：Base64toBytes
'機 能：base64文字列→Byte配列への変換（Windows標準MSXML使用）
'引 数：varStr (String) Base64エンコードされた文字列
'戻り値：デコードされたByte配列
'***************************************************************************************************
Private Function Base64toBytes(varStr As String) As Byte()
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
        .DataType = "bin.base64"
        .Text = varStr
        Base64toBytes = .nodeTypedValue
    End With
End Function

'***************************************************************************************************
'関数名：GetEncryptKey
'機 能：暗号化に用いる共通鍵を取得
'引 数：key (String) 鍵の元となる文字列（省略時は既定の共通鍵を使用）
'戻り値：16文字に整形された共通鍵文字列
'***************************************************************************************************
Private Function GetEncryptKey(key As String) As String
    GetEncryptKey = TRIPLE_DES_KEY
    If (key = "") Then Exit Function

    Const targetLength = 16
    Const padChar = "_"

    GetEncryptKey = Right(String(targetLength, padChar) & key, targetLength)
End Function
