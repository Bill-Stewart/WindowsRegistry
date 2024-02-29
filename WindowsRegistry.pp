{ Copyright (C) 2024 by Bill Stewart (bstewart at iname.com)

  This program is free software; you can redistribute it and/or modify it under
  the terms of the GNU Lesser General Public License as published by the Free
  Software Foundation; either version 3 of the License, or (at your option) any
  later version.

  This program is distributed in the hope that it will be useful, but WITHOUT
  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
  FOR A PARTICULAR PURPOSE. See the GNU General Lesser Public License for more
  details.

  You should have received a copy of the GNU Lesser General Public License
  along with this program. If not, see https://www.gnu.org/licenses/.

}

unit WindowsRegistry;

{$MODE OBJFPC}
{$MODESWITCH UNICODESTRINGS}

interface

uses
  Windows;

const
  HKEY_CURRENT_USER_64 = HKEY_CURRENT_USER or KEY_WOW64_64KEY;
  HKEY_CURRENT_USER_32 = HKEY_CURRENT_USER or KEY_WOW64_32KEY;
  HKEY_LOCAL_MACHINE_64 = HKEY_LOCAL_MACHINE or KEY_WOW64_64KEY;
  HKEY_LOCAL_MACHINE_32 = HKEY_LOCAL_MACHINE or KEY_WOW64_32KEY;

type
  LSTATUS = Integer;
  TByteArray = array of Byte;
  TStringArray = array of string;

function RegKeyExists(const ComputerName: string; RootKey: HKEY;
  const SubKeyName: string): LSTATUS;
function RegValueExists(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string): LSTATUS;

function RegCreateSubKey(const ComputerName: string; RootKey: HKEY;
  const SubKeyName: string): LSTATUS;

function RegGetSubKeyLastWriteTime(const ComputerName: string;
  RootKey: HKEY; const SubKeyName: string; var LastWriteTime: FILETIME): LSTATUS;
function RegGetValueType(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; out ValueType: DWORD): LSTATUS;

function RegGetSubKeyNames(const ComputerName: string; RootKey: HKEY;
  const SubKeyName: string; var Names: TStringArray): LSTATUS;
function RegGetValueNames(const ComputerName: string; RootKey: HKEY;
  const SubKeyName: string; var Names: TStringArray): LSTATUS;

//function RegRenameSubKey(const ComputerName: string; RootKey: HKEY;
//  const SubKeyName, NewName: string): LSTATUS;

function RegDeleteKeyIfEmpty(const ComputerName: string; RootKey: HKEY;
  const SubKeyName: string): LSTATUS;
function RegDeleteKeyIncludingSubKeys(const ComputerName: string;
  RootKey: HKEY; const SubKeyName: string): LSTATUS;
function RegDeleteValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string): LSTATUS;

function RegGetBinaryValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; var Bytes: TByteArray): LSTATUS;
function RegValueIsEmpty(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; out Empty: Boolean): LSTATUS;
function RegGetDWORDValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; out ValueData: DWORD): LSTATUS;
function RegGetStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; out ValueData: string): LSTATUS;
function RegGetExpandStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; out ValueData: string): LSTATUS;
function RegGetMultiStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; var Values: TStringArray): LSTATUS;

function RegSetBinaryValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; var Bytes: TByteArray): LSTATUS;
function RegSetDWORDValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; const ValueData: DWORD): LSTATUS;
function RegSetStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName, ValueData: string): LSTATUS;
function RegSetExpandStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName, ValueData: string): LSTATUS;
function RegSetMultiStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; var Values: TStringArray): LSTATUS;

implementation

//function RegRenameKey(RootKey: HKEY;
//  lpSubKeyName: LPCWSTR;
//  lpNewKeyName: LPCWSTR): LSTATUS;
//  stdcall; external 'advapi32.dll';

// Important: If this function succeeds (returns 0/ERROR_SUCCESS), caller
// must call RegCloseKey on the handle returned in the RegKeyHandle output
// parameter
function ConnectRegistry(ComputerName: string; RootKey: HKEY;
  out RegKeyHandle: HKEY): LSTATUS;
begin
  // Filter out KEY_WOW64_32KEY or KEY_WOW64_64KEY from RootKey if specified
  if (RootKey and KEY_WOW64_32KEY) <> 0 then
    RootKey := RootKey and (not KEY_WOW64_32KEY)
  else if (RootKey and KEY_WOW64_64KEY) <> 0 then
    RootKey := RootKey and (not KEY_WOW64_64KEY);
  result := RegConnectRegistryW(PChar(ComputerName),  // LPCWSTR lpMachineName
    RootKey,                                          // HKEY    hKey
    RegKeyHandle);                                    // PHKEY   phkResult
end;

// Updates AccessFlags (samDesired) to use 32-bit or 64-bit registry view if
// specified in RootKey
procedure UpdateAccessFlags(const RootKey: HKEY; var AccessFlags: REGSAM);
begin
  if (RootKey and KEY_WOW64_32KEY) <> 0 then
    AccessFlags := AccessFlags or KEY_WOW64_32KEY
  else if (RootKey and KEY_WOW64_64KEY) <> 0 then
    AccessFlags := AccessFlags or KEY_WOW64_64KEY;
end;

function RegKeyExists(const ComputerName: string; RootKey: HKEY;
  const SubKeyName: string): LSTATUS;
var
  RootKeyHandle, SubKeyHandle: HKEY;
  AccessFlags: REGSAM;
begin
  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
  if result <> ERROR_SUCCESS then
    exit;
  AccessFlags := KEY_READ;
  UpdateAccessFlags(RootKey, AccessFlags);
  result := RegOpenKeyExW(RootKeyHandle,  // HKEY    hKey
    PChar(SubKeyName),                    // LPCWSTR lpSubKey
    0,                                    // DWORD   ulOptions
    AccessFlags,                          // REGSAM  samDesired
    SubKeyHandle);                        // PHKEY   phkResult
  if result = ERROR_SUCCESS then
    RegCloseKey(SubKeyHandle);  // HKEY hKey
  RegCloseKey(RootKeyHandle);  // HKEY hKey
end;

function RegValueExists(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string): LSTATUS;
var
  RootKeyHandle, SubKeyHandle: HKEY;
  AccessFlags: REGSAM;
begin
  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
  if result <> ERROR_SUCCESS then
    exit;
  AccessFlags := KEY_READ;
  UpdateAccessFlags(RootKey, AccessFlags);
  result := RegOpenKeyExW(RootKeyHandle,  // HKEY    hKey
    PChar(SubKeyName),                    // LPCWSTR lpSubKey
    0,                                    // DWORD   ulOptions
    AccessFlags,                          // REGSAM  samDesired
    SubKeyHandle);                        // PHKEY   phkResult
  if result = ERROR_SUCCESS then
  begin
    result := RegQueryValueExW(SubKeyHandle,  // HKEY    hKey
      PChar(ValueName),                       // LPCWSTR lpValueName
      nil,                                    // LPDWORD lpReserved
      nil,                                    // LPDWORD lpType
      nil,                                    // LPBYTE  lpData
      nil);                                   // LPDWORD lpcbData
    RegCloseKey(SubKeyHandle);  // HKEY hKey
  end;
  RegCloseKey(RootKeyHandle);  // HKEY hKey
end;

function RegCreateSubKey(const ComputerName: string; RootKey: HKEY;
  const SubKeyName: string): LSTATUS;
var
  RootKeyHandle, SubKeyHandle, SubSubKeyHandle: HKEY;
  AccessFlags: REGSAM;
begin
  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
  if result <> ERROR_SUCCESS then
    exit;
  AccessFlags := KEY_ALL_ACCESS;
  UpdateAccessFlags(RootKey, AccessFlags);
  result := RegOpenKeyExW(RootKeyHandle,  // HKEY    hKey
    nil,                                  // LPCWSTR lpSubKey
    0,                                    // DWORD   ulOptions
    AccessFlags,                          // REGSAM  samDesired
    SubKeyHandle);                        // PHKEY   phkResult
  if result = ERROR_SUCCESS then
  begin
    result := RegCreateKeyW(SubKeyHandle,  // HKEY    hKey
      PChar(SubKeyName),                   // LPCWSTR lpSubKey
      SubSubKeyHandle);                    // PHKEY   phkResult
    if result = ERROR_SUCCESS then
      RegCloseKey(SubSubKeyHandle);  // HKEY hKey
    RegCloseKey(SubKeyHandle);  // HKEY hKey
  end;
  RegCloseKey(RootKeyHandle);  // HKEY hKey
end;

function RegGetSubKeyLastWriteTime(const ComputerName: string;
  RootKey: HKEY; const SubKeyName: string; var LastWriteTime: FILETIME): LSTATUS;
var
  RootKeyHandle, SubKeyHandle: HKEY;
  AccessFlags: REGSAM;
begin
  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
  if result <> ERROR_SUCCESS then
    exit;
  AccessFlags := KEY_READ;
  UpdateAccessFlags(RootKey, AccessFlags);
  result := RegOpenKeyExW(RootKeyHandle,  // HKEY    hKey
    PChar(SubKeyName),                    // LPCWSTR lpSubKey
    0,                                    // DWORD   ulOptions
    AccessFlags,                          // REGSAM  samDesired
    SubKeyHandle);                        // PHKEY   phkResult
  if result = ERROR_SUCCESS then
  begin
    result := RegQueryInfoKeyW(SubKeyHandle,  // HKEY      hKey
      nil,                                    // LPWSTR    lpClass
      nil,                                    // LPDWORD   lpcchClass
      nil,                                    // LPDWORD   lpReserved
      nil,                                    // LPDWORD   lpcSubKeys
      nil,                                    // LPDWORD   lpcbMaxSubKeyLen
      nil,                                    // LPDWORD   lpcbMaxClassLen
      nil,                                    // LPDWORD   lpcValues
      nil,                                    // LPDWORD   lpcbMaxValueNameLen
      nil,                                    // LPDWORD   lpcbMaxValueLen
      nil,                                    // LPDWORD   lpcbSecurityDescriptor
      @LastWriteTime);                        // PFILETIME lpftLastWriteTime
    RegCloseKey(SubKeyHandle);  // HKEY hKey
  end;
  RegCloseKey(RootKeyHandle);  // HKEY hKey
end;

function RegGetValueType(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; out ValueType: DWORD): LSTATUS;
var
  RootKeyHandle, SubKeyHandle: HKEY;
  AccessFlags: REGSAM;
begin
  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
  if result <> ERROR_SUCCESS then
    exit;
  AccessFlags := KEY_READ;
  UpdateAccessFlags(RootKey, AccessFlags);
  result := RegOpenKeyExW(RootKeyHandle,  // HKEY    hKey
    PChar(SubKeyName),                    // LPCWSTR lpSubKey
    0,                                    // DWORD   ulOptions
    AccessFlags,                          // REGSAM  samDesired
    SubKeyHandle);                        // PHKEY   phkResult
  if result = ERROR_SUCCESS then
  begin
    result := RegQueryValueExW(SubKeyHandle,  // HKEY    hKey
      PChar(ValueName),                       // LPCWSTR lpValueName
      nil,                                    // LPDWORD lpReserved
      @ValueType,                             // LPDWORD lpType
      nil,                                    // LPBYTE  lpData
      nil);                                   // LPDWORD lpcbData
    if result = ERROR_SUCCESS then
      RegCloseKey(SubKeyHandle);  // HKEY hKey
  end;
  RegCloseKey(RootKeyHandle);  // HKEY hKey
end;

function RegGetSubKeyNames(const ComputerName: string; RootKey: HKEY;
  const SubKeyName: string; var Names: TStringArray): LSTATUS;
var
  RootKeyHandle, SubKeyHandle: HKEY;
  AccessFlags: REGSAM;
  NumSubKeys, MaxSubKeyNameLen, I, MaxSubKeyNameLen2: DWORD;
  pName: PChar;
begin
  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
  if result <> ERROR_SUCCESS then
    exit;
  AccessFlags := KEY_READ;
  UpdateAccessFlags(RootKey, AccessFlags);
  result := RegOpenKeyExW(RootKeyHandle,  // HKEY    hKey
    PChar(SubKeyName),                    // LPCWSTR lpSubKey
    0,                                    // DWORD   ulOptions
    AccessFlags,                          // REGSAM  samDesired
    SubKeyHandle);                        // PHKEY   phkResult
  if result = ERROR_SUCCESS then
  begin
    result := RegQueryInfoKeyW(SubKeyHandle,  // HKEY      hKey
      nil,                                    // LPWSTR    lpClass
      nil,                                    // LPDWORD   lpcchClass
      nil,                                    // LPDWORD   lpReserved
      @NumSubKeys,                            // LPDWORD   lpcSubKeys
      @MaxSubKeyNameLen,                      // LPDWORD   lpcbMaxSubKeyLen
      nil,                                    // LPDWORD   lpcbMaxClassLen
      nil,                                    // LPDWORD   lpcValues
      nil,                                    // LPDWORD   lpcbMaxValueNameLen
      nil,                                    // LPDWORD   lpcbMaxValueLen
      nil,                                    // LPDWORD   lpcbSecurityDescriptor
      nil);                                   // PFILETIME lpftLastWriteTime
    if result = ERROR_SUCCESS then
    begin
      // Set dynamic array size
      SetLength(Names, NumSubKeys);
      if NumSubKeys > 0 then
      begin
        // lpcbMaxSubKeyLen doesn't include terminating null
        MaxSubKeyNameLen := (MaxSubKeyNameLen + 1) * SizeOf(Char);
        // Each call to RegEnumKeyEx will use this buffer
        GetMem(pName, MaxSubKeyNameLen);
        // Enumerate subkey names
        for I := 0 to NumSubKeys - 1 do
        begin
          // Use largest lpcchName for each call
          MaxSubKeyNameLen2 := MaxSubKeyNameLen;
          if RegEnumKeyExW(SubKeyHandle,  // HKEY      hKey
            I,                            // DWORD     dwIndex
            pName,                        // LPWSTR    lpName
            MaxSubKeyNameLen2,            // LPDWORD   lpcchName
            nil,                          // LPDWORD   lpReserved
            nil,                          // LPWSTR    lpClass
            nil,                          // LPDWORD   lpcchClass
            nil) = ERROR_SUCCESS then     // PFILETIME lpftLastWriteTime
            Names[I] := string(pName)
          else
            Names[I] := '';
        end;
        FreeMem(pName);
      end;
    end;
    RegCloseKey(SubKeyHandle);  // HKEY hKey
  end;
  RegCloseKey(RootKeyHandle);  // HKEY hKey
end;

function RegGetValueNames(const ComputerName: string; RootKey: HKEY;
  const SubKeyName: string; var Names: TStringArray): LSTATUS;
var
  RootKeyHandle, SubKeyHandle: HKEY;
  AccessFlags: REGSAM;
  NumValues, MaxValueNameLen, I, MaxValueNameLen2: DWORD;
  pName: PChar;
begin
  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
  if result <> ERROR_SUCCESS then
    exit;
  AccessFlags := KEY_READ;
  UpdateAccessFlags(RootKey, AccessFlags);
  result := RegOpenKeyExW(RootKeyHandle,  // HKEY    hKey
    PChar(SubKeyName),                    // LPCWSTR lpSubKey
    0,                                    // DWORD   ulOptions
    AccessFlags,                          // REGSAM  samDesired
    SubKeyHandle);                        // PHKEY   phkResult
  if result = ERROR_SUCCESS then
  begin
    result := RegQueryInfoKeyW(SubKeyHandle,  // HKEY      hKey
      nil,                                    // LPWSTR    lpClass
      nil,                                    // LPDWORD   lpcchClass
      nil,                                    // LPDWORD   lpReserved
      nil,                                    // LPDWORD   lpcSubKeys
      nil,                                    // LPDWORD   lpcbMaxSubKeyLen
      nil,                                    // LPDWORD   lpcbMaxClassLen
      @NumValues,                             // LPDWORD   lpcValues
      @MaxValueNameLen,                       // LPDWORD   lpcbMaxValueNameLen
      nil,                                    // LPDWORD   lpcbMaxValueLen
      nil,                                    // LPDWORD   lpcbSecurityDescriptor
      nil);                                   // PFILETIME lpftLastWriteTime
    if result = ERROR_SUCCESS then
    begin
      // Set dynamic array size
      SetLength(Names, NumValues);
      if NumValues > 0 then
      begin
        // lpcbMaxValueNameLen doesn't include terminating null
        MaxValueNameLen := (MaxValueNameLen + 1) * SizeOf(Char);
        // Each call to RegEnumValueW will use this buffer
        GetMem(pName, MaxValueNameLen);
        // Enumerate subkey names
        for I := 0 to NumValues - 1 do
        begin
          // Use largest lpcchName for each call
          MaxValueNameLen2 := MaxValueNameLen;
          if RegEnumValueW(SubKeyHandle,  // HKEY    hKey
            I,                            // DWORD   dwIndex
            pName,                        // LPWSTR  lpValueName
            MaxValueNameLen2,             // LPDWORD lpcchValueName
            nil,                          // LPDWORD lpReserved
            nil,                          // LPDWORD lpType
            nil,                          // LPBYTE  lpData
            nil) = ERROR_SUCCESS then     // LPDWORD lpcbData
            Names[I] := string(pName)
          else
            Names[I] := '';
        end;
        FreeMem(pName);
      end;
    end;
    RegCloseKey(SubKeyHandle);  // HKEY hKey
  end;
  RegCloseKey(RootKeyHandle);  // HKEY hKey
end;

//function RegRenameSubKey(const ComputerName: string; RootKey: HKEY;
//  const SubKeyName, NewName: string): LSTATUS;
//var
//  RootKeyHandle: HKEY;
//  AccessFlags: REGSAM;
//begin
//  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
//  if result <> ERROR_SUCCESS then
//    exit;
//  AccessFlags := KEY_WRITE;
//  UpdateAccessFlags(RootKey, AccessFlags);
//  result := RegRenameKey(RootKeyHandle,  // HKEY    hKey
//    PChar(SubKeyName),                   // LPCWSTR lpSubKeyName
//    PChar(NewName));                     // LPCWSTR lpNewKeyName
//  RegCloseKey(RootKeyHandle);  // HKEY hKey
//end;

function RegDeleteKeyIfEmpty(const ComputerName: string; RootKey: HKEY;
  const SubKeyName: string): LSTATUS;
var
  RootKeyHandle, SubKeyHandle: HKEY;
  AccessFlags: REGSAM;
  NumValues: DWORD;
begin
  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
  if result <> ERROR_SUCCESS then
    exit;
  AccessFlags := KEY_ALL_ACCESS;
  UpdateAccessFlags(RootKey, AccessFlags);
  result := RegOpenKeyExW(RootKeyHandle,  // HKEY    hKey
    PChar(SubKeyName),                    // LPCWSTR lpSubKey
    0,                                    // DWORD   ulOptions
    AccessFlags,                          // REGSSAM samDesired
    SubKeyHandle);                        // PHKEY   phkResult
  if result = ERROR_SUCCESS then
  begin
    result := RegQueryInfoKeyW(SubKeyHandle,  // HKEY      hKey
      nil,                                    // LPWSTR    lpClass
      nil,                                    // LPDWORD   lpcchClass
      nil,                                    // LPDWORD   lpReserved
      nil,                                    // LPDWORD   lpcSubKeys
      nil,                                    // LPDWORD   lpcbMaxSubKeyLen
      nil,                                    // LPDWORD   lpcbMaxClassLen
      @NumValues,                             // LPDWORD   lpcValues
      nil,                                    // LPDWORD   lpcbMaxValueNameLen
      nil,                                    // LPDWORD   lpcbMaxValueLen
      nil,                                    // LPDWORD   lpcbSecurityDescriptor
      nil);                                   // PFILETIME lpftLastWriteTime
    if result = ERROR_SUCCESS then
    begin
      RegCloseKey(SubKeyHandle);  // HKEY hKey
      if NumValues = 0 then
        result := RegDeleteKeyW(RootKeyHandle,  // HKEY    hKey
          PChar(SubKeyName))                    // LPCWSTR lpSubKey
      else
        result := ERROR_ACCESS_DENIED;
    end;
  end;
  RegCloseKey(RootKeyHandle);  // HKEY hKey
end;

function RegDeleteKeyIncludingSubKeys(const ComputerName: string;
  RootKey: HKEY; const SubKeyName: string): LSTATUS;
var
  Names: TStringArray;
  I: DWORD;
  RootKeyHandle: HKEY;
  AccessFlags: REGSAM;
begin
  result := RegGetSubKeyNames(ComputerName, RootKey, SubKeyName, Names);
  if (result = ERROR_SUCCESS) and (Length(Names) > 0) then
  begin
    for I := 0 to Length(Names) - 1 do
    begin
      result := RegDeleteKeyIncludingSubKeys(ComputerName, RootKey,
        SubKeyName + '\' + Names[I]);
      if result <> ERROR_SUCCESS then
        break;
    end;
  end;
  if result = ERROR_SUCCESS then
  begin
    result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
    if result = ERROR_SUCCESS then
    begin
      AccessFlags := KEY_ALL_ACCESS;
      UpdateAccessFlags(RootKey, AccessFlags);
      result := RegDeleteKeyW(RootKeyHandle,  // HKEY    hKey
        PChar(SubKeyName));                   // LPCWSTR lpSubKey
      RegCloseKey(RootKeyHandle);  // HKEY hKey
    end;
  end;
end;

function RegDeleteValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string): LSTATUS;
var
  RootKeyHandle, SubKeyHandle: HKEY;
  AccessFlags: REGSAM;
begin
  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
  if result <> ERROR_SUCCESS then
    exit;
  AccessFlags := KEY_ALL_ACCESS;
  UpdateAccessFlags(RootKey, AccessFlags);
  result := RegOpenKeyExW(RootKeyHandle,  // HKEY    hKey
    PChar(SubKeyName),                    // LPCWSTR lpSubKey
    0,                                    // DWORD   ulOptions
    AccessFlags,                          // REGSAM  samDesired
    SubKeyHandle);                        // PHKEY   phkResult
  if result = ERROR_SUCCESS then
  begin
    result := RegDeleteValueW(SubKeyHandle,  // HKEY    hKey
      PChar(ValueName));                     // LPCWSTR lpValueName
    RegCloseKey(SubKeyHandle);  // HKEY hKey
  end;
  RegCloseKey(RootKeyHandle);  // HKEY hKey
end;

function RegGetBinaryValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; var Bytes: TByteArray): LSTATUS;
var
  RootKeyHandle, SubKeyHandle: HKEY;
  AccessFlags: REGSAM;
  ValueSize: DWORD;
begin
  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
  if result <> ERROR_SUCCESS then
    exit;
  AccessFlags := KEY_READ;
  UpdateAccessFlags(RootKey, AccessFlags);
  result := RegOpenKeyExW(RootKeyHandle,  // HKEY    hKey
    PChar(SubKeyName),                    // LPCWSTR lpSubKey
    0,                                    // DWORD   ulOptions
    AccessFlags,                          // REGSAM  samDesired
    SubKeyHandle);                        // PHKEY   phkResult
  if result = ERROR_SUCCESS then
  begin
    result := RegQueryValueExW(SubKeyHandle,  // HKEY    hKey
      PChar(ValueName),                       // LPCWSTR lpValueName
      nil,                                    // LPDWORD lpReserved
      nil,                                    // LPDWORD lpType
      nil,                                    // LPBYTE  lpData
      @ValueSize);                            // LPDWORD lpcbData
    if result = ERROR_SUCCESS then
    begin
      SetLength(Bytes, ValueSize);
      if ValueSize > 0 then
      begin
        result := RegQueryValueExW(SubKeyHandle,  // HKEY    hKey
          PChar(ValueName),                       // LPCWSTR lpValueName
          nil,                                    // LPDWORD lpReserved
          nil,                                    // LPDWORD lpType
          @Bytes[0],                              // LPBYTE  lpData
          @ValueSize);                            // LPDWORD lpcbData
      end;
    end;
    RegCloseKey(SubKeyHandle);  // HKEY hKey
  end;
  RegCloseKey(RootKeyHandle);  // HKEY hKey
end;

function RegValueIsEmpty(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; out Empty: Boolean): LSTATUS;
var
  Bytes: TByteArray;
begin
  result := RegValueExists(ComputerName, RootKey, SubKeyName, ValueName);
  if result <> ERROR_SUCCESS then
    exit;
  RegGetBinaryValue(ComputerName, RootKey, SubKeyName, ValueName, Bytes);
  Empty := Length(Bytes) = 0;
end;

function RegGetDWORDValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; out ValueData: DWORD): LSTATUS;
var
  Bytes: TByteArray;
  DataLen: DWORD;
begin
  result := RegGetBinaryValue(ComputerName, RootKey, SubKeyName, ValueName, Bytes);
  if (result = ERROR_SUCCESS) and (Length(Bytes) > 0) then
  begin
    if Length(Bytes) < SizeOf(DWORD) then
      DataLen := Length(Bytes)
    else
      DataLen := SizeOf(DWORD);
    Move(Bytes[0], ValueData, DataLen);
  end;
end;

// Returns N as hex string of C length
function HexStr(N: LongInt; const C: Byte): string;
const
  HexChars: array[0..15] of Char = '0123456789ABCDEF';
var
  I: LongInt;
begin
  SetLength(result, C);
  for I := C downto 1 do
  begin
    result[I] := HexChars[N and $F];
    N := N shr 4;
  end;
end;

// Returns buffer content as a string (useful to validate string buf content)
function GetCharBufData(const pStrBuf: PChar; const BufSize: DWORD): string;
var
  pByteBuf: PByte;
  I: DWORD;
begin
  result := '';
  if (BufSize > 0) and Assigned(pStrBuf) then
  begin
    pByteBuf := PByte(pStrBuf);
    result := HexStr(pByteBuf^, 2);
    Inc(pByteBuf);
    for I := 1 to BufSize - 1 do
    begin
      if I mod SizeOf(Char) = 0 then
        result := result + ' ';
      result := result + HexStr(pByteBuf^, 2);
      Inc(pByteBuf);
    end;
  end;
end;

// Assumes dynamic byte array is a Char buffer and appends a terminating
// null if it doesn't have one
procedure NormalizeCharBuf(var Bytes: TByteArray);
var
  pStringData: PChar;
begin
  pStringData := PChar(@Bytes[0]);
  // If last character of buffer isn't a null...
  if pStringData[(Length(Bytes) div SizeOf(Char)) - 1] <> #0 then
  begin
    // ...increase dynamic array by a character and append null
    SetLength(Bytes, Length(Bytes) + SizeOf(Char));
    FillChar(Bytes[Length(Bytes) - SizeOf(Char)], SizeOf(Char), #0);
  end;
end;

function RegGetStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; out ValueData: string): LSTATUS;
var
  Bytes: TByteArray;
  pStringData: PChar;
begin
  result := RegGetBinaryValue(ComputerName, RootKey, SubKeyName, ValueName, Bytes);
  if result <> ERROR_SUCCESS then
    exit;
  // Not enough data for a string
  if Length(Bytes) < SizeOf(Char) then
  begin
    ValueData := '';
    exit;
  end;
  // Append terminating null if missing
  NormalizeCharBuf(Bytes);
  pStringData := PChar(@Bytes[0]);
  // Uncomment below line to dump buffer content
  //WriteLn(GetCharBufData(pStringData, Length(Bytes)));
  ValueData := string(pStringData);
end;

function ExpandEnvStrings(const S: string): string;
var
  NumChars, BufSize: DWORD;
  pBuffer: PChar;
begin
  NumChars := ExpandEnvironmentStringsW(PChar(S),  // LPCWSTR lpSrc
    nil,                                           // LPWSTR  lpDst
    0);                                            // DWORD   nSize
  if NumChars = 0 then
  begin
    result := S;  // If fail, return original string
    exit;
  end;
  BufSize := NumChars * SizeOf(Char);
  GetMem(pBuffer, BufSize);
  if ExpandEnvironmentStringsW(PChar(S),  // LPCWSTR lpSrc
    pBuffer,                              // LPWSTR  lpDst
    NumChars) > 0 then                    // DWORD   nSize
  begin
    result := string(pBuffer);
  end;
  FreeMem(pBuffer);
end;

function RegGetExpandStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; out ValueData: string): LSTATUS;
var
  Data: string;
begin
  result := RegGetStringValue(ComputerName, RootKey, SubKeyName, ValueName, Data);
  if result <> ERROR_SUCCESS then
    exit;
  ValueData := ExpandEnvStrings(Data);
end;

function RegGetMultiStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; var Values: TStringArray): LSTATUS;
var
  Bytes: TByteArray;
  BufSize, NumChars, NullCount, I: DWORD;
  pStringData: PChar;
begin
  result := RegGetBinaryValue(ComputerName, RootKey, SubKeyName, ValueName, Bytes);

  // Terminate if error
  if result <> ERROR_SUCCESS then
    exit;

  // Not enough data for a string
  if Length(Bytes) < SizeOf(Char) then
  begin
    SetLength(Values, 0);
    exit;
  end;

  // Append terminating null if missing
  NormalizeCharBuf(Bytes);
  // Get buffer size and number of characters
  BufSize := Length(Bytes);
  NumChars := BufSize div SizeOf(Char);
  // Point at first byte
  pStringData := PChar(@Bytes[0]);

  // Uncomment below line to dump buffer content
  //WriteLn(GetCharBufData(pStringData, Length(Bytes)));

  // Count # of nulls in buffer (i.e., # of null-terminated strings)
  NullCount := 0;
  for I := 0 to NumChars - 1 do
  begin
    if pStringData[0] = #0 then
      Inc(NullCount);
    Inc(pStringData);
  end;

  // REG_MULTI_SZ data should end with two null characters (but it might not);
  // if it does (typical case), don't count final "empty" string
  pStringData := PChar(@Bytes[0]);
  if (pStringData[NumChars - 2] = #0) and (pStringData[NumChars - 1] = #0) then
    Dec(NullCount);

  // Set number of strings in dynamic array
  SetLength(Values, NullCount);

  if NullCount > 0 then
  begin
    // Step through buffer and populate array
    for I := 0 to NullCount - 1 do
    begin
      Values[I] := pStringData;
      Inc(pStringData, Length(pStringData) + 1);
    end;
  end;
end;

// Writes dynamic byte array to registry using specified data type
function RegSetValueWithType(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; const DataType: DWORD;
  var Bytes: TByteArray): LSTATUS;
var
  RootKeyHandle, SubKeyHandle: HKEY;
  AccessFlags: REGSAM;
  DataSize: DWORD;
begin
  result := ConnectRegistry(ComputerName, RootKey, RootKeyHandle);
  if result <> ERROR_SUCCESS then
    exit;
  AccessFlags := KEY_ALL_ACCESS;
  UpdateAccessFlags(RootKey, AccessFlags);
  result := RegOpenKeyExW(RootKeyHandle,  // HKEY    hKey
    PChar(SubKeyName),                    // LPCWSTR lpSubKey
    0,                                    // DWORD   ulOptions
    AccessFlags,                          // REGSAM  samDesired
    SubKeyHandle);                        // PHKEY   phkResult
  if result = ERROR_SUCCESS then
  begin
    DataSize := SizeOf(Byte) * Length(Bytes);
    result := RegSetValueExW(SubKeyHandle,  // HKEY       hKey
      PChar(ValueName),                     // LPCWSTR    lpValueName
      0,                                    // DWORD      Reserved
      DataType,                             // DWORD      dwType
      @Bytes[0],                            // const BYTE *lpData
      DataSize);                            // DWORD      cbData
    RegCloseKey(SubKeyHandle);  // HKEY hKey
  end;
  RegCloseKey(RootKeyHandle);  // HKEY hKey
end;

function RegSetBinaryValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; var Bytes: TByteArray): LSTATUS;
begin
  result := RegSetValueWithType(ComputerName, RootKey, SubKeyName,
    ValueName, REG_BINARY, Bytes);
end;

function RegSetDWORDValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; const ValueData: DWORD): LSTATUS;
var
  Bytes: TByteArray;
begin
  SetLength(Bytes, SizeOf(DWORD));
  Move(ValueData, Bytes[0], SizeOf(DWORD));
  result := RegSetValueWithType(ComputerName, RootKey, SubKeyName,
    ValueName, REG_DWORD, Bytes);
end;

function RegSetSingleStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; const DataType: DWORD;
  const ValueData: string): LSTATUS;
var
  BufSize: DWORD;
  Bytes: TByteArray;
begin
  // Buffer size is # of Chars in value data + null
  BufSize := (Length(ValueData) + 1) * SizeOf(Char);

  // Set byte array length and copy string
  SetLength(Bytes, BufSize);
  if ValueData <> '' then
    Move(ValueData[1], Bytes[0], BufSize);

  // Uncomment below line to dump buffer content
  //WriteLn(GetCharBufData(PChar(@Bytes[0]), Length(Bytes)));

  result := RegSetValueWithType(ComputerName, RootKey, SubKeyName,
    ValueName, DataType, Bytes);
end;

function RegSetStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName, ValueData: string): LSTATUS;
begin
  result := RegSetSingleStringValue(ComputerName, RootKey, SubKeyName,
    ValueName, REG_SZ, ValueData);
end;

function RegSetExpandStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName, ValueData: string): LSTATUS;
begin
  result := RegSetSingleStringValue(ComputerName, RootKey, SubKeyName,
    ValueName, REG_EXPAND_SZ, ValueData);
end;

function RegSetMultiStringValue(const ComputerName: string; RootKey: HKEY;
  const SubKeyName, ValueName: string; var Values: TStringArray): LSTATUS;
var
  NumChars, I, BufSize: DWORD;
  pCharBuf, pStringData: PChar;
  Bytes: TByteArray;
begin
  // Buffer will contain at least 1 character
  NumChars := 1;

  // Get number of characters needed and buffer size
  if Length(Values) > 0 then
  begin
    for I := 0 to Length(Values) - 1 do
    begin
      Inc(NumChars, Length(Values[I]) + 1);  // Include null characters
    end;
  end;
  BufSize := NumChars * SizeOf(Char);

  // Allocate buffer for REG_MULTI_SZ data
  GetMem(pCharBuf, BufSize);

  // If there are strings to be copied, populate buffer
  if Length(Values) > 0 then
  begin
    // Point to head of buffer
    pStringData := pCharBuf;
    for I := 0 to Length(Values) - 1 do
    begin
      // Copy string to buffer
      Move(Values[I][1], pStringData^, Length(Values[I]) * SizeOf(Char));
      // Set null at end of string in buffer
      pStringData[Length(Values[I])] := #0;
      // Increment pointer for next string copy
      Inc(pStringData, Length(Values[I]) + 1);
    end;
  end;

  // Set final terminating null at end of buffer
  pCharBuf[NumChars - 1] := #0;

  // Uncomment below line to dump buffer content
  //WriteLn(GetCharBufData(pCharBuf, BufSize));

  // Set byte array length, copy buffer content to it, and free buffer
  SetLength(Bytes, BufSize);
  Move(pCharBuf^, Bytes[0], BufSize);
  FreeMem(pCharBuf);

  // Set in registry
  result := RegSetValueWithType(ComputerName, RootKey, SubKeyName,
    ValueName, REG_MULTI_SZ, Bytes);
end;

begin
end.
