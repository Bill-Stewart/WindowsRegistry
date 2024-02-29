<!-- omit in toc -->
# WindowsRegistry

**WindowsRegistry.pp** is a Free Pascal (FPC) unit that provides a set of functions for reading from and writing to the Windows registry.

<!-- omit in toc -->
# Table of Contents
- [License](#license)
- [Notable Features](#notable-features)
- [Example Program](#example-program)
- [Version History](#version-history)
- [Reference](#reference)
  - [Common Parameters](#common-parameters)
  - [Registry Value Types](#registry-value-types)
  - [RegKeyExists](#regkeyexists)
    - [Syntax](#syntax)
    - [Parameters](#parameters)
    - [Return Value](#return-value)
  - [RegValueExists](#regvalueexists)
    - [Syntax](#syntax-1)
    - [Parameters](#parameters-1)
    - [Return Value](#return-value-1)
  - [RegValueIsEmpty](#regvalueisempty)
    - [Syntax](#syntax-2)
    - [Parameters](#parameters-2)
    - [Return Value](#return-value-2)
  - [RegCreateSubKey](#regcreatesubkey)
    - [Syntax](#syntax-3)
    - [Parameters](#parameters-3)
    - [Return Value](#return-value-3)
    - [Remarks](#remarks)
  - [RegGetValueType](#reggetvaluetype)
    - [Syntax](#syntax-4)
    - [Parameters](#parameters-4)
    - [Return Value](#return-value-4)
    - [Remarks](#remarks-1)
  - [RegGetSubKeyNames](#reggetsubkeynames)
    - [Syntax](#syntax-5)
    - [Parameters](#parameters-5)
    - [Return Value](#return-value-5)
    - [Remarks](#remarks-2)
  - [RegGetValueNames](#reggetvaluenames)
    - [Syntax](#syntax-6)
    - [Parameters](#parameters-6)
    - [Return Value](#return-value-6)
    - [Remarks](#remarks-3)
  - [RegDeleteKeyIfEmpty](#regdeletekeyifempty)
    - [Syntax](#syntax-7)
    - [Parameters](#parameters-7)
    - [Return Value](#return-value-7)
    - [Remarks](#remarks-4)
  - [RegDeleteKeyIncludingSubKeys](#regdeletekeyincludingsubkeys)
    - [Syntax](#syntax-8)
    - [Parameters](#parameters-8)
    - [Return Value](#return-value-8)
  - [RegDeleteValue](#regdeletevalue)
    - [Syntax](#syntax-9)
    - [Parameters](#parameters-9)
    - [Return Value](#return-value-9)
  - [RegGetBinaryValue](#reggetbinaryvalue)
    - [Syntax](#syntax-10)
    - [Parameters](#parameters-10)
    - [Return Value](#return-value-10)
    - [Remarks](#remarks-5)
  - [RegGetDWORDValue](#reggetdwordvalue)
    - [Syntax](#syntax-11)
    - [Parameters](#parameters-11)
    - [Return Value](#return-value-11)
  - [RegGetStringValue](#reggetstringvalue)
    - [Syntax](#syntax-12)
    - [Parameters](#parameters-12)
    - [Return Value](#return-value-12)
    - [Remarks](#remarks-6)
  - [RegGetExpandStringValue](#reggetexpandstringvalue)
    - [Syntax](#syntax-13)
    - [Parameters](#parameters-13)
    - [Return Value](#return-value-13)
    - [Remarks](#remarks-7)
  - [RegGetMultiStringValue](#reggetmultistringvalue)
    - [Syntax](#syntax-14)
    - [Parameters](#parameters-14)
    - [Return Value](#return-value-14)
    - [Remarks](#remarks-8)
  - [RegSetBinaryValue](#regsetbinaryvalue)
    - [Syntax](#syntax-15)
    - [Parameters](#parameters-15)
    - [Return Value](#return-value-15)
    - [Remarks](#remarks-9)
  - [RegSetDWORDValue](#regsetdwordvalue)
    - [Syntax](#syntax-16)
    - [Parameters](#parameters-16)
    - [Return Value](#return-value-16)
  - [RegSetStringValue](#regsetstringvalue)
    - [Syntax](#syntax-17)
    - [Parameters](#parameters-17)
    - [Return Value](#return-value-17)
  - [RegSetExpandStringValue](#regsetexpandstringvalue)
    - [Syntax](#syntax-18)
    - [Parameters](#parameters-18)
    - [Return Value](#return-value-18)
    - [Remarks](#remarks-10)
  - [RegSetMultiStringValue](#regsetmultistringvalue)
    - [Syntax](#syntax-19)
    - [Parameters](#parameters-19)
    - [Return Value](#return-value-19)
    - [Remarks](#remarks-11)

-------------------------------------------------------------------------------

## License

**WindowsRegistry.pp** is covered by the GNU Lesser Public License (LPGL). See the file `LICENSE` for details.

## Notable Features

The following is a notable list of features the unit offers:

* Uses only the **windows** unit (very lightweight).

* All functions use only Unicode strings and characters.

* All functions return 0 for success or a non-zero Windows API error code value on failure.

* Uses dynamic arrays for `REG_BINARY` and `REG_MULTI_SZ` registry value types (greatly simplifies code for these value types).

* Supports 32-bit application access to the 64-bit portions of the registry (and 64-bit application access to the 32-bit portions of the registry).

* Automatically adds missing null terminators when reading string values from the registry (buffer overflow protection).

* Supports reading from and writing to the registry on a remote computer.

## Example Program

The following Free Pascal program demonstrates the unit's **RegGetMultiStringValue** function to retrieve the `PendingFileRenameOperations` registry value on the current computer:

    {$APPTYPE CONSOLE}
    {$MODE OBJFPC}
    {$MODESWITCH UNICODESTRINGS}

    uses
      windows,
      WindowsRegistry;

    const
      SUBKEY_PATH  = 'SYSTEM\CurrentControlSet\Control\Session Manager';
      SUBKEY_VALUE = 'PendingFileRenameOperations';

    var
      Values: TStringArray;
      I: Integer;

    begin
      ExitCode := RegGetMultiStringValue('', HKEY_LOCAL_MACHINE, SUBKEY_PATH,
        SUBKEY_VALUE, Values);
      if ExitCode = ERROR_SUCCESS then
      begin
        if Length(Values) > 0 then
        begin
          for I := 0 to Length(Values) - 1 do
            WriteLn(Values[I]);
        end
        else
          WriteLn(SUBKEY_VALUE, ' registry value exists, but is empty.');
      end
      else if ExitCode = ERROR_FILE_NOT_FOUND then
        WriteLn(SUBKEY_VALUE, ' registry value does not exist.')
      else
        WriteLn('Error code: ', ExitCode);
    end.

## Version History

**1.0.0 (2024-02-29)**

* Initial release.

## Reference

This section documents the registry functions.

> IMPORTANT: This unit enables FPC's UNICODESTRINGS mode, so all references to `string` are `UnicodeString`.

### Common Parameters

All of the function start with the same two parameters:

`ComputerName`

Specifies the name of a remote computer. Use an empty string (i.e., `''`) to specify the current computer.

`RootKey`

Specifies a predefined registry handle value. This parameter can be one of the following predefined keys:

* `HKEY_LOCAL_MACHINE`
* `HKEY_USERS`

If you are connecting to the the current computer's registry (i.e., the `ComputerName` parameter is an empty string), you can also specify `HKEY_CURRENT_USER` for the `RootKey` parameter. (You cannot specify `HKEY_CURRENT_USER` when connecting to a remote computer.)

The unit also supports the following registry handle values:

* `HKEY_CURRENT_USER_64`
* `HKEY_LOCAL_MACHINE_64`
* `HKEY_CURRENT_USER_32`
* `HKEY_LOCAL_MACHINE_32`

32-bit applications can use the the values with `_64` to access the 64-bit portions of the registry, and 64-bit applications can use the values with `_32` to access the 32-bit portions of the registry. For more information, see the topic **32-bit and 64-bit Application Data in the Registry** in the Microsoft documentation (currently https://learn.microsoft.com/en-us/windows/win32/sysinfo/32-bit-and-64-bit-application-data-in-the-registry). Just as with `HKEY_CURRENT_USER`, `HKEY_CURRENT_USER_64` and `HKEY_CURRENT_USER_32` are not available on a remote computer.

### Registry Value Types

The following table lists the most common registry value types:

Value           | Data Type
--------------- | -----------
`REG_SZ`        | String
`REG_EXPAND_SZ` | String that contains unexpanded environment variable references (e.g., _%Path%_)
`REG_BINARY`    | Binary
`REG_DWORD`     | 32-bit unsigned integer
`REG_MULTI_SZ`  | Multiple strings

> NOTE: There is no difference between the `REG_SZ` and `REG_EXPAND_SZ` types, except that the `REG_EXPAND_SZ` type is a "hint" to applications that the string might contain unexpanded environment variable references.

For convenience, the unit manages the `REG_BINARY` and `REG_MULTI_SZ` types using dynamic arrays:

    type
      TByteArray = array of Byte;
      TStringArray = array of string;

The use of dynamic arrays frees the programmer from having to release the memory that these arrays use, as well as avoiding potential memory leaks.

-------------------------------------------------------------------------------

### RegKeyExists

Tests whether a specified registry subkey exists.

#### Syntax

    function RegKeyExists(ComputerName: string; RootKey: HKEY; SubKeyName: string): Integer;

#### Parameters

`SubKeyName`

Specifies the name of the subkey.

#### Return Value

If the subkey exists, the function returns 0. If the subkey does not exist, the function's return value is a non-zero Windows error code.

-------------------------------------------------------------------------------

### RegValueExists

Tests whether a specified registry value exists.

#### Syntax

    function RegValueExists(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

#### Return Value

If the value exists, the function returns 0. If the value does not exist, the function's return value is a non-zero Windows error code.

-------------------------------------------------------------------------------

### RegValueIsEmpty

Tests whether a specified registry value is empty.

#### Syntax

    function RegValueIsEmpty(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string; out Empty: Boolean): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`Empty`

This parameter receives a value that indicates whether the registry value is empty (i.e., contains no data).

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code. The `Empty` value is not defined if the function fails.

-------------------------------------------------------------------------------

### RegCreateSubKey

Creates a registry subkey.

#### Syntax

    function RegCreateSubKey(ComputerName: string; RootKey: HKEY; SubKeyName: string): Integer;

#### Parameters

`SubKeyName`

Specifies the name of the subkey to create.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code.

#### Remarks

The **RegCreateKeySubkey** function creates all subkeys if you specify more than one subkey names separated by `\` characters. For example, the function creates a subkey three levels deep by specifying a string like the following for the `SubKeyName` parameter:

    subkey 1\subkey 2\subkey 3

-------------------------------------------------------------------------------

### RegGetValueType

Retrieves a registry value's data type.

#### Syntax

    function RegGetValueType(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string; out ValueType: DWORD): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`ValueType`

This parameter receives a value containing the registry value's specified type. See [Registry Value Types](#registry-value-types) for the types.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code. The `ValueType` value is not defined if the function fails.

#### Remarks

Every registry value is stored in binary format, but the value type lets programs know to interpret the value.

-------------------------------------------------------------------------------

### RegGetSubKeyNames

Retrieves an array of subkey names from a specified subkey.

#### Syntax

    function RegGetSubKeyNames(ComputerName: string; RootKey: HKEY; SubKeyName: string; var Names: TStringArray): Integer;

#### Parameters

`SubKeyName`

Specifies the name of the subkey. The function retrieves the names of the subkeys located at this subkey.

`Names`

This parameter receives an array of subkey names.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code. The content of the `Names` array is not defined if the function fails.

#### Remarks

See the [Example Program](#example-program) for an example of how to iterate the array of strings returned to the `Names` parameter.

-------------------------------------------------------------------------------

### RegGetValueNames

Retrieves an array of value names from a specified subkey.

#### Syntax

    function RegGetValueNames(ComputerName: string; RootKey: HKEY; SubKeyName: string; var Names: TStringArray): Integer;

#### Parameters

`SubKeyName`

Specifies the name of the subkey. The function retrieves the names of the values located at this subkey.

`Names`

This parameter receives an array of value names.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code. The content of the `Names` array is not defined if the function fails.

#### Remarks

See the [Example Program](#example-program) for an example of how to iterate the array of strings returned to the `Names` parameter.

-------------------------------------------------------------------------------

### RegDeleteKeyIfEmpty

Deletes a registry subkey if it contains no subkeys or values.

#### Syntax

    function RegDeleteKeyIfEmpty(ComputerName: string; RootKey: HKEY; SubKeyName: string): Integer;

#### Parameters

`SubKeyName`

Specifies the name of the subkey to delete.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code.

#### Remarks

The function will fail if the subkey contains any subkeys or values. To delete a subkey even if it contains subkeys and/or values, use the [**RegDeleteKeyIncludingSubKeys**](#regdeletekeyincludingsubkeys) function instead.

-------------------------------------------------------------------------------

### RegDeleteKeyIncludingSubKeys

Deletes a registry subkey, including all subkeys and values within it.

#### Syntax

    function RegDeleteKeyIncludingSubKeys(ComputerName: string; RootKey: HKEY; SubKeyName: string): Integer;

#### Parameters

`SubKeyName`

Specifies the name of the subkey to delete.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code.

-------------------------------------------------------------------------------

### RegDeleteValue

Deletes a registry value.

#### Syntax

    function RegDeleteValue(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the name of the value to delete.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code.

-------------------------------------------------------------------------------

### RegGetBinaryValue

Retrieves a `REG_BINARY` registry value as an array of bytes.

#### Syntax

    function RegGetBinaryValue(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string; var Bytes: TByteArray): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`Bytes`

This parameter receives an array of bytes.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code. The content of the `Bytes` array is not defined if the function fails.

#### Remarks

To iterate the array, use a numeric indexing variable and the **Length** function to iterate the array. For example:

    ...
    var
      Bytes: TByteArray;
      I: Integer;
    ...
    if (RegGetBinaryValue('', HKEY_CURRENT_USER, 'TestSubKey', 'TestValue',
      Bytes) = 0) and (Length(Bytes) > 0) then
    begin
      for I := 0 to Length(Bytes) - 1 do
      ...
    end;

All registry data is stored in ultimately stored in binary format, and in fact all of the other **RegGet**... functions use this function to retrieve registry values and interpret retrieved values as the type requested by the function. For example, the [**RegGetDWORDValue**](#reggetdwordvalue) function retrieves registry value data using **RegGetBinaryValue**, and then interprets the value data numerically.

-------------------------------------------------------------------------------

### RegGetDWORDValue

Retrieves a registry value as a `REG_DWORD` (unsigned 32-bit integer) value.

#### Syntax

    function RegGetDWORDValue(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string; out ValueData: DWORD): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`ValueData`

This parameter receives the value's data.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code. The `ValueData` value is not defined if the function fails.

-------------------------------------------------------------------------------

### RegGetStringValue

Retrieves a registry value as a `REG_SZ` (string) value.

#### Syntax

    function RegGetStringValue(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string; out ValueData: string): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`ValueData`

This parameter receives the value's data.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code. The `ValueData` value is not defined if the function fails.

#### Remarks

This function can read a `REG_EXPAND_SZ` value without expanding environment variable references in the value's data. If you want to read a `REG_EXPAND_SZ` value and automatically expand environment variable references in the value's data, use the [**RegGetExpandStringValue**](#reggetexpandstringvalue) function.

-------------------------------------------------------------------------------

### RegGetExpandStringValue

Retrieves a registry value as a `REG_EXPAND_SZ` value.

#### Syntax

    function RegGetExpandStringValue(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string; out ValueData: string): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`ValueData`

This parameter receives the value's data.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code. The `ValueData` value is not defined if the function fails.

#### Remarks

This function is the same as [**RegGetStringValue**](#reggetstringvalue) except that it expands environment variable references in the value's data. If you want to read a `REG_EXPAND_SZ` value without expanding environment variable references, use the **RegGetStringValue** function instead.

For example, if a registry value contains the string `%SystemRoot%`, this function will expand the environment variable reference to the Windows installation directory (e.g., `C:\Windows`). In contrast, the **RegGetStringValue** function will not expand the environment variable reference.

-------------------------------------------------------------------------------

### RegGetMultiStringValue

Retrieves a `REG_MULTI_SZ` registry value as an array of strings.

#### Syntax

    function RegGetMultiStringValue(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string; var Values: TStringArray): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`Values`

This parameter receives an array of strings.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code. The content of the `Values` array is not defined if the function fails.

#### Remarks

This function provides an easy-to-use interface for retrieving `REG_MULTI_SZ` registry values because it returns the strings as a dynamic array.

See the [Example Program](#example-program) for an example of how to iterate the array of strings returned to the `Values` parameter.

-------------------------------------------------------------------------------

### RegSetBinaryValue

Sets a `REG_BINARY` registry value using a dynamic array of bytes.

#### Syntax

    function RegSetBinaryValue(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string; var Bytes: TByteArray): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`Bytes`

Specifies the value's data.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code.

#### Remarks

This function provides an easy-to-use interface for setting `REG_BINARY` registry values because it uses a dynamic array of bytes.

To use the function, populate the dynamic array with as many bytes as needed, and then call the function to set the value. For example:

    var
      Bytes: TByteArray;
    ...
    SetLength(Bytes, 2);
    Bytes[0] := $0D;
    Bytes[1] := $0A;
    if RegSetBinaryValue('', HKEY_CURRENT_USER, 'TestSubKey', 'TestValue',
      Bytes) = 0 then
    begin
    ...
    end;

-------------------------------------------------------------------------------

### RegSetDWORDValue

Sets a `REG_DWORD` registry value.

#### Syntax

    function RegSetDWORDValue(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string; ValueData: DWORD): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`ValueData`

Specifies the value's data.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code.

-------------------------------------------------------------------------------

### RegSetStringValue

Sets a `REG_SZ` registry value.

#### Syntax

    function RegSetStringValue(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName, ValueData: string): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`ValueData`

Specifies the value's data.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code.

-------------------------------------------------------------------------------

### RegSetExpandStringValue

Sets a `REG_EXPAND_SZ` registry value.

#### Syntax

    function RegSetExpandStringValue(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName, ValueData: string): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`ValueData`

Specifies the value's data.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code.

#### Remarks

This function is the same as the [**RegSetStringValue**](#regsetstringvalue) function except that it sets the registry value's type to `REG_EXPAND_SZ` rather than `REG_SZ`.

-------------------------------------------------------------------------------

### RegSetMultiStringValue

Sets a `REG_MULTI_SZ` registry value using a dynamic array of strings.

#### Syntax

    function RegSetMultiStringValue(ComputerName: string; RootKey: HKEY; SubKeyName, ValueName: string; var Values: TStringArray): Integer;

#### Parameters

`SubKeyName`

Specifies the value's subkey.

`ValueName`

Specifies the value's name.

`Values`

Specifies the value's data.

#### Return Value

The function returns 0 if it succeeds. If it does not succeed, the return value is a non-zero Windows error code.

#### Remarks

This function provides an easy-to-use interface for setting `REG_MULTI_SZ` registry values because it uses a dynamic array of strings.

To use the function, populate the dynamic array with as many strings as needed, and then call the function to set the value. For example:

    var
      Items: TStringArray;
    ...
    SetLength(Items, 2);
    Items[0] := 'Value 1';
    Items[1] := 'Value 2';
    if RegSetMultiStringValue('', HKEY_CURRENT_USER, 'TestSubKey', 'TestValue',
      Items) = 0 then
    begin
    ...
    end;
