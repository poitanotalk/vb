## :books: Função WritePrivateProfileStringA - Kernel32.dll

#### :dependabot: Artigo · 11 Set 2023

Esta função serve para escrever ou definir uma cadeia de caracteres na seção e chave especificadas no perfil privado.

## Sintaxe

#### Visual Basic
```basic
Private Declare Ansi Function WritePrivateProfileString _
Lib "Kernel32.dll" _
Alias "WritePrivateProfileStringA" (
    ByVal lpAppName As System.String,    ' [in] LPCSTR
    ByVal lpKeyName As System.String,    ' [in] LPCSTR
    ByVal lpString As System.Byte(),     ' [in] LPCSTR
    ByVal lpFileName As System.String    ' [in] LPCSTR
) As System.Boolean                      ' [out] BOOL
```

#### Visual Basic .NET

```basic
<System.Runtime.InteropServices.DllImport(
"Kernel32.dll",
CharSet:=System.Runtime.InteropServices.CharSet.Ansi,
EntryPoint:="WritePrivateProfileStringA")>
Private Shared Function WritePrivateProfileString(
    ByVal lpAppName As System.String,    ' [in] LPCSTR
    ByVal lpKeyName As System.String,    ' [in] LPCSTR
    ByVal lpString As System.Byte(),     ' [in] LPCSTR
    ByVal lpFileName As System.String    ' [in] LPCSTR
) As System.Boolean                      ' [out] BOOL
    ' Deixar o corpo da função vazio.
End Function
```

## Parâmetros

|Nome                      |Descrição                                                                                                                                                                |
|:-------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|**lpAppName**             |Representa o nome da **Seção** na qual a cadeia de caracteres será escrita. O nome da **Seção** não será diferente de **SEÇÃO**, independente de maiúsculas e minúsculas.|
|**lpKeyName**             |Representa o nome da **Chave** para a qual a **Cadeia de Caracteres** será escrita. Se este parâmetro for **Nothing** toda a seção será excluída.                        |
|**lpString**              |Representa a **Cadeia de Caracteres** terminando em caractere **Null** que será escrita. Se este parâmetro for **Nothing** a chave será excluída.                        |
|**lpFileName**            |Representa o **Caminho** e o nome do **Arquivo** para a qual a **Cadeia de Caracteres** será escrita.                                                                    |
|**lpReturnedBool**        |Se a função for bem-sucedida, o valor retornado será **True** ou diferente de zero, caso contrário, o valor retornado será **False** ou igual a zero.                    |



## Exemplos

Se o *Arquivo* não existir ele será criado.

```basic
Dim lpReturnedBool As System.Boolean = WritePrivateProfileString(
    "Section1",
    "Key1",
    System.Text.Encoding.Default.GetBytes("Value1" & Chr(0)),
    System.IO.Path.GetFullPath(".\File.ini")
)
```

Se o valor da *Cadeia de Caracteres* for o caractere *Null* apenas o valor será excluído.

```basic
Dim lpReturnedBool As System.Boolean = WritePrivateProfileString(
    "Section1",
    "Key1",
    System.Text.Encoding.Default.GetBytes(Chr(0)),
    System.IO.Path.GetFullPath(".\File.ini")
)
```

Se o valor da *Cadeia de Caracteres* for igual a *Nothing* a *Chave* será excluída.

```basic
Dim lpReturnedBool As System.Boolean = WritePrivateProfileString(
    "Section1",
    "Key1",
    Nothing,
    System.IO.Path.GetFullPath(".\File.ini")
)
```

Se o valor da *chave*, independentemente do valor da *cadeia de caracteres*, for igual a *Nothing* a *Seção* toda será excluída.

```basic
Dim lpReturnedBool As System.Boolean = WritePrivateProfileString(
    "Section1",
    Nothing,
   System.Text.Encoding.Default.GetBytes("Value1" & Chr(0)),
   System.IO.Path.GetFullPath(".\File.ini")
)
```

## Pré-requisitos para uso da função

|  Requisito                      | Descrição                                                           |
| ------------------------------- | ------------------------------------------------------------------- |
| __Cliente mínimo com suporte__  |	Windows 2000 Professional [somente aplicativos da área de trabalho] |
| __Servidor mínimo com suporte__ | Windows 2000 Server [somente aplicativos da área de trabalho]       |
| __Plataforma de Destino__	      | Windows                                                             |
| __Cabeçalho	winbase.h__         | (incluir Windows.h)                                                 |
| __Biblioteca__                  | Kernel32.lib                                                        |
| __DLL__                         | Kernel32.dll                                                        |

## Confira também

GetPrivateProfileString

Writeprofilestring

----

Copyright © 2023 @Fabasa-Pro Developer.
