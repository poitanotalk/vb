## :books: Função WritePrivateProfileStringA - Kernel32.dll

#### :dependabot: Artigo · 11 de set de 2023

A função **WritePrivateProfileStringA** *grava* ou *define* uma **Cadeia de Caracteres** na codificação **Ansi** para a **Seção** e **Chave** especificadas do perfil privado de um **Arquivo** INI.

## Sintaxe

**`Visual Basic`**
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

**`Visual Basic .NET`**
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

|Nome                      |Descrição                                                                                                                                                                          |
|:-------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|**lpAppName**             |Representa o nome da **Seção** na qual a **Cadeia de Caracteres** será escrita. O nome **Seção** não será diferente de **SEÇÃO**, ou seja, independente de maiúsculas e minúsculas.|
|**lpKeyName**             |Representa o nome da **Chave** para a qual a **Cadeia de Caracteres** será escrita. Se este parâmetro for **Nothing** toda a **Seção** será excluída.                              |
|**lpString**              |Representa a **Cadeia de Caracteres** terminanda em **Null** que será escrita. Se este parâmetro for **Nothing** a **Chave** será excluída.                                        |
|**lpFileName**            |Representa o **Caminho** e o nome do **Arquivo** para a qual a **Cadeia de Caracteres** será escrita.                                                                              |
|**lpReturnedBool**        |Se a função for bem-sucedida, o valor retornado será **True** ou diferente de zero, caso contrário, o valor retornado será **False** ou igual a zero.                              |



## Exemplos

Este exemplo mostra a função *escrevendo* ou *definindo* um valor **"Value1" & Chr(0)** na **Section1** e **Key1** especificadas no perfil privado de um **Arquivo** *File.ini* no **Caminho** e **Diretório** existentes.

```basic
Dim lpReturnedBool As System.Boolean = WritePrivateProfileString(
    "Section1",
    "Key1",
    System.Text.Encoding.Default.GetBytes("Value1" & Chr(0)),
    System.IO.Path.GetFullPath(".\File.ini")
)
```


O exemplo a seguir *remove* ou *limpa* o espaço da **Cadeia de Caracteres**, que foi *escrita* ou *definida*, apenas escrevendo o caractere **Null**.

```basic
Dim lpReturnedBool As System.Boolean = WritePrivateProfileString(
    "Section1",
    "Key1",
    System.Text.Encoding.Default.GetBytes(Chr(0)),
    System.IO.Path.GetFullPath(".\File.ini")
)
```

Já este exemplo *exclui permanentemente* toda a **Chave**, na **Seção** e na **Chave** especificadas no perfil privado do **Arquivo**, se o valor da **Cadeia de Caracteres** for apenas **Nothing**.

```basic
Dim lpReturnedBool As System.Boolean = WritePrivateProfileString(
    "Section1",
    "Key1",
    Nothing,
    System.IO.Path.GetFullPath(".\File.ini")
)
```

E este exemplo *exclui permanentemente* toda a **Seção** se o valor da **Chave** for apenas **Nothing**, independentemente se o valor da **Cadeia de Caracteres** estiver **Nothing** .

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
