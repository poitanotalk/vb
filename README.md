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

`[in] lpAppName` Representa o nome da *seção* na qual a cadeia de caracteres será escrita. O nome da *seção* não será diferente de *SEÇÃO*, independente de maiúsculas e minúsculas.

`[in] lpKeyName` Representa o nome da *chave* para a qual a cadeia de caracteres será escrita. Se este parâmetro for **NULL** toda a seção será excluída.

`[in] lpString` Representa a *cadeia de caracteres* terminando em **NULL** que será escrita. Se este parâmetro for **NULL** a chave será excluída.

`[in] lpFileName` Representa o nome do *arquivo* para a qual a cadeia de caracteres será escrita. Se o arquivo for criado com codificação **Unicode**, a função escreve caracteres *Unicode*, caso contrário, a função escreve caracteres *Default* **ANSI**.

## Retorno

Se a função for bem-sucedida, o valor retornado será `True` ou diferente de zero, caso contrário, o valor retornado será `False` ou igual a zero.

## Exemplo 1







O arquivo [File.ini](https://github.com/fabasapro/fabasapro/blob/main/LICENSE.TXT) deve ter o seguinte formato:

```file
[Section1]
Key1=Value1
```

O arquivo [Module1.vb](https://github.com/fabasapro/fabasapro/blob/main/LICENSE.TXT) deve ter o seguinte formato:

```basic
' Licenciado sob a licença MIT.
' Copyright (C) 1991-2023 Fabasa-Pro Developer.
' Consulte LICENSE.TXT na raiz do projeto para obter informações.
Module Module1
    ' Esta função serve para escrever ou definir uma cadeia de caracteres na
    ' seção e chave especificadas no perfil privado.
    Private Declare Ansi Function WritePrivateProfileString _
        Lib "Kernel32.dll" _
        Alias "WritePrivateProfileStringA" (
        ByVal lpAppName As System.String,    ' [in] LPCSTR
        ByVal lpKeyName As System.String,    ' [in] LPCSTR
        ByVal lpString As System.Byte(),     ' [in] LPCSTR
        ByVal lpFileName As System.String    ' [in] LPCSTR
    ) As System.Boolean                      ' [out] BOOL
    Sub Main()
        ' Escrever ou definir uma cadeia de caracteres de 32767 bytes na seção e
        ' chave especificadas no perfil privado.
        Dim lpString As System.Byte() =
            System.Text.Encoding.Default.GetBytes("Value1")
        If WritePrivateProfileString(
                "Section1",
                "Key1",
                lpString,    ' Comprimento de 0 a 32767 bytes
                System.IO.Path.GetFullPath(".\File.ini")
            ) Then
            System.Console.WriteLine("Escrita bem-sucedida!")
        Else
            System.Console.WriteLine("Erro ao escrever!")
        End If
        System.Console.ReadKey(True)    ' Pausar
    End Sub
End Module
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

Os projetos da comunidade @Fabasa-Pro Developer adotam uma [licença MIT](https://github.com/fabasapro/fabasapro/blob/main/LICENSE). Consulte o arquivo [LICENSE.TXT](https://github.com/fabasapro/fabasapro/blob/main/LICENSE.TXT), para obter mais informações.

