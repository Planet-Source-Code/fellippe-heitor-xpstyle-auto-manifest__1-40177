Attribute VB_Name = "XPStyleAutoManifest"
Option Explicit
'Portugu�s do Brasil:
'   Este m�dulo faz com que sua aplica��o VB adquira
'nativamente os estilos visuais do Windows XP. N�o �
'necess�rio alterar seu c�digo original e qualquer
'programa pode ser adaptado. A �nica coisa que vc deve
'fazer � adicionar este m�dulo ao seu projeto e defini-lo
'como inicial (Project/Project Properties/Startup object/
'Sub Main). A �nica modifica��o necess�ria ser� se vc
'j� tiver um m�dulo com sub main... Altere essa sub e
'inclua a chamada para ela no fim da sub main DESTE
'M�DULO.
'   O processo consiste em gerar um arquivo .MANIFEST
'na mesma pasta de sua aplica��o. Se acidentalmente
'o usu�rio apagar o arquivo este ser� gerado novamente.
'O m�dulo tamb�m inicializa a fun��o InitCommonControls
'de COMCTL32.DLL (a vers�o 6, necess�ria para os Temas).
'E pronto! Seu programa com a cara de qualquer tema XP sem
'necessidade de reformula��o de design ou chamadas de API
'(al�m das que eu j� usei aqui!). Aproveite!!! ;-)
'
'Fellippe Heitor - Brasil
'fellippe.heitor@globo.com
'"Mude este c�digo como quiser e achar �til mas me
'me deixe saber (mande-o para mim) para que eu possa
'ver o que foi melhorado."
'-----------------------------
'English:
'   This module makes your VB application use natively WinXP
'Visual Styles. It's not necessary to change your source code
'and any program can be adapted. The only thing you must do is
'to add this module to your project and to define it as the
'startup object for your application (Project/Project
'Properties/Startup object/Sub Main). The only necessary
'change takes place if your current project already has a
'Sub Main()... Change this sub's name and add a call to it
'at the end of the Sub Main() IN THIS MODULE.
'   The process consists in creating a .MANIFEST file in the
'same folder of your program. If accidentaly the user erases
'this file it'll be created again (automatically). This module
'also initializes the function InitCommonControls from the library
'COMCTL32.DLL (version 6, needed to get themes working. And this
'is all! Your program with the same face as any winxp theme
'without the need of redesigning your forms and with no API
'calls (other than the ones I used here!). Enjoy!!! ;-)
'
'Fellippe Heitor - Brazil
'fellippe.heitor@globo.com
'"Change this code anyway you find useful but please
'let me know (send it to me) so I can see what is
'improved."
'
'
Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Sub Main()
    'PORTUGU�S: Este deve ser o objeto de inicializa��o do seu projeto
    '(Sub Main). Se for necess�rio chamar uma sub de inicializa��o para
    'o seu programa, voc� deve colocar esta chamada no fim deste procedimento.
    'Se seu programa tiver um formul�rio inicial ele deve ser chamado tamb�m
    'no fim deste procedimento (veja explica��o abaixo).
    '
    'ENGLISH: This must be the startup of your project. If any other
    'Sub should be called first, it'll be called after this one.
    'If your application has a startup form, then you must change
    'the last 2 lines of this Sub to set it as the form or sub to be
    'called.
    '
    
    Dim fs, sManifest As String, Exe As String, X As Long
    Dim Desc As String, MustRestart As Boolean
    
    Exe = App.EXEName + ".exe"
    Desc = App.Comments
    
        sManifest = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>"
        sManifest = sManifest & vbCrLf & "<assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & ">"
        sManifest = sManifest & vbCrLf & "<assemblyIdentity"
        sManifest = sManifest & vbCrLf & "  name=" & Chr(34) & Exe & Chr(34)
        sManifest = sManifest & vbCrLf & "  processorArchitecture=" & Chr(34) & "x86" & Chr(34)
        sManifest = sManifest & vbCrLf & "  version=" & Chr(34) & "1.0.0.0" & Chr(34)
        sManifest = sManifest & vbCrLf & "  type=" & Chr(34) & "win32" & Chr(34) & "/>"
        sManifest = sManifest & vbCrLf & "<description>" & Desc & "</description>"
        sManifest = sManifest & vbCrLf & "<dependency>"
        sManifest = sManifest & vbCrLf & "  <dependentAssembly>"
        sManifest = sManifest & vbCrLf & "    <assemblyIdentity"
        sManifest = sManifest & vbCrLf & "      type=" & Chr(34) & "win32" & Chr(34)
        sManifest = sManifest & vbCrLf & "      name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34)
        sManifest = sManifest & vbCrLf & "      version=" & Chr(34) & "6.0.0.0" & Chr(34)
        sManifest = sManifest & vbCrLf & "      processorArchitecture=" & Chr(34) & "x86" & Chr(34)
        sManifest = sManifest & vbCrLf & "      publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34)
        sManifest = sManifest & vbCrLf & "      language=" & Chr(34) & "*" & Chr(34)
        sManifest = sManifest & vbCrLf & "    />"
        sManifest = sManifest & vbCrLf & "  </dependentAssembly>"
        sManifest = sManifest & vbCrLf & "</dependency>"
        sManifest = sManifest & vbCrLf & "</assembly>"

    MustRestart = False
    X = InitCommonControls
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.fileexists(App.Path + "\" + App.EXEName + ".exe.MANIFEST") Then
        MustRestart = True
        Open App.Path + "\" + App.EXEName + ".exe.MANIFEST" For Binary As 1
        Put #1, 1, sManifest
        Close 1
    End If
    
    If MustRestart Then
        Shell App.Path + "\" + App.EXEName + ".exe", vbNormalFocus
        End
    End If
    
    'PORTUGU�S: Especifique aqui o formul�rio principal ou sub que deve ser
    'chamada para o seu programa. Altere "MyForm"
    '
    'ENGLISH: Specify here the main Form of your program
    'or any sub that should be called.
    'Change "MyForm":

    Load MyForm
    MyForm.Show
End Sub
