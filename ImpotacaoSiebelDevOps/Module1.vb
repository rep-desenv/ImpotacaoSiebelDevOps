Imports System.IO
Imports System.IO.Directory
Imports System.Security
Imports System.Text

Module Module1

    Public Conexao As String
    Dim Siebel As Object
    Dim arquivoErro As StreamWriter
    Public fs As FileStream
    Public mysw As StreamWriter

    Dim fsOut As FileStream
    Dim myswOut As StreamWriter

    Public errCode As Short
    Public sDiretorioErro As String


    Dim sCFG As String  ''  Caminho e arquivo CFG
    Public sUserName As String '' User DBF/CFG
    Dim sPassword As String '' Password DBF/CFG
    Dim sRepositorio As String
    Dim Trace As String = ""


    Dim ErrSiebel As String

    Public SiebelApplication As SiebelDataServer.SiebelApplication
    Public TheApplication As SiebelApplicationServer.SiebelApplication

    Public NomeArquivoLog As String = "Erro_Processo_DevOps.log"
    Dim NomeArquivoLogOut As String

    Dim NomeObjetoDestino As String

   

  

    Private Sub DeletaRepositorio(ByVal sArquivo As String)


        Dim Linha As String
        Dim Tamanho As Integer
        Dim NomeObjeto As String
        Dim IsRecord As Integer

        Dim BusObject As String = ""
        Dim BusComp As String = ""
        Dim Rep As String


        Dim oBORepo As SiebelDataServer.SiebelBusObject
        Dim oBCRepo As SiebelDataServer.SiebelBusComp

        Dim oBORep As SiebelDataServer.SiebelBusObject
        Dim oBCRep As SiebelDataServer.SiebelBusComp


        oBORep = SiebelApplication.GetBusObject("Repository Project", errCode)
        oBCRep = oBORep.GetBusComp("Repository Project", errCode)

        If InStr(sArquivo, "APT_") > 0 Then
            BusObject = "Repository Applet"
            BusComp = "Repository Applet"
        ElseIf InStr(sArquivo, "VIW_") > 0 Then
            BusObject = "Repository View"
            BusComp = "Repository View"
        ElseIf InStr(sArquivo, "BUC_") > 0 Then
            BusObject = "Repository Business Component"
            BusComp = "Repository Business Component"
        Else
            Exit Sub
        End If


        If BusObject.Length > 0 Then

            Dim texto As New StreamReader(sArquivo, System.Text.Encoding.Default)


            While Not texto.EndOfStream
                Linha = texto.ReadLine

                If InStr(1, Linha, "<Name>") > 0 Then
                    NomeObjeto = Linha.Substring(InStr(1, Linha, "<Name>") + 5)
                    Tamanho = InStr(1, NomeObjeto, "</Name>")
                    NomeObjeto = Mid(NomeObjeto, 1, Tamanho - 1)
                    Exit While
                End If

            End While

            oBORep = SiebelApplication.GetBusObject("Repository Project", errCode)
            oBCRep = oBORep.GetBusComp("Repository Project", errCode)



            With oBCRep
                .ClearToQuery(errCode)
                .SetViewMode(1, errCode)
                .ActivateField("Name", errCode)
                .SetSearchSpec("Name", "Siebel Repository", errCode)
                .ExecuteQuery(True, errCode)

                If .FirstRecord(errCode) <> 0 Then
                    Rep = .GetFieldValue("Id", errCode)
                End If

            End With

            oBCRep = Nothing
            oBORep = Nothing


            oBORepo = SiebelApplication.GetBusObject(BusObject, errCode)
            oBCRepo = oBORepo.GetBusComp(BusComp, errCode)

            With oBCRepo
                .ClearToQuery(errCode)
                .SetViewMode(1, errCode)
                .ActivateField("Name", errCode)
                .ActivateField("Repository Id", errCode)
                .SetSearchSpec("Name", NomeObjeto, errCode)
                .SetSearchSpec("Repository Id", Rep, errCode)
                .ExecuteQuery(False, errCode)
                IsRecord = .FirstRecord(errCode)

                If IsRecord <> 0 Then
                    While IsRecord <> 0
                        .DeleteRecord(errCode)
                        IsRecord = .NextRecord(errCode)
                    End While
                    Console.WriteLine("Apagou Objeto")
                End If

            End With

            oBORepo = Nothing
            oBCRepo = Nothing

        End If


    End Sub

   

    Private Function ValidaObjetoDestino(ByVal sArquivo As String) As String

        Dim Linha As String
        Dim Tamanho As Integer

        Dim SiebelAppObject As Object

        Dim BOObject As SiebelDataServer.SiebelBusObject
        Dim BCObject As SiebelDataServer.SiebelBusComp

        Dim BOAppli As SiebelApplicationServer.SiebelBusObject
        Dim BCAppli As SiebelApplicationServer.SiebelBusComp

        Console.WriteLine(sArquivo)

        ValidaObjetoDestino = "No"
        NomeObjetoDestino = ""


        Try

            Dim texto As New StreamReader(sArquivo, System.Text.Encoding.Default)
            While Not texto.EndOfStream
                Linha = texto.ReadLine

                If InStr(1, Linha, "<K>") > 0 Then
                    NomeObjetoDestino = Linha.Substring(InStr(1, Linha, "<K>") + 2)
                    Tamanho = InStr(1, NomeObjetoDestino, "</K>")
                    NomeObjetoDestino = Mid(NomeObjetoDestino, 1, Tamanho - 1)
                    Exit While
                End If

            End While

            If Conexao = "DataServer" Then

                BOObject = SiebelApplication.GetBusObject("Front Office Workflow", errCode)
                BCObject = BOObject.GetBusComp("Workflow Process Definition", errCode)

                With BCObject
                    .ClearToQuery(errCode)
                    .SetViewMode(1, errCode)
                    .ActivateField("Process Name", errCode)
                    .SetSearchSpec("Process Name", "'" & NomeObjetoDestino & "'", errCode)
                    .ExecuteQuery(True, errCode)


                    If .FirstRecord(errCode) <> 0 Then
                        ValidaObjetoDestino = "Yes"
                        Console.WriteLine("Existe WF")
                    End If

                End With

                BCObject = Nothing
                BOObject = Nothing

            Else

                Try
                    SiebelAppObject = GetObject("", "SiebelAppServer.ApplicationObject")


                    BOAppli = SiebelAppObject.GetBusObject("Front Office Workflow", errCode)
                    BCAppli = BOAppli.GetBusComp("Workflow Process Definition", errCode)

                    With BCAppli
                        .ClearToQuery(errCode)
                        .SetViewMode(1, errCode)
                        .ActivateField("Process Name", errCode)
                        .SetSearchSpec("Process Name", "'" & NomeObjetoDestino & "'", errCode)
                        .ExecuteQuery(True, errCode)

                        If .FirstRecord(errCode) <> 0 Then
                            ValidaObjetoDestino = "Yes"
                            Console.WriteLine("Existe WF")
                        End If


                    End With

                    BCAppli = Nothing
                    BOAppli = Nothing

                Catch ex As Exception

                    If InStr(ex.Message, "ActiveX") > 0 Then

                        Console.WriteLine("AppServer - Erro de conexão com Siebel Client: " + Err.Description)
                        Console.WriteLine("Favor abrir o Siebel Client e repetir a operação !")

                        fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                        mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                        mysw.WriteLine("[" & sUserName & "] " & Now + " - AppServer - Erro Exception de conexão com Siebel Client: " + Err.Description)
                        mysw.Close()

                        Exit Function

                    Else

                        fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                        mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                        mysw.WriteLine("[" & sUserName & "] " & Now + " - AppServer - Erro Exception ValidaObjetoDestino Descricao  : " + ex.Message)
                        mysw.Close()

                        Console.WriteLine("AppServer - Erro Exception ValidaObjetoDestino Descricao  : " + ex.Message)

                    End If

                End Try


            End If


        Catch ex As Exception

            ValidaObjetoDestino = "Erro"
            Console.WriteLine("XML  : " + sArquivo)
            Console.WriteLine("Erro Exception ValidaObjetoDestino  : " + ex.Message)
            Console.WriteLine("XML Erro ")

            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
            mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro ValidaObjetoDestino XML : " & sArquivo)
            mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro Exception ValidaObjetoDestino : " + Err.Description)
            mysw.Close()

        End Try

    End Function

   


  

    Private Function ValidaProjeto(ByVal sArquivo As String) As String

        Dim Linha As String
        Dim NomeProjeto As String = ""
        Dim Tamanho As Integer
        Dim Rep As String
        Dim Existe As Integer = 0
        Dim LockedProject As String = ""
        Dim CriouProjeto As String = ""

        Dim BC As SiebelDataServer.SiebelBusComp
        Dim BO As SiebelDataServer.SiebelBusObject

        Dim oBO As SiebelDataServer.SiebelBusObject
        Dim oBC As SiebelDataServer.SiebelBusComp

        Dim oBCDev As SiebelDataServer.SiebelBusComp
        Dim oBODev As SiebelDataServer.SiebelBusObject

        Dim BOApp As SiebelApplicationServer.SiebelBusObject
        Dim BCApp As SiebelApplicationServer.SiebelBusComp

        Dim oBCApp As SiebelApplicationServer.SiebelBusComp
        Dim oBOApp As SiebelApplicationServer.SiebelBusObject

        Dim oBCRep As SiebelApplicationServer.SiebelBusComp
        Dim oBORep As SiebelApplicationServer.SiebelBusObject




        Try

            If InStr(1, sArquivo, "PRJ_") > 0 Then '' É Projeto
                ValidaProjeto = "Ok"
                Exit Function
            End If

            Dim texto As New StreamReader(sArquivo, System.Text.Encoding.Default)
            While Not texto.EndOfStream
                Linha = texto.ReadLine

                If InStr(1, Linha, "<ProjectName>") > 0 Then
                    NomeProjeto = Linha.Substring(InStr(1, Linha, "<ProjectName>") + 12)
                    Tamanho = InStr(1, NomeProjeto, "</ProjectName>")
                    NomeProjeto = Mid(NomeProjeto, 1, Tamanho - 1)
                    Exit While
                End If

                'If InStr(1, Linha, "<ProjectLocked>Y</ProjectLocked>") > 0 Then
                '    LockedProject = "Y"
                'ElseIf InStr(1, Linha, "<ProjectLocked>N</ProjectLocked>") > 0 Then
                '    LockedProject = "N"
                'End If

            End While

            If NomeProjeto.Length > 0 Then

                If Conexao = "DataServer" Then

                    Dim sID = SiebelApplication.LoginId(errCode)

                    BO = SiebelApplication.GetBusObject("Repository Repository", errCode)
                    BC = BO.GetBusComp("Repository Repository", errCode)

                    oBODev = SiebelApplication.GetBusObject("Repository Project DevOps", errCode)
                    oBCDev = oBODev.GetBusComp("Repository Project DevOps", errCode)


                    oBO = SiebelApplication.GetBusObject("Repository Project", errCode)
                    oBC = oBO.GetBusComp("Repository Project", errCode)

                    With BC
                        .ClearToQuery(errCode)
                        .SetViewMode(1, errCode)
                        .ActivateField("Name", errCode)
                        .SetSearchSpec("Name", "Siebel Repository", errCode)
                        .ExecuteQuery(True, errCode)

                        If .FirstRecord(errCode) <> 0 Then
                            Rep = .GetFieldValue("Id", errCode)
                        End If

                    End With

                    BC = Nothing
                    BO = Nothing


                    With oBC
                        .ClearToQuery(errCode)
                        .SetViewMode(1, errCode)
                        .ActivateField("Name", errCode)
                        .ActivateField("Locked", errCode)
                        .ActivateField("Repository Id", errCode)
                        .ActivateField("Inactive", errCode)
                        .SetSearchSpec("Name", "'" & NomeProjeto & "'", errCode)
                        .SetSearchSpec("Repository Id", Rep, errCode)
                        .ExecuteQuery(True, errCode)

                        If .FirstRecord(errCode) <> 0 Then
                            Existe = 1
                            ValidaProjeto = "Ok"
                        End If

                        'If LockedProject = "Y" Then
                        '    Console.WriteLine("Vai Locar o Projeto : " & NomeProjeto)
                        '    .SetFieldValue("Repository Id", Rep, errCode)
                        '    .SetFieldValue("Locked", "Y", errCode)
                        '    .SetFieldValue("Inactive", "N", errCode)
                        '    .SetFieldValue("Locked By Id", sID, errCode)
                        '    .WriteRecord(errCode)
                        'End If

                    End With

                    oBO = Nothing
                    oBC = Nothing

                    If Existe = 0 Then

                        With oBCDev
                            .ActivateField("Locked By Id", errCode)
                            .NewRecord(1, errCode)
                            .SetFieldValue("Repository Id", Rep, errCode)
                            .SetFieldValue("Name", NomeProjeto, errCode)
                            .SetFieldValue("Locked", "Y", errCode)
                            .SetFieldValue("Inactive", "N", errCode)
                            .SetFieldValue("Locked By Id", sID, errCode)
                            .WriteRecord(errCode)
                            ValidaProjeto = "Ok"
                        End With

                        oBCDev = Nothing
                        oBODev = Nothing

                        Console.WriteLine("Foi criado o projeto: " & NomeProjeto)
                    End If

                Else

                    Try
                        Siebel = GetObject("", "SiebelAppServer.ApplicationObject")

                    Catch ex As Exception

                        If InStr(ex.Message, "ActiveX") > 0 Then

                            Console.WriteLine("AppServer - Erro de conexão com Siebel Client: " + Err.Description)
                            Console.WriteLine("Favor abrir o Siebel Client e repetir a operação !")
                            Console.WriteLine("Tecle enter para sair......")
                            ''Console.Read()

                            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                            mysw.WriteLine("[" & sUserName & "] " & Now + " - AppServer - Erro Exception de conexão com Siebel Client: " + Err.Description)
                            mysw.Close()

                            ValidaProjeto = "NOX"

                            Exit Function

                        Else

                            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                            mysw.WriteLine("[" & sUserName & "] " & Now + " - AppServer - Erro Exception ValidaProjeto Descricao  : " + ex.Message)
                            mysw.Close()

                            Console.WriteLine("AppServer - Erro Exception ValidaProjeto Descricao  : " + ex.Message)
                            '' Console.WriteLine("Tecle enter para sair......")
                            ''Console.Read()

                        End If

                    End Try

                    Dim sID = Siebel.LoginId(errCode)


                    BOApp = Siebel.GetBusObject("Repository Repository", errCode)
                    BCApp = BOApp.GetBusComp("Repository Repository", errCode)

                    With BCApp
                        .ClearToQuery(errCode)
                        .SetViewMode(1, errCode)
                        .ActivateField("Name", errCode)
                        .SetSearchSpec("Name", "Siebel Repository", errCode)
                        .ExecuteQuery(True, errCode)

                        If .FirstRecord(errCode) <> 0 Then
                            Rep = .GetFieldValue("Id", errCode)
                        End If

                    End With

                    BCApp = Nothing
                    BOApp = Nothing

                    oBORep = Siebel.GetBusObject("Repository Project", errCode)
                    oBCRep = oBORep.GetBusComp("Repository Project", errCode)

                    With oBCRep
                        .ClearToQuery(errCode)
                        .SetViewMode(1, errCode)
                        .ActivateField("Name", errCode)
                        .ActivateField("Locked", errCode)
                        .ActivateField("Repository Id", errCode)
                        .ActivateField("Inactive", errCode)
                        .SetSearchSpec("Name", "'" & NomeProjeto & "'", errCode)
                        .ExecuteQuery(True, errCode)

                        If .FirstRecord(errCode) <> 0 Then
                            Existe = 1
                            ValidaProjeto = "Ok"
                        End If

                        If LockedProject = "Y" Then
                            .SetFieldValue("Repository Id", Rep, errCode)
                            .SetFieldValue("Locked", "Y", errCode)
                            .SetFieldValue("Inactive", "N", errCode)
                            .SetFieldValue("Locked By Id", sID, errCode)
                            .WriteRecord(errCode)
                        End If

                    End With

                    oBORep = Nothing
                    oBCRep = Nothing

                    If Existe = 0 Then

                        oBOApp = Siebel.GetBusObject("Repository Project DevOps", errCode)
                        oBCApp = oBOApp.GetBusComp("Repository Project DevOps", errCode)

                        With oBCApp
                            .ActivateField("Locked By Id", errCode)
                            .NewRecord(1, errCode)
                            .SetFieldValue("Repository Id", Rep, errCode)
                            .SetFieldValue("Name", NomeProjeto, errCode)
                            .SetFieldValue("Locked", "Y", errCode)
                            .SetFieldValue("Inactive", "N", errCode)
                            .SetFieldValue("Locked By Id", sID, errCode)
                            .WriteRecord(errCode)
                            ValidaProjeto = "Ok"
                        End With

                        oBCApp = Nothing
                        oBOApp = Nothing

                        Console.WriteLine("Foi criado o projeto: " & NomeProjeto)
                    End If
                End If
            Else

                ValidaProjeto = "NOk"

                If sRepositorio.ToUpper = "IN" Then

                    NomeArquivoLogOut = "Erro_Arquivos_in.txt"

                    fsOut = New FileStream(sDiretorioErro + NomeArquivoLogOut, FileMode.Append)
                    myswOut = New StreamWriter(fsOut, System.Text.Encoding.UTF8)
                    myswOut.WriteLine("[" & Now & "] " & "XML : " & Trim$(sArquivo))
                    myswOut.WriteLine("ErroSiebel ValidaProjeto: Nome do Projeto não encontrado")
                    myswOut.WriteLine("")
                    myswOut.Close()


                End If

            End If


        Catch ex As Exception

            ValidaProjeto = "NOk"
            Console.WriteLine("XML  : " + sArquivo)
            Console.WriteLine("Erro Exception ValidaProjeto  : " + ex.Message)
            Console.WriteLine("XML Erro ")
            ''Console.WriteLine("Tecle enter para sair......")
            '' Console.Read()

            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
            mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro ValidaProjeto XML : " & sArquivo)
            mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro Exception ValidaProjeto : " + Err.Description)
            mysw.Close()


            If sRepositorio.ToUpper = "IN" Then

                NomeArquivoLogOut = "Erro_Arquivos_in.txt"

                fsOut = New FileStream(sDiretorioErro + NomeArquivoLogOut, FileMode.Append)
                myswOut = New StreamWriter(fsOut, System.Text.Encoding.UTF8)
                myswOut.WriteLine("[" & Now & "] " & "XML : " & Trim$(sArquivo))
                myswOut.WriteLine("ErroSiebel ValidaProjeto : " & Err.Description)
                myswOut.WriteLine("")
                myswOut.Close()


            End If

        Finally


        End Try


    End Function

    Private Sub ImportarObjetos(ByVal sDiretorioArquivoBatch As String)

        Dim linhaTexto As String
        Dim RetornoValida As String

        Dim RetornoValidacaoObjeto As String


        Console.WriteLine("Arquivo : " + sDiretorioArquivoBatch)
        Console.WriteLine("")


        Try

            Using sr As TextReader = New StreamReader(sDiretorioArquivoBatch, System.Text.Encoding.Default)

                linhaTexto = sr.ReadLine

                ''If Len(linhaTexto) > 0 Then
                If Conexao = "AppSever" Then
                    Try
                        Siebel = GetObject("", "SiebelAppServer.ApplicationObject")

                    Catch ex As Exception

                        If InStr(ex.Message, "ActiveX") > 0 Then

                            Console.WriteLine("AppServer - Erro de conexão com Siebel Client: " + Err.Description)
                            Console.WriteLine("Favor abrir o Siebel Client e repetir a operação !")
                            '' Console.WriteLine("Tecle enter para sair......")
                            '' Console.Read()

                            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                            mysw.WriteLine("[" & sUserName & "] " & Now + " - AppServer - Erro Exception de conexão com Siebel Client: " + Err.Description)
                            mysw.Close()

                            Exit Sub

                        Else

                            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                            mysw.WriteLine("[" & sUserName & "] " & Now + " - AppServer - Erro Exception Importacao Descricao  : " + ex.Message)
                            mysw.Close()

                            Console.WriteLine("AppServer - Erro Exception Importacao Descricao  : " + ex.Message)
                            '' Console.WriteLine("Tecle enter para sair......")
                            '' Console.Read()

                        End If

                    End Try

                    Dim SVC = Siebel.GetService("Workflow Process Manager", errCode)

                    If errCode <> 0 Then

                        Console.WriteLine("AppServer GetService Workflow Process Manager : " & errCode)
                        Console.WriteLine("ErroSiebel Descricao : " & Siebel.GetLastErrText)
                        ''Console.ReadLine()

                        fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                        mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                        mysw.WriteLine("[" & sUserName & "] " & Now + " - AppServer - ErroSiebel Workflow Process Manager : " & CStr(errCode))
                        mysw.WriteLine("ErroSiebel Descricao : " & Siebel.GetLastErrText)
                        mysw.Close()


                        GoTo Proximo

                    End If

                    Dim inputs = Siebel.NewPropertySet(errCode)
                    Dim outputs = Siebel.NewPropertySet(errCode)



                    inputs.SetProperty("ProcessName", "DevOps_Importa_XML_Object", errCode)


                    Do While linhaTexto <> Nothing
                        If Len(linhaTexto) > 0 Then

                            Try

                                If sRepositorio.ToUpper = "IN" Then
                                    RetornoValida = ValidaProjeto(Trim$(linhaTexto))
                                End If

                                If InStr(linhaTexto, "LOV_") <> 0 Then
                                    If DeletaLOV(linhaTexto) = False Then
                                        GoTo Proximo
                                    End If
                                End If


                                If InStr(linhaTexto, "RGN_") <> 0 Then
                                    If DeletaRGN(linhaTexto) = False Then
                                        GoTo Proximo
                                    End If
                                End If

                                '@@@ Deletando Mapa de Valores EAI @@@
                                If InStr(linhaTexto, "EVL_") <> 0 Then
                                    If DeletaEVL(linhaTexto) = False Then
                                        GoTo Proximo
                                    End If
                                End If
                               

                                If (RetornoValida = "Ok" Or sRepositorio.ToUpper = "OUT") Then
                                    If (InStr(linhaTexto, "PAR_") <> 0 Or InStr(linhaTexto, "SPR_") <> 0 Or InStr(linhaTexto, "TRD_") <> 0 Or InStr(linhaTexto, "EDM_") <> 0 Or InStr(linhaTexto, "LOV_") <> 0) Then
                                        inputs.SetProperty("Repositorio", "out", errCode)
                                        Console.WriteLine("InsertUpdate")
                                    End If


                                    Try


                                        If InStr(linhaTexto, "WKF_") <> 0 Then

                                            RetornoValidacaoObjeto = ValidaObjetoDestino(linhaTexto)

                                            If RetornoValidacaoObjeto = "Yes" Then
                                                DeletaWF(NomeObjetoDestino)
                                            ElseIf RetornoValidacaoObjeto = "Erro" Then
                                                GoTo Proximo
                                            End If

                                        End If


                                       

                                        inputs.SetProperty("XML Name", Trim$(linhaTexto), errCode)
                                        SVC.InvokeMethod("RunProcess", inputs, outputs, errCode)

                                       

                                    Catch ex As Exception

                                        fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                                        mysw = New StreamWriter(fs, System.Text.Encoding.UTF8)
                                        mysw.WriteLine("[" & Now & "]" & "XML : " & Trim$(linhaTexto))
                                        mysw.WriteLine("[" & sUserName & "] " & Now + " - ErroSiebel Descricao : " & ex.Message)
                                        mysw.Close()

                                        Console.WriteLine("XML : " & Trim$(linhaTexto))
                                        Console.WriteLine("Erro Exception: " & ex.Message)



                                        If sRepositorio.ToUpper = "IN" Then

                                            NomeArquivoLogOut = "Erro_Arquivos_in.txt"

                                            fsOut = New FileStream(sDiretorioErro + NomeArquivoLogOut, FileMode.Append)
                                            myswOut = New StreamWriter(fsOut, System.Text.Encoding.UTF8)
                                            myswOut.WriteLine("[" & Now & "]" & "XML : " & Trim$(linhaTexto))
                                            myswOut.WriteLine("ErroSiebel Exception : " & ex.Message)
                                            myswOut.WriteLine("")
                                            myswOut.Close()

                                        Else

                                            NomeArquivoLogOut = "Erro_Arquivos_out.txt"

                                            fsOut = New FileStream(sDiretorioErro + NomeArquivoLogOut, FileMode.Append)
                                            myswOut = New StreamWriter(fsOut, System.Text.Encoding.UTF8)
                                            myswOut.WriteLine("[" & Now & "] " & "XML : " & Trim$(linhaTexto))
                                            myswOut.WriteLine("ErroSiebel Exception: " & ex.Message)
                                            myswOut.WriteLine("")
                                            myswOut.Close()

                                        End If


                                        GoTo Proximo
                                    End Try

                                    If errCode <> 0 Then

                                        ErrSiebel = Trim$(Siebel.GetLastErrText.Replace(Chr(10), ""))

                                        Console.WriteLine("AppServer ErroSiebel Executando WorkFlow : " & errCode)
                                        Console.WriteLine("ErroSiebel Descricao : " & Siebel.GetLastErrText)
                                        '' Console.WriteLine("Tecle enter para sair......")
                                        '' Console.Read()

                                        fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                                        mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                                        mysw.WriteLine("[" & sUserName & "] " & Now + " - AppServer - Erro Siebel Exportação durante WorkFlow  : " + CStr(errCode))
                                        mysw.WriteLine("[" & sUserName & "] " & Now + " - ErroSiebel Descricao : " & Siebel.GetLastErrText)
                                        mysw.Close()


                                        If sRepositorio.ToUpper = "IN" Then

                                            NomeArquivoLogOut = "Erro_Arquivos_in.txt"

                                            fsOut = New FileStream(sDiretorioErro + NomeArquivoLogOut, FileMode.Append)
                                            myswOut = New StreamWriter(fsOut, System.Text.Encoding.Default)
                                            myswOut.WriteLine("[" & Now & "]" & "XML : " & Trim$(linhaTexto))
                                            myswOut.WriteLine("ErroSiebel Executando WorkFlow : " & ErrSiebel)
                                            myswOut.WriteLine("")
                                            myswOut.Close()

                                        Else

                                            NomeArquivoLogOut = "Erro_Arquivos_out.txt"

                                            fsOut = New FileStream(sDiretorioErro + NomeArquivoLogOut, FileMode.Append)
                                            myswOut = New StreamWriter(fsOut, System.Text.Encoding.Default)
                                            myswOut.WriteLine("[" & Now & "] " & "XML : " & Trim$(linhaTexto))
                                            myswOut.WriteLine("ErroSiebel Executando WorkFlow : " & ErrSiebel)
                                            myswOut.WriteLine("")
                                            myswOut.Close()

                                        End If


                                        GoTo Proximo

                                    Else
                                        Console.WriteLine(Trim$(linhaTexto))
                                        Console.WriteLine("XML Importado com sucesso!")
                                        Console.WriteLine("")
                                        ''Console.Read()

                                    End If

                                ElseIf RetornoValida = "NOX" Then
                                    Exit Do
                                    GoTo Erro
                                ElseIf RetornoValida = "NOk" Then

                                    Console.WriteLine("XML : " & Trim$(linhaTexto))
                                    Console.WriteLine("Erro: Retorno ValidaProjeto")

                                    GoTo Proximo

                                Else
                                    Console.WriteLine(" Erro retorno ValidaProjeto : " & Trim$(linhaTexto))

                                End If


                            Catch ex As Exception

                                Console.WriteLine("AppServer - Erro Exception durante WorkFlow  : " + ex.Message)
                                ''Console.WriteLine("Tecle enter para sair......")
                                '' Console.Read()


                                fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                                mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                                mysw.WriteLine("[" & sUserName & "]" & Now + " - AppServer - Erro Exception durante WorkFlow  : " + ex.Message)
                                mysw.Close()

                                Exit Sub

                            End Try

Proximo:
                            Console.WriteLine("")
                            linhaTexto = sr.ReadLine
                        End If
                    Loop

                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                Else
                    Try
                        SiebelApplication = CreateObject("SiebelDataServer.ApplicationObject")



                    Catch ex As Exception

                        Console.WriteLine("DataServer - Erro Exception Instanciando SiebelDataServer : " + ex.Message)


                        fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                        mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                        mysw.WriteLine("[" & sUserName & "]" & Now + " - DataServer - Erro Exception Instanciando SiebelDataServer : " + ex.Message)
                        mysw.Close()

                        Exit Sub

                    End Try

                    If Not SiebelApplication Is Nothing Then

                        SiebelApplication.LoadObjects(sCFG, errCode)

                        If errCode <> 0 Then


                            ErrSiebel = Trim$(SiebelApplication.GetLastErrText.Replace(Chr(10), ""))

                            Console.WriteLine("DataServer - ErroSiebel LoadObjects CFG : " & errCode)



                            SiebelApplication = Nothing

                            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                            mysw.WriteLine("[" & sUserName & "] " & Now + " - DataServer - ErroSiebel LoadObjects CFG : " & CStr(errCode))
                            mysw.Close()

                            Exit Sub

                        End If


                        SiebelApplication.Login(sUserName, sPassword, errCode)

                        If errCode <> 0 Then

                            ErrSiebel = Trim$(SiebelApplication.GetLastErrText.Replace(Chr(10), ""))

                            Console.WriteLine("DataServer - ErroSiebel Login  : " & errCode)
                            Console.WriteLine(ErrSiebel)
                            '' Console.WriteLine("Tecle enter para sair......")
                            '' Console.Read()



                            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                            mysw.WriteLine("[" & sUserName & "] " & Now + " - DataServer - ErroSiebel Login : " & CStr(errCode))
                            mysw.WriteLine("[" & sUserName & "] " & Now + " - ErroSiebel Descricao : " & ErrSiebel)
                            mysw.Close()

                            SiebelApplication = Nothing

                            Exit Sub

                        End If



                        Dim SVC = SiebelApplication.GetService("Workflow Process Manager", errCode)

                        If errCode <> 0 Then

                            ErrSiebel = Trim$(SiebelApplication.GetLastErrText.Replace(Chr(10), ""))

                            Console.WriteLine("DataServer ErroSiebel Workflow Process Manager : " & errCode)
                            '' Console.ReadLine()

                            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                            mysw.WriteLine("[" & sUserName & "] " & Now + " - DataServer - ErroSiebel Workflow Process Manager : " & CStr(errCode))
                            mysw.WriteLine("[" & sUserName & "] " & Now + " - ErroSiebel Descricao : " & ErrSiebel)
                            mysw.Close()

                            GoTo ProximoRegistro

                        End If

                        Dim inputs = SiebelApplication.NewPropertySet(errCode)
                        Dim outputs = SiebelApplication.NewPropertySet(errCode)

                        inputs.SetProperty("ProcessName", "DevOps_Importa_XML_Object", errCode)

                        Do While linhaTexto <> Nothing

                            If Len(linhaTexto) > 0 Then

                                If sRepositorio.ToUpper = "IN" Then
                                    RetornoValida = ValidaProjeto(Trim$(linhaTexto))
                                End If

                                If InStr(linhaTexto, "LOV_") <> 0 Then
                                    If DeletaLOV(linhaTexto) = False Then
                                        GoTo ProximoRegistro
                                    End If
                                End If


                                If InStr(linhaTexto, "RGN_") <> 0 Then
                                    If DeletaRGN(linhaTexto) = False Then
                                        GoTo ProximoRegistro
                                    End If
                                End If

                                If InStr(linhaTexto, "CST_") <> 0 Then
                                    If DeletaConsultaRGN(linhaTexto) = False Then
                                        GoTo ProximoRegistro
                                    End If
                                End If

                                If InStr(linhaTexto, "ACS_") <> 0 Then
                                    If DeletaAcoesRGN(linhaTexto) = False Then
                                        GoTo ProximoRegistro
                                    End If
                                End If

                                '@@@ Deletando Mapa de Valores EAI @@@
                                If InStr(linhaTexto, "EVL_") <> 0 Then
                                    If DeletaEVL(linhaTexto) = False Then
                                        GoTo Proximo
                                    End If
                                End If

                                If (RetornoValida = "Ok" Or sRepositorio.ToUpper = "OUT") Then

                                    If (InStr(linhaTexto, "PAR_") <> 0 Or InStr(linhaTexto, "SPR_") <> 0 Or InStr(linhaTexto, "TRD_") <> 0 Or InStr(linhaTexto, "LOV_") <> 0) Then '' Or InStr(linhaTexto, "EDM_") <> 0
                                        inputs.SetProperty("Repositorio", "out", errCode)
                                        Console.WriteLine("InsertUpdate")
                                    Else
                                        inputs.SetProperty("Repositorio", "in", errCode)
                                        Console.WriteLine("Overwrite")
                                    End If

                                    Try


                                        If InStr(linhaTexto, "WKF_") <> 0 Then

                                            RetornoValidacaoObjeto = ValidaObjetoDestino(linhaTexto)

                                            If RetornoValidacaoObjeto = "Yes" Then
                                                DeletaWF(NomeObjetoDestino)
                                            ElseIf RetornoValidacaoObjeto = "Erro" Then
                                                GoTo ProximoRegistro
                                            End If

                                        End If


                                        inputs.SetProperty("XML Name", Trim$(linhaTexto), errCode)


                                        ''    SiebelApplication.TraceOn("c:\Temp\TraceMattar.txt", "sql", "sql", errCode)
                                        '' Console.WriteLine("Vou executar" & errCode)
                                        ''Console.Read()
                                        SVC.InvokeMethod("RunProcess", inputs, outputs, errCode)
                                        '' Console.WriteLine("Voltei" & errCode)

                                        ''     SiebelApplication.TraceOff(errCode)

                                        '' Console.WriteLine("errCode " & errCode)
                                        '' Console.Read()



                                        If errCode <> 0 Then

                                            ErrSiebel = Trim$(SiebelApplication.GetLastErrText.Replace(Chr(10), ""))

                                            If InStr(ErrSiebel, "field values are unique") <> 0 Then

                                                DeletaRepositorio(Trim$(linhaTexto))
                                                inputs.SetProperty("XML Name", Trim$(linhaTexto), errCode)
                                                SVC.InvokeMethod("RunProcess", inputs, outputs, errCode)

                                                If errCode = 0 Then
                                                    ErrSiebel = ""
                                                    Console.WriteLine("XML : " & Trim$(linhaTexto))
                                                    Console.WriteLine("XML Importado com sucesso!")
                                                    GoTo ProximoRegistro
                                                End If

                                            End If

                                        End If


                                    Catch ex As Exception

                                        fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                                        mysw = New StreamWriter(fs, System.Text.Encoding.UTF8)
                                        mysw.WriteLine("[" & Now & "]" & "XML : " & Trim$(linhaTexto))
                                        mysw.WriteLine("[" & sUserName & "] " & Now + " - ErroSiebel Descricao : " & ex.Message)
                                        mysw.Close()

                                        Console.WriteLine("XML : " & Trim$(linhaTexto))
                                        Console.WriteLine("Erro Exception: " & ex.Message)


                                        If sRepositorio.ToUpper = "IN" Then

                                            NomeArquivoLogOut = "Erro_Arquivos_in.txt"

                                            fsOut = New FileStream(sDiretorioErro + NomeArquivoLogOut, FileMode.Append)
                                            myswOut = New StreamWriter(fsOut, System.Text.Encoding.UTF8)
                                            myswOut.WriteLine("[" & Now & "]" & "XML : " & Trim$(linhaTexto))
                                            myswOut.WriteLine("ErroSiebel Exception : " & ex.Message)
                                            myswOut.WriteLine("")
                                            myswOut.Close()

                                        Else

                                            NomeArquivoLogOut = "Erro_Arquivos_out.txt"

                                            fsOut = New FileStream(sDiretorioErro + NomeArquivoLogOut, FileMode.Append)
                                            myswOut = New StreamWriter(fsOut, System.Text.Encoding.UTF8)
                                            myswOut.WriteLine("[" & Now & "] " & "XML : " & Trim$(linhaTexto))
                                            myswOut.WriteLine("ErroSiebel Exception: " & ex.Message)
                                            myswOut.WriteLine("")
                                            myswOut.Close()

                                        End If

                                        GoTo ProximoRegistro
                                    End Try

                                    '' DeletaObjeto(Trim$(linhaTexto))

                                    If errCode <> 0 Then

                                        ErrSiebel = Trim$(SiebelApplication.GetLastErrText.Replace(Chr(10), ""))



                                        Console.WriteLine("DataServer ErroSiebel Executando WorkFlow : " & errCode)
                                        Console.WriteLine("ErroSiebel Descricao : " & ErrSiebel)

                                        fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                                        mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                                        mysw.WriteLine("[" & sUserName & "] " & Now + " - DataServer - ErroSiebel Executando WorkFlow : " & CStr(errCode))
                                        mysw.WriteLine("[" & sUserName & "] " & Now + " - ErroSiebel Descricao : " & ErrSiebel)
                                        mysw.Close()


                                        If sRepositorio.ToUpper = "IN" Then

                                            NomeArquivoLogOut = "Erro_Arquivos_in.txt"

                                            fsOut = New FileStream(sDiretorioErro + NomeArquivoLogOut, FileMode.Append)
                                            myswOut = New StreamWriter(fsOut, System.Text.Encoding.Default)
                                            myswOut.WriteLine("[" & Now & "]" & "XML : " & Trim$(linhaTexto))
                                            myswOut.WriteLine("ErroSiebel Executando WorkFlow : " & ErrSiebel)
                                            myswOut.WriteLine("")
                                            myswOut.Close()

                                        Else

                                            NomeArquivoLogOut = "Erro_Arquivos_out.txt"

                                            fsOut = New FileStream(sDiretorioErro + NomeArquivoLogOut, FileMode.Append)
                                            myswOut = New StreamWriter(fsOut, System.Text.Encoding.Default)
                                            myswOut.WriteLine("[" & Now & "] " & "XML : " & Trim$(linhaTexto))
                                            myswOut.WriteLine("ErroSiebel Executando WorkFlow : " & ErrSiebel)
                                            myswOut.WriteLine("")
                                            myswOut.Close()

                                        End If

                                        GoTo ProximoRegistro

                                    Else

                                        Console.WriteLine("XML : " & Trim$(linhaTexto))
                                        Console.WriteLine("XML Importado com sucesso!")
                                        '' Console.Read()

                                    End If

                                ElseIf RetornoValida = "NOk" Then

                                    Console.WriteLine("XML : " & Trim$(linhaTexto))
                                    Console.WriteLine("Erro: Retorno ValidaProjeto")

                                    GoTo ProximoRegistro
                                Else
                                    Console.WriteLine(" Erro no XML - retorno ValidaProjeto : " & Trim$(linhaTexto))
                                End If
                            End If


ProximoRegistro:
                    Console.WriteLine("")
                    linhaTexto = sr.ReadLine
                        Loop

                        SiebelApplication = Nothing

                    End If

                    End If

            End Using


            Exit Sub

Erro:


        Catch ex As Exception

            Console.WriteLine(Trim$(linhaTexto))
            Console.WriteLine("Erro exception na importa objetos: " + ex.Message)

            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
            mysw.WriteLine("[" & sUserName & "] " & Now + "Erro exception na importa objetos: " + ex.Message)
            mysw.Close()

        End Try

        Exit Sub

    End Sub

    Sub Main(ByVal args() As String)


        '' Versão 4.6 - Alterado para usar SiebelApp
        '' Versão 4.6.2 - Alterado para Apagar WF
        '' Versão 4.6.3 - Alterado para Apagar Regra de negócio
        '' Versão 4.6.4 - Alterado para Apagar Obj. Repositorio Applet - Unique
        '' Versão 4.6.6 - Alterado para Apagar Consulta de Regra de negócio
        '' Versão 4.6.7 - Alterado para Apagar Aoes de Regra de negócio
        ''Versão 4.6.8 - Alterado para Apagar Condicoes de Regra de negócio - BUG
        ''Versão 4.7 -  Alterado para Apagar Acoes de Regra de negócio - BUG
        ''Versão 4.7.2 -  Alterado para Apagar VIEW Unique
        ''Versão 4.7.3 -  Alterado para Apagar BC e BS Unique
        ''Versão 4.7.4 -  Alterado para colocar EDM como overwrite
        ''Versão 4.7.5 -  Alterado para colocar Business Service Client

        ''Versão 4.7.7 -  Alterado para Apagar Regras de Negócio - Ordem para deletes
        ''Versão 4.7.8 - Incluída versão para Apagar Mapas de Valores EAI

        ''Versão 4.7.9 - Ajuste no processo de delete de Propety da Ação de Regras de Negócio;
        ''               Ajuste na exibição da versão do programa;
        ''               Ajuste no processo de delete de RNG;
        ''               Ajuste no processo de delete de EVL;


        Dim sDiretorioArquivoBatch As String '' Arquivo a ser processado

        'Console.WriteLine("ImportacaoSiebelDevops Versão 4.7.8")

        With My.Application.Info.Version
            Console.WriteLine("ImportacaoSiebelDevops Versão " & .Major & "." & .Minor & "." & .Build)
        End With

        Try
            ''sCFG = "c:\sea630\client\bin\scomm_B10.cfg"
            'sCFG = "c:\sea630\client\bin\scomm_local.cfg"
            'sUserName = "E_CARVALHO"
            'sPassword = "E_CARVALHO"
            'sDiretorioArquivoBatch = "C:\Importacao\migra.txt"
            'sRepositorio = "OUT"

            sCFG = args(0)
            sUserName = args(1)
            sPassword = Trim$(args(2))
            sDiretorioArquivoBatch = Trim$(args(3))
            sRepositorio = Trim$(args(4))
            'Trace = Trim$(args(5))

            Console.WriteLine("CFG :  " + sCFG)
            Console.WriteLine("UserName : " + sUserName)
            Console.WriteLine("ArquivoBatch : " + sDiretorioArquivoBatch)
            Console.WriteLine("Repositorio : " + sRepositorio)


            sDiretorioErro = System.AppDomain.CurrentDomain.BaseDirectory()

            If (Len(sCFG) = 0 Or Len(sUserName) = 0 Or Len(sPassword) = 0 Or Len(sDiretorioArquivoBatch) = 0) Then
                Console.WriteLine("Parametros incorretos")

                fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                mysw.WriteLine("[" & sUserName & "] " & Now + " Parametros incorretos")
                mysw.Close()

                Exit Sub

            End If


        Catch ex As Exception

            Console.WriteLine("Exception Parâmetro de entada :  " + Err.Description)

            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
            mysw.WriteLine("[" & sUserName & "] " & Now + " Exception Parâmetro de entada :  " + Err.Description)
            mysw.Close()

            Exit Sub

        End Try

        FileClose()

        If InStr(sCFG.ToString.ToUpper, ".CFG") = 0 Then
            Console.WriteLine("Conexão = AppSever")
            Conexao = "AppSever"
        Else
            Console.WriteLine("Conexão = DataServer")
            Conexao = "DataServer"
        End If

        sDiretorioArquivoBatch = Replace(sDiretorioArquivoBatch, "@", " ")
        sCFG = Replace(sCFG, "@", " ")

        ImportarObjetos(sDiretorioArquivoBatch)


    End Sub

    Private Function sDiretoSystem() As Object
        Throw New NotImplementedException
    End Function

End Module
