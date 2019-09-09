Imports System.IO
Imports System.IO.Directory
Imports System.Security
Imports System.Text

Module ApagarObjetos

    Public Sub DeletaWF(ByVal sObjeto As String)

        Dim SiebelDeleta As Object

        Dim BOObject As SiebelDataServer.SiebelBusObject
        Dim BCObject As SiebelDataServer.SiebelBusComp

        Dim BOAppli As SiebelApplicationServer.SiebelBusObject
        Dim BCAppli As SiebelApplicationServer.SiebelBusComp

        If Conexao = "DataServer" Then

            BOObject = SiebelApplication.GetBusObject("Front Office Workflow", errCode)
            BCObject = BOObject.GetBusComp("Workflow Process Definition", errCode)


            With BCObject
                .ClearToQuery(errCode)
                .SetViewMode(1, errCode)
                .ActivateField("Process Name", errCode)
                .SetSearchSpec("Process Name", "'" & sObjeto & "'", errCode)
                .ExecuteQuery(True, errCode)

                If .FirstRecord(errCode) <> 0 Then
                    .DeleteRecord(errCode)
                    Console.WriteLine("Apagou WF ")
                End If

            End With

            BCObject = Nothing
            BCObject = Nothing

        Else

            SiebelDeleta = GetObject("", "SiebelAppServer.ApplicationObject")

            BOAppli = SiebelDeleta.GetBusObject("Front Office Workflow", errCode)
            BCAppli = BOAppli.GetBusComp("Workflow Process Definition", errCode)


            With BCAppli
                .ClearToQuery(errCode)
                .SetViewMode(1, errCode)
                .ActivateField("Process Name", errCode)
                .SetSearchSpec("Process Name", "'" & sObjeto & "'", errCode)
                .ExecuteQuery(True, errCode)

                If .FirstRecord(errCode) <> 0 Then
                    .DeleteRecord(errCode)
                    Console.WriteLine("Apagou WF ")
                End If

            End With

            BOAppli = Nothing
            BCAppli = Nothing

        End If


    End Sub

    Public Function DeletaRGN(ByVal sArq As String) As Boolean


        Dim RetList As Integer = 0
        Dim IsRecord As Integer
        Dim IsRecord1 As Integer
        Dim IsRecord2 As Integer
        Dim IsRecord3 As Integer
        Dim IsRecord4 As Integer
        Dim IdRegra As String
        Dim IdAcRegra As String
        Dim SiebelRGN As Object
        Dim Linha As String
        Dim Tamanho As Integer

        Dim oBORGN As SiebelDataServer.SiebelBusObject
        Dim oBCRGN As SiebelDataServer.SiebelBusComp
        Dim oBCRGNCond As SiebelDataServer.SiebelBusComp
        Dim oBCRGNEvento As SiebelDataServer.SiebelBusComp

        Dim oBCRGNAcaoRegra As SiebelDataServer.SiebelBusComp
        Dim oBCRGNInputsAcao As SiebelDataServer.SiebelBusComp

        Dim oBOAPPRGN As SiebelApplicationServer.SiebelBusObject
        Dim oBCAPPRGN As SiebelApplicationServer.SiebelBusComp


        DeletaRGN = True



        Dim texto As New StreamReader(sArq, System.Text.Encoding.Default)
        While Not texto.EndOfStream
            Linha = texto.ReadLine

            If InStr(1, Linha, "<PCSCodigo>") > 0 Then
                IdRegra = Linha.Substring(InStr(1, Linha, "<PCSCodigo>") + 10)
                Tamanho = InStr(1, IdRegra, "</PCSCodigo>")
                IdRegra = Mid(IdRegra, 1, Tamanho - 1)
                Exit While
            End If

        End While

        'Console.WriteLine(IdRegra)
        'Console.Read()


        Try

            If Conexao = "DataServer" Then


                oBORGN = SiebelApplication.GetBusObject("PCS RN - Regra", errCode)
                oBCRGN = oBORGN.GetBusComp("PCS RN - Regra", errCode)
                oBCRGNCond = oBORGN.GetBusComp("PCS RN - Regra Condicao", errCode)
                oBCRGNEvento = oBORGN.GetBusComp("PCS RN - Regra Evento", errCode)
                oBCRGNAcaoRegra = oBORGN.GetBusComp("PCS RN - Instancias de Acao em Regras", errCode)
                oBCRGNInputsAcao = oBORGN.GetBusComp("PCS RN - Inputs de Instancia de Acao", errCode)


                With oBCRGN
                    .ClearToQuery(errCode)
                    .SetViewMode(1, errCode)
                    oBCRGNCond.ActivateField("PCS Codigo", errCode)
                    .ActivateField("PCS Codigo", errCode)
                    .ActivateField("PCS Locked Flg", errCode)
                    .SetSearchSpec("PCS Codigo", IdRegra, errCode)
                    .ExecuteQuery(True, errCode)
                    IsRecord = .FirstRecord(errCode)

                    'Console.WriteLine(IsRecord)
                    'Console.Read()

                    If IsRecord <> 0 Then
                        While IsRecord <> 0
                            .SetFieldValue("PCS Locked Flg", "Y", errCode)
                            .WriteRecord(errCode)

                            With oBCRGNCond

                                .ClearToQuery(errCode)
                                .SetViewMode(1, errCode)
                                .SetSearchSpec("PCS Codigo Pai", IdRegra, errCode)
                                .ExecuteQuery(True, errCode)
                                IsRecord1 = .FirstRecord(errCode)

                                If IsRecord1 <> 0 Then
                                    While IsRecord1 <> 0

                                        With oBCRGNEvento

                                            .ClearToQuery(errCode)
                                            .SetViewMode(1, errCode)

                                            .SetSearchSpec("PCS Codigo Pai", IdRegra, errCode)
                                            .SetSearchSpec("PCS Codigo Pai", IdRegra, errCode)
                                            .ExecuteQuery(True, errCode)
                                            IsRecord2 = .FirstRecord(errCode)

                                            If IsRecord2 <> 0 Then
                                                While IsRecord2 <> 0
                                                    .DeleteRecord(errCode)
                                                    IsRecord2 = .NextRecord(errCode)
                                                End While
                                            End If

                                        End With

                                        .DeleteRecord(errCode)
                                        IsRecord1 = .NextRecord(errCode)
                                    End While

                                End If

                               
                            End With

                            With oBCRGNAcaoRegra

                                .ClearToQuery(errCode)
                                .SetViewMode(1, errCode)
                                .SetSearchSpec("PCS Regra Id", IdRegra, errCode)
                                .ExecuteQuery(True, errCode)
                                IsRecord3 = .FirstRecord(errCode)

                                If IsRecord <> 0 Then

                                    While IsRecord3 <> 0

                                        IdAcRegra = .GetFieldValue("PCS Codigo", errCode)

                                        With oBCRGNInputsAcao
                                            .ClearToQuery(errCode)
                                            .SetViewMode(1, errCode)
                                            .SetSearchSpec("PCS Instancia Acao Id", IdAcRegra, errCode)
                                            .ExecuteQuery(True, errCode)
                                            IsRecord4 = .FirstRecord(errCode)

                                            If IsRecord4 <> 0 Then
                                                While IsRecord4 <> 0
                                                    .DeleteRecord(errCode)
                                                    IsRecord4 = .NextRecord(errCode)
                                                End While
                                            End If
                                        End With

                                        .DeleteRecord(errCode)
                                        IsRecord3 = .NextRecord(errCode)

                                    End While
                                End If

                                '@@@ movemos para cá
                                .DeleteRecord(errCode)
                                Console.WriteLine("Apagou RGN")

                            End With

                            '@@@ retiramos daqui
                            '.DeleteRecord(errCode)
                            'Console.WriteLine("Apagou RGN")
                            IsRecord = .NextRecord(errCode)
                        End While
                    End If

                End With

                oBCRGN = Nothing
                oBCRGNEvento = Nothing
                oBCRGNCond = Nothing
                oBORGN = Nothing
                oBCRGNAcaoRegra = Nothing
                oBCRGNInputsAcao = Nothing


            Else  ''''''''' NÃO ESTÁ SENDO MAIS USADO '''''''

                SiebelRGN = GetObject("", "SiebelAppServer.ApplicationObject")

                oBOAPPRGN = SiebelRGN.GetBusObject("PCS RN - Regra", errCode)
                oBCAPPRGN = oBOAPPRGN.GetBusComp("PCS RN - Regra", errCode)

                With oBCAPPRGN
                    .ClearToQuery(errCode)
                    .SetViewMode(1, errCode)
                    .ActivateField("PCS Codigo", errCode)
                    .ActivateField("PCS Locked Flg", errCode)
                    .SetSearchSpec("PCS Codigo", IdRegra, errCode)
                    .ExecuteQuery(False, errCode)
                    IsRecord = .FirstRecord(errCode)

                    If IsRecord <> 0 Then
                        While IsRecord <> 0
                            .SetFieldValue("PCS Locked Flg", "Y", errCode)
                            .WriteRecord(errCode)
                            .DeleteRecord(errCode)
                            Console.WriteLine("Apagou RGN")
                            IsRecord = .NextRecord(errCode)
                        End While
                    End If

                End With

                oBCAPPRGN = Nothing
                oBOAPPRGN = Nothing

            End If

        Catch ex As Exception

            If InStr(ex.Message, "ActiveX") > 0 Then

                Console.WriteLine("Erro de conexão com Siebel Client: " + Err.Description)
                Console.WriteLine("Favor abrir o Siebel Client e repetir a operação !")


                fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro Exception de conexão com Siebel Client: " + Err.Description)
                mysw.Close()

                DeletaRGN = False

                Exit Function

            Else

                fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro Exception DeletaRGN Descricao  : " + ex.Message)
                mysw.Close()

                Console.WriteLine(" Erro Exception DeletaRGN Descricao  : " + ex.Message)

                DeletaRGN = False

                Exit Function

            End If

        End Try


    End Function

    Public Function DeletaLOV(ByVal sArq As String) As Boolean

        Dim RetList As Integer = 0
        Dim sLOV As String
        Dim sLista As String
        Dim SiebelApp As Object
        Dim IsRecord As Integer

        Dim oBOLOV As SiebelApplicationServer.SiebelBusObject
        Dim oBCLOV As SiebelApplicationServer.SiebelBusComp

        Dim oBCDev As SiebelDataServer.SiebelBusComp
        Dim oBODev As SiebelDataServer.SiebelBusObject


        DeletaLOV = True


        Try

            sLista = sArq
            RetList = InStr(sLista, "LOV_")
            sLista = Mid(sLista, RetList + 4)
            sLOV = sLista.Substring(0, InStr(sLista.ToUpper, ".XML") - 1)

            If Conexao = "DataServer" Then


                oBODev = SiebelApplication.GetBusObject("PCS List Of Values IO", errCode)
                oBCDev = oBODev.GetBusComp("PCS List Of Values IO", errCode)



                With oBCDev
                    .ClearToQuery(errCode)
                    .SetViewMode(1, errCode)
                    .SetSearchSpec("Type", sLOV, errCode)
                    .ExecuteQuery(False, errCode)
                    IsRecord = .FirstRecord(errCode)

                    If IsRecord <> 0 Then
                        While IsRecord <> 0
                            .DeleteRecord(errCode)
                            IsRecord = .NextRecord(errCode)
                        End While
                    End If

                End With

                oBCDev = Nothing
                oBODev = Nothing

            Else

                Try


                    SiebelApp = GetObject("", "SiebelAppServer.ApplicationObject")

                    oBOLOV = SiebelApp.GetBusObject("PCS List Of Values IO", errCode)
                    oBCLOV = oBOLOV.GetBusComp("PCS List Of Values IO", errCode)



                    With oBCLOV
                        .ClearToQuery(errCode)
                        .SetViewMode(1, errCode)
                        .SetSearchSpec("Type", sLOV, errCode)
                        .ExecuteQuery(False, errCode)
                        IsRecord = .FirstRecord(errCode)


                        '' If Trace.Length > 0 Then
                        ''Siebel.TraceOn(Trace, "SQL", "TESTE", errCode)
                        '' End If

                        If IsRecord <> 0 Then
                            While IsRecord <> 0
                                Console.WriteLine("Apaguei M ")
                                .DeleteRecord(errCode)
                                IsRecord = .NextRecord(errCode)
                            End While
                        End If

                    End With

                    ''If Trace.Length > 0 Then
                    ''Siebel.TraceOff(errCode)
                    '' End If


                    oBCLOV = Nothing
                    oBOLOV = Nothing

                Catch ex As Exception

                    If InStr(ex.Message, "ActiveX") > 0 Then

                        Console.WriteLine("AppServer - Erro de conexão com Siebel Client: " + Err.Description)
                        Console.WriteLine("Favor abrir o Siebel Client e repetir a operação !")

                        fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                        mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                        mysw.WriteLine("[" & sUserName & "] " & Now + " - AppServer - Erro Exception de conexão com Siebel Client: " + Err.Description)
                        mysw.Close()

                        DeletaLOV = False

                        Exit Function

                    Else

                        fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                        mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                        mysw.WriteLine("[" & sUserName & "] " & Now + " - AppServer - Erro Exception DeletaLOV Descricao  : " + ex.Message)
                        mysw.Close()

                        Console.WriteLine("AppServer - Erro Exception DeletaLOV Descricao  : " + ex.Message)

                        DeletaLOV = False

                    End If

                End Try

            End If


        Catch ex As Exception


            Console.WriteLine("XML  : " + sArq)
            Console.WriteLine("Erro Exception DeletaLOV  : " + ex.Message)


            fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
            mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro Exception DeletaLOV : " & sArq)
            mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro Exception DeletaLOV : " + Err.Description)
            mysw.Close()

            DeletaLOV = False

        End Try


    End Function

    Public Function DeletaAcoesRGN(ByVal sArq As String) As Boolean


        Dim RetList As Integer = 0
        Dim IsRecord As Integer
        Dim IsRecord1 As Integer
        Dim IsRecord2 As Integer
        Dim IsRecord3 As Integer


        Dim IdAcao As String
        Dim SiebelRGN As Object

        Dim Linha As String
        Dim Tamanho As Integer
        Dim ParamId As String

        Dim oBOARGN As SiebelDataServer.SiebelBusObject
        Dim oBCARGN As SiebelDataServer.SiebelBusComp
        Dim oBCOperacao As SiebelDataServer.SiebelBusComp
        Dim oBCVariavel As SiebelDataServer.SiebelBusComp


        Dim oBOAPPCRGN As SiebelApplicationServer.SiebelBusObject
        Dim oBCAPPCRGN As SiebelApplicationServer.SiebelBusComp

        DeletaAcoesRGN = True

        Dim texto As New StreamReader(sArq, System.Text.Encoding.Default)
        While Not texto.EndOfStream
            Linha = texto.ReadLine

            If InStr(1, Linha, "<PCSCodigo>") > 0 Then
                IdAcao = Linha.Substring(InStr(1, Linha, "<PCSCodigo>") + 10)
                Tamanho = InStr(1, IdAcao, "</PCSCodigo>")
                IdAcao = Mid(IdAcao, 1, Tamanho - 1)
                Exit While
            End If

        End While

        'Console.WriteLine(IdAcao)
        'Console.Read()

        Try

            If Conexao = "DataServer" Then


                oBOARGN = SiebelApplication.GetBusObject("DevOps FRep PCS RN - Acoes Migracao", errCode)
                oBCARGN = oBOARGN.GetBusComp("DevOps FRep PCS RN - Acoes Migracao", errCode)
                oBCOperacao = oBOARGN.GetBusComp("DevOps FRep PCS RN - Acao Migracao", errCode)
                oBCVariavel = oBOARGN.GetBusComp("DevOps FRep PCS RN - Propertys de Acao Migracao", errCode)

                With oBCARGN
                    .ClearToQuery(errCode)
                    .SetViewMode(1, errCode)
                    oBCOperacao.ActivateField("PCS Codigo", errCode)
                    .ActivateField("PCS Codigo", errCode)
                    .SetSearchSpec("PCS Codigo", IdAcao, errCode)
                    .ExecuteQuery(True, errCode)
                    IsRecord = .FirstRecord(errCode)

                    'Console.WriteLine(IsRecord)
                    'Console.Read()

                    If IsRecord <> 0 Then
                        While IsRecord <> 0
                            While oBCOperacao.FirstRecord(errCode)

                                ParamId = oBCOperacao.GetFieldValue("PCS Codigo", errCode)

                                'Console.WriteLine("ParamId = " & ParamId)
                                'Console.Read()

                                With oBCVariavel
                                    .ClearToQuery(errCode)
                                    .SetViewMode(1, errCode)
                                    .SetSearchSpec("PCS Action Id", ParamId, errCode)
                                    .ExecuteQuery(True, errCode)
                                    IsRecord2 = .FirstRecord(errCode)

                                    If IsRecord2 <> 0 Then
                                        'Console.WriteLine("existe")
                                        'Console.Read()
                                        While IsRecord2 <> 0
                                            oBCVariavel.DeleteRecord(errCode)
                                            IsRecord2 = .NextRecord(errCode)
                                            'Console.WriteLine("apagou")
                                        End While

                                    End If

                                End With

                                'Console.WriteLine("2")

                                oBCOperacao.DeleteRecord(errCode)

                            End While

                            .DeleteRecord(errCode)
                            Console.WriteLine("Apagou Acao RGN")
                            IsRecord = .NextRecord(errCode)
                        End While
                    End If
                End With

                oBCARGN = Nothing
                oBOARGN = Nothing

            Else

                SiebelRGN = GetObject("", "SiebelAppServer.ApplicationObject")

                oBOAPPCRGN = SiebelRGN.GetBusObject("DevOps FRep PCS RN - Acoes Migracao", errCode)
                oBCAPPCRGN = oBOAPPCRGN.GetBusComp("DevOps FRep PCS RN - Acoes Migracao", errCode)

                With oBCAPPCRGN
                    .ClearToQuery(errCode)
                    .SetViewMode(1, errCode)
                    .ActivateField("PCS Codigo", errCode)
                    .SetSearchSpec("PCS Codigo", IdAcao, errCode)
                    .ExecuteQuery(False, errCode)
                    IsRecord = .FirstRecord(errCode)

                    If IsRecord <> 0 Then
                        While IsRecord <> 0
                            .DeleteRecord(errCode)
                            Console.WriteLine("Apagou Acao RGN")
                            IsRecord = .NextRecord(errCode)
                        End While
                    End If

                End With

                oBCAPPCRGN = Nothing
                oBOAPPCRGN = Nothing

            End If

        Catch ex As Exception

            If InStr(ex.Message, "ActiveX") > 0 Then

                Console.WriteLine("Erro de conexão com Siebel Client: " + Err.Description)
                Console.WriteLine("Favor abrir o Siebel Client e repetir a operação !")

                fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro Exception de conexão com Siebel Client: " + Err.Description)
                mysw.Close()

                DeletaAcoesRGN = False
                Exit Function
            Else

                fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro Exception DeletaAcoesRGN Descricao  : " + ex.Message)
                mysw.Close()

                Console.WriteLine(" Erro Exception DeletaAcoesRGN Descricao  : " + ex.Message)

                DeletaAcoesRGN = False

                Exit Function

            End If

        End Try

    End Function

    Public Function DeletaConsultaRGN(ByVal sArq As String) As Boolean


        Dim RetList As Integer = 0

        Dim IsRecord As Integer
        Dim IsRecord2 As Integer
        Dim IdConsulta As String

        Dim SiebelRGN As Object

        Dim Linha As String
        Dim Tamanho As Integer

        Dim ParamId As String

        Dim oBOCRGN As SiebelDataServer.SiebelBusObject
        Dim oBCCRGN As SiebelDataServer.SiebelBusComp

        Dim oBCConfig As SiebelDataServer.SiebelBusComp
        Dim oBCCampos As SiebelDataServer.SiebelBusComp


        Dim oBOAPPCRGN As SiebelApplicationServer.SiebelBusObject
        Dim oBCAPPCRGN As SiebelApplicationServer.SiebelBusComp


        DeletaConsultaRGN = True


        Dim texto As New StreamReader(sArq, System.Text.Encoding.Default)
        While Not texto.EndOfStream
            Linha = texto.ReadLine

            If InStr(1, Linha, "<PCSCodigo>") > 0 Then
                IdConsulta = Linha.Substring(InStr(1, Linha, "<PCSCodigo>") + 10)
                Tamanho = InStr(1, IdConsulta, "</PCSCodigo>")
                IdConsulta = Mid(IdConsulta, 1, Tamanho - 1)
                Exit While
            End If

        End While

        Try

            If Conexao = "DataServer" Then

                oBOCRGN = SiebelApplication.GetBusObject("DevOps FRep PCS RN - Consulta Migracao", errCode)
                oBCCRGN = oBOCRGN.GetBusComp("DevOps FRep PCS RN - Consulta Migracao", errCode)
                oBCConfig = oBOCRGN.GetBusComp("DevOps FRep PCS RCL - Consulta Configuracao Migracao", errCode)
                oBCCampos = oBOCRGN.GetBusComp("DevOps FRep PCS RCL - Consulta Field Migracao", errCode)

                With oBCCRGN
                    .ClearToQuery(errCode)
                    .SetViewMode(1, errCode)
                    .ActivateField("PCS Codigo", errCode)
                    oBCConfig.ActivateField("PCS Codigo", errCode)
                    .SetSearchSpec("PCS Codigo", IdConsulta, errCode)
                    .ExecuteQuery(False, errCode)
                    IsRecord = .FirstRecord(errCode)

                    If IsRecord <> 0 Then
                        While IsRecord <> 0

                            While oBCConfig.FirstRecord(errCode)

                                ParamId = oBCConfig.GetFieldValue("PCS Codigo", errCode)

                                With oBCCampos
                                    .ClearToQuery(errCode)
                                    .SetViewMode(1, errCode)
                                    .SetSearchSpec("PCS Codigo Pai", ParamId, errCode)
                                    .ExecuteQuery(True, errCode)
                                    IsRecord2 = .FirstRecord(errCode)

                                    If IsRecord2 <> 0 Then
                                        While .FirstRecord(errCode)
                                            .DeleteRecord(errCode)
                                            IsRecord2 = .NextRecord(errCode)
                                        End While
                                    End If


                                End With

                                oBCConfig.DeleteRecord(errCode)
                            End While

                            .DeleteRecord(errCode)
                            Console.WriteLine("Apagou Consulta RGN")
                            IsRecord = .NextRecord(errCode)
                        End While
                    End If


                End With

                oBCCRGN = Nothing
                oBOCRGN = Nothing
                oBCConfig = Nothing
                oBCCampos = Nothing


            Else

                SiebelRGN = GetObject("", "SiebelAppServer.ApplicationObject")

                oBOAPPCRGN = SiebelRGN.GetBusObject("DevOps FRep PCS RN - Consulta Migracao", errCode)
                oBCAPPCRGN = oBOAPPCRGN.GetBusComp("DevOps FRep PCS RN - Consulta Migracao", errCode)

                With oBCAPPCRGN
                    .ClearToQuery(errCode)
                    .SetViewMode(1, errCode)
                    .ActivateField("PCS Codigo", errCode)
                    .SetSearchSpec("PCS Codigo", IdConsulta, errCode)
                    .ExecuteQuery(False, errCode)
                    IsRecord = .FirstRecord(errCode)

                    If IsRecord <> 0 Then
                        While IsRecord <> 0
                            .DeleteRecord(errCode)
                            Console.WriteLine("Apagou Consulta RGN")
                            IsRecord = .NextRecord(errCode)
                        End While
                    End If

                End With

                oBCAPPCRGN = Nothing
                oBOAPPCRGN = Nothing

            End If

        Catch ex As Exception

            If InStr(ex.Message, "ActiveX") > 0 Then

                Console.WriteLine("Erro de conexão com Siebel Client: " + Err.Description)
                Console.WriteLine("Favor abrir o Siebel Client e repetir a operação !")


                fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro Exception de conexão com Siebel Client: " + Err.Description)
                mysw.Close()

                DeletaConsultaRGN = False

                Exit Function

            Else

                fs = New FileStream(sDiretorioErro + NomeArquivoLog, FileMode.Append)
                mysw = New StreamWriter(fs, System.Text.Encoding.Default)
                mysw.WriteLine("[" & sUserName & "] " & Now + " - Erro Exception DeletaConsultaRGN Descricao  : " + ex.Message)
                mysw.Close()

                Console.WriteLine(" Erro Exception DeletaRGN Descricao  : " + ex.Message)

                DeletaConsultaRGN = False

                Exit Function

            End If

        End Try
    End Function

End Module
