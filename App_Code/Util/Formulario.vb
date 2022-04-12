Imports System.Data
Imports System.Net
Imports System.IO
Imports CepFacil.API

Public Module Formulario

    Public Sub CarregarComboTabela(ByRef Controle As Object, ByRef objClasse As Object, Optional ByVal PrimeiroItem As String = "")
        With Controle
            .Items.Clear()

            .DataValueField = "CODIGO"
            .DataTextField = "DESCRICAO"

            .DataSource = objClasse.ObterTabela()
            .DataBind()

            If TypeOf Controle Is DropDownList Then
                .Items.Insert(0, New ListItem(PrimeiroItem, 0))
            End If
        End With

        'objClasse.Encerrar()
        objClasse = Nothing
    End Sub

    Public Sub CarregarComboTabelaRelacionada(ByRef Controle As Object, ByRef objClasse As Object, ByVal CodigoChave As Integer, Optional ByVal PrimeiroItem As String = "")

        With Controle
            .Items.Clear()

            .DataValueField = "CODIGO"
            .DataTextField = "DESCRICAO"

            .DataSource = objClasse.ObterTabela(CodigoChave)
            .DataBind()

            If TypeOf Controle Is DropDownList Then
                .Items.Insert(0, New ListItem(PrimeiroItem, 0))
            End If
        End With

    End Sub


    Public Sub CarregarComboTabelaSecundaria(ByRef Controle As Object, ByRef objClasse As Object, ByVal CodigoChave As Integer, ByVal CodigoChave2 As Integer, Optional ByVal PrimeiroItem As String = "")

        With Controle
            .Items.Clear()

            .DataValueField = "CODIGO"
            .DataTextField = "DESCRICAO"

            .DataSource = objClasse.ObterTabela(CodigoChave, CodigoChave2)
            .DataBind()

            If TypeOf Controle Is DropDownList Then
                .Items.Insert(0, New ListItem(PrimeiroItem, 0))
            End If
        End With

    End Sub


    Public Sub CarregarComboTabelaRelacionadaEstado(ByRef Controle As Object, ByRef objClasse As Object, ByVal CodigoChave As Integer, Optional ByVal PrimeiroItem As String = "")

        With Controle
            .Items.Clear()

            .DataValueField = "CODIGO"
            .DataTextField = "DESCRICAO"

            .DataSource = objClasse.ObterTabelaEstado(CodigoChave)
            .DataBind()

            If TypeOf Controle Is DropDownList Then
                .Items.Insert(0, New ListItem(PrimeiroItem, 0))
            End If
        End With

    End Sub

    'Public Sub CarregarComboTabelaRelacionada(ByRef Controle As Object, ByRef objClasse As Object, ByVal CodigoChave As String, Optional ByVal PrimeiroItem As String = "")

    '    With Controle
    '        .Items.Clear()

    '        .DataValueField = "CODIGO"
    '        .DataTextField = "DESCRICAO"

    '        .DataSource = objClasse.ObterTabela(CodigoChave)
    '        .DataBind()

    '        If TypeOf Controle Is DropDownList Then
    '            .Items.Insert(0, New ListItem(PrimeiroItem, 0))
    '        End If
    '    End With

    'End Sub


    'Public Sub CarregarComboTabelaRelacionada(ByRef Controle As Object, ByRef objClasse As Object, ByVal CodigoChave As Integer, ByVal CodigoChave2 As Integer, Optional ByVal PrimeiroItem As String = "")

    '    With Controle
    '        .Items.Clear()

    '        .DataValueField = "CODIGO"
    '        .DataTextField = "DESCRICAO"

    '        .DataSource = objClasse.ObterTabela(CodigoChave, CodigoChave2)
    '        .DataBind()

    '        If TypeOf Controle Is DropDownList Then
    '            .Items.Insert(0, New ListItem(PrimeiroItem, 0))
    '        End If
    '    End With

    'End Sub
    'Public Sub CarregarComboTabelaRelacionada(ByRef Controle As Object, ByRef objClasse As Object, ByVal CodigoChave As Integer, ByVal CodigoChave2 As Integer, ByVal CodigoChave3 As Integer, Optional ByVal PrimeiroItem As String = "")

    '    With Controle
    '        .Items.Clear()

    '        .DataValueField = "CODIGO"
    '        .DataTextField = "DESCRICAO"

    '        .DataSource = objClasse.ObterTabela(CodigoChave, CodigoChave2, CodigoChave3)
    '        .DataBind()

    '        If TypeOf Controle Is DropDownList Then
    '            .Items.Insert(0, New ListItem(PrimeiroItem, 0))
    '        End If
    '    End With

    'End Sub

    Public Sub CarregarComboSimNao(ByVal Controle As Object, Optional ByVal ValorPadrao As String = "-1", Optional ByVal ValorSim As String = "1", Optional ByVal ValorNao As String = "0")
        With Controle
            .Items.Add(New ListItem("SIM", ValorSim))
            .Items.Add(New ListItem("NÃO", ValorNao))

            If TypeOf Controle Is DropDownList Then
                .Items.Insert(0, New ListItem("", ValorPadrao))
            End If
        End With
    End Sub

    Public Sub SelecionarCombo(ByRef Controle As Object, ByVal Codigo As Object, Optional ByVal PesquisarPorTexto As Boolean = False)
        Controle.ClearSelection()
        If Not PesquisarPorTexto Then
            If Not Controle.Items.FindByValue(Codigo) Is Nothing Then
                Controle.Items.FindByValue(Codigo).Selected = True
            End If
        Else
            If Not Controle.Items.FindByText(Codigo) Is Nothing Then
                Controle.Items.FindByText(Codigo).Selected = True
            End If
        End If
    End Sub

    Public Sub SelecionarComboRelacionado(ByRef Controle As Object, ByRef ControlePai As Object, ByRef objClasse As Object, ByRef PropriedadeClassePai As String, ByRef Codigo As Integer, Optional ByVal PesquisarPorTexto As Boolean = False)
        objClasse.Obter(Codigo)

        SelecionarCombo(ControlePai, CallByName(objClasse, PropriedadeClassePai, CallType.Method), PesquisarPorTexto)
        CarregarComboTabelaRelacionada(Controle, objClasse, CallByName(objClasse, PropriedadeClassePai, CallType.Method))
        SelecionarCombo(Controle, Codigo, PesquisarPorTexto)

        'objClasse.Encerrar()
        objClasse = Nothing
    End Sub

    Public Function ObterRegistroRelacionado(ByVal objClasse As Object, ByVal objClassePai As Object, ByVal PropriedadeChave As String, ByVal PropriedadeRetorno As String, ByVal PropriedadeRetornoPai As String, ByVal Codigo As Integer) As String
        Dim strRetorno As String

        If Codigo <= 0 Then
            Return ""
        End If

        objClasse.Obter(Codigo)
        objClassePai.Obter(CallByName(objClasse, PropriedadeChave, CallType.Method))

        strRetorno = CallByName(objClasse, PropriedadeRetorno, CallType.Method) & " - " & CallByName(objClassePai, PropriedadeRetornoPai, CallType.Method)

        'objClasse.Encerrar()
        objClasse = Nothing

        objClassePai.Encerrar()
        objClassePai = Nothing

        Return strRetorno
    End Function

    Public Sub ObterCheckListBox(ByRef Controle As Object, ByRef objClasse As Object, ByRef Codigo As Integer)
        Controle.ClearSelection()

        For Each Item As ListItem In Controle.Items
            Item.Selected = objClasse.Obter(Codigo, Item.Value)
        Next

        'objClasse.Encerrar()
        objClasse = Nothing
    End Sub

    Public Sub SalvarCheckListBox(ByRef Controle As Object, ByRef objClasse As Object, ByRef Codigo As Integer)
        For Each Item As ListItem In Controle.Items
            If Item.Selected Then
                objClasse.Salvar(Codigo, Item.Value)
            Else
                objClasse.Excluir(Codigo, Item.Value)
            End If
        Next

        'objClasse.Encerrar()
        objClasse = Nothing

    End Sub

    Public Enum eTipoValor As Short
        CHAVE = 0
        DATA = 1
        TEXTO = 2
        NUMERO_INTEIRO = 3
        NUMERO_DECIMAL = 4
        DATA_COMPLETA = 5
        BOOLEANO = 6
        TEXTO_LIVRE = 7
        MONETARIO = 8
    End Enum

    Public Function ProBanco(ByRef Valor As Object, ByVal TipoValor As eTipoValor) As Object
        Dim Campo As Object = Nothing

        Select Case TipoValor
            Case eTipoValor.CHAVE
                If Valor > 0 Then
                    Campo = Valor
                Else
                    Campo = DBNull.Value
                End If
            Case eTipoValor.DATA, eTipoValor.DATA_COMPLETA
                If IsDate(Valor) Then
                    Campo = Convert.ToDateTime(Valor)
                Else
                    Campo = DBNull.Value
                End If
            Case eTipoValor.TEXTO
                If Valor <> "" Then
                    Campo = RemoverAcento(Valor.ToString.ToUpper.Trim.Replace("  ", " "))
                Else
                    Campo = ""
                End If
            Case eTipoValor.TEXTO_LIVRE
                If Valor <> "" Then
                    Campo = Valor
                Else
                    Campo = ""
                End If
            Case eTipoValor.NUMERO_INTEIRO
                If IsNumeric(Valor) Then
                    Campo = Convert.ToInt64(Valor)
                Else
                    Campo = DBNull.Value
                End If
            Case eTipoValor.NUMERO_DECIMAL
                If IsNumeric(Valor) Then
                    Campo = Convert.ToDouble(Valor)
                Else
                    Campo = DBNull.Value
                End If
            Case eTipoValor.MONETARIO
                If IsNumeric(Replace(Valor, ".", ",")) Then
                    Campo = Convert.ToDouble(Replace(Valor, ".", ","))
                Else
                    Campo = DBNull.Value
                End If
            Case eTipoValor.BOOLEANO
                If Valor = True Then
                    Campo = 1
                ElseIf Valor = False Then
                    Campo = 0
                End If

        End Select

        Return Campo
    End Function

    Public Function DoBanco(ByRef Campo As Object, ByVal TipoValor As eTipoValor) As Object
        Dim Valor As Object = Nothing

        Select Case TipoValor
            Case eTipoValor.CHAVE
                If Not IsDBNull(Campo) Then
                    Valor = Campo
                Else
                    Valor = 0
                End If
            Case eTipoValor.DATA
                If Not IsDBNull(Campo) Then
                    Valor = Convert.ToDateTime(Campo).ToString("dd/MM/yyyy")
                Else
                    Valor = ""
                End If
            Case eTipoValor.DATA_COMPLETA
                If Not IsDBNull(Campo) Then
                    Valor = Convert.ToDateTime(Campo).ToString("dd/MM/yyyy HH:mm:ss")
                Else
                    Valor = ""
                End If
            Case eTipoValor.TEXTO, eTipoValor.TEXTO_LIVRE
                If Not IsDBNull(Campo) Then
                    Valor = Campo
                Else
                    Valor = ""
                End If
            Case eTipoValor.NUMERO_INTEIRO
                If Not IsDBNull(Campo) Then
                    Valor = Campo
                Else
                    Valor = 0
                End If
            Case eTipoValor.NUMERO_DECIMAL
                If Not IsDBNull(Campo) Then
                    Valor = Convert.ToDouble(Campo).ToString("#0.00")
                Else
                    Valor = 0
                End If
            Case eTipoValor.MONETARIO
                If Not IsDBNull(Campo) Then
                    Valor = Convert.ToDouble(Campo).ToString("#0.00")
                Else
                    Valor = 0
                End If
            Case eTipoValor.BOOLEANO
                If Not IsDBNull(Campo) Then
                    Valor = Convert.ToBoolean(Campo)
                Else
                    Valor = False
                End If

        End Select

        Return Valor
    End Function

    Public Function RemoverAcento(ByVal Palavra As String, Optional ByVal ApenasAZ As Boolean = False) As String
        Dim Antes As String = "ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïñòóôõöùúûüýÿ"
        Dim Depois As String = "AAAAAACEEEEIIIINOOOOOUUUUYaaaaaaceeeeiiiinooooouuuuuu"
        Dim SemAcento As String = ""

        For x As Integer = 0 To Antes.Length - 1
            Palavra = Palavra.Replace(Antes.Substring(x, 1), Depois.Substring(x, 1))
        Next

        Palavra = LTrim(RTrim((Palavra)))

        Palavra = Replace(Palavra, "  ", " ")

        Palavra = UCase(Palavra)

        If ApenasAZ Then
            'O que não for de A a Z, remove
            For x As Integer = 0 To Palavra.Length - 1
                If (Asc(Palavra.Substring(x, 1)) >= 65 And Asc(Palavra.Substring(x, 1)) <= 90) Or Asc(Palavra.Substring(x, 1)) = 32 Then
                    SemAcento += Palavra.Substring(x, 1)
                End If
            Next
        Else
            SemAcento = Palavra
        End If

        Return SemAcento
    End Function

    Public Function ValidarTexto(ByVal Controle As TextBox, ByVal Descricao As String) As Boolean
        If RemoverAcento(Controle.Text) = "" Then
            MsgBox("Campo " & Descricao & " Obrigatório!")
            Return False
        Else
            Return True
        End If
    End Function

    Public Function ValidarData(ByVal Controle As TextBox, ByVal Descricao As String) As Boolean
        If Not ValidarTexto(Controle, Descricao) Then
            Return False
        End If

        If Not IsDate(Controle.Text.Trim) Then
            MsgBox("Campo " & Descricao & " Incorreto!")
            Return False
        Else
            Return True
        End If
    End Function

    Public Function ValidarNumero(ByVal Controle As TextBox, ByVal Descricao As String) As Boolean
        If Not ValidarTexto(Controle, Descricao) Then
            Return False
        End If

        If Not IsNumeric(Controle.Text.Trim) Then
            MsgBox("Campo " & Descricao & " Incorreto!")
            Return False
        Else
            Return True
        End If
    End Function

    Public Function ValidarCombo(ByVal Controle As DropDownList, ByVal Descricao As String) As Boolean
        If Controle.SelectedIndex <= 0 Then
            MsgBox("Campo " & Descricao & " Obrigatório!")
            Return False
        End If

        If Controle.SelectedValue = "0" Then
            MsgBox("Campo " & Descricao & " Obrigatório!")
            Return False
        Else
            Return True
        End If
    End Function

    Public Function BuscarCep(ByVal Cep As String) As Data.DataTable
        Dim ds As New DataSet()
        Dim dt As New DataTable

        'Fonte de varios webservices de CEP http://www.pinceladasdaweb.com.br/blog/2013/02/15/apis-para-consulta-de-cep/

        'Dim localidade As Localidade = CepFacilBusca.BuscarLocalidade("13214080", "SEU CÓDIGO DE FILIAÇÃO")
        'Dim Endereco As String = RequestDadosWeb("http://api.postmon.com.br/v1/cep/" & Replace(Replace(Cep, ".", ""), "-", "") & "?format=xml")
        Dim Endereco As String = RequestDadosWeb("http://cep.republicavirtual.com.br/web_cep.php?cep=" & Replace(Replace(Cep, ".", ""), "-", "") & "&formato=xml")
        'Dim Endereco As String = RequestDadosWeb("http://correios.w2info.com.br/Service.svc/Buscar?cep=" & Replace(Replace(Cep, ".", ""), "-", ""))
        'Dim Endereco As String = RequestDadosWeb("http://cep.osgestor.info/" & Replace(Replace(Cep, ".", ""), "-", "") & ".xml")

        'Endereco = Mid$(Endereco, 11)
        'Endereco = Left(Endereco, Endereco.Length - 11)

        'Endereco = "<enderecos>" + Endereco + "</enderecos>"

        ds.ReadXml(New StringReader(Endereco))

        dt = ds.Tables(0)

        Return dt
    End Function

    'Public Function BuscarCep(ByVal Cep As String) As Data.DataTable
    '    Dim dt As New Data.DataTable
    '    Dim dr As Data.DataRow

    '    Dim parametros As String = "cepEntrada=" + Cep.Replace("-", "").Trim() + "&tipoCep=&cepTemp=&metodo=buscarCep"

    '    Dim request As WebRequest = WebRequest.Create(Convert.ToString("http://m.correios.com.br/movel/buscaCepConfirma.do?") & parametros)
    '    Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)

    '    Dim stream As New StreamReader(response.GetResponseStream(), Encoding.GetEncoding("ISO-8859-1"))

    '    Dim dados As String = stream.ReadToEnd()
    '    Dim count As Integer = 0
    '    Dim ExpressaoRegular As String = "<span class=""respostadestaque"">(.*?)</span>"
    '    Dim endereco As MatchCollection = Regex.Matches(dados, ExpressaoRegular, RegexOptions.Singleline Or RegexOptions.IgnoreCase)

    '    dt.Columns.Add("logradouro")
    '    dt.Columns.Add("bairro")
    '    dt.Columns.Add("tipo_logradouro")
    '    dt.Columns.Add("cidade")
    '    dt.Columns.Add("uf")

    '    dr = dt.NewRow

    '    For Each resultado As Match In endereco
    '        count += 1

    '        Select Case count
    '            Case 1
    '                dr("tipo_logradouro") = ""
    '                dr("logradouro") = RemoverCaracteres(resultado.Groups(1).Value)

    '                Exit Select
    '            Case 2
    '                dr("bairro") = RemoverCaracteres(resultado.Groups(1).Value)

    '                Exit Select
    '            Case 3
    '                Try
    '                    dr("cidade") = RemoverCaracteres(resultado.Groups(1).Value.Trim().Split("/"c)(0))
    '                    dr("uf") = RemoverCaracteres(resultado.Groups(1).Value.Trim().Split("/"c)(1))
    '                Catch ex As Exception
    '                End Try

    '                Exit Select
    '        End Select
    '    Next

    '    dt.Rows.Add(dr)

    '    Return dt
    'End Function

    Private Function RemoverCaracteres(texto As String) As String
        Dim resultado As String = texto

        resultado = resultado.Replace(vbLf, "")
        resultado = resultado.Replace(vbCr, "")
        resultado = resultado.Replace(vbTab, "")
        resultado = resultado.Trim()

        Return resultado
    End Function

    Public Function BuscarCPFBCfone(ByVal CPF As String) As Data.DataSet
        Dim ds As New DataSet()
        Dim Pesquisa As String = RequestDadosWeb("http://www.grupobci.com.br/sistema/webservices/index.php?Usuario=16220&Senha=5d4417244a466cac3418b01be27474fc&Campo=CPF&Dado=" & Replace(Replace(CPF, ".", ""), "-", "") & "&Produto=64")

        ds.ReadXml(New StringReader(Pesquisa))

        Return ds

    End Function

    Public Function RequestDadosWeb(ByVal pstrURL As String) As String
        Dim oWebRequest As WebRequest
        Dim oWebResponse As WebResponse = Nothing
        Dim strBuffer As String = ""
        Dim objSR As StreamReader = Nothing
        'conecta com o website
        Try
            oWebRequest = HttpWebRequest.Create(pstrURL)
            oWebResponse = oWebRequest.GetResponse()
            'Le a resposta do web site e armazena em uma stream
            objSR = New StreamReader(oWebResponse.GetResponseStream, Encoding.GetEncoding("ISO-8859-1"))
            'objSR = New StreamReader(oWebResponse.GetResponseStream)

            strBuffer = objSR.ReadToEnd
        Catch ex As Exception
            Throw ex
        Finally
            objSR.Close()
            oWebResponse.Close()
        End Try
        Return strBuffer
    End Function

End Module
