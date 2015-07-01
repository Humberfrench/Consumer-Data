Imports Microsoft.VisualBasic.CompilerServices
Imports System.ComponentModel
Namespace Security
    ''' <summary>
    ''' EncriptData - Classe que executa a criptografia de dados de forma simples.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Class Encript

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Função que faz a decriptografia de números
        ''' </summary>
        ''' <param name="strTexto">Texto contendo os números criptografados</param>
        ''' <returns>Inteiro</returns>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function DecryptNumber(ByVal strTexto As String) As Integer

            Dim intNumber1 As Integer
            Dim intNumber2 As Integer
            Dim strSenha() As Integer = {23, 19, 27, 21, 20, 30, 26, 31, 25, 33}
            Dim strBuff As String = ""

            For intNumber1 = 1 To Len(strTexto)
                intNumber2 = Asc(Mid$(strTexto, intNumber1, 1))
                intNumber2 = intNumber2 - strSenha(intNumber1 Mod 10)
                strBuff = strBuff & Chr(intNumber2)
            Next intNumber1

            Dim offset As Int16 = strBuff.LastIndexOf("/")
            If offset = -1 Then
                Return -1
            Else
                Return CInt(Left(strBuff, offset))
            End If

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Funçao que faz a decriptografia de querystring.
        ''' </summary>
        ''' <param name="strTexto">Texto a ser Decriptografado</param>
        ''' <returns>Texto Original</returns>
        ''' <remarks>
        ''' Usado para retirar da querystring o valor.
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function DecryptQueryString(ByVal strTexto As String) As String

            Dim intNumber1 As Integer
            Dim intNumber2 As Integer
            Dim intOffSet As Integer
            Dim strSenha() As Integer = {9, -1, 4, -5, 7, -9, 5, -7, 2, -8, 3, -2, 6, -4, 1, -3, 8, -6}
            Dim strBuff As String = ""
            Dim strAuxNum As String = ""
            Dim blnConcatenar As Boolean = False

            For intNumber1 = 1 To strTexto.Length
                If intNumber1 > strTexto.Length Then Exit For
                If Mid$(strTexto, intNumber1, 1) = "A" Then
                    If Not blnConcatenar Then
                        intOffSet = intNumber1
                        strAuxNum = ""
                    Else
                        strTexto = strTexto.Substring(0, intOffSet) & strTexto.Substring(intNumber1)
                        intNumber1 = intOffSet
                        intNumber2 = CInt(strAuxNum) - 11 - strSenha(intNumber1 Mod 18)
                        strBuff = strBuff & Chr(intNumber2)
                    End If
                    blnConcatenar = Not blnConcatenar
                ElseIf Mid$(strTexto, intNumber1, 1) = "B" Then
                    If Not blnConcatenar Then
                        intOffSet = intNumber1
                        strAuxNum = ""
                    Else
                        strTexto = strTexto.Substring(0, intOffSet) & strTexto.Substring(intNumber1)
                        intNumber1 = intOffSet
                        intNumber2 = CInt(strAuxNum) + 11 - strSenha(intNumber1 Mod 18)
                        strBuff = strBuff & Chr(intNumber2)
                    End If
                    blnConcatenar = Not blnConcatenar
                ElseIf blnConcatenar Then
                    strAuxNum = strAuxNum & Mid$(strTexto, intNumber1, 1)
                Else
                    intNumber2 = Asc(Mid$(strTexto, intNumber1, 1))
                    intNumber2 = intNumber2 - strSenha(intNumber1 Mod 18)
                    strBuff = strBuff & Chr(intNumber2)
                End If
            Next intNumber1

            intOffSet = strBuff.LastIndexOf("/")
            If intOffSet = -1 Then
                Return ""
            Else
                Return strBuff.Substring(0, intOffSet)
            End If

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Decriptografia normal do texto
        ''' </summary>
        ''' <param name="strTexto">Valor criptografado</param>
        ''' <returns>String, original.</returns>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function Decrypt(ByVal strTexto As String) As String

            Dim intNumber1 As Integer
            Dim intNumber2 As Integer
            Dim strSenha As String = "jnhtazkwzkanwqhtislk"
            Dim strBuff As String = ""
            Dim intOffSet As Int16 = 0
            Dim intSkipCharCount As Int16 = 0

            strTexto = strTexto.Substring(3)

            If Len(strSenha) Then
                For intNumber1 = 1 To Len(strTexto)
                    intNumber2 = Asc(Mid$(strTexto, intNumber1, 1))
                    If intNumber2 = 33 Then
                        intOffSet = 40
                        intSkipCharCount += 1
                    ElseIf intNumber2 = 35 Then
                        intOffSet = 200
                        intSkipCharCount += 1
                    Else
                        intNumber2 = intNumber2 + intOffSet
                        intNumber2 = intNumber2 - Asc(Mid$(strSenha, ((intNumber1 - intSkipCharCount) Mod Len(strSenha)) + 1, 1))
                        strBuff = strBuff & Chr(intNumber2 And &HFF)
                        intOffSet = 0
                    End If
                Next intNumber1
            Else
                strBuff = strTexto
            End If

            Return strBuff
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Decriptografa String
        ''' </summary>
        ''' <param name="strTexto">Texto a ser decriptografado</param>
        ''' <returns>String, original</returns>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function DecryptString(ByVal strTexto As String) As String

            Dim intNumber1 As Integer
            Dim intNumber2 As Integer
            Dim strSenha() As Integer = {23, 19, 27, 21, 20, 30, 26, 31, 25, 33}
            Dim strBuff As String = ""

            For intNumber1 = 1 To Len(strTexto)
                intNumber2 = Asc(Mid$(strTexto, intNumber1, 1))
                intNumber2 = intNumber2 - strSenha(intNumber1 Mod 10)
                strBuff = strBuff & Chr(intNumber2)
            Next intNumber1

            Dim offset As Int16 = strBuff.LastIndexOf("/")
            If offset = -1 Then
                Return ""
            Else
                Return Left(strBuff, offset)
            End If

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Executa a criptografia de um número
        ''' </summary>
        ''' <param name="auxNumero">Número a ser criptografado</param>
        ''' <returns>Uma String criptografada equivalente ao original</returns>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function EncryptNumber(ByVal auxNumero As Integer) As String

            Dim intNumber1 As Integer
            Dim intNumber2 As Integer
            Dim strTexto As String
            Dim strSenha() As Integer = {23, 19, 27, 21, 20, 30, 26, 31, 25, 33}
            Dim strBuff As String = ""

            strTexto = auxNumero.ToString

            Randomize()
            strTexto &= "/"
            For intNumber1 = strTexto.Length To 7
                strTexto &= Int(9 * Rnd()).ToString
            Next

            For intNumber1 = 1 To Len(strTexto)
                intNumber2 = Asc(Mid$(strTexto, intNumber1, 1))
                intNumber2 = intNumber2 + strSenha(intNumber1 Mod 10)
                strBuff = strBuff & Chr(intNumber2)
            Next intNumber1

            Return strBuff

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Executa a criptografia de um querystring
        ''' </summary>
        ''' <returns>Uma String criptografada equivalente ao original</returns>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function EncryptQueryString(ByVal auxString As String) As String

            Dim intNumber1 As Integer
            Dim intNumber2 As Integer
            Dim strTexto As String
            Dim strSenha() As Integer = {9, -1, 4, -5, 7, -9, 5, -7, 2, -8, 3, -2, 6, -4, 1, -3, 8, -6}
            Dim strBuff As String = ""
            strTexto = auxString

            Randomize()
            strTexto &= "/"
            For intNumber1 = strTexto.Length To 9
                strTexto &= Int(9 * Rnd()).ToString
            Next

            For intNumber1 = 1 To strTexto.Length
                intNumber2 = Asc(Mid$(strTexto, intNumber1, 1))
                intNumber2 = intNumber2 + strSenha(intNumber1 Mod 18)
                If (intNumber2 < 48) Or (intNumber2 > 57 And intNumber2 < 67) Then  'Caracteres especiais
                    strBuff = strBuff & "A" & intNumber2 + 11 & "A"
                ElseIf (intNumber2 > 47 And intNumber2 < 58) Or (intNumber2 > 66 And intNumber2 < 91) Or (intNumber2 > 96 And intNumber2 < 123) Then '0-9 C-Z a-z
                    strBuff = strBuff & Chr(intNumber2)
                ElseIf (intNumber2 > 90 And intNumber2 < 97) Or (intNumber2 > 122) Then 'Caracteres especiais
                    strBuff = strBuff & "B" & intNumber2 - 11 & "B"
                End If
            Next intNumber1
            Return strBuff
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Executa a criptografia de um texto comum
        ''' </summary>
        ''' <returns>Uma String criptografada equivalente ao original</returns>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function Encrypt(ByVal strTexto As String, Optional ByVal caseSensitive As Boolean = False) As String

            Dim intNumber1 As Integer
            Dim intNumber2 As Integer
            Dim strSenha As String = "jnhtazkwzkanwqhtislk"
            Dim strBuff As String = ""
            'Este comando faz a senha ficar Case Insensitive
            If Not caseSensitive Then strTexto = strTexto.ToLower()

            If Len(strSenha) Then
                'Colocando este prefixo gera uma identificação de qual o 
                'tipo de encriptação que foi utilizada para esta palavra
                'strBuff = "!@!"
                strBuff = ""
                For intNumber1 = 1 To Len(strTexto)
                    intNumber2 = Asc(Mid$(strTexto, intNumber1, 1))
                    intNumber2 = intNumber2 + Asc(Mid$(strSenha, (intNumber1 Mod Len(strSenha)) + 1, 1))
                    If intNumber2 >= 127 And intNumber2 <= 160 Then
                        strBuff = strBuff & Chr(33) & Chr(intNumber2 - 40)
                    ElseIf intNumber2 >= 255 Then
                        strBuff = strBuff & Chr(35) & Chr(intNumber2 - 200)
                    Else
                        strBuff = strBuff & Chr(intNumber2 And &HFF)
                    End If
                Next intNumber1
            Else
                strBuff = strTexto
            End If

            Return strBuff

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Executa a criptografia de uma string
        ''' </summary>
        ''' <returns>Uma String criptografada equivalente ao original</returns>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Shared Function EncryptString(ByVal auxString As String) As String

            Dim intNumber1 As Integer
            Dim intNumber2 As Integer
            Dim strTexto As String
            Dim senha() As Integer = {23, 19, 27, 21, 20, 30, 26, 31, 25, 33}
            Dim strBuff As String = ""

            strTexto = auxString

            Randomize()
            strTexto &= "/"
            For intNumber1 = strTexto.Length To 9
                strTexto &= Int(9 * Rnd()).ToString
            Next

            For intNumber1 = 1 To Len(strTexto)
                intNumber2 = Asc(Mid$(strTexto, intNumber1, 1))
                intNumber2 = intNumber2 + senha(intNumber1 Mod 10)
                strBuff = strBuff & Chr(intNumber2)
            Next intNumber1

            Return strBuff

        End Function

    End Class
End Namespace
