<#
.SYNOPSIS
    Envia e-mails utilizando o Microsoft Outlook clássico via PowerShell.

.DESCRIPTION
    Esta função envia e-mails utilizando o Outlook clássico instalado
    no Windows através da interface COM.

    O arquivo é OPCIONAL.

    Comportamentos:

      1. Sem arquivo:
         - Exibe um alerta.
         - Ignora todas as etapas relacionadas ao arquivo.
         - Continua normalmente com a validação do Outlook.
         - Envia o e-mail sem arquivo.

      2. Com arquivo válido:
         - Sem -SendAsAttachment:
             O conteúdo do arquivo é colocado no corpo do e-mail.

         - Com -SendAsAttachment:
             O arquivo é enviado como anexo.

      3. Arquivo inexistente:
         - Exibe um alerta.
         - Ignora as etapas relacionadas ao arquivo.
         - Continua o processamento.
         - Envia o e-mail sem arquivo.

      4. Arquivo vazio:
         - Exibe um alerta.
         - Ignora as etapas relacionadas ao arquivo.
         - Continua o processamento.
         - Envia o e-mail sem arquivo.

      5. Arquivo incompatível com corpo de texto:
         - Exibe um alerta.
         - Ignora as etapas relacionadas ao arquivo.
         - Continua o processamento.
         - Envia o e-mail sem arquivo.

    Observação:
      - Arquivos incompatíveis ainda podem ser enviados como anexo,
        pois não precisam ser interpretados como texto.

.NOTES
    Requisitos:
      - Windows
      - Microsoft Outlook clássico
      - Conta configurada no Outlook
      - PowerShell com suporte à automação COM

    O novo Outlook para Windows não oferece a mesma interface COM
    utilizada por este script.

.EXAMPLES

    # ------------------------------------------------------
    # Exemplo 1 - E-mail sem arquivo
    # ------------------------------------------------------

    Send-OutlookEmail `
        -To "destinatario@empresa.com" `
        -Subject "Processamento concluído"


    # ------------------------------------------------------
    # Exemplo 2 - Arquivo TXT no corpo do e-mail
    # ------------------------------------------------------

    Send-OutlookEmail `
        -To "destinatario@empresa.com" `
        -Subject "Relatório diário" `
        -FilePath "C:\Relatorios\relatorio.txt"


    # ------------------------------------------------------
    # Exemplo 3 - Arquivo como anexo
    # ------------------------------------------------------

    Send-OutlookEmail `
        -To "destinatario@empresa.com" `
        -Subject "Relatório diário" `
        -FilePath "C:\Relatorios\relatorio.pdf" `
        -SendAsAttachment


    # ------------------------------------------------------
    # Exemplo 4 - Arquivo inexistente
    #
    # O e-mail será enviado sem arquivo.
    # ------------------------------------------------------

    Send-OutlookEmail `
        -To "destinatario@empresa.com" `
        -Subject "Aviso" `
        -FilePath "C:\Relatorios\arquivo-inexistente.txt"


    # ------------------------------------------------------
    # Exemplo 5 - Sem arquivo + SendAsAttachment
    #
    # O parâmetro será ignorado porque não existe arquivo.
    # O e-mail será enviado normalmente.
    # ------------------------------------------------------

    Send-OutlookEmail `
        -To "destinatario@empresa.com" `
        -Subject "Aviso" `
        -SendAsAttachment

#>


function Send-OutlookEmail {

    param (

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$To,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Subject,

        [Parameter(Mandatory = $false)]
        [string]$FilePath,

        [Parameter(Mandatory = $false)]
        [switch]$SendAsAttachment
    )


    # ==========================================================
    # VARIÁVEIS
    # ==========================================================

    $Outlook  = $null
    $Mail     = $null
    $Session  = $null
    $Accounts = $null
    $Account  = $null

    $ArquivoValido = $false
    $ArquivoParaCorpo = $false
    $ArquivoParaAnexo = $false


    # ==========================================================
    # FUNÇÃO DE MENSAGEM
    # ==========================================================

    function Write-Message {

        param (
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [ValidateSet(
                "INFO",
                "WARNING",
                "ERROR",
                "SUCCESS"
            )]
            [string]$Level = "INFO"
        )

        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        $Text = "[$Timestamp] [$Level] $Message"

        switch ($Level) {

            "INFO" {
                Write-Host $Text
            }

            "WARNING" {
                Write-Host $Text -ForegroundColor Yellow
            }

            "ERROR" {
                Write-Host $Text -ForegroundColor Red
            }

            "SUCCESS" {
                Write-Host $Text -ForegroundColor Green
            }
        }
    }


    # ==========================================================
    # EXTENSÕES ACEITAS PARA USO COMO CORPO
    #
    # Esses arquivos são tratados como texto.
    # ==========================================================

    $TextFileExtensions = @(
        ".txt",
        ".log",
        ".csv",
        ".json",
        ".xml",
        ".html",
        ".htm",
        ".md",
        ".ps1",
        ".sql",
        ".ini",
        ".cfg",
        ".conf"
    )


    # ==========================================================
    # INÍCIO
    # ==========================================================

    Write-Message "=========================================================="
    Write-Message "INÍCIO DO PROCESSAMENTO"
    Write-Message "=========================================================="

    Write-Message "Destinatário: $To"
    Write-Message "Assunto: $Subject"


    # ==========================================================
    # 1. VALIDAÇÃO DO ARQUIVO
    # ==========================================================

    if ([string]::IsNullOrWhiteSpace($FilePath)) {

        Write-Message `
            "Nenhum arquivo foi informado." `
            -Level WARNING

        Write-Message `
            "As etapas relacionadas ao arquivo serão ignoradas." `
            -Level INFO
    }
    else {

        Write-Message "Arquivo informado: $FilePath"


        # ------------------------------------------------------
        # Verifica existência
        # ------------------------------------------------------

        if (-not (Test-Path -Path $FilePath -PathType Leaf)) {

            Write-Message `
                "O arquivo não foi encontrado." `
                -Level WARNING

            Write-Message `
                "As etapas relacionadas ao arquivo serão ignoradas." `
                -Level WARNING
        }
        else {

            try {

                $FileInfo = Get-Item `
                    -Path $FilePath `
                    -ErrorAction Stop


                # --------------------------------------------------
                # Verifica se está vazio
                # --------------------------------------------------

                if ($FileInfo.Length -eq 0) {

                    Write-Message `
                        "O arquivo está vazio." `
                        -Level WARNING

                    Write-Message `
                        "As etapas relacionadas ao arquivo serão ignoradas." `
                        -Level WARNING
                }
                else {

                    $ArquivoValido = $true

                    Write-Message `
                        "Arquivo encontrado: $($FileInfo.Name)" `
                        -Level SUCCESS

                    Write-Message `
                        "Tamanho: $($FileInfo.Length) bytes"


                    # --------------------------------------------------
                    # Identifica extensão
                    # --------------------------------------------------

                    $Extension = $FileInfo.Extension.ToLowerInvariant()

                    Write-Message `
                        "Extensão detectada: $Extension"


                    # --------------------------------------------------
                    # Determina como o arquivo será utilizado
                    # --------------------------------------------------

                    if ($SendAsAttachment) {

                        # Qualquer arquivo válido pode ser anexo.

                        $ArquivoParaAnexo = $true

                        Write-Message `
                            "Modo: arquivo será enviado como ANEXO."
                    }
                    else {

                        # Para ser usado no corpo, precisa ser um arquivo
                        # reconhecidamente textual.

                        if ($TextFileExtensions -contains $Extension) {

                            $ArquivoParaCorpo = $true

                            Write-Message `
                                "Modo: conteúdo do arquivo será utilizado no CORPO do e-mail."
                        }
                        else {

                            $ArquivoValido = $false

                            Write-Message `
                                "O tipo de arquivo '$Extension' não é compatível com leitura como texto." `
                                -Level WARNING

                            Write-Message `
                                "As etapas relacionadas ao arquivo serão ignoradas." `
                                -Level WARNING
                        }
                    }
                }
            }
            catch {

                $ArquivoValido = $false

                Write-Message `
                    "Não foi possível acessar o arquivo: $($_.Exception.Message)" `
                    -Level WARNING

                Write-Message `
                    "As etapas relacionadas ao arquivo serão ignoradas." `
                    -Level WARNING
            }
        }
    }


    # ==========================================================
    # 2. RESUMO DO MODO DE ENVIO
    # ==========================================================

    if ($ArquivoParaAnexo) {

        Write-Message `
            "Modo de processamento: ANEXO."
    }
    elseif ($ArquivoParaCorpo) {

        Write-Message `
            "Modo de processamento: CONTEÚDO NO CORPO."
    }
    else {

        Write-Message `
            "Modo de processamento: SEM ARQUIVO." `
            -Level WARNING
    }


    try {

        # ======================================================
        # 3. LOCALIZAÇÃO DO OUTLOOK
        # ======================================================

        Write-Message `
            "Procurando o Microsoft Outlook."


        $OutlookPaths = @()


        # ------------------------------------------------------
        # Instalação Microsoft 365 / Click-to-Run
        # ------------------------------------------------------

        $OutlookPaths += @(
            "$env:ProgramFiles\Microsoft Office\root\Office16\OUTLOOK.EXE",
            "$env:ProgramFiles(x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
        )


        # ------------------------------------------------------
        # Instalação MSI tradicional
        # ------------------------------------------------------

        $OutlookPaths += @(
            "$env:ProgramFiles\Microsoft Office\Office16\OUTLOOK.EXE",
            "$env:ProgramFiles(x86)\Microsoft Office\Office16\OUTLOOK.EXE",

            "$env:ProgramFiles\Microsoft Office\Office15\OUTLOOK.EXE",
            "$env:ProgramFiles(x86)\Microsoft Office\Office15\OUTLOOK.EXE",

            "$env:ProgramFiles\Microsoft Office\Office14\OUTLOOK.EXE",
            "$env:ProgramFiles(x86)\Microsoft Office\Office14\OUTLOOK.EXE"
        )


        # ------------------------------------------------------
        # Verifica registro do Windows
        # ------------------------------------------------------

        $RegistryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
        )


        foreach ($RegistryPath in $RegistryPaths) {

            try {

                if (Test-Path $RegistryPath) {

                    $RegistryValue = Get-ItemProperty `
                        -Path $RegistryPath `
                        -ErrorAction SilentlyContinue

                    if ($RegistryValue.'(default)') {

                        $OutlookPaths += $RegistryValue.'(default)'
                    }
                }
            }
            catch {
                # Não interrompe o processamento.
            }
        }


        # ------------------------------------------------------
        # Remove caminhos duplicados
        # ------------------------------------------------------

        $OutlookPaths = $OutlookPaths |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_)
            } |
            Select-Object -Unique


        # ------------------------------------------------------
        # Procura o executável
        # ------------------------------------------------------

        $OutlookExe = $null

        foreach ($Path in $OutlookPaths) {

            if (Test-Path -Path $Path -PathType Leaf) {

                $OutlookExe = $Path

                break
            }
        }


        # ------------------------------------------------------
        # Tenta também pelo PATH do Windows
        # ------------------------------------------------------

        if ($null -eq $OutlookExe) {

            try {

                $Command = Get-Command `
                    "outlook.exe" `
                    -ErrorAction SilentlyContinue

                if ($null -ne $Command) {

                    $OutlookExe = $Command.Source
                }
            }
            catch {
                # Continua para tentativa via COM.
            }
        }


        if ($null -ne $OutlookExe) {

            Write-Message `
                "Outlook encontrado: $OutlookExe" `
                -Level SUCCESS
        }
        else {

            Write-Message `
                "O executável do Outlook não foi localizado nos caminhos conhecidos." `
                -Level WARNING

            Write-Message `
                "Será feita uma tentativa de inicialização através do componente COM." `
                -Level INFO
        }


        # ======================================================
        # 4. INICIALIZAÇÃO DO OUTLOOK VIA COM
        # ======================================================

        Write-Message `
            "Tentando inicializar o Outlook."


        try {

            $Outlook = New-Object `
                -ComObject Outlook.Application `
                -ErrorAction Stop


            Write-Message `
                "Outlook inicializado com sucesso." `
                -Level SUCCESS
        }
        catch {

            Write-Message `
                "Não foi possível inicializar o Outlook." `
                -Level ERROR

            Write-Message `
                "Detalhes: $($_.Exception.Message)" `
                -Level ERROR

            Write-Message `
                "Verifique se o Outlook clássico está instalado e disponível." `
                -Level ERROR

            return
        }


        # ======================================================
        # 5. SESSÃO DO OUTLOOK
        # ======================================================

        Write-Message `
            "Obtendo sessão do Outlook."


        try {

            $Session = $Outlook.Session

            if ($null -eq $Session) {

                Write-Message `
                    "Não foi possível obter a sessão do Outlook." `
                    -Level ERROR

                return
            }


            Write-Message `
                "Sessão do Outlook obtida com sucesso." `
                -Level SUCCESS
        }
        catch {

            Write-Message `
                "Erro ao obter a sessão do Outlook: $($_.Exception.Message)" `
                -Level ERROR

            return
        }


        # ======================================================
        # 6. CONTAS CONFIGURADAS
        # ======================================================

        Write-Message `
            "Verificando contas configuradas no Outlook."


        try {

            $Accounts = $Session.Accounts

            $AccountCount = $Accounts.Count


            Write-Message `
                "Quantidade de contas encontradas: $AccountCount"
        }
        catch {

            Write-Message `
                "Não foi possível acessar as contas do Outlook: $($_.Exception.Message)" `
                -Level ERROR

            return
        }


        if ($AccountCount -eq 0) {

            Write-Message `
                "Nenhuma conta de e-mail está configurada no Outlook." `
                -Level ERROR

            return
        }


        # ======================================================
        # 7. LOCALIZA UMA CONTA COM SMTP
        # ======================================================

        Write-Message `
            "Procurando uma conta com endereço de e-mail válido."


        foreach ($Item in $Accounts) {

            try {

                $DisplayName = $Item.DisplayName
                $SmtpAddress = $Item.SmtpAddress


                Write-Message `
                    "Conta encontrada: $DisplayName | $SmtpAddress"


                if (-not [string]::IsNullOrWhiteSpace($SmtpAddress)) {

                    $Account = $Item

                    break
                }
            }
            catch {

                Write-Message `
                    "Não foi possível obter informações de uma das contas." `
                    -Level WARNING
            }
        }


        if ($null -eq $Account) {

            Write-Message `
                "Nenhuma conta com endereço de e-mail válido foi encontrada." `
                -Level ERROR

            return
        }


        Write-Message `
            "Conta selecionada: $($Account.SmtpAddress)" `
            -Level SUCCESS


        # ======================================================
        # 8. CRIAÇÃO DA MENSAGEM
        # ======================================================

        Write-Message `
            "Criando nova mensagem."


        try {

            $Mail = $Outlook.CreateItem(0)


            if ($null -eq $Mail) {

                Write-Message `
                    "O Outlook não conseguiu criar a mensagem." `
                    -Level ERROR

                return
            }


            Write-Message `
                "Mensagem criada com sucesso." `
                -Level SUCCESS
        }
        catch {

            Write-Message `
                "Erro ao criar a mensagem: $($_.Exception.Message)" `
                -Level ERROR

            return
        }


        # ======================================================
        # 9. CONFIGURAÇÃO DA MENSAGEM
        # ======================================================

        Write-Message `
            "Configurando destinatário."


        $Mail.To = $To


        Write-Message `
            "Destinatário configurado: $To" `
            -Level SUCCESS


        Write-Message `
            "Configurando assunto."


        $Mail.Subject = $Subject


        Write-Message `
            "Assunto configurado: $Subject" `
            -Level SUCCESS


        # ======================================================
        # 10. DEFINE A CONTA DE ENVIO
        # ======================================================

        try {

            $Mail.SendUsingAccount = $Account


            Write-Message `
                "Conta de envio configurada: $($Account.SmtpAddress)" `
                -Level SUCCESS
        }
        catch {

            Write-Message `
                "Não foi possível definir explicitamente a conta de envio." `
                -Level WARNING

            Write-Message `
                "O Outlook utilizará a conta padrão para o envio." `
                -Level WARNING
        }


        # ======================================================
        # 11. CONFIGURAÇÃO DO ARQUIVO
        # ======================================================

        if ($ArquivoParaAnexo) {

            # --------------------------------------------------
            # ANEXO
            # --------------------------------------------------

            Write-Message `
                "Adicionando arquivo como anexo."


            $Mail.Body = @"
Olá,

Segue o arquivo em anexo.

Atenciosamente,
"@


            try {

                $Mail.Attachments.Add($FilePath) |
                    Out-Null


                Write-Message `
                    "Arquivo anexado com sucesso: $FilePath" `
                    -Level SUCCESS
            }
            catch {

                Write-Message `
                    "Não foi possível adicionar o arquivo como anexo: $($_.Exception.Message)" `
                    -Level ERROR

                Write-Message `
                    "O e-mail não será enviado." `
                    -Level ERROR

                return
            }
        }
        elseif ($ArquivoParaCorpo) {

            # --------------------------------------------------
            # CONTEÚDO NO CORPO
            # --------------------------------------------------

            Write-Message `
                "Lendo conteúdo do arquivo."


            try {

                $Content = Get-Content `
                    -Path $FilePath `
                    -Raw `
                    -Encoding UTF8 `
                    -ErrorAction Stop


                if ([string]::IsNullOrEmpty($Content)) {

                    Write-Message `
                        "O arquivo não possui conteúdo." `
                        -Level WARNING

                    Write-Message `
                        "O e-mail será enviado sem conteúdo de arquivo." `
                        -Level WARNING

                    $Mail.Body = ""
                }
                else {

                    $Mail.Body = $Content


                    Write-Message `
                        "Conteúdo do arquivo inserido no corpo do e-mail." `
                        -Level SUCCESS

                    Write-Message `
                        "Quantidade de caracteres: $($Content.Length)"
                }
            }
            catch {

                Write-Message `
                    "Não foi possível ler o conteúdo do arquivo: $($_.Exception.Message)" `
                    -Level WARNING

                Write-Message `
                    "O e-mail será enviado sem o conteúdo do arquivo." `
                    -Level WARNING

                $Mail.Body = ""
            }
        }
        else {

            # --------------------------------------------------
            # SEM ARQUIVO
            # --------------------------------------------------

            $Mail.Body = ""


            Write-Message `
                "Nenhum arquivo válido disponível." `
                -Level WARNING

            Write-Message `
                "Etapas relacionadas ao arquivo foram ignoradas." `
                -Level INFO

            Write-Message `
                "O e-mail será enviado sem arquivo." `
                -Level INFO
        }


        # ======================================================
        # 12. ENVIO
        # ======================================================

        Write-Message `
            "Preparando envio do e-mail."


        try {

            $Mail.Send()


            Write-Message `
                "Comando de envio executado com sucesso." `
                -Level SUCCESS


            Write-Message `
                "E-mail enviado para: $To" `
                -Level SUCCESS


            Write-Message `
                "Conta utilizada: $($Account.SmtpAddress)" `
                -Level SUCCESS
        }
        catch {

            Write-Message `
                "Erro ao enviar o e-mail." `
                -Level ERROR

            Write-Message `
                "Detalhes: $($_.Exception.Message)" `
                -Level ERROR

            return
        }


        # ======================================================
        # FINALIZAÇÃO
        # ======================================================

        Write-Message "=========================================================="
        Write-Message "PROCESSAMENTO CONCLUÍDO"
        Write-Message "=========================================================="
    }
    catch {

        Write-Message `
            "Erro inesperado durante o processamento." `
            -Level ERROR

        Write-Message `
            "Detalhes: $($_.Exception.Message)" `
            -Level ERROR
    }
    finally {

        # ======================================================
        # LIBERAÇÃO DOS OBJETOS COM
        # ======================================================

        Write-Message `
            "Liberando recursos do Outlook."


        if ($null -ne $Mail) {

            try {

                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Mail) |
                    Out-Null
            }
            catch {
                # Ignora erro durante liberação.
            }

            $Mail = $null
        }


        if ($null -ne $Account) {

            try {

                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Account) |
                    Out-Null
            }
            catch {
                # Ignora erro durante liberação.
            }

            $Account = $null
        }


        if ($null -ne $Accounts) {

            try {

                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Accounts) |
                    Out-Null
            }
            catch {
                # Ignora erro durante liberação.
            }

            $Accounts = $null
        }


        if ($null -ne $Session) {

            try {

                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Session) |
                    Out-Null
            }
            catch {
                # Ignora erro durante liberação.
            }

            $Session = $null
        }


        if ($null -ne $Outlook) {

            try {

                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) |
                    Out-Null
            }
            catch {
                # Ignora erro durante liberação.
            }

            $Outlook = $null
        }


        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()


        Write-Message `
            "Recursos do Outlook liberados." `
            -Level INFO
    }
}


# ==========================================================
# CHAMADA DA FUNÇÃO
# ==========================================================

Send-OutlookEmail `
    -To "destinatario@empresa.com" `
    -Subject "Aviso de processamento"