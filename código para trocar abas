<#
comando para rodar
PS C:\temp\app_tv_v0> .\script-app_tv_v0_11.ps1 -RestartEdgeIfOpen -ForceKillEdge -EdgeSilent -Kiosk -KioskType fullscreen
True


.SYNOPSIS
  Abre Edge com 2 abas (Link1 e Link2) e alterna entre elas a cada N segundos,
  atualizando a aba seguinte em segundo plano antes de trocar (CDP/Page.reload + bringToFront).

.DESCRIPTION
  - Suporta configuração "hardcoded" no próprio script (bloco CONFIG)
  - Ainda aceita parâmetros via linha de comando (parâmetros sobrescrevem o CONFIG)
  - Inicia o Edge em modo silencioso opcional (minimizado + flags anti-prompts)
  - Usa Edge DevTools Protocol (CDP) via --remote-debugging-port e endpoint /json/list.

  - Opcional: Detectar Edge do usuário e reiniciar (fechar janelas).
  - Opcional: Arquivar logs a cada 24h em ZIP (compressão alta) e excluir logs.
  - Opcional: "Kiosk" aplicado DEPOIS via CDP (fullscreen após detectar 2 abas).

.PARAMETER Kiosk
  Quando habilitado, após detectar as 2 abas via CDP, o script coloca a janela do Edge em fullscreen via CDP.

.PARAMETER KioskType
  Atualmente suportado neste modo pós-CDP: "fullscreen" (ou "full screen").
  Outros tipos exigiriam flags na inicialização do Edge (não aplicáveis depois).

.NOTES
  - O fullscreen via CDP equivale ao modo tela cheia.
  - Se EdgeSilent estiver habilitado e Kiosk também, o script desabilita o start-minimized para permitir fullscreen corretamente.
#>

# ==========================================================
# IMPORTANTE:
# Em scripts (.ps1), [CmdletBinding()] e param() devem vir
# antes de qualquer código executável (variáveis, funções, etc).
# ==========================================================
[CmdletBinding()]
param(
    [string]$Link1,
    [string]$Link2,

    [ValidateRange(1, 86400)]
    [int]$IntervaloSegundos,

    # NOVO: tempo de espera (segundos) entre atualizar a próxima aba e efetivamente trocar para ela
    [ValidateRange(1, 3600)]
    [int]$periodo_troca_abas,

    [string]$LogPath,

    [switch]$RunHidden,
    [switch]$EdgeSilent,

    [switch]$RestartEdgeIfOpen,
    [switch]$ForceKillEdge,

    [switch]$ArchiveLogsEvery24h,

    # NOVO: Kiosk (aplicar fullscreen via CDP depois de encontrar as abas)
    [switch]$Kiosk,

    # NOVO: KioskType (suportado: fullscreen / "full screen")
    [string]$KioskType,

    # interno: evita loop de relaunch
    [switch]$HiddenChild
)

# =========================
# ======= CONFIG ==========
# =========================

$CFG_Link1             = "https://app.powerbi.com/view?r=eyJrIjoiZDM5MTZkMjEtZGVlOC00OTNiLTkwNWMtMDhhMzdkZTk0NzgzIiwidCI6Ijk2MzgxOWY2LWUxYTMtNDkxYy1hMWNjLTUwOTZmOTE0Y2Y0YiJ9"
$CFG_Link2             = "https://app.powerbi.com/view?r=eyJrIjoiNTU3NGNhOTQtOWE2Ni00MTMyLTk2MTktMTQ5ZGM5NDZiYTczIiwidCI6Ijk2MzgxOWY2LWUxYTMtNDkxYy1hMWNjLTUwOTZmOTE0Y2Y0YiJ9"


<# TESTE
    $CFG_Link1             = "https://g1.globo.com/politica/blog/andreia-sadi/post/2026/03/13/vorcaro-troca-de-advogado-apos-maioria-no-stf-e-abre-caminho-para-delacao-premiada.ghtml"
    $CFG_Link2             = "https://vendaonline.bradescard.com.br/?canal=531&origem=913&tpopto=3&ptovda=17&campa=068&prod=7837"
#>


#Atualiza a cada X segundo para troca de aba
$CFG_IntervaloSegundos = 60

# Aguardar X segundos após atualizar a próxima aba (reload) antes de enviar para frente - Segura para atualizar o BI e depois troca
$CFG_periodo_troca_abas = 5

# LOG EM: C:\ProgramData\app\logs\logs_teste_yyyy_MM_dd_HH_mm.log
$CFG_LogDir            = "C:\ProgramData\app\logs"
$CFG_LogPath           = Join-Path $CFG_LogDir ("logs_teste_{0}.log" -f (Get-Date -Format "yyyy_MM_dd_HH_mm"))

$CFG_EdgeSilentDefault = $true
$CFG_RestartEdgeIfOpenDefault = $true
$CFG_ForceKillEdgeDefault     = $false

$CFG_ArchiveLogsEvery24hDefault = $true
$CFG_ZipPrefix = "log_app_tv"

# Defaults do Kiosk
$CFG_KioskDefault     = $false
$CFG_KioskTypeDefault = "fullscreen"
# =========================

# -------------------------
# RESOLVE DEFAULTS (CONFIG vs PARAM)
# -------------------------
if ([string]::IsNullOrWhiteSpace($Link1)) { $Link1 = $CFG_Link1 }
if ([string]::IsNullOrWhiteSpace($Link2)) { $Link2 = $CFG_Link2 }

if (-not $PSBoundParameters.ContainsKey('IntervaloSegundos')) { $IntervaloSegundos = $CFG_IntervaloSegundos }

# NOVO: resolve default do periodo_troca_abas
if (-not $PSBoundParameters.ContainsKey('periodo_troca_abas')) { $periodo_troca_abas = $CFG_periodo_troca_abas }

if (-not $PSBoundParameters.ContainsKey('LogPath'))          { $LogPath = $CFG_LogPath }

if (-not $PSBoundParameters.ContainsKey('EdgeSilent')) {
    $EdgeSilent = [bool]$CFG_EdgeSilentDefault
}

if (-not $PSBoundParameters.ContainsKey('RestartEdgeIfOpen')) {
    $RestartEdgeIfOpen = [bool]$CFG_RestartEdgeIfOpenDefault
}
if (-not $PSBoundParameters.ContainsKey('ForceKillEdge')) {
    $ForceKillEdge = [bool]$CFG_ForceKillEdgeDefault
}

if (-not $PSBoundParameters.ContainsKey('ArchiveLogsEvery24h')) {
    $ArchiveLogsEvery24h = [bool]$CFG_ArchiveLogsEvery24hDefault
}

if (-not $PSBoundParameters.ContainsKey('Kiosk')) {
    $Kiosk = [bool]$CFG_KioskDefault
}
if ([string]::IsNullOrWhiteSpace($KioskType)) {
    $KioskType = $CFG_KioskTypeDefault
}

# Se Kiosk estiver ligado, não faz sentido iniciar minimizado (senão fullscreen pode não aparecer)
if ($Kiosk) { $EdgeSilent = $false }

# manter log em escopo de script (permite trocar/rotacionar durante execução)
$script:LogPath = $LogPath
$script:LogDir  = Split-Path -Parent $script:LogPath

# -------------------------
# VALIDAÇÕES (sem quebrar estrutura)
# - Capturamos mensagens agora e registramos no log depois que Write-Log existir.
# -------------------------
$script:ValidationMessages = New-Object System.Collections.Generic.List[string]

# 1) Clamp (protege CONFIG errado, pois ValidateRange só atua bem via CLI)
if ($periodo_troca_abas -lt 1) {
    $script:ValidationMessages.Add("Validação: periodo_troca_abas ($periodo_troca_abas) < 1. Ajustando para 1.")
    $periodo_troca_abas = 1
}
elseif ($periodo_troca_abas -gt 3600) {
    $script:ValidationMessages.Add("Validação: periodo_troca_abas ($periodo_troca_abas) > 3600. Ajustando para 3600.")
    $periodo_troca_abas = 3600
}

# 2) Coerência: periodo_troca_abas não pode ser maior que IntervaloSegundos
if ($periodo_troca_abas -gt $IntervaloSegundos) {
    $script:ValidationMessages.Add("Validação: periodo_troca_abas ($periodo_troca_abas) > IntervaloSegundos ($IntervaloSegundos). Ajustando periodo_troca_abas para $IntervaloSegundos.")
    $periodo_troca_abas = $IntervaloSegundos
}

# -------------------------
# LOG
# -------------------------
function Initialize-Log {
    param([string]$Path)

    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

    # Rotação simples: se > 5MB, renomeia para .1 (sobrescreve)
    if (Test-Path $Path) {
        $len = (Get-Item $Path).Length
        if ($len -gt 5MB) {
            $bak = "$Path.1"
            if (Test-Path $bak) { Remove-Item $bak -Force -ErrorAction SilentlyContinue }
            Rename-Item -Path $Path -NewName (Split-Path -Leaf $bak) -Force
        }
    }
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO"
    )
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
    $line = "[$ts][$Level] $Message"
    try {
        Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
    } catch { }
}

# -------------------------
# NOVO: Funções de arquivamento ZIP (a cada 24h)
# -------------------------
function New-LogFilePath {
    param([string]$Dir)
    if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir | Out-Null }
    return (Join-Path $Dir ("logs_teste_{0}.log" -f (Get-Date -Format "yyyy_MM_dd_HH_mm")))
}

function Get-ZipPathForToday {
    param([string]$Dir, [string]$ZipPrefix)
    if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir | Out-Null }
    $zipName = "{0}_{1}.zip" -f $ZipPrefix, (Get-Date -Format "yyyy_MM_dd")
    return (Join-Path $Dir $zipName)
}

function Add-FilesToZipHighCompression {
    param([string[]]$Files, [string]$ZipPath)

    if (-not $Files -or $Files.Count -eq 0) { return $false }

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null

    if (-not (Test-Path $ZipPath)) {
        $null = New-Item -ItemType File -Path $ZipPath -Force
        $zipCreate = [System.IO.Compression.ZipFile]::Open($ZipPath, [System.IO.Compression.ZipArchiveMode]::Create)
        $zipCreate.Dispose()
    }

    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::Open($ZipPath, [System.IO.Compression.ZipArchiveMode]::Update)
        $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal

        foreach ($f in $Files) {
            if (-not (Test-Path $f)) { continue }
            $entryName = (Split-Path $f -Leaf)

            $existing = $zip.GetEntry($entryName)
            if ($existing) { $existing.Delete() }

            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $zip, $f, $entryName, $compressionLevel
            ) | Out-Null
        }
        return $true
    } finally {
        if ($zip) { $zip.Dispose() }
    }
}

function Invoke-ArchiveLogsLast24h {
    param([string]$LogDir, [string]$ZipPrefix)

    $now = Get-Date
    $windowStart = $now.AddHours(-24)

    # Rotaciona o log atual
    $oldLog = $script:LogPath
    $script:LogPath = New-LogFilePath -Dir $LogDir
    Initialize-Log -Path $script:LogPath
    Write-Log "Log rotacionado para arquivamento 24h. Antigo='$oldLog' Novo='$($script:LogPath)'" "INFO"

    $logFiles = Get-ChildItem -Path $LogDir -Filter "*.log" -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $windowStart -and $_.FullName -ne $script:LogPath } |
        Sort-Object LastWriteTime

    if (-not $logFiles -or $logFiles.Count -eq 0) {
        Write-Log "Arquivamento 24h: nenhum .log encontrado nas últimas 24h em '$LogDir'." "INFO"
        return
    }

    $zipPath = Get-ZipPathForToday -Dir $LogDir -ZipPrefix $ZipPrefix

    $manifest = Join-Path $env:TEMP ("{0}_manifest_{1}.txt" -f $ZipPrefix, ($now.ToString("yyyy_MM_dd_HH_mm_ss")))
    $manifestContent = New-Object System.Collections.Generic.List[string]
    $manifestContent.Add("Manifest de arquivamento")
    $manifestContent.Add("Gerado em: $($now.ToString('yyyy-MM-dd HH:mm:ss'))")
    $manifestContent.Add("Janela:   $($windowStart.ToString('yyyy-MM-dd HH:mm:ss')) -> $($now.ToString('yyyy-MM-dd HH:mm:ss'))")
    $manifestContent.Add("")
    $manifestContent.Add("Arquivos incluídos:")
    foreach ($f in $logFiles) { $manifestContent.Add(" - " + $f.FullName) }
    Set-Content -Path $manifest -Value $manifestContent -Encoding UTF8

    Write-Log "Arquivamento 24h: adicionando $($logFiles.Count) arquivos ao ZIP '$zipPath' (Compression=Optimal) e removendo .log após sucesso." "INFO"

    foreach ($f in $logFiles) {
        Write-Log "  incluir: $($f.Name) (LastWrite=$($f.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')))" "INFO"
    }

    $filesToZip = @($logFiles.FullName + $manifest)

    try {
        $ok = Add-FilesToZipHighCompression -Files $filesToZip -ZipPath $zipPath
        if (-not $ok) { throw "Falha ao compactar: rotina de ZIP retornou falso." }

        $removed = 0
        foreach ($f in $logFiles) {
            try {
                Remove-Item -Path $f.FullName -Force -ErrorAction Stop
                $removed++
            } catch {
                Write-Log "Falha ao excluir '$($f.FullName)': $($_.Exception.Message)" "WARN"
            }
        }

        Write-Log "Arquivamento 24h concluído. ZIP='$zipPath'. Logs removidos com sucesso: $removed de $($logFiles.Count)." "INFO"
    } catch {
        Write-Log "Arquivamento 24h falhou: $($_.Exception.Message). Logs NÃO foram apagados." "ERROR"
    } finally {
        if (Test-Path $manifest) { Remove-Item $manifest -Force -ErrorAction SilentlyContinue }
    }
}

Initialize-Log -Path $script:LogPath

# registra avisos de validação (se existirem)
if ($script:ValidationMessages -and $script:ValidationMessages.Count -gt 0) {
    foreach ($m in $script:ValidationMessages) {
        Write-Log $m "WARN"
    }
}

Write-Log "Iniciando. Link1='$Link1' Link2='$Link2' IntervaloSegundos=$IntervaloSegundos periodo_troca_abas=$periodo_troca_abas RunHidden=$RunHidden EdgeSilent=$EdgeSilent RestartEdgeIfOpen=$RestartEdgeIfOpen ForceKillEdge=$ForceKillEdge ArchiveLogsEvery24h=$ArchiveLogsEvery24h Kiosk=$Kiosk KioskType='$KioskType' HiddenChild=$HiddenChild"

# -------------------------
# EXECUÇÃO EM SEGUNDO PLANO (CONSOLE OCULTO)
# -------------------------
if ($RunHidden -and -not $HiddenChild) {
    try {
        $argList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"")
        $argList += @("-Link1", "`"$Link1`"")
        $argList += @("-Link2", "`"$Link2`"")
        $argList += @("-IntervaloSegundos", "$IntervaloSegundos")
        $argList += @("-periodo_troca_abas", "$periodo_troca_abas")
        $argList += @("-LogPath", "`"$script:LogPath`"")

        if ($EdgeSilent)          { $argList += @("-EdgeSilent") }
        if ($RestartEdgeIfOpen)   { $argList += @("-RestartEdgeIfOpen") }
        if ($ForceKillEdge)       { $argList += @("-ForceKillEdge") }
        if ($ArchiveLogsEvery24h) { $argList += @("-ArchiveLogsEvery24h") }
        if ($Kiosk)               { $argList += @("-Kiosk") }
        if (-not [string]::IsNullOrWhiteSpace($KioskType)) { $argList += @("-KioskType", "`"$KioskType`"") }

        $argList += @("-HiddenChild")

        Write-Log "Reexecutando em segundo plano (PS oculto). Args: $($argList -join ' ')" "INFO"
        Start-Process -FilePath "powershell.exe" -ArgumentList $argList -WindowStyle Hidden | Out-Null
        Write-Log "Processo em segundo plano iniciado. Encerrando processo atual." "INFO"
    } catch {
        Write-Log "Falha ao reexecutar em segundo plano: $($_.Exception.Message)" "ERROR"
        throw
    }
    exit
}

# -------------------------
# EDGE / CDP
# -------------------------
$EdgeProfileDir = Join-Path $env:TEMP "Edge_Automation_Profile_CDP"

function Get-EdgePath {
    $candidatos = @(
        (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe"),
        "msedge.exe"
    )
    foreach ($p in $candidatos) {
        if ($p -eq "msedge.exe") {
            $cmd = Get-Command msedge.exe -ErrorAction SilentlyContinue
            if ($cmd) { return $cmd.Source }
        } else {
            if (Test-Path $p) { return $p }
        }
    }
    throw "Não encontrei o executável do Edge (msedge.exe)."
}

function Get-FreeTcpPort {
    $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    $listener.Start()
    $port = ($listener.LocalEndpoint).Port
    $listener.Stop()
    return $port
}

function Start-EdgeWithTabsCDP {
    param(
        [string]$EdgeExe,
        [string]$Url1,
        [string]$Url2,
        [int]$Port,
        [string]$ProfileDir,
        [switch]$Silent
    )

    if (!(Test-Path $ProfileDir)) { New-Item -ItemType Directory -Path $ProfileDir | Out-Null }

    $quietFlags = @(
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-notifications",
        "--mute-audio"
    )

    $args = @(
        "--remote-debugging-port=$Port",
        "--user-data-dir=`"$ProfileDir`"",
        "--new-window"
    )

    if ($Silent) {
        $args += "--start-minimized"
        $args += $quietFlags
    }

    $args += @(
        "`"$Url1`"",
        "`"$Url2`""
    )

    Write-Log "Abrindo Edge (CDP porta $Port) com 2 abas. Silent=$Silent (Kiosk será aplicado depois via CDP se habilitado)" "INFO"

    $windowStyle = if ($Silent) { "Minimized" } else { "Normal" }
    $proc = Start-Process -FilePath $EdgeExe -ArgumentList $args -WindowStyle $windowStyle -PassThru
    return $proc
}

function Wait-ForCDP {
    param([int]$Port)
    $endpoint = "http://127.0.0.1:$Port/json/list"
    for ($i=0; $i -lt 200; $i++) {
        try {
            $r = Invoke-RestMethod -Uri $endpoint -Method GET -TimeoutSec 2
            if ($r) { return $true }
        } catch { }
        Start-Sleep -Milliseconds 150
    }
    throw "Não foi possível conectar ao endpoint CDP em $endpoint."
}

function Get-TabWebSocketUrls {
    param(
        [int]$Port,
        [string]$Url1,
        [string]$Url2
    )
    $endpoint = "http://127.0.0.1:$Port/json/list"
    $pages = Invoke-RestMethod -Uri $endpoint -Method GET

    $tabs = $pages | Where-Object { $_.type -eq "page" -and $_.url -and $_.url -notlike "devtools:*" }
    if (-not $tabs -or $tabs.Count -lt 2) {
        throw "Não consegui identificar 2 abas do tipo 'page' no CDP."
    }

    $tab1 = $tabs | Where-Object { $_.url -like "$Url1*" } | Select-Object -First 1
    $tab2 = $tabs | Where-Object { $_.url -like "$Url2*" } | Select-Object -First 1

    if (-not $tab1 -or -not $tab2) {
        $tab1 = $tabs | Select-Object -First 1
        $tab2 = $tabs | Select-Object -Skip 1 -First 1
        Write-Log "Aviso: não consegui casar URLs exatamente; usando as duas primeiras abas detectadas." "WARN"
    }

    [pscustomobject]@{
        Tab1Id  = $tab1.id
        Tab2Id  = $tab2.id
        Tab1Url = $tab1.url
        Tab2Url = $tab2.url
        Tab1Ws  = $tab1.webSocketDebuggerUrl
        Tab2Ws  = $tab2.webSocketDebuggerUrl
    }
}

function Connect-WebSocket {
    param([string]$WsUrl)

    Add-Type -AssemblyName System.Net.Http | Out-Null
    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $cts = New-Object System.Threading.CancellationTokenSource
    $ws.ConnectAsync([Uri]$WsUrl, $cts.Token).Wait()

    if ($ws.State -ne [System.Net.WebSockets.WebSocketState]::Open) {
        throw "Falha ao abrir WebSocket: $WsUrl"
    }
    return $ws
}

# >>> ALTERADO: agora retorna a resposta do Id (para usarmos Browser.getWindowForTarget)
function Send-CDP {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Ws,
        [string]$Method,
        [hashtable]$Params = $null,
        [int]$Id
    )

    $payload = @{ id = $Id; method = $Method }
    if ($Params) { $payload.params = $Params }

    $json = ($payload | ConvertTo-Json -Compress)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $seg = New-Object System.ArraySegment[byte] -ArgumentList @(,$bytes)

    $ct = [System.Threading.CancellationToken]::None
    $Ws.SendAsync($seg, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $ct).Wait()

    $buffer = New-Object byte[] 65536
    while ($true) {
        $recvSeg = New-Object System.ArraySegment[byte] -ArgumentList @(,$buffer)
        $result = $Ws.ReceiveAsync($recvSeg, $ct).Result

        if ($result.Count -gt 0) {
            $text = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $result.Count)
            try {
                $obj = $text | ConvertFrom-Json
                if ($obj.id -eq $Id) { return $obj }
            } catch {
                # evento ou fragmento; ignora
            }
        }
    }
}

# -------------------------
# NOVO: Browser WebSocket (/json/version) + aplicar fullscreen após achar as abas
# -------------------------
function Get-BrowserWebSocketUrl {
    param([int]$Port)
    $endpoint = "http://127.0.0.1:$Port/json/version"
    $info = Invoke-RestMethod -Uri $endpoint -Method GET
    if (-not $info.webSocketDebuggerUrl) {
        throw "Não encontrei webSocketDebuggerUrl em /json/version (porta $Port)."
    }
    return $info.webSocketDebuggerUrl
}

function Normalize-KioskType {
    param([string]$Value)
    $v = ($Value ?? "").Trim().ToLowerInvariant()
    if ($v -eq "full screen") { return "fullscreen" }
    return $v
}

function Apply-KioskAfterTabs {
    param(
        [int]$Port,
        [string]$AnyTabTargetId,
        [string]$KioskType
    )

    $kt = Normalize-KioskType -Value $KioskType
    if ($kt -ne "fullscreen") {
        Write-Log "KioskType='$KioskType' solicitado, mas no modo pós-CDP suportamos apenas 'fullscreen'. Outros tipos exigem flags no start do msedge." "WARN"
        return
    }

    $browserWsUrl = Get-BrowserWebSocketUrl -Port $Port
    $browserWs = $null

    try {
        $browserWs = Connect-WebSocket -WsUrl $browserWsUrl
        $idLocal = 9000

        # 1) Descobre windowId da janela que contém a aba
        $resp = Send-CDP -Ws $browserWs -Method "Browser.getWindowForTarget" -Params @{ targetId = $AnyTabTargetId } -Id $idLocal
        $idLocal++

        $windowId = $resp.result.windowId
        if (-not $windowId) { throw "Browser.getWindowForTarget não retornou windowId." }

        # 2) Coloca janela em fullscreen
        $null = Send-CDP -Ws $browserWs -Method "Browser.setWindowBounds" -Params @{ windowId = $windowId; bounds = @{ windowState = "fullscreen" } } -Id $idLocal
        $idLocal++

        Write-Log "Kiosk aplicado via CDP: windowState=fullscreen (windowId=$windowId) após detectar abas." "INFO"
    }
    catch {
        Write-Log "Falha ao aplicar fullscreen via CDP após detectar abas: $($_.Exception.Message)" "WARN"
    }
    finally {
        if ($browserWs) { $browserWs.Dispose() }
    }
}

# -------------------------
# NOVO: DETECTAR EDGE "ABERTO" E REINICIAR (OPCIONAL)
# -------------------------
function Get-EdgeUiProcesses {
    try {
        $procs = Get-Process -Name msedge -ErrorAction SilentlyContinue
        if (-not $procs) { return @() }

        $ui = $procs | Where-Object {
            ($_.MainWindowHandle -ne 0) -or (-not [string]::IsNullOrWhiteSpace($_.MainWindowTitle))
        }
        return @($ui)
    } catch {
        return @()
    }
}

function Stop-EdgeGracefully {
    param(
        [int]$WaitSeconds = 8,
        [switch]$ForceKillAll
    )

    $ui = Get-EdgeUiProcesses
    if (-not $ui -or $ui.Count -eq 0) {
        Write-Log "Não há janelas do Edge abertas (apenas background, ou nada). Nada a reiniciar." "INFO"
        return
    }

    Write-Log "Detectadas $($ui.Count) janelas do Edge abertas. Fechando para reiniciar..." "WARN"

    foreach ($p in $ui) {
        try {
            Write-Log "Solicitando fechamento gracioso do Edge. PID=$($p.Id) Title='$($p.MainWindowTitle)'" "INFO"
            $null = $p.CloseMainWindow()
        } catch {
            Write-Log "Falha ao chamar CloseMainWindow no PID=$($p.Id): $($_.Exception.Message)" "WARN"
        }
    }

    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $WaitSeconds) {
        Start-Sleep -Milliseconds 300
        $stillUi = Get-EdgeUiProcesses
        if (-not $stillUi -or $stillUi.Count -eq 0) { break }
    }
    $sw.Stop()

    $stillUi = Get-EdgeUiProcesses
    if ($stillUi -and $stillUi.Count -gt 0) {
        Write-Log "Ainda existem janelas do Edge após $WaitSeconds s. Forçando encerramento dos PIDs: $($stillUi.Id -join ',')" "WARN"
        try { $stillUi | Stop-Process -Force -ErrorAction SilentlyContinue } catch {}
    }

    if ($ForceKillAll) {
        $all = Get-Process -Name msedge -ErrorAction SilentlyContinue
        if ($all) {
            Write-Log "ForceKillEdge habilitado: encerrando TODOS os processos msedge (inclui background)." "WARN"
            try { $all | Stop-Process -Force -ErrorAction SilentlyContinue } catch {}
        }
    }

    Write-Log "Edge do usuário fechado (abas/janelas). Prosseguindo com reinício pelo script." "INFO"
}

# -------------------------
# SINGLE INSTANCE (mutex)
# -------------------------
$createdNew = $false
$mutex = $null

try {
    $mutexName = "Global\EdgeTabSwapMutex"
    $mutex = New-Object System.Threading.Mutex($true, $mutexName, [ref]$createdNew)
} catch {
    $mutexName = "Local\EdgeTabSwapMutex"
    $mutex = New-Object System.Threading.Mutex($true, $mutexName, [ref]$createdNew)
}

if (-not $createdNew) {
    Write-Log "Outra instância do script já está rodando. Encerrando esta instância." "WARN"
    exit 2
}

if ($RestartEdgeIfOpen) {
    Stop-EdgeGracefully -WaitSeconds 8 -ForceKillAll:$ForceKillEdge
}

# -------------------------
# EXECUÇÃO PRINCIPAL
# -------------------------
$ws1 = $null
$ws2 = $null
$edgeProc = $null
$DebugPort = Get-FreeTcpPort
$EdgeProfileDir = Join-Path $env:TEMP "Edge_Automation_Profile_CDP"

$nextArchiveAt = (Get-Date).AddHours(24)

try {
    $edgeExe = Get-EdgePath
    $edgeProc = Start-EdgeWithTabsCDP -EdgeExe $edgeExe -Url1 $Link1 -Url2 $Link2 -Port $DebugPort -ProfileDir $EdgeProfileDir -Silent:$EdgeSilent

    Write-Log "Edge iniciado PID=$($edgeProc.Id). Aguardando CDP na porta $DebugPort..." "INFO"

    Wait-ForCDP -Port $DebugPort
    $tabInfo = Get-TabWebSocketUrls -Port $DebugPort -Url1 $Link1 -Url2 $Link2

    Write-Log "Aba 1 detectada: Id=$($tabInfo.Tab1Id) Url=$($tabInfo.Tab1Url)" "INFO"
    Write-Log "Aba 2 detectada: Id=$($tabInfo.Tab2Id) Url=$($tabInfo.Tab2Url)" "INFO"

    # >>> AQUI: aplicar Kiosk/Fullscreen APÓS detectar as 2 abas
    if ($Kiosk) {
        Write-Log "Kiosk habilitado: aplicando fullscreen via CDP após detectar as abas..." "INFO"
        Apply-KioskAfterTabs -Port $DebugPort -AnyTabTargetId $tabInfo.Tab1Id -KioskType $KioskType
    }

    $ws1 = Connect-WebSocket -WsUrl $tabInfo.Tab1Ws
    $ws2 = Connect-WebSocket -WsUrl $tabInfo.Tab2Ws
    Write-Log "WebSockets conectados (Tab1 e Tab2)." "INFO"

    $id = 1
    $null = Send-CDP -Ws $ws1 -Method "Page.enable" -Id $id; $id++
    $null = Send-CDP -Ws $ws2 -Method "Page.enable" -Id $id; $id++

    $null = Send-CDP -Ws $ws1 -Method "Page.bringToFront" -Id $id; $id++
    Write-Log "Iniciando na Aba 1." "INFO"

    while ($true) {

        if ($ArchiveLogsEvery24h -and (Get-Date) -ge $nextArchiveAt) {
            Write-Log "Disparando rotina de arquivamento 24h..." "INFO"
            Invoke-ArchiveLogsLast24h -LogDir $script:LogDir -ZipPrefix $CFG_ZipPrefix
            $nextArchiveAt = (Get-Date).AddHours(24)
        }

        Start-Sleep -Seconds $IntervaloSegundos

        Write-Log "Ciclo: Atualizar Aba 2 (background), aguardar $periodo_troca_abas para enviar para frente." "INFO"
        $null = Send-CDP -Ws $ws2 -Method "Page.reload" -Params @{ ignoreCache = $true } -Id $id; $id++
        Start-Sleep -Seconds $periodo_troca_abas
        $null = Send-CDP -Ws $ws2 -Method "Page.bringToFront" -Id $id; $id++

        Start-Sleep -Seconds $IntervaloSegundos

        Write-Log "Ciclo: Atualizar Aba 1 (background), aguardar $periodo_troca_abas para enviar para frente." "INFO"
        $null = Send-CDP -Ws $ws1 -Method "Page.reload" -Params @{ ignoreCache = $true } -Id $id; $id++
        Start-Sleep -Seconds $periodo_troca_abas
        $null = Send-CDP -Ws $ws1 -Method "Page.bringToFront" -Id $id; $id++
    }
}
catch {
    Write-Log "Erro fatal: $($_.Exception.Message)" "ERROR"
    throw
}
finally {
    if ($ws1) { $ws1.Dispose() }
    if ($ws2) { $ws2.Dispose() }

    if ($edgeProc -and -not $edgeProc.HasExited) {
        Write-Log "Encerrando Edge do script (PID=$($edgeProc.Id))..." "INFO"
        try { $edgeProc.CloseMainWindow() | Out-Null } catch {}
        Start-Sleep -Milliseconds 800
        try { if (-not $edgeProc.HasExited) { $edgeProc | Stop-Process -Force } } catch {}
    }

    if ($mutex) {
        try { $mutex.ReleaseMutex() | Out-Null } catch {}
        $mutex.Dispose()
    }
    Write-Log "Script finalizado (cleanup ok)." "INFO"
}