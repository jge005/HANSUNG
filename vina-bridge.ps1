param(
  [int]$Port = 18765
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Web

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:$Port/")
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$fillScript = Join-Path $scriptRoot 'vina-fill-only.ps1'
$runtimePayload = Join-Path $scriptRoot '_vina_runtime_payload.json'

function Write-JsonResponse($context, [int]$statusCode, $payload) {
  $json = $payload | ConvertTo-Json -Depth 10
  $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
  $response = $context.Response
  $response.StatusCode = $statusCode
  $response.ContentType = 'application/json; charset=utf-8'
  $response.Headers['Access-Control-Allow-Origin'] = '*'
  $response.Headers['Access-Control-Allow-Headers'] = 'Content-Type'
  $response.Headers['Access-Control-Allow-Methods'] = 'GET,POST,OPTIONS'
  $response.OutputStream.Write($buffer, 0, $buffer.Length)
  $response.OutputStream.Close()
}

function Get-DefaultOutputPath($payload) {
  $documents = [Environment]::GetFolderPath('MyDocuments')
  $shipDate = if ($payload.shipDate) { [datetime]$payload.shipDate } else { Get-Date }
  $transport = if ([string]::IsNullOrWhiteSpace([string]$payload.transport)) { 'AIR' } elseif ([string]$payload.transport -match 'VSL') { 'VSL' } else { 'AIR' }
  return (Join-Path $documents ("[HANSUNG] {0} VINA ({1}) filled detail.xlsx" -f $shipDate.ToString('yyMMdd'), $transport))
}

function Invoke-VinaFill($payload) {
  $payload | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $runtimePayload -Encoding UTF8
  $hasOutputPath = $payload.PSObject.Properties.Name -contains 'outputPath'
  $outputPath = if ($hasOutputPath -and -not [string]::IsNullOrWhiteSpace([string]$payload.outputPath)) {
    [string]$payload.outputPath
  } else {
    Get-DefaultOutputPath $payload
  }
  $outputDir = Split-Path -Parent $outputPath
  if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
    [void](New-Item -ItemType Directory -Path $outputDir -Force)
  }
  $stdout = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $fillScript -PayloadPath $runtimePayload -OutputPath $outputPath 2>&1 | Out-String
  if ($LASTEXITCODE -ne 0) {
    throw ($stdout.Trim())
  }
  return @{
    outputPath = $outputPath
    log = $stdout.Trim()
  }
}

try {
  while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request

    if ($request.HttpMethod -eq 'OPTIONS') {
      Write-JsonResponse $context 200 @{ ok = $true }
      continue
    }

    try {
      if ($request.HttpMethod -eq 'GET' -and $request.Url.AbsolutePath -eq '/health') {
        Write-JsonResponse $context 200 @{
          ok = $true
          bridge = 'vina'
          fillScript = $fillScript
        }
        continue
      }

      if ($request.HttpMethod -eq 'POST' -and $request.Url.AbsolutePath -eq '/vina/fill') {
        $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
        $body = $reader.ReadToEnd()
        $reader.Close()
        $payload = $body | ConvertFrom-Json
        $result = Invoke-VinaFill $payload
        Write-JsonResponse $context 200 @{
          ok = $true
          outputPath = $result.outputPath
          log = $result.log
        }
        continue
      }

      Write-JsonResponse $context 404 @{
        ok = $false
        error = 'Not found'
      }
    } catch {
      Write-JsonResponse $context 500 @{
        ok = $false
        error = $_.Exception.Message
      }
    }
  }
}
finally {
  if ($listener.IsListening) {
    $listener.Stop()
  }
  $listener.Close()
}
