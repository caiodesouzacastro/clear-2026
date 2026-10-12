# ============================================================
#  noite_run_all.ps1  (sifilis_congenita, roda sem supervisão)
#  1. Portão R4: --check-cols, limpeza dos .txt de sessão, árvore limpa
#  2. ANTES:  infra/acervo        -> rodar_run_all.ps1 -Rotulo antes
#  3. DEPOIS: fix/obito-neonatal  -> rodar_run_all.ps1 -Rotulo depois
#  4. tests/test_neonatal.R
#  Para na primeira falha. Antes de cada run_all espera até 8 h por
#  8 GB de memória livre. Git com até 5 tentativas (30 s).
#  Log único: 05_explore\logs\noite_run_all_<carimbo>.log
# ============================================================

$PROJ = 'D:\Dropbox\Papers\sifilis_congenita'
$R    = 'C:\Program Files\R\R-4.5.3\bin\Rscript.exe'
$X    = 'D:\wt\sifilis_explore\05_explore'
Set-Location $PROJ
$env:SIFILIS_N_WORKERS = '2'
$CARIMBO = Get-Date -Format 'yyyyMMdd_HHmmss'
$L = "$PROJ\05_explore\logs\noite_run_all_$CARIMBO.log"

function Log($msg) {
  "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg" | Tee-Object -FilePath $L -Append
}

function G {
  param([Parameter(ValueFromRemainingArguments = $true)][string[]]$a)
  for ($i = 1; $i -le 5; $i++) {
    $out = & git @a 2>&1
    if ($LASTEXITCODE -eq 0) {
      if ($out) { $out | ForEach-Object { Log "  git: $_" } }
      return $out
    }
    Log "git $($a -join ' ') falhou (tentativa $i/5): $($out -join ' | ')"
    Start-Sleep -Seconds 30
  }
  throw "git $($a -join ' ') falhou 5 vezes"
}

function EsperaMemoria($gb) {
  for ($i = 1; $i -le 96; $i++) {
    $livre = (Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory / 1MB
    if ($livre -ge $gb) {
      Log ("memória livre {0:N1} GB (limite {1} GB): ok" -f $livre, $gb)
      return
    }
    $top = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 3 |
      ForEach-Object { "{0} {1:N1} GB" -f $_.Name, ($_.WorkingSet64 / 1GB) }
    Log ("memória livre {0:N1} GB < {1} GB; esperando 5 min ({2}/96); maiores: {3}" -f $livre, $gb, $i, ($top -join ', '))
    Start-Sleep -Seconds 300
  }
  throw "memória abaixo de $gb GB por 8 horas"
}

function ArvoreLimpa($etapa) {
  $sujo = & git status --short --untracked-files=no 2>&1
  if ($LASTEXITCODE -ne 0 -or $sujo) {
    Log "Árvore com rastreados modificados ($etapa): $($sujo -join ' | ')"
    throw "árvore suja em $etapa"
  }
  Log "árvore limpa ($etapa)"
}

function ConfereBranch($esperado) {
  $b = (& git branch --show-current).Trim()
  if ($b -ne $esperado) { throw "branch atual $b, esperado $esperado" }
  $c = (& git rev-parse --short HEAD).Trim()
  Log "branch $b @ $c"
}

try {
  Log "INICIO noite_run_all (N_WORKERS=$env:SIFILIS_N_WORKERS)"

  # ---- 1. Portão R4 ---------------------------------------------------
  Log "R4: check-cols"
  $o = & $R 02_scripts/19_datasus_fetch.R --check-cols 2>&1
  $ec = $LASTEXITCODE
  $o | ForEach-Object { "$_" } | Out-File -FilePath "$PROJ\05_explore\logs\check_cols_pre_run_all_$CARIMBO.log" -Encoding utf8
  ($o | Select-Object -Last 5) | ForEach-Object { Log "  $_" }
  if ($ec -ne 0 -or -not ($o -match '0 arquivo\(s\) fora do contrato')) { throw "check-cols falhou (exit $ec)" }
  Log "R4: check-cols ok"

  Log "R4: guardando e limpando .txt de sessão"
  $mod = & git status --short --untracked-files=no -- 03_output/logs/sessions/
  if ($mod) {
    Get-ChildItem 03_output\logs\sessions\*.txt | ForEach-Object {
      Copy-Item $_.FullName "$PROJ\05_explore\logs\sessao_pre_run_all_$($_.BaseName)_$CARIMBO.log"
    }
    G restore 03_output/logs/sessions/ | Out-Null
  }
  $nao_rastreados = & git ls-files --others --exclude-standard -- 03_output/
  foreach ($f in $nao_rastreados) {
    $destino = "$PROJ\05_explore\logs\sessao_pre_run_all_$([IO.Path]::GetFileNameWithoutExtension($f))_$CARIMBO.log"
    Move-Item -Force $f $destino
    Log "  movido $f -> $destino"
  }
  $sobra = & git ls-files --others --exclude-standard -- 03_output/
  if ($sobra) { throw "não rastreados ainda em 03_output: $($sobra -join ', ')" }
  ArvoreLimpa "R4"

  # infra/acervo está preso no worktree D:/wt/sifilis_infra: soltar o branch
  # sem apagar nada (o worktree fica em HEAD destacado no mesmo commit)
  $wt = & git worktree list --porcelain
  if ($wt -match 'branch refs/heads/infra/acervo') {
    Log "soltando infra/acervo do worktree D:/wt/sifilis_infra (switch --detach)"
    G -C D:/wt/sifilis_infra switch --detach | Out-Null
  }

  # ---- 2. ANTES -------------------------------------------------------
  Log "ANTES: trocando para infra/acervo"
  G switch infra/acervo | Out-Null
  ConfereBranch 'infra/acervo'
  ArvoreLimpa "pré-ANTES"
  EsperaMemoria 8
  Log "ANTES: INICIO run_all"
  pwsh -NoProfile -File "$X\rodar_run_all.ps1" -Rotulo antes 2>&1 | Tee-Object -FilePath $L -Append
  $ec = $LASTEXITCODE
  Log "ANTES: FIM run_all (exit $ec)"
  if ($ec -ne 0) { throw "ANTES terminou com exit $ec; DEPOIS não roda" }
  ArvoreLimpa "pós-ANTES"

  # ---- 3. DEPOIS ------------------------------------------------------
  Log "DEPOIS: trocando para fix/obito-neonatal"
  G switch fix/obito-neonatal | Out-Null
  ConfereBranch 'fix/obito-neonatal'
  ArvoreLimpa "pré-DEPOIS"
  EsperaMemoria 8
  Log "DEPOIS: INICIO run_all"
  pwsh -NoProfile -File "$X\rodar_run_all.ps1" -Rotulo depois 2>&1 | Tee-Object -FilePath $L -Append
  $ec = $LASTEXITCODE
  Log "DEPOIS: FIM run_all (exit $ec)"
  if ($ec -ne 0) { throw "DEPOIS terminou com exit $ec" }

  # ---- 4. Teste -------------------------------------------------------
  Log "TESTE: tests/test_neonatal.R"
  & $R tests/test_neonatal.R 2>&1 | Tee-Object -FilePath "$PROJ\05_explore\logs\test_neonatal_completo_$CARIMBO.log"
  $ec = $LASTEXITCODE
  Log "TESTE: FIM (exit $ec)"
  if ($ec -ne 0) { throw "test_neonatal.R falhou (exit $ec)" }

  Log "TUDO OK: ANTES, DEPOIS e teste concluídos"
}
catch {
  Log "PAROU: $($_.Exception.Message)"
}
finally {
  Log "FIM noite_run_all. Branch: $((& git branch --show-current).Trim())"
}
