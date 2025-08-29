$ErrorActionPreference='Stop'
$sshDir = Join-Path $env:USERPROFILE '.ssh'
if (!(Test-Path $sshDir)) { New-Item -ItemType Directory -Path $sshDir | Out-Null }
$keyPath = Join-Path $sshDir 'id_ed25519_customer_risk'
if (!(Test-Path ($keyPath + '.pub'))) {
  ssh-keygen -q -t ed25519 -C "codex-cli-customer-risk" -f $keyPath -N ""
}
Get-Content ($keyPath + '.pub')
