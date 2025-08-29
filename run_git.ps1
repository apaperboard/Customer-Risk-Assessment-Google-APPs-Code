$ErrorActionPreference='Stop'
function Find-Git {
  $candidates = @(
    'C:\Program Files\Git\cmd\git.exe',
    'C:\Program Files\Git\bin\git.exe',
    'C:\Program Files (x86)\Git\cmd\git.exe'
  )
  foreach ($p in $candidates) { if (Test-Path $p) { return $p } }
  throw 'git not found'
}

$git = Find-Git
& $git init -q
& $git config user.name 'AR Tool Bot'
& $git config user.email 'bot@example.com'
& $git add -A
try {
  & $git commit -m 'Initial commit: Apps Script + SPA + Pages workflow'
} catch {}
& $git branch -M main
& $git remote remove origin 2>$null
& $git remote add origin 'git@github.com:apaperboard/Customer-Risk-Assessment-Google-APPs-Code.git'
& $git push -u origin main
