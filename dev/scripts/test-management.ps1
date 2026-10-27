param([Parameter(Mandatory)][string]$Executable, [switch]$StartForTest)
$ErrorActionPreference = 'Stop'
$officeExecutable = (Resolve-Path -LiteralPath $Executable).Path
$officeSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$officePipeName = "KotorinToys.$officeSid.office-assist"
function Send-OfficeRequest([string]$Request) {
    $pipe = [System.IO.Pipes.NamedPipeClientStream]::new('.', $officePipeName, [System.IO.Pipes.PipeDirection]::InOut)
    try {
        $pipe.Connect(2000)
        $writer = [System.IO.StreamWriter]::new($pipe, [System.Text.UTF8Encoding]::new($false), 1024, $true)
        $reader = [System.IO.StreamReader]::new($pipe, [System.Text.Encoding]::UTF8, $true, 1024, $true)
        try {
            $writer.AutoFlush = $true
            $writer.WriteLine($Request)
            $reader.ReadLineAsync().WaitAsync([TimeSpan]::FromSeconds(4)).GetAwaiter().GetResult()
        } finally {
            $reader.Dispose()
            # 不正JSONはサーバーが応答せず切断する。送信済みwriterの後始末だけを許容する。
            try { $writer.Dispose() } catch [System.IO.IOException] { }
        }
    } finally { $pipe.Dispose() }
}
$owned = $null
try {
if ($StartForTest) {
    if (Get-Process OfficeAssist -ErrorAction SilentlyContinue) { throw '起動中のOfficeAssistがあるため専用起動を中止します。' }
    $owned = Start-Process -FilePath $officeExecutable -ArgumentList '--startup' -WindowStyle Hidden -PassThru
}
$before = Send-OfficeRequest 'status' | ConvertFrom-Json
$officeProcess = Get-Process -Id $before.pid
if ($officeProcess.Path -ne $officeExecutable) { throw '試験対象と起動中の実行ファイルが異なります。' }
if ($before.protocolVersion -ne 2 -or !$before.ready -or $before.version -ne '0.1.0') { throw '管理状態が一致しません。' }
'PASS: 起動中の実行ファイル・管理契約2・版・readyを確認'
$refused = Send-OfficeRequest ('{"command":"stop","expectedPid":' + ($before.pid + 1) + '}') | ConvertFrom-Json
if ($refused.result -ne 'refused') { throw '不一致PIDの停止要求が拒否されません。' }
'PASS: 不一致PIDの停止要求を拒否'
$null = Send-OfficeRequest '{broken'
$afterInvalid = Send-OfficeRequest 'status' | ConvertFrom-Json
if ($afterInvalid.pid -ne $before.pid -or !$afterInvalid.ready) { throw '不正要求後に受付が継続していません。' }
'PASS: 不正JSONの後も状態取得を継続'
$duplicate = Start-Process -FilePath $officeExecutable -ArgumentList '--startup' -WindowStyle Hidden -PassThru
try {
    if (!$duplicate.WaitForExit(5000) -or $duplicate.ExitCode -ne 0) { throw '無表示の二重起動が正常終了しません。' }
} finally { $duplicate.Dispose() }
$after = Send-OfficeRequest 'status' | ConvertFrom-Json
if ($after.pid -ne $before.pid -or !$after.ready) { throw '二重起動後に既存PIDが保持されません。' }
'PASS: 無表示の二重起動が終了し、既存PIDを維持'
if ($owned) {
    $owned.Refresh()
    if ($owned.MainWindowHandle -ne [IntPtr]::Zero) { throw '無表示起動でメイン画面が表示されています。' }
    'PASS: startupと無表示の二重起動でメイン画面を表示しない'
}
} finally {
    if ($owned) {
        try {
            if (!$owned.HasExited) {
                $stopped = Send-OfficeRequest ('{"command":"stop","expectedPid":' + $owned.Id + '}') | ConvertFrom-Json
                if ($stopped.result -ne 'accepted' -or !$owned.WaitForExit(5000) -or $owned.ExitCode -ne 0) { throw '試験所有プロセスを正常終了できません。' }
                'PASS: 試験所有PIDだけを管理要求で正常終了'
            }
        } finally { $owned.Dispose() }
    }
}
