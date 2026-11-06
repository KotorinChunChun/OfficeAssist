param([Parameter(Mandatory)][string]$Before,[Parameter(Mandatory)][string]$After,[Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$baseline=Get-Content -LiteralPath $Before -Raw | ConvertFrom-Json
$updated=Get-Content -LiteralPath $After -Raw | ConvertFrom-Json
if($baseline.schemaVersion -ne 1 -or $updated.schemaVersion -ne 1 -or $baseline.repeats -ne $updated.repeats){throw '測定形式または反復回数が一致しません。'}
$comparison=@(foreach($previous in $baseline.results){
    $next=@($updated.results|Where-Object Name -eq $previous.Name)
    if($next.Count -ne 1 -or $next[0].Rows -ne $previous.Rows -or $next[0].Columns -ne $previous.Columns -or $next[0].Features -ne $previous.Features){throw "比較する条件が不一致です: $($previous.Name)"}
    [pscustomobject]@{name=$previous.Name;beforeMs=$previous.MedianMs;afterMs=$next[0].MedianMs;beforeFirstMs=$previous.Milliseconds[0];afterFirstMs=$next[0].Milliseconds[0];speedup=if($next[0].MedianMs -gt 0){$previous.MedianMs/$next[0].MedianMs}else{0};sameResult=$previous.ResultHash -eq $next[0].ResultHash}
})
if(Test-Path -LiteralPath $OutputDirectory){throw '比較結果の保存先は未使用フォルダーを指定してください。'}
New-Item -ItemType Directory -Path $OutputDirectory|Out-Null
$comparison|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $OutputDirectory 'comparison.json') -Encoding utf8
$comparison|Export-Csv -LiteralPath (Join-Path $OutputDirectory 'comparison.csv') -NoTypeInformation -Encoding utf8
$comparison|Format-Table name,@{n='改修前ms';e={'{0:F2}' -f $_.beforeMs}},@{n='改修後ms';e={'{0:F2}' -f $_.afterMs}},@{n='倍率';e={'{0:F2}' -f $_.speedup}},sameResult
if(@($comparison|Where-Object {-not $_.sameResult}).Count -gt 0){throw '処理結果が一致しないケースがあります。comparison.jsonを確認してください。'}
