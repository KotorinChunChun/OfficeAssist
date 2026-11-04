param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$folder = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $folder -Force | Out-Null
[Text.Encoding]::RegisterProvider([Text.CodePagesEncodingProvider]::Instance)
$culture = [Globalization.CultureInfo]::InvariantCulture
$kinds = @('番号','日付','金額','日本語','改行','引用符','数式風','空欄')
function Format-Field([string]$Value, [char]$Separator) {
    if ($Value.Contains($Separator) -or $Value.Contains('"') -or $Value.Contains("`r") -or $Value.Contains("`n")) {
        return '"' + $Value.Replace('"','""') + '"'
    }
    return $Value
}
foreach ($columnCount in @(40,64)) {
    $separator = if ($columnCount -eq 40) { [char]',' } else { [char]9 }
    $extension = if ($columnCount -eq 40) { 'csv' } else { 'tsv' }
    $name = if ($columnCount -eq 40) { '01-wide-40cols.csv' } else { '02-wide-64cols.tsv' }
    $path = Join-Path $folder $name
    if (Test-Path -LiteralPath $path) { throw "既存のサンプルは上書きしません: $path" }
    $lines = [Collections.Generic.List[string]]::new()
    $fields = for ($c = 0; $c -lt $columnCount; $c++) { '列{0:D2}_{1}' -f ($c + 1), $kinds[$c % 8] }
    $lines.Add(($fields -join $separator))
    for ($r = 1; $r -le 1500; $r++) {
        $fields = for ($c = 0; $c -lt $columnCount; $c++) {
            $value = switch ($c % 8) {
                0 { $r.ToString('D6',$culture) }
                1 { [datetime]::new(2026,11,5).AddDays($r % 31).ToString('yyyy/MM/dd',$culture) }
                2 { ([double]($r * ($c + 1)) / 10).ToString('0.00',$culture) }
                3 { "サンプル$r,列$($c + 1)" }
                4 { "前半$r`n後半$r" }
                5 { "引用符`"$r`"" }
                6 { "=SUM($r,1)" }
                7 { '' }
            }
            Format-Field ([string]$value) $separator
        }
        $lines.Add(($fields -join $separator))
    }
    $encoding = if ($columnCount -eq 40) { [Text.UTF8Encoding]::new($true) } else { [Text.Encoding]::GetEncoding(932) }
    [IO.File]::WriteAllLines($path,$lines,$encoding)
    [pscustomobject]@{File=$path;Columns=$columnCount;DataRows=1500;Cells=($columnCount*1501);Encoding=$encoding.WebName}
}
