function Get-Order {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$InputExcel,

        [Parameter(Mandatory = $true)]
        [string]$OutputWord,

        [string]$Text
    )

    $exePath = Join-Path $PSScriptRoot "order.exe"
    $defaultXml = Join-Path $PSScriptRoot "pizza.xml"

    if (-not (Test-Path $exePath)) {
        throw "order.exe not found in module folder: $exePath"
    }

    if (-not $Text) {
        $Text = $defaultXml
    }

    if (-not (Test-Path $Text)) {
        throw "Text XML file not found: $Text"
    }

    $args = @(
        "--input-excel", "`"$InputExcel`"",
        "--output-word", "`"$OutputWord`"",
        "--text", "`"$Text`""
    )

    & $exePath @args
}
