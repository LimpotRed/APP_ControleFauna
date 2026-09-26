<#
Gera a planilha-base do aplicativo de monitoramento ambiental.
O arquivo de saída é criado apenas se ainda não existir, para evitar sobrescrever
um trabalho já feito.
#>

[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'MONITORAMENTO_AMBIENTAL_APPSHEET.xlsx'),
    [switch]$ReplaceEmptyOutput
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$fileMode = [System.IO.FileMode]::CreateNew
if (Test-Path -LiteralPath $OutputPath) {
    $existingFile = Get-Item -LiteralPath $OutputPath
    if (-not $ReplaceEmptyOutput -or $existingFile.Length -ne 0) {
        throw "O arquivo de saída já existe e não será sobrescrito: $OutputPath"
    }
    $fileMode = [System.IO.FileMode]::Create
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function ConvertTo-XmlText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    # Windows PowerShell can read a UTF-8 script without BOM using the legacy
    # code page. This repair keeps Portuguese accents correct in the .xlsx.
    if ($text -match '[^\x00-\x7F]') {
        $text = [System.Text.Encoding]::UTF8.GetString(
            [System.Text.Encoding]::GetEncoding(1252).GetBytes($text)
        )
    }
    return [System.Security.SecurityElement]::Escape($text)
}

function Get-ExcelColumnName {
    param([int]$Index)
    $name = ''
    while ($Index -gt 0) {
        $remainder = ($Index - 1) % 26
        $name = [char](65 + $remainder) + $name
        $Index = [math]::Floor(($Index - 1) / 26)
    }
    return $name
}

function New-CellXml {
    param(
        [string]$Reference,
        [AllowNull()][object]$Value,
        [bool]$IsHeader,
        [bool]$IsNumeric
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return ''
    }

    if ($IsNumeric) {
        $numericValue = ([string]$Value).Replace(',', '.')
        return "<c r='$Reference'><v>$numericValue</v></c>"
    }

    $style = if ($IsHeader) { " s='1'" } else { '' }
    $escaped = ConvertTo-XmlText $Value
    return "<c r='$Reference'$style t='inlineStr'><is><t>$escaped</t></is></c>"
}

function New-WorksheetXml {
    param([pscustomobject]$Sheet)

    $header = $Sheet.Rows[0]
    $lastColumn = Get-ExcelColumnName $header.Count
    $lastRow = $Sheet.Rows.Count
    $columnXml = [System.Text.StringBuilder]::new()
    for ($columnIndex = 0; $columnIndex -lt $header.Count; $columnIndex++) {
        $width = if ($Sheet.Widths.Count -gt $columnIndex) { $Sheet.Widths[$columnIndex] } else { 18 }
        [void]$columnXml.Append("<col min='$($columnIndex + 1)' max='$($columnIndex + 1)' width='$width' customWidth='1'/>")
    }

    $rowsXml = [System.Text.StringBuilder]::new()
    for ($rowIndex = 0; $rowIndex -lt $Sheet.Rows.Count; $rowIndex++) {
        $rowNumber = $rowIndex + 1
        $rowAttributes = if ($rowIndex -eq 0) { " r='$rowNumber' ht='24' customHeight='1'" } else { " r='$rowNumber'" }
        [void]$rowsXml.Append("<row$rowAttributes>")
        $row = $Sheet.Rows[$rowIndex]
        for ($columnIndex = 0; $columnIndex -lt $header.Count; $columnIndex++) {
            $columnName = [string]$header[$columnIndex]
            $isNumeric = ($rowIndex -gt 0) -and ($Sheet.NumericColumns -contains $columnName)
            $cell = New-CellXml -Reference "$(Get-ExcelColumnName ($columnIndex + 1))$rowNumber" -Value $row[$columnIndex] -IsHeader ($rowIndex -eq 0) -IsNumeric $isNumeric
            [void]$rowsXml.Append($cell)
        }
        [void]$rowsXml.Append('</row>')
    }

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>$columnXml</cols>
  <sheetData>$rowsXml</sheetData>
  <autoFilter ref="A1:$lastColumn$lastRow"/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"@
}

function Add-ZipText {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$Path,
        [string]$Content
    )

    $entry = $Archive.CreateEntry($Path, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
    try {
        $writer.Write($Content)
    }
    finally {
        $writer.Dispose()
    }
}

$worksheets = @(
    [pscustomobject]@{
        Name = 'Fauna_Rodovia'
        NumericColumns = @('Km')
        Widths = @(16, 13, 10, 13, 10, 24, 25, 19, 20, 14, 18, 22, 30, 16)
        Rows = @(
            @('ID_Ocorrencia', 'Data', 'Hora', 'Rodovia', 'Km', 'Localizacao', 'Especie', 'Grupo_Faunistico', 'Tipo_Ocorrencia', 'Condicao', 'Fotografia', 'Entorno', 'Observacoes', 'Status_Revisao'),
            @('F001', '2026-09-10', '08:30', 'BR-153', '512.4', '-16.6800, -49.2500', 'Tamanduá-bandeira', 'Mamífero', 'Atropelamento', 'Carcaça', '', 'Pastagem', 'Registro de demonstração.', 'Pendente'),
            @('F002', '2026-09-11', '14:10', 'BR-060', '178.2', '-16.2500, -49.3000', 'Tucano', 'Ave', 'Animal vivo', 'Vivo', '', 'Rio', 'Registro de demonstração.', 'Revisado'),
            @('F003', '2026-09-12', '07:45', 'BR-153', '515.1', '-16.6900, -49.2600', 'Capivara', 'Mamífero', 'Atropelamento', 'Carcaça', '', 'Floresta', 'Registro de demonstração.', 'Pendente'),
            @('F004', '2026-09-12', '16:20', 'BR-060', '181.7', '-16.2700, -49.3200', 'Serpente', 'Réptil', 'Animal vivo', 'Vivo', '', 'Floresta / pastagem', 'Registro de demonstração.', 'Revisado'),
            @('F005', '2026-09-13', '09:15', 'BR-153', '519.3', '-16.7000, -49.2700', 'Tatu-galinha', 'Mamífero', 'Atropelamento', 'Carcaça', '', 'Agricultura', 'Registro de demonstração.', 'Pendente'),
            @('F006', '2026-09-14', '10:05', 'BR-153', '520.8', '-16.7050, -49.2750', 'Sapo-cururu', 'Anfíbio', 'Vestígio', 'Vestígio', '', 'Agricultura', 'Registro de demonstração.', 'Pendente'),
            @('F007', '2026-09-14', '17:40', 'BR-060', '183.1', '-16.2810, -49.3300', 'Gambá', 'Mamífero', 'Atropelamento', 'Carcaça', '', 'Agricultura', 'Registro de demonstração.', 'Revisado'),
            @('F008', '2026-09-15', '06:55', 'BR-153', '522.0', '-16.7100, -49.2800', 'Coruja', 'Ave', 'Animal vivo', 'Vivo', '', 'Agricultura', 'Registro de demonstração.', 'Pendente')
        )
    },
    [pscustomobject]@{
        Name = 'Flora_Risco'
        NumericColumns = @('Km', 'DAP_cm', 'Altura_m')
        Widths = @(14, 16, 10, 13, 10, 24, 24, 12, 12, 24, 16, 20, 18, 35, 16)
        Rows = @(
            @('ID_Arvore', 'Data_Inspecao', 'Hora', 'Rodovia', 'Km', 'Localizacao', 'Especie', 'DAP_cm', 'Altura_m', 'Estado_Fitossanitario', 'Risco_Queda', 'Alvo_Proximo', 'Fotografia', 'Recomendacao', 'Status_Revisao'),
            @('A001', '2026-09-10', '09:10', 'BR-153', '512.6', '-16.6810, -49.2520', 'Ipê-amarelo', '42.0', '12.0', 'Bom', 'Baixo', 'Nenhum', '', 'Monitoramento de rotina.', 'Revisado'),
            @('A002', '2026-09-11', '10:30', 'BR-060', '178.6', '-16.2510, -49.3020', 'Figueira', '86.0', '18.0', 'Regular', 'Médio', 'Via', '', 'Inspecionar novamente em 30 dias.', 'Pendente'),
            @('A003', '2026-09-12', '08:50', 'BR-153', '515.8', '-16.6920, -49.2620', 'Eucalipto', '110.0', '24.0', 'Ruim', 'Alto', 'Rede elétrica', '', 'Solicitar avaliação técnica priorizada.', 'Pendente'),
            @('A004', '2026-09-13', '15:35', 'BR-060', '182.4', '-16.2780, -49.3250', 'Angico', '55.0', '15.0', 'Bom', 'Baixo', 'Pedestre', '', 'Manter acompanhamento anual.', 'Revisado'),
            @('A005', '2026-09-14', '11:20', 'BR-153', '521.4', '-16.7080, -49.2790', 'Sibipiruna', '73.0', '17.0', 'Crítico', 'Alto', 'Via', '', 'Isolar área e encaminhar para avaliação técnica.', 'Pendente')
        )
    },
    [pscustomobject]@{
        Name = 'Rodovias'
        NumericColumns = @()
        Widths = @(14, 42)
        Rows = @(
            @('Rodovia', 'Descricao'),
            @('BR-060', 'Rodovia de demonstração'),
            @('BR-153', 'Rodovia de demonstração'),
            @('GO-010', 'Rodovia de demonstração'),
            @('GO-020', 'Rodovia de demonstração')
        )
    },
    [pscustomobject]@{
        Name = 'Grupos_Fauna'
        NumericColumns = @()
        Widths = @(24)
        Rows = @(
            @('Grupo_Faunistico'),
            @('Mamífero'),
            @('Ave'),
            @('Réptil'),
            @('Anfíbio'),
            @('Outro')
        )
    },
    [pscustomobject]@{
        Name = 'Condicao_Fauna'
        NumericColumns = @()
        Widths = @(20)
        Rows = @(
            @('Condicao'),
            @('Vivo'),
            @('Carcaça'),
            @('Vestígio')
        )
    },
    [pscustomobject]@{
        Name = 'Risco_Queda'
        NumericColumns = @()
        Widths = @(22)
        Rows = @(
            @('Risco_Queda'),
            @('Baixo'),
            @('Médio'),
            @('Alto'),
            @('Não determinado')
        )
    },
    [pscustomobject]@{
        Name = 'Estado_Fitossanitario'
        NumericColumns = @()
        Widths = @(27)
        Rows = @(
            @('Estado_Fitossanitario'),
            @('Bom'),
            @('Regular'),
            @('Ruim'),
            @('Crítico')
        )
    },
    [pscustomobject]@{
        Name = 'Usuarios'
        NumericColumns = @()
        Widths = @(38, 26, 18, 12)
        Rows = @(
            @('Email', 'Nome', 'Perfil', 'Ativo'),
            @('substitua-por-email-real@exemplo.com', 'Substitua este cadastro', 'Administrador', 'Sim')
        )
    }
)

$sheetDefinitions = [System.Text.StringBuilder]::new()
$worksheetRelationships = [System.Text.StringBuilder]::new()
for ($sheetIndex = 0; $sheetIndex -lt $worksheets.Count; $sheetIndex++) {
    $sheet = $worksheets[$sheetIndex]
    $relationshipId = $sheetIndex + 1
    [void]$sheetDefinitions.Append("<sheet name='$($sheet.Name)' sheetId='$relationshipId' r:id='rId$relationshipId'/>")
    [void]$worksheetRelationships.Append("<Relationship Id='rId$relationshipId' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet$relationshipId.xml'/>")
}

$contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
$(for ($sheetIndex = 0; $sheetIndex -lt $worksheets.Count; $sheetIndex++) { "  <Override PartName='/xl/worksheets/sheet$($sheetIndex + 1).xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/>" })
</Types>
"@

$rootRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@

$workbook = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView/></bookViews>
  <sheets>$sheetDefinitions</sheets>
</workbook>
"@

$workbookRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  $worksheetRelationships
  <Relationship Id="rId$($worksheets.Count + 1)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@

$styles = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"@

$createdAt = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
$coreProperties = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>MONITORAMENTO_AMBIENTAL_APPSHEET</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$createdAt</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$createdAt</dcterms:modified>
</cp:coreProperties>
"@

$applicationProperties = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>$($worksheets.Count)</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="$($worksheets.Count)" baseType="lpstr">
$(foreach ($sheet in $worksheets) { "      <vt:lpstr>$($sheet.Name)</vt:lpstr>" })
    </vt:vector>
  </TitlesOfParts>
</Properties>
"@

$fileStream = [System.IO.FileStream]::new($OutputPath, $fileMode)
$archive = [System.IO.Compression.ZipArchive]::new($fileStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)
try {
    Add-ZipText -Archive $archive -Path '[Content_Types].xml' -Content $contentTypes
    Add-ZipText -Archive $archive -Path '_rels/.rels' -Content $rootRelationships
    Add-ZipText -Archive $archive -Path 'xl/workbook.xml' -Content $workbook
    Add-ZipText -Archive $archive -Path 'xl/_rels/workbook.xml.rels' -Content $workbookRelationships
    Add-ZipText -Archive $archive -Path 'xl/styles.xml' -Content $styles
    Add-ZipText -Archive $archive -Path 'docProps/core.xml' -Content $coreProperties
    Add-ZipText -Archive $archive -Path 'docProps/app.xml' -Content $applicationProperties
    for ($sheetIndex = 0; $sheetIndex -lt $worksheets.Count; $sheetIndex++) {
        Add-ZipText -Archive $archive -Path "xl/worksheets/sheet$($sheetIndex + 1).xml" -Content (New-WorksheetXml -Sheet $worksheets[$sheetIndex])
    }
}
finally {
    $archive.Dispose()
    $fileStream.Dispose()
}

Write-Host "Planilha criada em: $OutputPath"
