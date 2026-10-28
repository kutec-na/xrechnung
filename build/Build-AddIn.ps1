<#
  Baut das Word-Add-In XRechnung.dotm aus den Quelldateien in ..\src.
  Nur für die Entwicklung nötig - das fertige Add-In enthält den gesamten Code als VBA.

  Voraussetzung: In Word ist "Zugriff auf das VBA-Projektobjektmodell vertrauen"
  aktiviert (Trust Center > Makroeinstellungen).

  Ergebnis:
    ..\XRechnung.dotm       Add-In mit Menüband-Schaltfläche "XRechnung erzeugen"
    ..\export\*.bas|cls|frm  importierbare VBA-Dateien (Windows-1252) für manuelle Installation
#>
param(
    [string]$Out = (Join-Path $PSScriptRoot '..\XRechnung.dotm')
)
$ErrorActionPreference = 'Stop'
$root   = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$src    = Join-Path $root 'src'
$export = Join-Path $root 'export'
$Out    = [IO.Path]::GetFullPath($Out)

$components = @(
    @{ Type = 3; Name = 'frmXRechnung'; File = 'frmXRechnung.frm.txt' },
    @{ Type = 2; Name = 'clsBTRow';     File = 'clsBTRow.cls' },
    @{ Type = 2; Name = 'clsMapRow';    File = 'clsMapRow.cls' },
    @{ Type = 1; Name = 'modXRechnung'; File = 'modXRechnung.bas' }
)

if (Test-Path $Out) { Remove-Item $Out }

# Word mit neutralem TEMP-Ordner starten: der VBA-Editor legt für Formulare einen Verweis
# auf %TEMP%\VBE\MSForms.exd im Projekt ab - ohne das stünde dort der Benutzername.
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Bitte Word vor dem Build schließen.' }
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class XrRot {
  [DllImport("ole32.dll")] static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string p, out Guid g);
  [DllImport("oleaut32.dll")] static extern int GetActiveObject(ref Guid g, IntPtr r, [MarshalAs(UnmanagedType.IUnknown)] out object o);
  public static object Get(string progId) { Guid g; CLSIDFromProgID(progId, out g); object o; return GetActiveObject(ref g, IntPtr.Zero, out o) == 0 ? o : null; }
}
"@
$neutralTemp = Join-Path $env:PUBLIC 'Temp'
New-Item -ItemType Directory -Force $neutralTemp | Out-Null
$exe = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe').'(default)'
$proc = Start-Process $exe -ArgumentList '/q', '/n', '/automation' -Environment @{ TEMP = $neutralTemp; TMP = $neutralTemp } -PassThru
$word = $null
$deadline = (Get-Date).AddSeconds(60)
while (-not $word -and (Get-Date) -lt $deadline) {
    $word = [XrRot]::Get('Word.Application')
    if (-not $word) { Start-Sleep -Milliseconds 500 }
}
if (-not $word) { $proc | Stop-Process; throw 'Word konnte nicht gestartet werden.' }
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc  = $word.Documents.Add()
    $proj = $doc.VBProject
    $proj.Name = 'XRechnungAddIn'
    # Forms-Verweis vorab auf die System-DLL setzen, sonst landet ein Pfad im Benutzerprofil im Projekt
    [void]$proj.References.AddFromFile((Join-Path $env:WINDIR 'System32\FM20.DLL'))

    foreach ($c in $components) {
        $comp = $proj.VBComponents.Add($c.Type)
        $comp.Name = $c.Name
        $cm = $comp.CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        # VBE erwartet CRLF; ein abschließendes LF würde als ungültiges Zeichen übernommen
        $code = [IO.File]::ReadAllText((Join-Path $src $c.File), [Text.Encoding]::UTF8)
        $code = ($code -replace "`r?`n", "`r`n").TrimEnd()
        $cm.AddFromString($code)
        if ($c.Type -eq 3) { $comp.Properties.Item('Caption').Value = 'XRechnung erzeugen' }
    }

    # Verweise: Microsoft Office (IRibbonControl) und Microsoft Forms 2.0
    $names = @($proj.References | ForEach-Object { $_.Name })
    if ($names -notcontains 'Office') { [void]$proj.References.AddFromGuid('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 2, 0) }
    if ($names -notcontains 'MSForms') { [void]$proj.References.AddFromGuid('{0D452EE1-E08F-101A-852E-02608C4D0BB4}', 2, 0) }
    'Verweise: ' + ((@($proj.References | ForEach-Object { $_.Name })) -join ', ')

    # Export für manuelle Installation
    New-Item -ItemType Directory -Force $export | Out-Null
    Get-ChildItem $export -File | Remove-Item
    foreach ($c in $components) {
        $ext = @{ 1 = '.bas'; 2 = '.cls'; 3 = '.frm' }[$c.Type]
        $proj.VBComponents.Item($c.Name).Export((Join-Path $export ($c.Name + $ext)))
    }

    $doc.SaveAs2($Out, 15)      # wdFormatXMLTemplateMacroEnabled
    $doc.Close(0)
}
finally {
    $word.Quit(0)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}

# Menüband (customUI) in die .dotm einfügen
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($Out, 'Update')
try {
    $entry = $zip.CreateEntry('customUI/customUI.xml')
    $w = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding($false)))
    $w.Write([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'customUI.xml'), [Text.Encoding]::UTF8))
    $w.Close()

    $relsEntry = $zip.GetEntry('_rels/.rels')
    $r = New-Object IO.StreamReader($relsEntry.Open())
    $rels = [xml]$r.ReadToEnd()
    $r.Close()
    $ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $rel = $rels.CreateElement('Relationship', $ns)
    $rel.SetAttribute('Id', 'rIdXRechnungUI')
    $rel.SetAttribute('Type', 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility')
    $rel.SetAttribute('Target', 'customUI/customUI.xml')
    [void]$rels.DocumentElement.AppendChild($rel)
    $relsEntry.Delete()
    $e2 = $zip.CreateEntry('_rels/.rels')
    $w = New-Object IO.StreamWriter($e2.Open(), (New-Object Text.UTF8Encoding($false)))
    $rels.Save($w)
    $w.Close()

    # Metadaten entfernen: Autor, Bearbeitungszeit, Vorlage, Bearbeitungs-IDs (rsid)
    $utf8 = New-Object Text.UTF8Encoding($false)
    $clean = @{
        'docProps/core.xml' = { param($x) $x -replace '<dc:creator>[^<]*</dc:creator>', '<dc:creator></dc:creator>' `
                                             -replace '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', '<cp:lastModifiedBy></cp:lastModifiedBy>' `
                                             -replace '<dcterms:(created|modified) ([^>]*)>[^<]*<', '<dcterms:$1 $2>2026-01-01T00:00:00Z<' }
        'docProps/app.xml'  = { param($x) $x -replace '<Template>[^<]*</Template>', '' -replace '<TotalTime>[^<]*</TotalTime>', '' }
        'word/document.xml' = { param($x) $x -replace ' w:rsid\w*="[0-9A-F]+"', '' }
        'word/settings.xml' = { param($x) $x -replace '<w:rsids>.*?</w:rsids>', '' -replace '<w1[45]:docId [^>]*/>', '' }
    }
    foreach ($name in $clean.Keys) {
        $entry = $zip.GetEntry($name)
        if (-not $entry) { continue }
        $r = New-Object IO.StreamReader($entry.Open(), $utf8)
        $xml = & $clean[$name] $r.ReadToEnd()
        $r.Close()
        $entry.Delete()
        $w = New-Object IO.StreamWriter($zip.CreateEntry($name).Open(), $utf8)
        $w.Write($xml)
        $w.Close()
    }
}
finally { $zip.Dispose() }

"Erstellt: $Out"
