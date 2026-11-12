# Validates MathPaste OMML output inside the real Microsoft Word (desktop) through COM.
# Range.InsertXML uses the same Flat-OPC importer as Office.js Range.insertOoxml.
# Run: npm run test:word   (generates fixtures, then runs this script)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$fixtures = Join-Path $root 'tests\word-out'
# Word documents are written to a local temp folder (saving into a synced OneDrive folder can
# raise invisible dialogs in a hidden Word instance), then copied next to the fixtures.
$out = [string](Join-Path $env:TEMP ('mathpaste-word-test-' + (Get-Date -Format 'yyyyMMdd-HHmmss')))
New-Item -ItemType Directory -Force $out | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding($false)
function Load($name) { [IO.File]::ReadAllText((Join-Path $fixtures $name), $utf8) }

$failures = 0
function Check($cond, $label) {
  if ($cond) { Write-Host "  [OK]   $label" -ForegroundColor Green }
  else { Write-Host "  [FAIL] $label" -ForegroundColor Red; $script:failures++ }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$opened = @()
try {
  # ---------------------------------------------------------------- 1. three equations in one insertion
  Write-Host "1. Insertion simultanee des 3 equations de reference"
  $doc = $word.Documents.Add(); $opened += $doc
  $doc.Content.InsertXML((Load 'ref.xml'))
  Check ($doc.OMaths.Count -eq 3) "3 equations natives (OMaths.Count = $($doc.OMaths.Count))"
  $types = @()
  foreach ($i in 1..$doc.OMaths.Count) {
    $m = $doc.OMaths.Item($i)
    $t = @(); foreach ($f in $m.Functions) { $t += $f.Type }
    $types += ,$t
    Check ($m.Type -eq 0) "equation $i en mode affichage (display)"
  }
  # wdOMathFunction: Acc=1 Bar=2 Box=3 BorderBox=4 Delim=5 EqArray=6 Frac=7 Func=8 GroupChar=9 LimLow=10 LimUpp=11 Mat=12 Nary=13 Phantom=14 ScrPre=15 Rad=16 ScrSub=17 ScrSubSup=18 ScrSup=19 Text=20
  Check (($types[0] -contains 17) -and ($types[0] -contains 7)) "C_m1 : indice (ScrSub) + fraction (Frac) au premier niveau"
  $frac = $doc.OMaths.Item(1).Functions | Where-Object { $_.Type -eq 7 } | Select-Object -First 1
  $denTypes = @(); foreach ($f in $frac.Frac.Den.Functions) { $denTypes += $f.Type }
  Check ($denTypes -contains 16) "C_m1 : racine carree (Rad) au denominateur"
  $numTypes = @(); foreach ($f in $frac.Frac.Num.Functions) { $numTypes += $f.Type }
  Check (($numTypes -contains 19) -and ($numTypes -contains 17)) "C_m1 : exposant r^2 et indice epsilon_0 au numerateur"
  Check ($types[2] -contains 5) "C_m3 : parentheses extensibles (Delim)"
  $delim = $doc.OMaths.Item(3).Functions | Where-Object { $_.Type -eq 5 } | Select-Object -First 1
  $inner = @(); foreach ($f in $delim.Delim.E.Item(1).Functions) { $inner += $f.Type }
  Check ($inner -contains 8) "C_m3 : ln est une fonction native (Func)"

  # Linear form of each equation (what Word's own editor shows when linearised)
  foreach ($i in 1..3) {
    $r = $doc.OMaths.Item($i).Range
    Write-Host ("     lineaire {0}: {1}" -f $i, $r.Text)
  }
  $doc.Close(0)
  # Fresh document for save / export (COM sub-objects released)
  $docx = [string](Join-Path $out 'ref.docx'); $pdf = [string](Join-Path $out 'ref.pdf')
  $doc = $word.Documents.Add(); $opened += $doc
  $doc.Content.InsertXML((Load 'ref.xml'))
  $doc.ExportAsFixedFormat($pdf, 17)
  $doc.SaveAs2($docx, 16)
  $doc.Close(0)
  # Reopen: the document must still contain 3 native, editable equations
  $doc = $word.Documents.Open($docx, $false, $false, $false); $opened += $doc
  Check ($doc.OMaths.Count -eq 3) "document sauvegarde puis rouvert : 3 equations"
  # Edit one equation individually: linearise then build up again
  $doc.OMaths.Item(2).Linearize()
  $lin = $doc.OMaths.Item(2).Range.Text
  $doc.OMaths.Item(2).BuildUp()
  Check ($lin -match "_" -and $lin -match "=" -and $doc.OMaths.Count -eq 3) "equation 2 editable individuellement (linearisee: $lin)"
  $doc.Close(0)

  # ---------------------------------------------------------------- 2. coverage sheet
  Write-Host "2. Feuille de couverture LaTeX"
  $doc = $word.Documents.Add(); $opened += $doc
  $doc.Content.InsertXML((Load 'coverage.xml'))
  Check ($doc.OMaths.Count -eq 12) "12 equations de couverture (OMaths.Count = $($doc.OMaths.Count))"
  $doc.ExportAsFixedFormat([string](Join-Path $out 'coverage.pdf'), 17)
  $doc.SaveAs2([string](Join-Path $out 'coverage.docx'), 16)
  $doc.Close(0)

  # ---------------------------------------------------------------- 3. replace ONLY the 2nd expression
  Write-Host "3. Scenario : selection et conversion de la 2e expression uniquement"
  $latex = (Load 'ref-latex.txt') -split "`n"
  $doc = $word.Documents.Add(); $opened += $doc
  $doc.Content.Text = "Texte avant`r" + ($latex -join "`r") + "`rTexte apres"
  $before = @(); foreach ($p in $doc.Paragraphs) { $before += $p.Range.Text }
  $target = $doc.Paragraphs.Item(3).Range
  $target.MoveEnd(1, -1) | Out-Null   # exclude the paragraph mark, like the add-in does
  $target.Select()
  $word.Selection.Range.InsertXML((Load 'sel2.xml'))
  $after = @(); foreach ($p in $doc.Paragraphs) { $after += $p.Range.Text }
  Check ($doc.Paragraphs.Count -eq $before.Count) "nombre de paragraphes inchange ($($doc.Paragraphs.Count))"
  Check ($doc.OMaths.Count -eq 1) "une seule equation creee"
  Check ($after[1] -eq $before[1]) "expression 1 inchangee"
  Check ($after[3] -eq $before[3]) "expression 3 inchangee"
  Check ($after[0] -eq $before[0] -and $after[4] -eq $before[4]) "texte avant/apres inchange"
  Check ($doc.OMaths.Item(1).Range.Paragraphs.Item(1).Range.Start -eq $doc.Paragraphs.Item(3).Range.Start) "l'equation est dans le 3e paragraphe"
  $doc.ExportAsFixedFormat([string](Join-Path $out 'scenario.pdf'), 17)
  $doc.Close(0)

  # ---------------------------------------------------------------- 4. inline replacement inside text
  Write-Host "4. Texte mixte : seule la formule est remplacee"
  $doc = $word.Documents.Add(); $opened += $doc
  $prefix = "La capacite est "; $suffix = " ou r est le rayon."
  $doc.Content.Text = $prefix + $latex[1] + $suffix
  $doc.Paragraphs.Item(1).Range.Font.Size = 14
  $r = $doc.Range($prefix.Length, $prefix.Length + $latex[1].Length)
  $r.InsertXML((Load 'inline.xml'))
  $txt = $doc.Paragraphs.Item(1).Range.Text
  Check ($doc.OMaths.Count -eq 1 -and $doc.Paragraphs.Count -eq 1) "1 equation, 1 paragraphe"
  Check ($txt.StartsWith($prefix) -and $txt.TrimEnd("`r").EndsWith($suffix)) "texte environnant conserve"
  Check ($doc.OMaths.Item(1).Type -eq 1) "equation en ligne (inline)"
  $doc.ExportAsFixedFormat([string](Join-Path $out 'inline.pdf'), 17)
  $doc.Close(0)
}
finally {
  foreach ($d in $opened) { try { $d.Close(0) } catch {} }
  if ($word.Documents.Count -eq 0) { $word.Quit() }
  Copy-Item (Join-Path $out '*.pdf'), (Join-Path $out '*.docx') $fixtures -Force -ErrorAction SilentlyContinue
}
if ($failures -gt 0) { Write-Host "$failures verification(s) en echec" -ForegroundColor Red; exit 1 }
Write-Host "Toutes les verifications Word sont passees. PDF dans tests\word-out" -ForegroundColor Green
