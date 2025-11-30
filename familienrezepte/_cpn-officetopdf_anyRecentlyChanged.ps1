# Ensure we can run everything
Set-ExecutionPolicy Bypass -Scope Process -Force

Write-Output "PowerShell Version: $($PSVersionTable.PSVersion)"
Import-Module "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1" -ErrorAction SilentlyContinue
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8 # Set code page to UTF-8

$PSNativeCommandArgumentPassing = 'Legacy'
if (-not $PSScriptRoot) { $PSScriptRoot = Get-Location }
Set-Location -LiteralPath $PSScriptRoot
Write-Host "== $PSScriptRoot =="

# *********************************************************************
# author: Christoph P. Neumann (patched)
# *********************************************************************
$recursive = $false

$patterns = @('*.pptx','*.docx','*.xlsx')
foreach ($pattern in $patterns) {
    $gciParams = @{
        LiteralPath = $PSScriptRoot
        Filter      = $pattern
    }
    if ($recursive) { $gciParams.Recurse = $true }
    Get-ChildItem @gciParams | Sort-Object Name | ForEach-Object {
        $srcFilePath = $_.FullName
        $tgtFilePath = [System.IO.Path]::ChangeExtension($srcFilePath, ".pdf")
        $needsAction = $false

        if (!(Test-Path -LiteralPath $tgtFilePath)) {
            $needsAction = $true
        } else {
            $srcFileLastWriteTime = (Get-Item -LiteralPath $srcFilePath).LastWriteTime
            $tgtFileLastWriteTime = (Get-Item -LiteralPath $tgtFilePath).LastWriteTime
            if ($srcFileLastWriteTime -gt $tgtFileLastWriteTime) {
                $needsAction = $true
            }
        }

        if ($needsAction) {
            $srcDisplay = Resolve-Path -LiteralPath $srcFilePath -Relative
            $tgtDisplay = [System.IO.Path]::ChangeExtension($srcDisplay, ".pdf")
            Write-Host "🅢: $srcDisplay" -ForegroundColor Cyan
            Write-Host "🅣: $tgtDisplay" -ForegroundColor Cyan

            # Remove any stale target
            Remove-Item -LiteralPath $tgtFilePath -Force -ErrorAction SilentlyContinue

            # Prepend \\?\ for long path support
            $srcSafe = "\\?\$srcFilePath"
            $tgtSafe = "\\?\$tgtFilePath"

            # cognidox OfficeToPDF: https://github.com/cognidox/OfficeToPDF
            # Build a robust argument array and avoid fragile quoting/backticks
            $otfArgs = @(
                $srcSafe,
                $tgtSafe,
                "/noquit",                 # do not exit running Office applications
                "/hidden",                 # attempts to minimise the Office application when converting
                "/readonly",               # attempts to open the source document in read-only mode
                "/print",                  # create high-quality PDFs optimised for print
                #"/pdfa",                  # produce ISO 19005-1 (PDF/A) compliant PDFs
                "/excludeprops",           # do not include properties in generated PDF
                "/excludetags",            # do not include tags in generated PDF
                "/pdf_clean_meta","full",  # removes author, keywords, creator, subject, and the title
                "/pdf_layout","single",    # single = show pages one at a time
                "/pdf_page_mode","none"    # none = the PDF will open without the navigation bar visible
            )

            # ensure case-insensitive extension checks
            $ext = $_.Extension.ToLowerInvariant()

            if ($ext -in '.xlsx', '.pptx') {
                # NOT for WORD due to Font-Post-Processing via GhostScript:
                $otfArgs += @(
                    "/pdf_owner_pass","KlzPhCIBlBeEE2hO5NcM",
                    "/pdf_restrict_modify",  # prevent modification without the owner password
                    "/pdf_restrict_assembly" # prevent rotation, removal or insertion of pages without the owner password
                )
            }
            if ($ext -eq '.pptx') {
                # For PPTX, prevent extractions:
                $otfArgs += @(
                    "/pdf_restrict_extraction",               # prevent content extraction without the owner password
                    "/pdf_restrict_accessibility_extraction"  # prevent all content extraction without the owner password
                )
            }

            # Work on temp copies (safer when source is on network drives)
            $tempSrc = Join-Path $env:TEMP ([IO.Path]::GetFileName($srcFilePath))
            $tempTgt = [System.IO.Path]::ChangeExtension($tempSrc, ".pdf")

            try {
                Copy-Item -LiteralPath $srcFilePath -Destination $tempSrc -Force -ErrorAction Stop

                # Prepare arguments for calling officetopdf.exe with the temp paths:
                # Replace the first two entries of $otfArgs (originally $srcSafe/$tgtSafe) with $tempSrc/$tempTgt
                $callArgs = $otfArgs.Clone()
                $callArgs[0] = $tempSrc
                $callArgs[1] = $tempTgt

                # Call OfficeToPDF reliably (no fragile embedded quotes)
                $exe = "officetopdf.exe"
                Write-Host "→ Running: $exe $($callArgs -join ' ')" -ForegroundColor DarkGray
                & $exe @callArgs
                $last = $LASTEXITCODE

                if ($last -eq 0 -and (Test-Path $tempTgt)) {
                    Move-Item -LiteralPath $tempTgt -Destination $tgtFilePath -Force
                    Write-Host "✓ Conversion succeeded → $tgtDisplay" -ForegroundColor Green

                    if ($ext -eq '.docx') {
                        # Ghostscript optimization (fonts)
                        $optimizedPdf = "$tgtFilePath.gsopt.pdf"

                        $gsExe = "gswin64c"
                        $gsArgs = @(
                            "-sDEVICE=pdfwrite",
                            "-dCompatibilityLevel=1.6",
                            "-dPDFSETTINGS=/printer",
                            # "-dEmbedAllFonts=true", # Only embed fonts that are actually used so NO EmbedAllFonts
                            "-dSubsetFonts=true",
                            "-dCompressFonts=true",
                            "-dNOPAUSE",
                            "-dQUIET",
                            "-dBATCH",
                            "-sOutputFile=$optimizedPdf",
                            $tgtFilePath
                        )
                        try {
                            Write-Host "→ Running Ghostscript to optimize fonts..." -ForegroundColor DarkGray
                            & $gsExe @gsArgs
                        }
                        catch {
                            Write-Warning "Ghostscript execution failed: $($_.Exception.Message)"
                        }
                        if (Test-Path $optimizedPdf) {
                            try {
                                Move-Item -Force $optimizedPdf $tgtFilePath
                                Write-Host "✓ Ghostscript optimized fonts → $tgtDisplay" -ForegroundColor Yellow
                            } catch {
                                Write-Warning "Failed to replace original PDF with Ghostscript-optimized PDF: $($_.Exception.Message)"
                            }
                        } else {
                            Write-Warning "Ghostscript did not produce optimized PDF for $srcDisplay"
                        }

                        $protectedPdf = "$tgtFilePath.protected.pdf"
                        Write-Host "→ Running: pdftk (set permissions)" -ForegroundColor DarkGray
                        try {
                            & pdftk $tgtFilePath output $protectedPdf owner_pw "KlzPhCIBlBeEE2hO5NcM" allow "Printing" "DegradedPrinting" "CopyContents" "ScreenReaders" "ModifyAnnotations"
                        } catch {
                            Write-Warning "pdftk execution failed: $($_.Exception.Message)"
                        }
                        if (Test-Path $protectedPdf) {
                            try {
                                Move-Item -Force $protectedPdf $tgtFilePath
                                Write-Host "✓ PDF permissions set → $tgtDisplay" -ForegroundColor Yellow
                            } catch {
                                Write-Warning "Failed to move pdftk output into place: $($_.Exception.Message)"
                            }
                        } else {
                            Write-Warning "❌ pdftk failed to set permissions for $srcDisplay"
                        }
                    }

                    $srcSize = (Get-Item -LiteralPath $srcFilePath).Length
                    $tgtSize = (Get-Item -LiteralPath $tgtFilePath).Length
                    if ($tgtSize -gt 2 * $srcSize) {
                        Write-Warning "⚠ The PDF '$tgtDisplay' is more than double the size of the source file ($([math]::Round($srcSize/1KB,0)) KB → $([math]::Round($tgtSize/1KB,0)) KB)."
                        if ($_.Extension -eq '.pptx') {
                            Write-Warning "⚠ This might be caused by Emojis in Word documents (Segoe UI Emoji) or format templates with hidden font dependencies"
                        }
                    }

                } else {
                    Write-Error "❌ Conversion failed for $srcDisplay (exit code: $last)"
                    if (Test-Path $tempTgt) {
                        Write-Host "Temp target exists at $tempTgt — keeping for inspection." -ForegroundColor Yellow
                    }
                }
            }
            catch {
                Write-Error "Exception during conversion of $srcDisplay : $($_.Exception.Message)"
            }
            finally {
                # Cleanup temp files if still present
                Remove-Item -LiteralPath $tempSrc -Force -ErrorAction SilentlyContinue
                Remove-Item -LiteralPath $tempTgt -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

Write-Host "Press ENTER to continue..."
cmd /c Pause | Out-Null
