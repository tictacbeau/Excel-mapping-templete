<#
.SYNOPSIS
    Remittance Allocation Engine — Build Script
    Creates a deployable .xlsm workbook from VBA source files.

.DESCRIPTION
    This PowerShell script:
    1. Creates a new Excel XLSM workbook
    2. Imports all VBA modules (.bas), class modules (.cls), and forms (.frm)
       from the src/vba directory
    3. Injects the CustomUI ribbon XML
    4. Configures workbook properties
    5. Saves the final .xlsm file

    Requirements:
    - Windows 10/11 with Excel 2016+ installed
    - PowerShell 5.0+
    - Excel must be available (COM automation)

.PARAMETER OutputPath
    Path for the output .xlsm file (default: ../dist/RemittanceAllocationEngine.xlsm)

.PARAMETER SkipRibbon
    Skip CustomUI injection (requires Office Open XML tools)

.EXAMPLE
    .\build.ps1
    .\build.ps1 -OutputPath "C:\MyTools\RemittanceEngine.xlsm"
#>

param(
    [string]$OutputPath = "",
    [switch]$SkipRibbon = $false,
    [switch]$Verbose = $false
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ===========================================================
#  CONFIGURATION
# ===========================================================
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot  = Split-Path -Parent $ScriptDir
$SrcVba    = Join-Path $RepoRoot "src\vba"
$SrcUI     = Join-Path $RepoRoot "src\customui"
$DistDir   = Join-Path $RepoRoot "dist"

if ($OutputPath -eq "") {
    $OutputPath = Join-Path $DistDir "RemittanceAllocationEngine.xlsm"
}

$ModulesDir = Join-Path $SrcVba "modules"
$ClassesDir = Join-Path $SrcVba "classes"
$FormsDir   = Join-Path $SrcVba "forms"
$CustomUiXml = Join-Path $SrcUI "customUI14.xml"

# ===========================================================
#  HELPERS
# ===========================================================
function Write-Step([string]$msg) {
    Write-Host "  [BUILD] $msg" -ForegroundColor Cyan
}

function Write-Success([string]$msg) {
    Write-Host "  [OK]    $msg" -ForegroundColor Green
}

function Write-Warn([string]$msg) {
    Write-Host "  [WARN]  $msg" -ForegroundColor Yellow
}

function Write-Fail([string]$msg) {
    Write-Host "  [FAIL]  $msg" -ForegroundColor Red
}

# ===========================================================
#  MAIN BUILD
# ===========================================================
Write-Host ""
Write-Host "=============================================" -ForegroundColor White
Write-Host " Remittance Allocation Engine — Build Script" -ForegroundColor White
Write-Host "=============================================" -ForegroundColor White
Write-Host ""

# Ensure dist directory exists
if (!(Test-Path $DistDir)) {
    New-Item -ItemType Directory -Path $DistDir | Out-Null
    Write-Step "Created dist directory: $DistDir"
}

# Delete existing output file
if (Test-Path $OutputPath) {
    Remove-Item $OutputPath -Force
    Write-Step "Removed existing output: $OutputPath"
}

# ===========================================================
#  STEP 1: Launch Excel via COM
# ===========================================================
Write-Step "Starting Excel..."
$excel = $null
$wb    = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    Write-Success "Excel started (version $($excel.Version))"

    # ===========================================================
    #  STEP 2: Create new XLSM workbook
    # ===========================================================
    Write-Step "Creating new macro-enabled workbook..."
    $wb = $excel.Workbooks.Add()
    Write-Success "Workbook created"

    # ===========================================================
    #  STEP 3: Import VBA Modules (.bas)
    # ===========================================================
    Write-Step "Importing VBA modules..."
    $imported = 0
    $moduleFiles = @(
        "modConstants.bas",
        "modUtils.bas",
        "modFileConverter.bas",
        "modPayorRegistry.bas",
        "modTemplateEngine.bas",
        "modDataLookup.bas",
        "modMain.bas",
        "modRibbon.bas"
    )

    foreach ($fname in $moduleFiles) {
        $fpath = Join-Path $ModulesDir $fname
        if (Test-Path $fpath) {
            $wb.VBProject.VBComponents.Import($fpath) | Out-Null
            $imported++
            if ($Verbose) { Write-Success "  Imported module: $fname" }
        } else {
            Write-Warn "Module not found: $fpath"
        }
    }
    Write-Success "Imported $imported VBA module(s)"

    # ===========================================================
    #  STEP 4: Import Class Modules (.cls)
    # ===========================================================
    Write-Step "Importing class modules..."
    $classFiles = @(
        "clsCellFormat.cls",
        "clsPayor.cls",
        "clsRemittance.cls",
        "clsTemplate.cls"
    )
    $importedCls = 0
    foreach ($fname in $classFiles) {
        $fpath = Join-Path $ClassesDir $fname
        if (Test-Path $fpath) {
            $wb.VBProject.VBComponents.Import($fpath) | Out-Null
            $importedCls++
            if ($Verbose) { Write-Success "  Imported class: $fname" }
        } else {
            Write-Warn "Class not found: $fpath"
        }
    }
    Write-Success "Imported $importedCls class module(s)"

    # ===========================================================
    #  STEP 5: Import UserForms (.frm)
    # ===========================================================
    Write-Step "Importing UserForms..."
    $formFiles = @(
        "frmProcessRemittance.frm",
        "frmPayorWizard.frm",
        "frmTemplateBuilder.frm",
        "frmPayorManager.frm",
        "frmDataRefresh.frm"
    )
    $importedFrm = 0
    foreach ($fname in $formFiles) {
        $fpath = Join-Path $FormsDir $fname
        if (Test-Path $fpath) {
            $wb.VBProject.VBComponents.Import($fpath) | Out-Null
            $importedFrm++
            if ($Verbose) { Write-Success "  Imported form: $fname" }
        } else {
            Write-Warn "Form not found: $fpath"
        }
    }
    Write-Success "Imported $importedFrm UserForm(s)"

    # ===========================================================
    #  STEP 6: Configure ThisWorkbook VBA event handlers
    # ===========================================================
    Write-Step "Configuring ThisWorkbook event handlers..."
    $thisWbCode = $wb.VBProject.VBComponents("ThisWorkbook").CodeModule

    $openHandler = @"
Private Sub Workbook_Open()
    modMain.Application_Initialize
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If Not ThisWorkbook.Saved Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox("Save changes to the Remittance Allocation Engine workbook?", _
                        vbYesNoCancel + vbQuestion, "Remittance Allocation Engine")
        If answer = vbYes Then
            ThisWorkbook.Save
        ElseIf answer = vbCancel Then
            Cancel = True
        End If
    End If
    On Error GoTo 0
End Sub
"@

    $thisWbCode.InsertLines(1, $openHandler)
    Write-Success "Workbook event handlers configured"

    # ===========================================================
    #  STEP 7: Set workbook properties
    # ===========================================================
    Write-Step "Setting workbook properties..."
    $wb.CustomDocumentProperties | Out-Null
    try {
        $wb.BuiltinDocumentProperties("Title").Value = "Remittance Allocation Engine"
        $wb.BuiltinDocumentProperties("Subject").Value = "AR Remittance Allocation Automation"
        $wb.BuiltinDocumentProperties("Company").Value = ""
        $wb.BuiltinDocumentProperties("Comments").Value = "Version 1.0 | Built $(Get-Date -Format 'yyyy-MM-dd')"
    } catch {
        Write-Warn "Could not set all document properties (non-critical)"
    }
    Write-Success "Workbook properties set"

    # ===========================================================
    #  STEP 8: Save as XLSM
    # ===========================================================
    Write-Step "Saving as XLSM: $OutputPath"
    # xlOpenXMLWorkbookMacroEnabled = 52
    $wb.SaveAs($OutputPath, 52)
    Write-Success "Workbook saved"

    # ===========================================================
    #  STEP 9: Inject CustomUI ribbon (post-save via Open XML)
    # ===========================================================
    if (!$SkipRibbon -and (Test-Path $CustomUiXml)) {
        Write-Step "Injecting CustomUI ribbon XML..."
        $wb.Close($false)
        $wb = $null
        $success = Inject-CustomUI -XlsmPath $OutputPath -CustomUiPath $CustomUiXml
        if ($success) {
            Write-Success "Ribbon XML injected successfully"
        } else {
            Write-Warn "Ribbon injection skipped (requires .NET ZIP support)"
        }
    } else {
        if ($wb) { $wb.Close($false); $wb = $null }
        Write-Warn "CustomUI injection skipped (SkipRibbon flag or file not found)"
    }

} finally {
    # Cleanup COM objects
    if ($wb) {
        try { $wb.Close($false) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# ===========================================================
#  INJECT CUSTOMUI (using .NET System.IO.Compression)
# ===========================================================
function Inject-CustomUI {
    param([string]$XlsmPath, [string]$CustomUiPath)

    try {
        Add-Type -AssemblyName "System.IO.Compression.FileSystem"
        $xmlContent = Get-Content $CustomUiPath -Raw -Encoding UTF8

        # Work on a temp copy
        $tmpPath = [System.IO.Path]::GetTempFileName() + ".xlsm"
        Copy-Item $XlsmPath $tmpPath -Force

        $zip = [System.IO.Compression.ZipFile]::Open($tmpPath,
               [System.IO.Compression.ZipArchiveMode]::Update)

        # Add customUI folder and file
        $entry = $zip.CreateEntry("customUI/customUI14.xml")
        $stream = $entry.Open()
        $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
        $writer.Write($xmlContent)
        $writer.Close()
        $stream.Close()

        # Update [Content_Types].xml to include customUI
        $ctEntry = $zip.GetEntry("[Content_Types].xml")
        $ctStream = $ctEntry.Open()
        $reader = New-Object System.IO.StreamReader($ctStream)
        $ctXml = $reader.ReadToEnd()
        $reader.Close()
        $ctStream.Close()

        if ($ctXml -notlike "*customUI*") {
            $ctXml = $ctXml -replace "</Types>",
                '<Override PartName="/customUI/customUI14.xml" ContentType="application/vnd.ms-office.activeX+xml"/></Types>'
            $ctStream = $ctEntry.Open()
            $ctStream.SetLength(0)
            $ctWriter = New-Object System.IO.StreamWriter($ctStream, [System.Text.Encoding]::UTF8)
            $ctWriter.Write($ctXml)
            $ctWriter.Close()
            $ctStream.Close()
        }

        # Add relationship for customUI in _rels/.rels
        $relsEntry = $zip.GetEntry("_rels/.rels")
        if ($relsEntry) {
            $relsStream = $relsEntry.Open()
            $relsReader = New-Object System.IO.StreamReader($relsStream)
            $relsXml = $relsReader.ReadToEnd()
            $relsReader.Close()
            $relsStream.Close()

            if ($relsXml -notlike "*customUI*") {
                $relsXml = $relsXml -replace "</Relationships>",
                    '<Relationship Id="rIdCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>'
                $relsStream = $relsEntry.Open()
                $relsStream.SetLength(0)
                $relsWriter = New-Object System.IO.StreamWriter($relsStream, [System.Text.Encoding]::UTF8)
                $relsWriter.Write($relsXml)
                $relsWriter.Close()
                $relsStream.Close()
            }
        }

        $zip.Dispose()
        Copy-Item $tmpPath $XlsmPath -Force
        Remove-Item $tmpPath -Force
        return $true

    } catch {
        Write-Warn "CustomUI injection error: $_"
        return $false
    }
}

# ===========================================================
#  RESULTS
# ===========================================================
Write-Host ""
Write-Host "=============================================" -ForegroundColor White
if (Test-Path $OutputPath) {
    $size = (Get-Item $OutputPath).Length
    Write-Host " BUILD SUCCESSFUL" -ForegroundColor Green
    Write-Host " Output: $OutputPath" -ForegroundColor White
    Write-Host " Size:   $([math]::Round($size/1024, 1)) KB" -ForegroundColor White
} else {
    Write-Host " BUILD FAILED — output file not created" -ForegroundColor Red
}
Write-Host "=============================================" -ForegroundColor White
Write-Host ""
