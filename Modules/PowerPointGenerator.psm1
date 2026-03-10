#Requires -Version 5.1
<#
.SYNOPSIS
    PowerPoint Generator Module for Conditional Access Policy Documenter

.DESCRIPTION
    Generates PowerPoint (.pptx) presentations using Open XML format without
    requiring third-party libraries. A PPTX file is a ZIP archive containing
    XML files following the Office Open XML specification.
#>

# Add required .NET types
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Script-level variables
$script:SlideCount = 0
$script:SlideRels = @()
$script:Theme = @{
    PrimaryColor   = "0078D4"
    EnabledColor   = "107C10"
    DisabledColor  = "A80000"
    ReportOnlyColor = "FFB900"
    TextColor      = "333333"
    LightGray      = "F5F5F5"
    White          = "FFFFFF"
}

function New-PptxDocument {
    <#
    .SYNOPSIS
        Creates a new PowerPoint document structure

    .PARAMETER Title
        The title for the presentation

    .PARAMETER Author
        The author name for the presentation metadata
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [string]$Author = "Conditional Access Documenter"
    )

    $script:SlideCount = 0
    $script:SlideRels = @()

    return @{
        Title    = $Title
        Author   = $Author
        Created  = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
        Slides   = @()
    }
}

function Add-TitleSlide {
    <#
    .SYNOPSIS
        Adds a title slide to the presentation

    .PARAMETER Document
        The presentation document object

    .PARAMETER Title
        The main title text

    .PARAMETER Subtitle
        The subtitle text
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Document,

        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [string]$Subtitle = ""
    )

    $script:SlideCount++

    $slide = @{
        Type     = "Title"
        Number   = $script:SlideCount
        Title    = $Title
        Subtitle = $Subtitle
    }

    $Document.Slides += $slide
    return $Document
}

function Add-PolicySlide {
    <#
    .SYNOPSIS
        Adds a policy detail slide to the presentation

    .PARAMETER Document
        The presentation document object

    .PARAMETER Policy
        The parsed policy object
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Document,

        [Parameter(Mandatory = $true)]
        [object]$Policy
    )

    $script:SlideCount++

    $slide = @{
        Type   = "Policy"
        Number = $script:SlideCount
        Policy = $Policy
    }

    $Document.Slides += $slide
    return $Document
}

function Add-SummarySlide {
    <#
    .SYNOPSIS
        Adds a summary slide to the presentation

    .PARAMETER Document
        The presentation document object

    .PARAMETER Policies
        Array of all policies for statistics
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Document,

        [Parameter(Mandatory = $true)]
        [array]$Policies
    )

    $script:SlideCount++

    $stats = @{
        Total      = $Policies.Count
        Enabled    = ($Policies | Where-Object { $_.StateRaw -eq "enabled" }).Count
        Disabled   = ($Policies | Where-Object { $_.StateRaw -eq "disabled" }).Count
        ReportOnly = ($Policies | Where-Object { $_.StateRaw -eq "enabledForReportingButNotEnforced" }).Count
    }

    $slide = @{
        Type   = "Summary"
        Number = $script:SlideCount
        Stats  = $stats
    }

    $Document.Slides += $slide
    return $Document
}

function Save-PptxDocument {
    <#
    .SYNOPSIS
        Saves the presentation to a .pptx file

    .PARAMETER Document
        The presentation document object

    .PARAMETER Path
        The output file path
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Document,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    # Ensure output directory exists
    $outputDir = Split-Path -Parent $Path
    if ($outputDir -and -not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    # Create temp directory for building the PPTX structure
    $tempDir = Join-Path $env:TEMP "pptx_$(Get-Random)"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    try {
        # Create directory structure
        $dirs = @(
            "_rels",
            "docProps",
            "ppt",
            "ppt/_rels",
            "ppt/slideLayouts",
            "ppt/slideLayouts/_rels",
            "ppt/slideMasters",
            "ppt/slideMasters/_rels",
            "ppt/slides",
            "ppt/slides/_rels",
            "ppt/theme"
        )

        foreach ($dir in $dirs) {
            New-Item -ItemType Directory -Path (Join-Path $tempDir $dir) -Force | Out-Null
        }

        # Write [Content_Types].xml
        Write-ContentTypesXml -TempDir $tempDir -SlideCount $Document.Slides.Count

        # Write _rels/.rels
        Write-RootRelsXml -TempDir $tempDir

        # Write docProps/app.xml
        Write-AppXml -TempDir $tempDir -Document $Document

        # Write docProps/core.xml
        Write-CoreXml -TempDir $tempDir -Document $Document

        # Write ppt/presentation.xml
        Write-PresentationXml -TempDir $tempDir -SlideCount $Document.Slides.Count

        # Write ppt/_rels/presentation.xml.rels
        Write-PresentationRelsXml -TempDir $tempDir -SlideCount $Document.Slides.Count

        # Write theme
        Write-ThemeXml -TempDir $tempDir

        # Write slide master and layout
        Write-SlideMasterXml -TempDir $tempDir
        Write-SlideLayoutXml -TempDir $tempDir

        # Write each slide
        foreach ($slide in $Document.Slides) {
            Write-SlideXml -TempDir $tempDir -Slide $slide
        }

        # Create the PPTX file (ZIP)
        if (Test-Path $Path) {
            Remove-Item $Path -Force
        }

        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $Path)

        Write-Host "PowerPoint saved to: $Path" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to create PowerPoint: $_"
        return $false
    }
    finally {
        # Cleanup temp directory
        if (Test-Path $tempDir) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Write-ContentTypesXml {
    param([string]$TempDir, [int]$SlideCount)

    $slideParts = ""
    for ($i = 1; $i -le $SlideCount; $i++) {
        $slideParts += "`n  <Override PartName=`"/ppt/slides/slide$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.presentationml.slide+xml`"/>"
    }

    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>$slideParts
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"@

    $xml | Out-File -LiteralPath (Join-Path $TempDir "[Content_Types].xml") -Encoding UTF8
}

function Write-RootRelsXml {
    param([string]$TempDir)

    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@

    $xml | Out-File -FilePath (Join-Path $TempDir "_rels\.rels") -Encoding UTF8
}

function Write-AppXml {
    param([string]$TempDir, [hashtable]$Document)

    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Conditional Access Documenter</Application>
  <Slides>$($Document.Slides.Count)</Slides>
  <Company>$([System.Security.SecurityElement]::Escape($Document.Author))</Company>
</Properties>
"@

    $xml | Out-File -FilePath (Join-Path $TempDir "docProps\app.xml") -Encoding UTF8
}

function Write-CoreXml {
    param([string]$TempDir, [hashtable]$Document)

    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>$([System.Security.SecurityElement]::Escape($Document.Title))</dc:title>
  <dc:creator>$([System.Security.SecurityElement]::Escape($Document.Author))</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">$($Document.Created)</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$($Document.Created)</dcterms:modified>
</cp:coreProperties>
"@

    $xml | Out-File -FilePath (Join-Path $TempDir "docProps\core.xml") -Encoding UTF8
}

function Write-PresentationXml {
    param([string]$TempDir, [int]$SlideCount)

    $slideIdList = ""
    for ($i = 1; $i -le $SlideCount; $i++) {
        $slideId = 255 + $i
        $slideIdList += "`n      <p:sldId id=`"$slideId`" r:id=`"rId$($i + 2)`"/>"
    }

    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>$slideIdList
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
"@

    $xml | Out-File -FilePath (Join-Path $TempDir "ppt\presentation.xml") -Encoding UTF8
}

function Write-PresentationRelsXml {
    param([string]$TempDir, [int]$SlideCount)

    $slideRels = ""
    for ($i = 1; $i -le $SlideCount; $i++) {
        $slideRels += "`n  <Relationship Id=`"rId$($i + 2)`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`" Target=`"slides/slide$i.xml`"/>"
    }

    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>$slideRels
</Relationships>
"@

    $xml | Out-File -FilePath (Join-Path $TempDir "ppt\_rels\presentation.xml.rels") -Encoding UTF8
}

function Write-ThemeXml {
    param([string]$TempDir)

    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="$($script:Theme.PrimaryColor)"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Segoe UI Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Segoe UI"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>
"@

    $xml | Out-File -FilePath (Join-Path $TempDir "ppt\theme\theme1.xml") -Encoding UTF8
}

function Write-SlideMasterXml {
    param([string]$TempDir)

    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>
"@

    $xml | Out-File -FilePath (Join-Path $TempDir "ppt\slideMasters\slideMaster1.xml") -Encoding UTF8

    # Write slide master rels
    $relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
"@

    $relsXml | Out-File -FilePath (Join-Path $TempDir "ppt\slideMasters\_rels\slideMaster1.xml.rels") -Encoding UTF8
}

function Write-SlideLayoutXml {
    param([string]$TempDir)

    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
"@

    $xml | Out-File -FilePath (Join-Path $TempDir "ppt\slideLayouts\slideLayout1.xml") -Encoding UTF8

    # Write slide layout rels
    $relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
"@

    $relsXml | Out-File -FilePath (Join-Path $TempDir "ppt\slideLayouts\_rels\slideLayout1.xml.rels") -Encoding UTF8
}

function Write-SlideXml {
    param([string]$TempDir, [hashtable]$Slide)

    $slideNum = $Slide.Number
    $slideXml = ""

    switch ($Slide.Type) {
        "Title" {
            $slideXml = Get-TitleSlideXml -Slide $Slide
        }
        "Policy" {
            $slideXml = Get-PolicySlideXml -Slide $Slide
        }
        "Summary" {
            $slideXml = Get-SummarySlideXml -Slide $Slide
        }
    }

    $slideXml | Out-File -FilePath (Join-Path $TempDir "ppt\slides\slide$slideNum.xml") -Encoding UTF8

    # Write slide rels
    $relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>
"@

    $relsXml | Out-File -FilePath (Join-Path $TempDir "ppt\slides\_rels\slide$slideNum.xml.rels") -Encoding UTF8
}

function Get-TitleSlideXml {
    param([hashtable]$Slide)

    $title = [System.Security.SecurityElement]::Escape($Slide.Title)
    $subtitle = [System.Security.SecurityElement]::Escape($Slide.Subtitle)

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="$($script:Theme.PrimaryColor)"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="609600" y="2286000"/>
            <a:ext cx="10972800" cy="1524000"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="4800" b="1">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                <a:latin typeface="Segoe UI Light"/>
              </a:rPr>
              <a:t>$title</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Subtitle"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="609600" y="4000000"/>
            <a:ext cx="10972800" cy="762000"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="2400">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                <a:latin typeface="Segoe UI"/>
              </a:rPr>
              <a:t>$subtitle</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@
}

function Get-PolicySlideXml {
    param([hashtable]$Slide)

    $policy = $Slide.Policy
    $policyName = [System.Security.SecurityElement]::Escape($policy.DisplayName)
    $state = $policy.State
    $stateColor = switch ($policy.StateRaw) {
        "enabled" { $script:Theme.EnabledColor }
        "disabled" { $script:Theme.DisabledColor }
        "enabledForReportingButNotEnforced" { $script:Theme.ReportOnlyColor }
        default { $script:Theme.TextColor }
    }

    # Build content sections
    $usersContent = Get-UsersContentXml -Conditions $policy.Conditions
    $appsContent = Get-AppsContentXml -Conditions $policy.Conditions
    $conditionsContent = Get-ConditionsContentXml -Conditions $policy.Conditions
    $controlsContent = Get-ControlsContentXml -GrantControls $policy.GrantControls -SessionControls $policy.SessionControls

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <!-- Header Bar -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Header"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="12192000" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="$($script:Theme.PrimaryColor)"/></a:solidFill>
        </p:spPr>
      </p:sp>
      <!-- Policy Title -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Title"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="304800" y="152400"/>
            <a:ext cx="9144000" cy="609600"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="2800" b="1">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                <a:latin typeface="Segoe UI"/>
              </a:rPr>
              <a:t>$policyName</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <!-- State Badge -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="State"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="10287000" y="304800"/>
            <a:ext cx="1600000" cy="381000"/>
          </a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="$stateColor"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="1400" b="1">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              </a:rPr>
              <a:t>$state</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <!-- Users Section -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="5" name="Users"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="304800" y="1066800"/>
            <a:ext cx="5486400" cy="2286000"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="F5F5F5"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="E0E0E0"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="182880" tIns="182880" rIns="182880" bIns="91440"/>
          <a:lstStyle/>
$usersContent
        </p:txBody>
      </p:sp>
      <!-- Applications Section -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="6" name="Apps"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="6096000" y="1066800"/>
            <a:ext cx="5791200" cy="2286000"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="F5F5F5"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="E0E0E0"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="182880" tIns="182880" rIns="182880" bIns="91440"/>
          <a:lstStyle/>
$appsContent
        </p:txBody>
      </p:sp>
      <!-- Conditions Section -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="7" name="Conditions"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="304800" y="3505200"/>
            <a:ext cx="5486400" cy="2895600"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="F5F5F5"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="E0E0E0"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="182880" tIns="182880" rIns="182880" bIns="91440"/>
          <a:lstStyle/>
$conditionsContent
        </p:txBody>
      </p:sp>
      <!-- Controls Section -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="8" name="Controls"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="6096000" y="3505200"/>
            <a:ext cx="5791200" cy="2895600"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="F5F5F5"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="E0E0E0"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="182880" tIns="182880" rIns="182880" bIns="91440"/>
          <a:lstStyle/>
$controlsContent
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@
}

function Get-UsersContentXml {
    param([hashtable]$Conditions)

    $users = $Conditions.Users
    $lines = @()

    # Header
    $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1600`" b=`"1`"><a:solidFill><a:srgbClr val=`"$($script:Theme.PrimaryColor)`"/></a:solidFill></a:rPr><a:t>USERS &amp; GROUPS</a:t></a:r></a:p>"

    # Include Users
    if ($users.IncludeUsers.Count -gt 0) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Include Users:</a:t></a:r></a:p>"
        foreach ($user in ($users.IncludeUsers | Select-Object -First 5)) {
            $escapedUser = [System.Security.SecurityElement]::Escape($user)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - $escapedUser</a:t></a:r></a:p>"
        }
        if ($users.IncludeUsers.Count -gt 5) {
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`" i=`"1`"/><a:t>  ... and $($users.IncludeUsers.Count - 5) more</a:t></a:r></a:p>"
        }
    }

    # Include Groups
    if ($users.IncludeGroups.Count -gt 0) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Include Groups:</a:t></a:r></a:p>"
        foreach ($group in ($users.IncludeGroups | Select-Object -First 5)) {
            $escapedGroup = [System.Security.SecurityElement]::Escape($group)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - $escapedGroup</a:t></a:r></a:p>"
        }
        if ($users.IncludeGroups.Count -gt 5) {
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`" i=`"1`"/><a:t>  ... and $($users.IncludeGroups.Count - 5) more</a:t></a:r></a:p>"
        }
    }

    # Include Roles
    if ($users.IncludeRoles.Count -gt 0) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Include Roles:</a:t></a:r></a:p>"
        foreach ($role in ($users.IncludeRoles | Select-Object -First 3)) {
            $escapedRole = [System.Security.SecurityElement]::Escape($role)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - $escapedRole</a:t></a:r></a:p>"
        }
        if ($users.IncludeRoles.Count -gt 3) {
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`" i=`"1`"/><a:t>  ... and $($users.IncludeRoles.Count - 3) more</a:t></a:r></a:p>"
        }
    }

    # Exclusions
    $hasExclusions = ($users.ExcludeUsers.Count -gt 0) -or ($users.ExcludeGroups.Count -gt 0) -or ($users.ExcludeRoles.Count -gt 0)
    if ($hasExclusions) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"><a:solidFill><a:srgbClr val=`"$($script:Theme.DisabledColor)`"/></a:solidFill></a:rPr><a:t>Exclusions:</a:t></a:r></a:p>"

        foreach ($user in ($users.ExcludeUsers | Select-Object -First 3)) {
            $escapedUser = [System.Security.SecurityElement]::Escape($user)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - $escapedUser (User)</a:t></a:r></a:p>"
        }
        foreach ($group in ($users.ExcludeGroups | Select-Object -First 3)) {
            $escapedGroup = [System.Security.SecurityElement]::Escape($group)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - $escapedGroup (Group)</a:t></a:r></a:p>"
        }
    }

    if ($lines.Count -eq 1) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`" i=`"1`"/><a:t>No users configured</a:t></a:r></a:p>"
    }

    return $lines -join "`n"
}

function Get-AppsContentXml {
    param([hashtable]$Conditions)

    $apps = $Conditions.Applications
    $lines = @()

    # Header
    $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1600`" b=`"1`"><a:solidFill><a:srgbClr val=`"$($script:Theme.PrimaryColor)`"/></a:solidFill></a:rPr><a:t>APPLICATIONS</a:t></a:r></a:p>"

    # Include Applications
    if ($apps.IncludeApplications.Count -gt 0) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Include Apps:</a:t></a:r></a:p>"
        foreach ($app in ($apps.IncludeApplications | Select-Object -First 6)) {
            $escapedApp = [System.Security.SecurityElement]::Escape($app)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - $escapedApp</a:t></a:r></a:p>"
        }
        if ($apps.IncludeApplications.Count -gt 6) {
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`" i=`"1`"/><a:t>  ... and $($apps.IncludeApplications.Count - 6) more</a:t></a:r></a:p>"
        }
    }

    # User Actions
    if ($apps.IncludeUserActions.Count -gt 0) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>User Actions:</a:t></a:r></a:p>"
        foreach ($action in $apps.IncludeUserActions) {
            $escapedAction = [System.Security.SecurityElement]::Escape($action)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - $escapedAction</a:t></a:r></a:p>"
        }
    }

    # Exclude Applications
    if ($apps.ExcludeApplications.Count -gt 0) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"><a:solidFill><a:srgbClr val=`"$($script:Theme.DisabledColor)`"/></a:solidFill></a:rPr><a:t>Exclude Apps:</a:t></a:r></a:p>"
        foreach ($app in ($apps.ExcludeApplications | Select-Object -First 4)) {
            $escapedApp = [System.Security.SecurityElement]::Escape($app)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - $escapedApp</a:t></a:r></a:p>"
        }
    }

    if ($lines.Count -eq 1) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`" i=`"1`"/><a:t>No applications configured</a:t></a:r></a:p>"
    }

    return $lines -join "`n"
}

function Get-ConditionsContentXml {
    param([hashtable]$Conditions)

    $lines = @()

    # Header
    $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1600`" b=`"1`"><a:solidFill><a:srgbClr val=`"$($script:Theme.PrimaryColor)`"/></a:solidFill></a:rPr><a:t>CONDITIONS</a:t></a:r></a:p>"

    # Platforms
    if ($Conditions.Platforms.IncludePlatforms.Count -gt 0) {
        $platforms = ($Conditions.Platforms.IncludePlatforms -join ", ")
        $escapedPlatforms = [System.Security.SecurityElement]::Escape($platforms)
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Platforms: </a:t></a:r><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>$escapedPlatforms</a:t></a:r></a:p>"
    }

    # Locations
    if ($Conditions.Locations.IncludeLocations.Count -gt 0) {
        $locations = ($Conditions.Locations.IncludeLocations | Select-Object -First 3) -join ", "
        $escapedLocations = [System.Security.SecurityElement]::Escape($locations)
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Locations: </a:t></a:r><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>$escapedLocations</a:t></a:r></a:p>"
    }

    # Client App Types
    if ($Conditions.ClientAppTypes.Count -gt 0) {
        $clientApps = ($Conditions.ClientAppTypes -join ", ")
        $escapedClientApps = [System.Security.SecurityElement]::Escape($clientApps)
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Client Apps: </a:t></a:r><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>$escapedClientApps</a:t></a:r></a:p>"
    }

    # Sign-in Risk
    if ($Conditions.SignInRiskLevels.Count -gt 0) {
        $riskLevels = ($Conditions.SignInRiskLevels -join ", ")
        $escapedRiskLevels = [System.Security.SecurityElement]::Escape($riskLevels)
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Sign-in Risk: </a:t></a:r><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>$escapedRiskLevels</a:t></a:r></a:p>"
    }

    # User Risk
    if ($Conditions.UserRiskLevels.Count -gt 0) {
        $userRiskLevels = ($Conditions.UserRiskLevels -join ", ")
        $escapedUserRiskLevels = [System.Security.SecurityElement]::Escape($userRiskLevels)
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>User Risk: </a:t></a:r><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>$escapedUserRiskLevels</a:t></a:r></a:p>"
    }

    # Device Filter
    if ($Conditions.Devices.DeviceFilter) {
        $filter = $Conditions.Devices.DeviceFilter
        $mode = [System.Security.SecurityElement]::Escape($filter.Mode)
        $rule = [System.Security.SecurityElement]::Escape($filter.Rule)
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Device Filter ($mode):</a:t></a:r></a:p>"
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1000`"/><a:t>  $rule</a:t></a:r></a:p>"
    }

    if ($lines.Count -eq 1) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`" i=`"1`"/><a:t>No additional conditions</a:t></a:r></a:p>"
    }

    return $lines -join "`n"
}

function Get-ControlsContentXml {
    param([hashtable]$GrantControls, [hashtable]$SessionControls)

    $lines = @()

    # Header
    $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1600`" b=`"1`"><a:solidFill><a:srgbClr val=`"$($script:Theme.PrimaryColor)`"/></a:solidFill></a:rPr><a:t>ACCESS CONTROLS</a:t></a:r></a:p>"

    # Grant Controls
    if ($GrantControls.BuiltInControls.Count -gt 0) {
        $operator = if ($GrantControls.Operator -eq "AND") { "Require ALL of:" } else { "Require ONE of:" }
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Grant Controls ($operator)</a:t></a:r></a:p>"

        foreach ($control in $GrantControls.BuiltInControls) {
            $escapedControl = [System.Security.SecurityElement]::Escape($control)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - $escapedControl</a:t></a:r></a:p>"
        }
    }

    # Authentication Strength
    if ($GrantControls.AuthenticationStrength) {
        $authStrength = [System.Security.SecurityElement]::Escape($GrantControls.AuthenticationStrength.DisplayName)
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Auth Strength: </a:t></a:r><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>$authStrength</a:t></a:r></a:p>"
    }

    # Session Controls
    $hasSession = $SessionControls.SignInFrequency -or $SessionControls.PersistentBrowser -or
                  $SessionControls.CloudAppSecurity -or $SessionControls.ApplicationEnforcedRestrictions

    if ($hasSession) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1200`" b=`"1`"/><a:t>Session Controls:</a:t></a:r></a:p>"

        if ($SessionControls.SignInFrequency) {
            $sif = [System.Security.SecurityElement]::Escape($SessionControls.SignInFrequency)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - Sign-in frequency: $sif</a:t></a:r></a:p>"
        }

        if ($SessionControls.PersistentBrowser) {
            $pb = [System.Security.SecurityElement]::Escape($SessionControls.PersistentBrowser)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - Persistent browser: $pb</a:t></a:r></a:p>"
        }

        if ($SessionControls.CloudAppSecurity) {
            $cas = [System.Security.SecurityElement]::Escape($SessionControls.CloudAppSecurity)
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - Cloud App Security: $cas</a:t></a:r></a:p>"
        }

        if ($SessionControls.ApplicationEnforcedRestrictions) {
            $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`"/><a:t>  - App enforced restrictions: Enabled</a:t></a:r></a:p>"
        }
    }

    if ($lines.Count -eq 1) {
        $lines += "          <a:p><a:r><a:rPr lang=`"en-US`" sz=`"1100`" i=`"1`"/><a:t>No access controls configured</a:t></a:r></a:p>"
    }

    return $lines -join "`n"
}

function Get-SummarySlideXml {
    param([hashtable]$Slide)

    $stats = $Slide.Stats
    $date = Get-Date -Format "MMMM d, yyyy"

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <!-- Header -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Header"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="12192000" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="$($script:Theme.PrimaryColor)"/></a:solidFill>
        </p:spPr>
      </p:sp>
      <!-- Title -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Title"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="304800" y="228600"/>
            <a:ext cx="11582400" cy="457200"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="3200" b="1">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              </a:rPr>
              <a:t>Summary - Conditional Access Policies</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <!-- Total Box -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Total"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="609600" y="1371600"/>
            <a:ext cx="2438400" cy="1828800"/>
          </a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="$($script:Theme.PrimaryColor)"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="6000" b="1">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              </a:rPr>
              <a:t>$($stats.Total)</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="1800">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              </a:rPr>
              <a:t>Total Policies</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <!-- Enabled Box -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="5" name="Enabled"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="3352800" y="1371600"/>
            <a:ext cx="2438400" cy="1828800"/>
          </a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="$($script:Theme.EnabledColor)"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="6000" b="1">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              </a:rPr>
              <a:t>$($stats.Enabled)</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="1800">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              </a:rPr>
              <a:t>Enabled</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <!-- Report-Only Box -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="6" name="ReportOnly"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="6096000" y="1371600"/>
            <a:ext cx="2438400" cy="1828800"/>
          </a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="$($script:Theme.ReportOnlyColor)"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="6000" b="1">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              </a:rPr>
              <a:t>$($stats.ReportOnly)</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="1800">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              </a:rPr>
              <a:t>Report-Only</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <!-- Disabled Box -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="7" name="Disabled"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="8839200" y="1371600"/>
            <a:ext cx="2438400" cy="1828800"/>
          </a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="$($script:Theme.DisabledColor)"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="6000" b="1">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              </a:rPr>
              <a:t>$($stats.Disabled)</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="1800">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              </a:rPr>
              <a:t>Disabled</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <!-- Date -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="8" name="Date"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="304800" y="6096000"/>
            <a:ext cx="11582400" cy="457200"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="1400" i="1">
                <a:solidFill><a:srgbClr val="666666"/></a:solidFill>
              </a:rPr>
              <a:t>Generated on $date</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@
}

function Set-PptxTheme {
    <#
    .SYNOPSIS
        Sets the color theme for generated presentations

    .PARAMETER PrimaryColor
        Primary color in hex (without #)

    .PARAMETER EnabledColor
        Color for enabled state

    .PARAMETER DisabledColor
        Color for disabled state

    .PARAMETER ReportOnlyColor
        Color for report-only state
    #>
    [CmdletBinding()]
    param(
        [string]$PrimaryColor,
        [string]$EnabledColor,
        [string]$DisabledColor,
        [string]$ReportOnlyColor
    )

    if ($PrimaryColor) { $script:Theme.PrimaryColor = $PrimaryColor -replace '^#', '' }
    if ($EnabledColor) { $script:Theme.EnabledColor = $EnabledColor -replace '^#', '' }
    if ($DisabledColor) { $script:Theme.DisabledColor = $DisabledColor -replace '^#', '' }
    if ($ReportOnlyColor) { $script:Theme.ReportOnlyColor = $ReportOnlyColor -replace '^#', '' }
}

# Export functions
Export-ModuleMember -Function @(
    'New-PptxDocument',
    'Add-TitleSlide',
    'Add-PolicySlide',
    'Add-SummarySlide',
    'Save-PptxDocument',
    'Set-PptxTheme'
)
