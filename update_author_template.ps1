param(
  [string]$SourcePath,
  [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

if (-not $SourcePath) {
  $projectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
  $SourcePath = Join-Path $projectRoot 'Manuscript POC_PE_for_Life_Author_Submission_Template.docx'
}

if (-not $OutputPath) {
  $OutputPath = Join-Path $PSScriptRoot 'Manuscript POC_PE_for_Life_Author_Submission_Template_Navigation_Aligned.docx'
}

$wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$xmlNs = 'http://www.w3.org/XML/1998/namespace'

Add-Type -AssemblyName System.IO.Compression.FileSystem

function New-WElement {
  param([System.Xml.XmlDocument]$Doc, [string]$Name)
  $Doc.CreateElement('w', $Name, $script:wNs)
}

function Set-WAttribute {
  param([System.Xml.XmlElement]$Node, [string]$Name, [string]$Value)
  $null = $Node.SetAttribute($Name, $script:wNs, $Value)
}

function Add-Run {
  param(
    [System.Xml.XmlDocument]$Doc,
    [System.Xml.XmlElement]$Paragraph,
    [string]$Text,
    [switch]$Bold,
    [switch]$Italic,
    [string]$Color = ''
  )
  $run = New-WElement $Doc 'r'
  if ($Bold -or $Italic -or $Color) {
    $runProperties = New-WElement $Doc 'rPr'
    if ($Bold) {
      $null = $runProperties.AppendChild((New-WElement $Doc 'b'))
    }
    if ($Italic) {
      $null = $runProperties.AppendChild((New-WElement $Doc 'i'))
    }
    if ($Color) {
      $colorNode = New-WElement $Doc 'color'
      Set-WAttribute $colorNode 'val' $Color
      $null = $runProperties.AppendChild($colorNode)
    }
    $null = $run.AppendChild($runProperties)
  }
  $textNode = New-WElement $Doc 't'
  $null = $textNode.SetAttribute('space', $script:xmlNs, 'preserve')
  $resolvedText = $Text.Replace('[[EN]]', [string][char]0x2013)
  $resolvedText = $resolvedText.Replace('[[EM]]', [string][char]0x2014)
  $resolvedText = $resolvedText.Replace('[[BOX]]', [string][char]0x2610)
  $resolvedText = $resolvedText.Replace('[[BULLET]]', [string][char]0x2022)
  $textNode.InnerText = $resolvedText
  $null = $run.AppendChild($textNode)
  $null = $Paragraph.AppendChild($run)
}

function New-Paragraph {
  param(
    [System.Xml.XmlDocument]$Doc,
    [string]$Text,
    [string]$Style = 'Normal',
    [switch]$Bold,
    [switch]$Italic,
    [string]$Color = ''
  )
  $paragraph = New-WElement $Doc 'p'
  if ($Style) {
    $paragraphProperties = New-WElement $Doc 'pPr'
    $styleNode = New-WElement $Doc 'pStyle'
    Set-WAttribute $styleNode 'val' $Style
    $null = $paragraphProperties.AppendChild($styleNode)
    $null = $paragraph.AppendChild($paragraphProperties)
  }
  $null = Add-Run $Doc $paragraph $Text -Bold:$Bold -Italic:$Italic -Color $Color
  $paragraph
}

function New-FieldParagraph {
  param([System.Xml.XmlDocument]$Doc, [string]$Label, [string]$Prompt)
  $paragraph = New-WElement $Doc 'p'
  $paragraphProperties = New-WElement $Doc 'pPr'
  $styleNode = New-WElement $Doc 'pStyle'
  Set-WAttribute $styleNode 'val' 'Normal'
  $null = $paragraphProperties.AppendChild($styleNode)
  $null = $paragraph.AppendChild($paragraphProperties)
  $null = Add-Run $Doc $paragraph ($Label + ' ') -Bold
  $null = Add-Run $Doc $paragraph $Prompt -Italic -Color '6B7777'
  $paragraph
}

function New-PageBreak {
  param([System.Xml.XmlDocument]$Doc)
  $paragraph = New-WElement $Doc 'p'
  $run = New-WElement $Doc 'r'
  $break = New-WElement $Doc 'br'
  Set-WAttribute $break 'type' 'page'
  $null = $run.AppendChild($break)
  $null = $paragraph.AppendChild($run)
  $paragraph
}

function New-Table {
  param(
    [System.Xml.XmlDocument]$Doc,
    [string[]]$Headers,
    [System.Collections.IEnumerable]$Rows
  )
  $table = New-WElement $Doc 'tbl'
  $tableProperties = New-WElement $Doc 'tblPr'
  $style = New-WElement $Doc 'tblStyle'
  Set-WAttribute $style 'val' 'TableGrid'
  $null = $tableProperties.AppendChild($style)
  $width = New-WElement $Doc 'tblW'
  Set-WAttribute $width 'w' '0'
  Set-WAttribute $width 'type' 'auto'
  $null = $tableProperties.AppendChild($width)
  $null = $table.AppendChild($tableProperties)

  $grid = New-WElement $Doc 'tblGrid'
  foreach ($unused in $Headers) {
    $column = New-WElement $Doc 'gridCol'
    Set-WAttribute $column 'w' ([string][Math]::Floor(9000 / $Headers.Count))
    $null = $grid.AppendChild($column)
  }
  $null = $table.AppendChild($grid)

  $allRows = [System.Collections.ArrayList]::new()
  $null = $allRows.Add($Headers)
  foreach ($row in $Rows) {
    $null = $allRows.Add($row)
  }

  for ($rowIndex = 0; $rowIndex -lt $allRows.Count; $rowIndex++) {
    $rowValues = [string[]]($allRows[$rowIndex])
    $tableRow = New-WElement $Doc 'tr'
    for ($columnIndex = 0; $columnIndex -lt $Headers.Count; $columnIndex++) {
      $cell = New-WElement $Doc 'tc'
      $cellProperties = New-WElement $Doc 'tcPr'
      $cellWidth = New-WElement $Doc 'tcW'
      Set-WAttribute $cellWidth 'w' ([string][Math]::Floor(9000 / $Headers.Count))
      Set-WAttribute $cellWidth 'type' 'dxa'
      $null = $cellProperties.AppendChild($cellWidth)
      if ($rowIndex -eq 0) {
        $shading = New-WElement $Doc 'shd'
        Set-WAttribute $shading 'fill' 'DDEBE7'
        $null = $cellProperties.AppendChild($shading)
      }
      $null = $cell.AppendChild($cellProperties)
      $cellText = if ($columnIndex -lt $rowValues.Count) { $rowValues[$columnIndex] } else { '' }
      $null = $cell.AppendChild((New-Paragraph $Doc $cellText -Bold:($rowIndex -eq 0)))
      $null = $tableRow.AppendChild($cell)
    }
    $null = $table.AppendChild($tableRow)
  }
  $table
}

function Add-Node {
  param([System.Xml.XmlElement]$Body, [System.Xml.XmlNode]$Node)
  $null = $Body.AppendChild($Node)
}

function Add-NamedTable {
  param(
    [System.Xml.XmlElement]$Body,
    [System.Xml.XmlDocument]$Doc,
    [string]$Name
  )
  $rows = [System.Collections.ArrayList]::new()
  switch ($Name) {
    'PageTypes' {
      $headers = @('Page Type','Navigation Location','Author Section')
      $null = $rows.Add([string[]]@('K-2 Unit Overview','Elementary > K-2 > Unit > Overview','Complete Section C once per unit.'))
      $null = $rows.Add([string[]]@('K-2 Content Set','Elementary > K-2 > Unit > Content Set I, II, or III','Complete Section D once per Content Set.'))
      $null = $rows.Add([string[]]@('Activity Unit','Elementary > 3-5 > Activity Group > Activity, or Secondary > Category > Activity','Complete Section E once per activity page.'))
      $null = $rows.Add([string[]]@('Resource Collection','Resources > Resource button','Complete Section F once per resource page.'))
    }
    'Standards' {
      $headers = @('National Standard','Grade-Span Indicator','Unit Goal/Outcome','Assessment')
      1..4 | ForEach-Object {
        $null = $rows.Add([string[]]@('[Enter standard]','[Enter indicator]','[Enter goal/outcome]','[Enter type and timing]'))
      }
    }
    'LessonActivities' {
      $headers = @('Time','Activity','Description','Diagram/File Reference')
      $null = $rows.Add([string[]]@('[Enter time]','Introductory / Instant Activity','[Enter teacher-directed procedures]','[Enter reference]'))
      $null = $rows.Add([string[]]@('[Enter time]','Transition','[Enter procedures]','[Enter reference]'))
      $null = $rows.Add([string[]]@('[Enter time]','Set Induction','[Enter procedures]','[Enter reference]'))
      $null = $rows.Add([string[]]@('[Enter time]','Skill/Concept Development','[Enter procedures, cues, and differentiation]','[Enter reference]'))
      $null = $rows.Add([string[]]@('[Enter time]','Application Activity / Assessment','[Enter procedures and assessment connection]','[Enter reference]'))
      $null = $rows.Add([string[]]@('[Enter time]','Closure','[Enter closure and reflection prompts]','[Enter reference]'))
    }
    'ResourceFiles' {
      $headers = @('Resource Title','Audience','Brief Description','File Name/URL','Accessibility Notes')
      1..5 | ForEach-Object {
        $null = $rows.Add([string[]]@('[Enter title]','[Teacher / Student / Home / All]','[Enter description]','[Enter exact file name or URL]','[Enter alt-text notes]'))
      }
    }
    'Inventory' {
      $headers = @('File Name','Type','Page Section','Description','Accessibility','Source/Permission')
      1..8 | ForEach-Object {
        $null = $rows.Add([string[]]@('[Enter exact file name]','[PDF / image / video / link / other]','[Enter section]','[Enter description]','[Enter alt text or transcript notes]','[Enter permission status]'))
      }
    }
    default {
      throw "Unknown table: $Name"
    }
  }
  Add-Node $Body (New-Table $Doc $headers $rows)
}

function Add-ResourceEntries {
  param([System.Xml.XmlElement]$Body, [System.Xml.XmlDocument]$Doc)
  foreach ($audience in @('Teacher','Student','Home')) {
    Add-Node $Body (New-Paragraph $Doc $audience 'Heading3')
    Add-Node $Body (New-FieldParagraph $Doc 'Resource title:' '[Enter response here.]')
    Add-Node $Body (New-FieldParagraph $Doc 'Brief description:' '[Enter response here.]')
    Add-Node $Body (New-FieldParagraph $Doc 'File name or URL:' '[Enter response here.]')
    Add-Node $Body (New-FieldParagraph $Doc 'Accessibility/alt-text notes:' '[Enter response here.]')
  }
}

function Add-LessonTemplate {
  param([System.Xml.XmlElement]$Body, [System.Xml.XmlDocument]$Doc)
  Add-Node $Body (New-Paragraph $Doc 'Lesson Day #: [Enter day/lesson number.]' 'Heading3')
  Add-Node $Body (New-FieldParagraph $Doc 'Lesson title:' '[Enter response here.]')
  Add-Node $Body (New-FieldParagraph $Doc 'Estimated lesson length:' '[Enter response here.]')
  Add-Node $Body (New-Paragraph $Doc 'A. Lesson Outcomes/Objectives' 'Heading3')
  1..5 | ForEach-Object {
    Add-Node $Body (New-FieldParagraph $Doc ("$_.") '[Enter measurable outcome/objective.]')
  }
  Add-Node $Body (New-Paragraph $Doc 'B. Equipment' 'Heading3')
  Add-Node $Body (New-Paragraph $Doc '[List equipment and quantities.]' -Italic -Color '6B7777')
  Add-Node $Body (New-Paragraph $Doc 'C. Teaching Resources' 'Heading3')
  Add-Node $Body (New-Paragraph $Doc '[List teaching resources and exact file names or URLs.]' -Italic -Color '6B7777')
  Add-Node $Body (New-Paragraph $Doc 'D. Lesson Activities' 'Heading3')
  Add-NamedTable $Body $Doc 'LessonActivities'
  Add-Node $Body (New-Paragraph $Doc 'E. Assessment Integration' 'Heading3')
  Add-Node $Body (New-FieldParagraph $Doc 'Diagnostic/formative assessment(s):' '[Enter response here.]')
  Add-Node $Body (New-FieldParagraph $Doc 'Summative assessment(s):' '[Enter response here.]')
  Add-Node $Body (New-FieldParagraph $Doc 'How results will be recorded or graded:' '[Enter response here.]')
}

$sourceZip = [System.IO.Compression.ZipFile]::OpenRead($SourcePath)
try {
  $entry = $sourceZip.GetEntry('word/document.xml')
  $reader = [System.IO.StreamReader]::new($entry.Open())
  try {
    [xml]$document = $reader.ReadToEnd()
  }
  finally {
    $reader.Dispose()
  }
}
finally {
  $sourceZip.Dispose()
}

$ns = [System.Xml.XmlNamespaceManager]::new($document.NameTable)
$ns.AddNamespace('w', $wNs)
$body = $document.SelectSingleNode('//w:body', $ns)
$sourceParagraphs = @($body.SelectNodes('./w:p', $ns))
$sectionProperties = $body.SelectSingleNode('./w:sectPr', $ns).CloneNode($true)
$titleOne = $sourceParagraphs[0].CloneNode($true)
$titleTwo = $sourceParagraphs[1].CloneNode($true)
$instruction = $sourceParagraphs[2].CloneNode($true)

function Set-ParagraphText {
  param([System.Xml.XmlNode]$Paragraph, [string]$Text, [System.Xml.XmlNamespaceManager]$NamespaceManager)
  $nodes = @($Paragraph.SelectNodes('.//w:t', $NamespaceManager))
  $nodes[0].InnerText = $Text
  for ($i = 1; $i -lt $nodes.Count; $i++) {
    $nodes[$i].InnerText = ''
  }
}

Set-ParagraphText $titleOne 'PE for Life' $ns
Set-ParagraphText $titleTwo 'Navigation-Aligned Author Content Submission Template' $ns
Set-ParagraphText $instruction 'Instructions: Complete Sections A and B, then complete only the page-type section(s) matching the selected navigation destination. Do not delete headings. Duplicate the K-2 Content Set section once for each Content Set and duplicate the Lesson Day subsection for every lesson. Provide exact file names for all PDFs, diagrams, and media.' $ns

$outline = @'
P|Template version: July 2026

H1|A. Author and Submission Information
F|Author name(s):|[Enter response here.]
F|Affiliation/role:|[Enter response here.]
F|Email:|[Enter response here.]
F|Date submitted:|[Enter response here.]
F|Content owner/contact:|[Enter response here.]
F|Submission status:|[[BOX]] New page   [[BOX]] Revision to an existing page
F|Working unit/activity title:|[Enter response here.]

H1|B. Navigation Placement
I|Use the navigation labels exactly as they appear below. The web team will use this information to connect submitted content to the correct HTML page and navigation button.
H2|1. Selected Navigation Destination
F|Level 1:|[[BOX]] Elementary   [[BOX]] Secondary   [[BOX]] Resources
F|Level 2 (grade band/category):|[Enter exact button label.]
F|Level 3 (unit/activity group):|[Enter exact button label.]
F|Level 4 (page/button, if applicable):|[Enter exact button label.]
F|Full navigation breadcrumb:|Example: Elementary > K-2 > Rhythm & Dance > Content Set I
F|Page type:|[[BOX]] K-2 Unit Overview   [[BOX]] K-2 Content Set   [[BOX]] 3-5/Secondary Activity Unit   [[BOX]] Resource Collection
F|Content Set number (if applicable):|[[BOX]] I   [[BOX]] II   [[BOX]] III
F|Navigation button label:|[Enter exact label.]
F|Proposed HTML file name (optional):|[Enter response here.]
H2|2. Page-Type Guide
T|PageTypes
H2|3. Current Navigation Button Reference
H3|Elementary
B|[[BULLET]] K-2
P|  [[BULLET]] Locomotor/nonlocomotory [[EM]] Overview; Content Set I; Content Set II; Content Set III
P|  [[BULLET]] Manipulative [[EM]] Overview; Content Set I; Content Set II; Content Set III
P|  [[BULLET]] Rhythm & Dance [[EM]] Overview; Content Set I; Content Set II; Content Set III
P|  [[BULLET]] Developmental Gymnastics [[EM]] Overview; Content Set I; Content Set II; Content Set III
B|[[BULLET]] 3-5
P|  [[BULLET]] Application of skills/concepts to beginning Soccer [[EM]] Soccer
P|  [[BULLET]] Application of skills/concepts to beginning Basketball [[EM]] Basketball
P|  [[BULLET]] Application of skills/concepts to beginning Target Activities [[EM]] Target Activities
P|  [[BULLET]] Application of skills/concepts to beginning Striking Activities [[EM]] Striking Activities
P|  [[BULLET]] Application of skills/concepts to beginning Volleying Activities [[EM]] Volleying
P|  [[BULLET]] Dance/Rhythms [[EM]] Placeholder
P|  [[BULLET]] Developmental Gymnastics #2 [[EM]] Placeholder
H3|Secondary
P|[[BULLET]] Territory/Invasion Games [[EM]] Basketball (beginner and advanced); Field Hockey; Flag football; Lacrosse; Soccer; Team Handball; Ultimate
P|[[BULLET]] Net/Wall [[EM]] Badminton; Pickleball; Tennis; Volleyball
P|[[BULLET]] Striking/Fielding [[EM]] Softball
P|[[BULLET]] Target Games [[EM]] Disc golf; Golf
P|[[BULLET]] Lifetime/Leisure & Wellness [[EM]] Fitness Walking; Strength & Conditioning; Track & Field; Yoga
P|[[BULLET]] Rhythm & Dance [[EM]] Country Western Line Dance; Creative Dance; Cultural Dance; Line Dance
P|[[BULLET]] Adventure/Outdoor Education [[EM]] Adventure based learning
H3|Resources
P|[[BULLET]] Warmups/cooldowns
P|[[BULLET]] Other potential resources

PB|
H1|C. K-2 Unit Overview Page
I|Complete this section for an Overview button under a K-2 unit.
F|Unit/page title:|[Enter response here.]
F|Navigation breadcrumb:|[Enter response here.]
H2|1. Statement of Purpose
PROMPT|[Enter content here. Add additional paragraphs as needed.]
H2|2. Brief Description of Unit
I|Summarize the unit and describe the focus and progression of Content Sets I, II, and III.
PROMPT|[Enter content here. Add additional paragraphs as needed.]
H2|3. Rationale for Unit
I|Explain developmental appropriateness, foundational importance, and connections to later learning.
PROMPT|[Enter content here. Add additional paragraphs as needed.]
H2|4. Grading Resources
F|Resource title:|[Enter response here.]
F|Brief description:|[Enter response here.]
F|File name or URL:|[Enter response here.]
F|Accessibility/alt-text notes:|[Enter response here.]

PB|
H1|D. K-2 Content Set Page
I|Duplicate this entire section once for each Content Set (I, II, and III) being submitted.
F|Unit title:|[Enter response here.]
F|Content Set:|[[BOX]] I   [[BOX]] II   [[BOX]] III
F|Content Set page title:|[Enter response here.]
F|Navigation breadcrumb:|[Enter response here.]
F|Number of lessons:|[Enter response here.]
H2|1. Essential Questions
QUESTIONS|8|[Enter essential question.]
H2|2. Standards and Unit Goals
T|Standards
H3|Optional Instructional Focus
F|Prerequisite skills/movement concepts:|[Enter response here.]
F|Skills introduced or refined:|[Enter response here.]
F|Movement concepts (space, effort, relationships, body awareness):|[Enter response here.]
H2|3. Equipment & Materials
F|Equipment and quantities:|[Enter response here.]
F|Teaching materials/resources:|[Enter response here.]
F|Facilities/space:|[Enter response here.]
F|Lesson length/time:|[Enter response here.]
F|Teacher expertise/training/certification:|[Enter response here.]
F|Safety and accessibility considerations:|[Enter response here.]
H2|4. Block Plan
F|Block Plan file name (PDF):|[Enter response here.]
F|Brief description:|[Enter response here.]
F|Alternative text/accessibility notes:|[Enter response here.]
H2|5. Lesson Plan
I|Duplicate the Lesson Day subsection below for every lesson in this Content Set.
LESSON|
H2|6. Assessments & Grading Criteria
F|Assessment resource title:|[Enter response here.]
F|Assessment type and timing:|[Enter response here.]
F|Scoring/grading criteria:|[Enter response here.]
F|File name or URL:|[Enter response here.]
F|Brief description:|[Enter response here.]
H2|7. Resources
RESOURCES|

PB|
H1|E. 3-5 / Secondary Activity Unit Page
I|Complete this section for a single activity button under Elementary 3-5 or Secondary.
F|Activity/page title:|[Enter response here.]
F|Grade band:|[Enter response here.]
F|Navigation breadcrumb:|[Enter response here.]
H2|1. Overview of Unit
PROMPT|[Enter content here. Add additional paragraphs as needed.]
H2|2. Rationale
I|Explain age appropriateness, the central organizer for the unit, and its value for student learning.
PROMPT|[Enter content here. Add additional paragraphs as needed.]
H2|3. Essential Questions
QUESTIONS|8|[Enter essential question.]
H2|4. Standards and Unit Goals
T|Standards
H2|5. Equipment & Materials
F|Equipment and quantities:|[Enter response here.]
F|Teaching materials/resources:|[Enter response here.]
F|Facilities/space:|[Enter response here.]
F|Lesson length/time:|[Enter response here.]
F|Safety and accessibility considerations:|[Enter response here.]
H2|6. Block Plan
F|Block Plan file name (PDF):|[Enter response here.]
F|Brief description:|[Enter response here.]
F|Alternative text/accessibility notes:|[Enter response here.]
H2|7. Lesson Plan
I|Duplicate the Lesson Day subsection below for every lesson in this activity unit.
LESSON|
H2|8. Assessments & Grading Criteria
F|Assessment resource title:|[Enter response here.]
F|Assessment type and timing:|[Enter response here.]
F|Scoring/grading criteria:|[Enter response here.]
F|File name or URL:|[Enter response here.]
F|Brief description:|[Enter response here.]
H2|9. Resources
RESOURCES|

PB|
H1|F. Resource Collection Page
I|Complete this section for a button under the Resources menu.
F|Resource collection/page title:|[Enter response here.]
F|Navigation breadcrumb:|[Enter response here.]
H2|1. Overview
PROMPT|[Enter content here. Add additional paragraphs as needed.]
H2|2. Resource Files
T|ResourceFiles
H2|3. Related Resources
F|Related PE for Life page/title:|[Enter response here.]
F|Navigation breadcrumb or URL:|[Enter response here.]
F|Reason for linking:|[Enter response here.]

PB|
H1|G. File and Media Inventory
I|List every file supplied with this submission. File names must match the submitted files exactly.
T|Inventory

PB|
H1|H. Author Final Check
P|[[BOX]] Sections A and B are complete.
P|[[BOX]] The navigation breadcrumb uses the exact current button labels.
P|[[BOX]] The correct page-type section is complete.
P|[[BOX]] K-2 submissions include a separate copy of Section D for each Content Set.
P|[[BOX]] Every lesson has a completed Lesson Day subsection.
P|[[BOX]] Lesson outcomes align with standards and unit goals.
P|[[BOX]] Assessments and grading criteria are clearly identified.
P|[[BOX]] Teacher, Student, and Home resources are separated and labeled.
P|[[BOX]] All PDFs, diagrams, images, and media use exact file names.
P|[[BOX]] Accessibility notes, alt text, and source/permission information are included.
P|[[BOX]] Language is teacher-facing, clear, and professional.
F|Author signature/name:|[Enter response here.]
F|Date:|[Enter response here.]
'@

$body.RemoveAll()
Add-Node $body $titleOne
Add-Node $body $titleTwo
Add-Node $body $instruction

$outlineLines = [System.Text.RegularExpressions.Regex]::Split($outline, '\r?\n')
foreach ($line in $outlineLines) {
  if ([string]::IsNullOrWhiteSpace($line)) {
    continue
  }
  $parts = $line.Split('|')
  $command = $parts[0]
  switch ($command) {
    'H1' { Add-Node $body (New-Paragraph $document $parts[1] 'Heading1') }
    'H2' { Add-Node $body (New-Paragraph $document $parts[1] 'Heading2') }
    'H3' { Add-Node $body (New-Paragraph $document $parts[1] 'Heading3') }
    'P' { Add-Node $body (New-Paragraph $document $parts[1]) }
    'B' { Add-Node $body (New-Paragraph $document $parts[1] -Bold) }
    'I' { Add-Node $body (New-Paragraph $document $parts[1] -Italic -Color '6B7777') }
    'F' { Add-Node $body (New-FieldParagraph $document $parts[1] $parts[2]) }
    'PROMPT' { Add-Node $body (New-Paragraph $document $parts[1] -Italic -Color '6B7777') }
    'PB' { Add-Node $body (New-PageBreak $document) }
    'T' { Add-NamedTable $body $document $parts[1] }
    'QUESTIONS' {
      $count = [int]$parts[1]
      1..$count | ForEach-Object {
        Add-Node $body (New-FieldParagraph $document ("$_.") $parts[2])
      }
    }
    'LESSON' { Add-LessonTemplate $body $document }
    'RESOURCES' { Add-ResourceEntries $body $document }
    default { throw "Unknown outline command: $command" }
  }
}

Add-Node $body $sectionProperties

Copy-Item -LiteralPath $SourcePath -Destination $OutputPath -Force
$targetZip = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
  $oldEntry = $targetZip.GetEntry('word/document.xml')
  $oldEntry.Delete()
  $newEntry = $targetZip.CreateEntry('word/document.xml', [System.IO.Compression.CompressionLevel]::Optimal)
  $stream = $newEntry.Open()
  try {
    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
    try {
      $document.Save($writer)
    }
    finally {
      $writer.Dispose()
    }
  }
  finally {
    $stream.Dispose()
  }
}
finally {
  $targetZip.Dispose()
}

Get-Item -LiteralPath $OutputPath | Select-Object FullName,Length,LastWriteTime
