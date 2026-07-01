$logoPath   = "C:\Users\GregHulsey\OneDrive - Hulsey Medical Management\Documents\Hulsey Medical\HMM Logo 2 PNG.png"
$outputPath = "C:\Users\GregHulsey\OneDrive - Hulsey Medical Management\Desktop\AI Work\HMM Website\HMM One Pager.docx"

# Word BGR color values (Blue*65536 + Green*256 + Red)
$darkGreen  = 69*65536  + 94*256  + 29   # #1D5E45
$limeGreen  = 52*65536  + 194*256 + 109  # #6DC234
$lightGreen = 225*65536 + 245*256 + 225  # very light mint — header shading
$white      = 16777215
$black      = 0

# Characters built at runtime to avoid encoding issues
$iconEmail   = [char]0x2709   # ✉
$iconPhone   = [char]0x260E   # ☎
$iconWeb     = [char]0x2295   # ⊕
$eDash       = [char]0x2014   # —
$accentArrow = [char]0x25B6   # ▶  lime green accent before each service

$services = @(
    @("New Practice Start Up",
      "Setting up a new practice is an extremely time and labor intensive process with many moving pieces. I take on that burden so you can focus on patient care."),
    @("Professional Practice Management",
      "Management services can be provided on either a full or fractional basis depending on your needs and budget."),
    @("Strategic Planning",
      "Failure to plan is planning to fail. I engage providers in a formal strategic planning process to identify goals and objectives, then develop tactics to achieve them."),
    @("Human Resources Management",
      "The most valuable resource of any organization is its staff, and a key to success is to attract, hire, develop, and retain the right people. I perform those tasks as well as disciplinary action, removing you from functions that are essential but can be difficult to execute."),
    @("Revenue Cycle Management",
      "Cash flow is the key to any successful business. An analysis of your front-end processes, accounts receivables, and collection efforts help ensure that you are receiving every dollar you are entitled to in reimbursement."),
    @("Credentialing",
      "Without proper credentialing with third party payors, you are working for free. I manage the burdensome process of maintaining your credentials for both insurance companies and hospital medical staff privileges."),
    @("Vendor Resource Alignment",
      "Practices rely on a variety of vendors to assist with their operations, from IT services to collection agencies to benefit brokers. Let me utilize my knowledge and network of contacts to get the right resources in place for what you need."),
    @("Provider Recruitment",
      "As your practice grows, it is essential that you identify new providers to help meet the needs of your patient base. I conduct a thorough and vetted search to find the right candidate who aligns with you personally and professionally."),
    @("Regulatory Compliance",
      "The practice of medicine is filled with the complexity of keeping up with evolving rules and regulations. I stay current with those changes and make sure that your practice is following them, allowing you to focus on patients."),
    @("Staff Training",
      "It is not enough to just know what the newest regulations are; to be truly compliant, your staff must be trained so that they understand such rules and know how to follow them. I do this in an easy-to-follow format that both educates your staff and formally documents your compliance efforts.")
)

# Intro split so the call-to-action sentence can be bold + italic
$introMain = "In today's environment of ever increasing regulatory and financial pressures, physician practices that want to remain independent need a resource that can help them navigate these turbulent times. I'm Greg Hulsey, and I have almost 30 years of hands-on management experience with physician practices and ambulatory surgery centers across a wide spectrum of specialties. I take the management side of your practice off your plate so that you can spend your time treating patients and generating revenue. "
$introCTA  = "Call me today for a free consultation to discuss how I can help you and your practice succeed."

$word = New-Object -ComObject Word.Application
$word.Visible = $false

try {
    $doc = $word.Documents.Add()

    $doc.PageSetup.TopMargin    = $word.InchesToPoints(0.5)
    $doc.PageSetup.BottomMargin = $word.InchesToPoints(0.5)
    $doc.PageSetup.LeftMargin   = $word.InchesToPoints(0.5)
    $doc.PageSetup.RightMargin  = $word.InchesToPoints(0.5)

    $sel = $word.Selection

    # ── LOGO ──────────────────────────────────────────────────────────
    $sel.ParagraphFormat.Alignment      = 1
    $sel.ParagraphFormat.SpaceBefore    = 0
    $sel.ParagraphFormat.SpaceAfter     = 2
    $sel.ParagraphFormat.LineSpacingRule = 0
    $pic = $sel.InlineShapes.AddPicture($logoPath, $false, $true)
    $pic.LockAspectRatio = $true
    $pic.Width = $word.InchesToPoints(2.0)
    $sel.TypeParagraph()

    # ── NAME / TITLE ──────────────────────────────────────────────────
    $sel.ParagraphFormat.Alignment      = 1
    $sel.ParagraphFormat.SpaceBefore    = 0
    $sel.ParagraphFormat.SpaceAfter     = 6
    $sel.ParagraphFormat.LineSpacingRule = 0
    $sel.Font.Name   = "Times New Roman"
    $sel.Font.Size   = 11
    $sel.Font.Bold   = $true
    $sel.Font.Italic = $false
    $sel.Font.Color  = $darkGreen
    $sel.TypeText("Greg Hulsey  |  Principal & Founder")
    $sel.TypeParagraph()

    # ── INTRO PARAGRAPH ───────────────────────────────────────────────
    $sel.ParagraphFormat.Alignment      = 3
    $sel.ParagraphFormat.SpaceBefore    = 0
    $sel.ParagraphFormat.SpaceAfter     = 6
    $sel.ParagraphFormat.LineSpacingRule = 0
    $sel.Font.Name   = "Times New Roman"
    $sel.Font.Size   = 11
    $sel.Font.Bold   = $false
    $sel.Font.Italic = $false
    $sel.Font.Color  = $black
    $sel.TypeText($introMain)
    # Call-to-action sentence — bold italic
    $sel.Font.Bold   = $true
    $sel.Font.Italic = $true
    $sel.TypeText($introCTA)
    $sel.Font.Bold   = $false
    $sel.Font.Italic = $false
    $sel.TypeParagraph()

    # ── SERVICES OFFERED HEADER (light green shading) ─────────────────
    $sel.ParagraphFormat.Alignment      = 1
    $sel.ParagraphFormat.SpaceBefore    = 2
    $sel.ParagraphFormat.SpaceAfter     = 6
    $sel.ParagraphFormat.LineSpacingRule = 0
    $sel.ParagraphFormat.Shading.BackgroundPatternColor = $lightGreen
    $sel.Font.Name   = "Times New Roman"
    $sel.Font.Size   = 13
    $sel.Font.Bold   = $true
    $sel.Font.Italic = $false
    $sel.Font.Color  = $darkGreen
    $sel.TypeText("Services Offered:")
    $sel.TypeParagraph()

    # ── SERVICES LIST — lime green ▶ accent + bold name + description ─
    foreach ($svc in $services) {
        $sel.ParagraphFormat.Alignment      = 3
        $sel.ParagraphFormat.SpaceBefore    = 0
        $sel.ParagraphFormat.SpaceAfter     = 7
        $sel.ParagraphFormat.LineSpacingRule = 0
        $sel.ParagraphFormat.Shading.BackgroundPatternColor = $white

        # Lime green accent arrow (Segoe UI Symbol for reliable rendering)
        $sel.Font.Name   = "Segoe UI Symbol"
        $sel.Font.Size   = 9
        $sel.Font.Bold   = $false
        $sel.Font.Italic = $false
        $sel.Font.Color  = $limeGreen
        $sel.TypeText($accentArrow + " ")

        # Service name — bold dark green Georgia
        $sel.Font.Name  = "Times New Roman"
        $sel.Font.Size  = 11
        $sel.Font.Bold  = $true
        $sel.Font.Color = $darkGreen
        $sel.TypeText($svc[0] + " ")

        # Description — regular black Georgia
        $sel.Font.Bold  = $false
        $sel.Font.Color = $black
        $sel.TypeText($eDash + " " + $svc[1])
        $sel.TypeParagraph()
    }

    # ── CONTACT INFO — top border = dark green divider line ───────────
    $sel.ParagraphFormat.Alignment      = 1
    $sel.ParagraphFormat.SpaceBefore    = 8
    $sel.ParagraphFormat.SpaceAfter     = 0
    $sel.ParagraphFormat.LineSpacingRule = 0
    $sel.ParagraphFormat.Shading.BackgroundPatternColor = $white
    $sel.ParagraphFormat.Borders.Enable = $false          # clear any inherited borders
    $topBdr = $sel.ParagraphFormat.Borders.Item(-1)       # wdBorderTop
    $topBdr.LineStyle = 1 ; $topBdr.Color = $darkGreen ; $topBdr.LineWidth = 8
    $sel.ParagraphFormat.Borders.DistanceFromTop = 6
    $sel.Font.Bold   = $true
    $sel.Font.Italic = $false
    $sel.Font.Color  = $darkGreen

    $sel.Font.Name = "Segoe UI Symbol" ; $sel.Font.Size = 11
    $sel.TypeText($iconEmail)
    $sel.Font.Name = "Times New Roman" ; $sel.Font.Size = 11
    $sel.TypeText("  greg@hulseymedical.com")

    $sel.TypeText("          ")

    $sel.Font.Name = "Segoe UI Symbol" ; $sel.Font.Size = 11
    $sel.TypeText($iconPhone)
    $sel.Font.Name = "Times New Roman" ; $sel.Font.Size = 11
    $sel.TypeText("  205-746-7352")

    $sel.TypeText("          ")

    $sel.Font.Name = "Segoe UI Symbol" ; $sel.Font.Size = 11
    $sel.TypeText($iconWeb)
    $sel.Font.Name = "Times New Roman" ; $sel.Font.Size = 11
    $sel.TypeText("  www.hulseymedical.com")

    $sel.TypeParagraph()

    # ── SAVE ──────────────────────────────────────────────────────────
    $doc.SaveAs2($outputPath, 12)
    $doc.Close($false)
    Write-Output "SUCCESS -- saved to: $outputPath"

} catch {
    Write-Error "Failed: $_"
} finally {
    $word.Quit()
}
