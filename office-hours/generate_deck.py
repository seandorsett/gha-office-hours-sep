import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


OUT = Path(__file__).with_name("GitHub-Actions-Office-Hours-Aug-Sep-2026.pptx")

W = Inches(13.333)
H = Inches(7.5)
BG = RGBColor(13, 17, 23)
PANEL = RGBColor(22, 27, 34)
PANEL_2 = RGBColor(33, 38, 45)
WHITE = RGBColor(240, 246, 252)
MUTED = RGBColor(139, 148, 158)
BLUE = RGBColor(88, 166, 255)
PURPLE = RGBColor(188, 140, 255)
GREEN = RGBColor(63, 185, 80)
ORANGE = RGBColor(210, 153, 34)
RED = RGBColor(248, 81, 73)
MONO = "Cascadia Mono"
SANS = "Aptos"


def rect(slide, x, y, w, h, fill, radius=False, line=None):
    kind = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line or fill
    return shape


def text(slide, value, x, y, w, h, size=20, color=WHITE, bold=False,
         font=SANS, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, margin=0.04):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = Inches(margin)
    frame.margin_top = frame.margin_bottom = Inches(margin)
    frame.vertical_anchor = valign
    p = frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = value
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def rich_text(slide, runs, x, y, w, h, size=20, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.margin_left = frame.margin_right = Inches(0.04)
    p = frame.paragraphs[0]
    p.alignment = align
    for value, color, bold, font in runs:
        run = p.add_run()
        run.text = value
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
    return box


def bullet_list(slide, items, x, y, w, h, size=22, color=WHITE, gap=10):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.05)
    frame.margin_right = Inches(0.04)
    for idx, item in enumerate(items):
        p = frame.paragraphs[0] if idx == 0 else frame.add_paragraph()
        p.text = f"•  {item}"
        p.font.name = SANS
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.space_after = Pt(gap)
    return box


def code_box(slide, code, x, y, w, h, size=15):
    rect(slide, x, y, w, h, PANEL, radius=True, line=PANEL_2)
    text(slide, code, x + 0.25, y + 0.22, w - 0.5, h - 0.4,
         size=size, color=WHITE, font=MONO)


def badge(slide, label, x, y, color=BLUE, width=1.05):
    rect(slide, x, y, width, 0.34, color, radius=True)
    text(slide, label, x, y + 0.01, width, 0.28, size=11, color=BG,
         bold=True, align=PP_ALIGN.CENTER)


def title(slide, heading, kicker=None, number=None):
    if kicker:
        text(slide, kicker.upper(), 0.7, 0.32, 7, 0.3, size=11, color=BLUE, bold=True)
    text(slide, heading, 0.7, 0.73, 11.8, 0.72, size=30, bold=True)
    rect(slide, 0.7, 1.55, 1.0, 0.05, PURPLE)
    if number is not None:
        text(slide, f"{number:02}", 12.15, 0.34, 0.5, 0.3, size=11,
             color=MUTED, bold=True, align=PP_ALIGN.RIGHT)


def footer(slide, label="GitHub Actions Office Hour · Sep 2026"):
    text(slide, label, 0.7, 7.13, 6.4, 0.2, size=9, color=MUTED)


def add_notes(slide, timing, summary, notes):
    frame = slide.notes_slide.notes_text_frame
    frame.clear()
    frame.word_wrap = True

    heading = frame.paragraphs[0]
    heading.text = "PRESENTER NOTES"
    heading.font.name = SANS
    heading.font.size = Pt(20)
    heading.font.bold = True
    heading.font.color.rgb = RGBColor(31, 78, 121)
    heading.space_after = Pt(8)

    timing_line = frame.add_paragraph()
    timing_line.text = f"Timing: {timing}"
    timing_line.font.name = SANS
    timing_line.font.size = Pt(16)
    timing_line.font.bold = True
    timing_line.font.color.rgb = RGBColor(46, 125, 50)
    timing_line.space_after = Pt(12)

    summary_label = frame.add_paragraph()
    summary_label.text = "Feature summary"
    summary_label.font.name = SANS
    summary_label.font.size = Pt(16)
    summary_label.font.bold = True
    summary_label.font.color.rgb = RGBColor(31, 78, 121)
    summary_label.space_after = Pt(5)

    summary_text = frame.add_paragraph()
    summary_text.text = summary
    summary_text.font.name = SANS
    summary_text.font.size = Pt(15)
    summary_text.font.color.rgb = RGBColor(32, 32, 32)
    summary_text.space_after = Pt(12)
    summary_text.line_spacing = 1.1

    label = frame.add_paragraph()
    label.text = "Presenter cues"
    label.font.name = SANS
    label.font.size = Pt(16)
    label.font.bold = True
    label.font.color.rgb = RGBColor(31, 78, 121)
    label.space_after = Pt(5)

    for sentence in re.split(r"(?<=[.!?])\s+(?=[A-Z])", notes.strip()):
        point = frame.add_paragraph()
        point.text = f"• {sentence}"
        point.font.name = SANS
        point.font.size = Pt(15)
        point.font.color.rgb = RGBColor(32, 32, 32)
        point.space_after = Pt(7)
        point.line_spacing = 1.1


def new_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = BG
    return slide


def card(slide, x, y, w, h, eyebrow, heading, body, accent=BLUE):
    rect(slide, x, y, w, h, PANEL, radius=True, line=PANEL_2)
    rect(slide, x, y, 0.07, h, accent)
    text(slide, eyebrow.upper(), x + 0.28, y + 0.25, w - 0.5, 0.3,
         size=10, color=accent, bold=True)
    text(slide, heading, x + 0.28, y + 0.65, w - 0.56, 0.72,
         size=19, bold=True)
    text(slide, body, x + 0.28, y + 1.42, w - 0.56, h - 1.66,
         size=13, color=MUTED, margin=0.02)


def build():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    prs.core_properties.title = "GitHub Actions Office Hour — August and September 2026"
    prs.core_properties.subject = "GitHub Actions product updates, demos, and performance tips"
    prs.core_properties.author = "GitHub Office Hours"

    # 1
    s = new_slide(prs)
    badge(s, "OFFICE HOUR", 0.75, 0.63, PURPLE, 1.4)
    text(s, "GitHub Actions", 0.75, 1.35, 11.5, 0.9, size=44, bold=True)
    text(s, "What changed in August + September 2026", 0.75, 2.22, 11.5, 0.65,
         size=28, color=BLUE, bold=True)
    text(s, "Six updates · two live demos · five performance moves", 0.78, 3.25,
         9.5, 0.45, size=18, color=MUTED)
    rect(s, 0.78, 4.05, 11.7, 1.35, PANEL, radius=True, line=PANEL_2)
    text(s, "45 min", 1.1, 4.42, 1.6, 0.45, size=26, color=GREEN, bold=True)
    text(s, "presentation + demos", 2.5, 4.48, 3.0, 0.35, size=16)
    text(s, "15 min", 6.55, 4.42, 1.6, 0.45, size=26, color=PURPLE, bold=True)
    text(s, "Q&A", 8.0, 4.48, 2.0, 0.35, size=16)
    text(s, "September 2026", 0.78, 6.75, 3, 0.3, size=12, color=MUTED)
    add_notes(s, "1 minute", "This session covers six GitHub Actions updates from August and early September 2026, two practical demonstrations, and five workflow performance techniques.", "Welcome everyone. Set the promise: practical changes, two small demos, and actions attendees can take this week. The Q&A clock starts after the 45-minute presentation.")

    # 2
    s = new_slide(prs)
    title(s, "Run of show", "Today", 2)
    segments = [
        ("00–05", "Why these updates matter", "The operating themes"),
        ("05–17", "August releases", "Runners, retention, reporting"),
        ("17–29", "September releases", "API, permissions, identity"),
        ("29–40", "Two demos", "Small YAML, visible result"),
        ("40–45", "Performance tips", "Five high-leverage moves"),
        ("45–60", "Q&A", "Your workflows and constraints"),
    ]
    for i, (time, head, body) in enumerate(segments):
        x = 0.8 + (i % 3) * 4.15
        y = 1.95 + (i // 3) * 2.15
        card(s, x, y, 3.75, 1.65, time, head, body, [BLUE, PURPLE, GREEN][i % 3])
    footer(s)
    add_notes(s, "2 minutes", "The 60-minute office hour reserves 45 minutes for product updates, demonstrations, and performance guidance, followed by 15 minutes for audience questions.", "Describe the cadence. Tell attendees the demos are intentionally short and reproducible. Invite them to save detailed environment-specific questions for Q&A.")

    # 3
    s = new_slide(prs)
    title(s, "The release theme: know more, grant less", "Executive summary", 3)
    card(s, 0.8, 1.95, 3.75, 3.8, "Observe", "More identity", "Reusable jobs can identify the exact workflow file, repository, ref, and SHA that defined them.", BLUE)
    card(s, 4.8, 1.95, 3.75, 3.8, "Control", "Less privilege", "GITHUB_TOKEN can read Dependabot alerts through a narrow, purpose-built permission.", PURPLE)
    card(s, 8.8, 1.95, 3.75, 3.8, "Operate", "Clearer lifecycle", "Runner deprecation dates, broader retention rules, and cleaner Code Quality attribution reduce surprises.", GREEN)
    text(s, "The practical outcome: safer reusable automation with better operational signals.", 1.1, 6.35, 11.1, 0.45,
         size=20, bold=True, align=PP_ALIGN.CENTER)
    footer(s)
    add_notes(s, "2 minutes", "The releases improve three operational qualities: workflow provenance, least-privilege access, and predictable lifecycle management for runners and workflow data.", "Frame all six announcements as an operating-model improvement rather than a feature dump. The common thread is explicit provenance, narrower access, and predictable lifecycle management.")

    # 4
    s = new_slide(prs)
    title(s, "August: three operational changes", "Release timeline", 4)
    items = [
        ("AUG 20", "Windows 11 Arm64 + VS 2026", "GA image and migration path", GREEN),
        ("AUG 20", "Code Quality workflow identity", "Dedicated path and actor", BLUE),
        ("AUG 27", "Retention scope expands", "Checks, runs, statuses join logs and artifacts", ORANGE),
    ]
    for i, (date, head, body, color) in enumerate(items):
        y = 1.95 + i * 1.45
        rect(s, 0.9, y, 1.3, 0.55, color, radius=True)
        text(s, date, 0.9, y + 0.12, 1.3, 0.28, size=12, color=BG, bold=True, align=PP_ALIGN.CENTER)
        rect(s, 2.55, y + 0.25, 0.5, 0.05, PANEL_2)
        text(s, head, 3.25, y - 0.02, 4.8, 0.42, size=21, bold=True)
        text(s, body, 3.25, y + 0.47, 7.7, 0.35, size=15, color=MUTED)
    footer(s)
    add_notes(s, "3 minutes", "August introduced a new Windows Arm64 runner image, distinct reporting identity for Code Quality, and a broader Actions retention policy.", "Preview the three August announcements. Emphasize that two require review of existing assumptions: VS 2022 dependencies and long-lived run history.")

    # 5
    s = new_slide(prs)
    title(s, "Windows 11 Arm64 moves to Visual Studio 2026", "August · Hosted runners", 5)
    badge(s, "GA", 0.85, 1.93, GREEN, 0.7)
    code_box(s, "jobs:\n  build-arm:\n    runs-on: windows-11-vs2026-arm\n    steps:\n      - uses: actions/checkout@v4\n      - run: dotnet build", 0.85, 2.45, 5.2, 3.2, 16)
    bullet_list(s, [
        "Available on standard and larger hosted runners",
        "windows-11-arm migrates September 21–30",
        "VS 2022 dependencies are the breaking-change risk",
        "Pin the preview label now to test before migration",
    ], 6.55, 2.05, 5.8, 3.7, size=19, gap=13)
    rect(s, 6.55, 5.88, 5.7, 0.72, PANEL_2, radius=True)
    rich_text(s, [("ACTION: ", ORANGE, True, SANS), ("test native toolchains and installed workloads.", WHITE, False, SANS)],
              6.82, 6.08, 5.1, 0.3, size=15)
    footer(s)
    add_notes(s, "3 minutes", "The Windows 11 Arm64 image with Visual Studio 2026 is generally available as windows-11-vs2026-arm, while windows-11-arm migrates automatically during September 21–30.", "Explain that Arm64 is the architecture and VS 2026 is the image change. Teams using C++, .NET workloads, SDK components, or hard-coded Visual Studio paths should test explicitly. The migration is automatic for windows-11-arm.")

    # 6
    s = new_slide(prs)
    title(s, "One retention dial will govern five data types", "August · Retention", 6)
    labels = [("Checks", BLUE), ("Workflow runs", PURPLE), ("Statuses", GREEN), ("Artifacts", ORANGE), ("Logs", RED)]
    for i, (label, color) in enumerate(labels):
        x = 0.85 + i * 2.47
        rect(s, x, 2.0, 2.1, 0.85, PANEL, radius=True, line=color)
        text(s, label, x, 2.26, 2.1, 0.3, size=16, bold=True, align=PP_ALIGN.CENTER)
    text(s, "Your configured Actions retention period", 2.3, 3.35, 8.7, 0.45,
         size=26, bold=True, align=PP_ALIGN.CENTER)
    text(s, "↓", 6.2, 3.85, 0.8, 0.6, size=34, color=PURPLE, bold=True, align=PP_ALIGN.CENTER)
    rect(s, 3.2, 4.55, 6.9, 1.15, PANEL_2, radius=True)
    text(s, "Default: 90 days", 3.4, 4.8, 3.0, 0.38, size=22, color=GREEN, bold=True)
    text(s, "Effective October 1, 2026", 6.45, 4.86, 3.35, 0.3, size=16, color=WHITE)
    text(s, "Public repositories: 90-day maximum · Deleted history is not restored by later increases", 1.2, 6.15, 10.9, 0.4,
         size=15, color=MUTED, align=PP_ALIGN.CENTER)
    footer(s)
    add_notes(s, "4 minutes", "Starting October 1, 2026, one Actions retention setting controls checks, workflow runs, statuses, artifacts, and logs; the default is 90 days.", "Previously, checks, runs, and statuses could remain for 400+ days regardless of artifact/log retention. Starting October 1, the same setting governs all five. Ask: does anyone rely on old run URLs for audits, support, or release evidence? Export what must outlive the configured window. Metadata is not billed, but artifacts and logs are.")

    # 7
    s = new_slide(prs)
    title(s, "Code Quality runs get their own identity", "August · Reporting", 7)
    text(s, "Before", 0.9, 1.95, 5.4, 0.4, size=18, color=MUTED, bold=True)
    code_box(s, "path:  dynamic/github-code-scanning/codeql\nactor: github-advanced-security", 0.9, 2.45, 5.35, 1.35, 15)
    text(s, "→", 6.3, 2.75, 0.7, 0.5, size=32, color=PURPLE, bold=True, align=PP_ALIGN.CENTER)
    text(s, "Now", 7.1, 1.95, 5.3, 0.4, size=18, color=GREEN, bold=True)
    code_box(s, "path:  dynamic/github-code-quality/codeql\nactor: github-code-quality", 7.1, 2.45, 5.35, 1.35, 15)
    bullet_list(s, [
        "Run history and usage reports now separate quality from security",
        "Update billing filters, dashboards, and actor-based scripts",
        "No Code Quality repository reconfiguration is required",
    ], 1.2, 4.35, 11.0, 1.75, size=19, gap=12)
    footer(s)
    add_notes(s, "2 minutes", "GitHub Code Quality Actions runs now use a dedicated dynamic path and github-code-quality actor, separating quality analysis from code-scanning activity in reports.", "This is primarily a reporting and automation compatibility change. Existing analysis keeps running. The migration task is to update anything that classifies a run by the old dynamic path or github-advanced-security actor.")

    # 8
    s = new_slide(prs)
    title(s, "September: visibility and least privilege", "Month to date · Sep 6", 8)
    card(s, 0.8, 1.95, 3.75, 3.65, "API", "Runner EOL dates", "Ask GitHub when a runner version stops registering and running.", BLUE)
    card(s, 4.8, 1.95, 3.75, 3.65, "Token", "Dependabot read scope", "Read vulnerability alerts without granting a broader permission.", PURPLE)
    card(s, 8.8, 1.95, 3.75, 3.65, "Context", "Reusable job source", "Identify the reusable workflow that actually defines the job.", GREEN)
    text(s, "All three reduce hidden assumptions in shared automation.", 1.3, 6.25, 10.7, 0.45,
         size=21, bold=True, align=PP_ALIGN.CENTER)
    footer(s)
    add_notes(s, "2 minutes", "Early September added runner end-of-life visibility, a narrow Dependabot-alert permission, and source-identity properties for reusable workflows.", "Introduce the September group. These are additive features and can be adopted independently. The last two are the demo targets because they produce a visible result with very little YAML.")

    # 9
    s = new_slide(prs)
    title(s, "Automate runner upgrade planning", "September · REST API", 9)
    code_box(s, "GET /orgs/{org}/actions/runners/\n    deprecations/{version}", 0.85, 2.0, 5.0, 1.45, 18)
    code_box(s, '{\n  "runner_version": "2.x.y",\n  "registration_deprecates_at": "...",\n  "runtime_deprecates_at": "..."\n}', 0.85, 3.75, 5.0, 2.05, 15)
    bullet_list(s, [
        "Repository, organization, or enterprise scope",
        "Registration cutoff: new runner connections",
        "Runtime cutoff: existing runner execution",
        "Feed dates into inventory alerts and upgrade SLAs",
    ], 6.4, 2.02, 5.8, 3.5, size=19, gap=14)
    rect(s, 6.45, 5.75, 5.55, 0.72, PANEL_2, radius=True)
    text(s, "Best use: proactive fleet health checks", 6.7, 5.96, 5.0, 0.3,
         size=16, color=GREEN, bold=True, align=PP_ALIGN.CENTER)
    footer(s)
    add_notes(s, "3 minutes", "A new REST API reports registration and runtime deprecation dates for a specific Actions runner version at repository, organization, or enterprise scope.", "Differentiate runner binary versions from hosted runner image labels. This endpoint is most valuable for self-hosted fleets. A scheduled inventory process can query every version in use and alert before either cutoff.")

    # 10
    s = new_slide(prs)
    title(s, "Read alerts without opening the whole security drawer", "September · GITHUB_TOKEN", 10)
    code_box(s, "permissions:\n  contents: read\n  vulnerability-alerts: read", 0.9, 2.05, 5.25, 1.75, 19)
    text(s, "Purpose-built access", 6.65, 2.05, 5.1, 0.45, size=23, bold=True)
    bullet_list(s, [
        "Supports read or none",
        "Works with the workflow-scoped GITHUB_TOKEN",
        "Avoids a PAT and reduces permission surface",
        "Can be set at workflow or job scope",
    ], 6.6, 2.68, 5.5, 2.6, size=18, gap=12)
    rect(s, 0.9, 4.35, 5.25, 1.15, PANEL_2, radius=True)
    text(s, "Default closed", 1.15, 4.58, 1.6, 0.3, size=16, color=PURPLE, bold=True)
    text(s, "Declare only where needed.", 2.65, 4.58, 3.1, 0.3, size=16)
    rich_text(s, [("Pattern: ", BLUE, True, SANS), ("workflow defaults + job-level elevation", WHITE, False, SANS)],
              1.2, 6.02, 10.8, 0.35, size=19, align=PP_ALIGN.CENTER)
    footer(s)
    add_notes(s, "4 minutes", "The new vulnerability-alerts permission gives GITHUB_TOKEN read-only access to Dependabot alerts without requiring a personal access token or broader security scope.", "Show the permission name and emphasize job-level scoping. A good pattern is contents: read at workflow level, then grant vulnerability-alerts: read only to the reporting job. This is simpler to audit than a PAT and limits blast radius.")

    # 11
    s = new_slide(prs)
    title(s, "Reusable workflows can identify their own source", "September · Job context", 11)
    props = [
        ("job.workflow_ref", "full workflow ref"),
        ("job.workflow_sha", "defining commit"),
        ("job.workflow_repository", "owner/repo"),
        ("job.workflow_file_path", "repo-relative path"),
    ]
    for i, (key, desc) in enumerate(props):
        x = 0.85 + (i % 2) * 6.05
        y = 1.95 + (i // 2) * 1.45
        rect(s, x, y, 5.55, 1.1, PANEL, radius=True, line=PANEL_2)
        text(s, key, x + 0.25, y + 0.2, 3.6, 0.3, size=16, color=BLUE, bold=True, font=MONO)
        text(s, desc, x + 0.25, y + 0.63, 4.8, 0.25, size=14, color=MUTED)
    text(s, "Caller identity", 1.05, 5.1, 2.0, 0.3, size=15, color=MUTED, bold=True)
    text(s, "github.workflow_ref", 3.0, 5.1, 3.0, 0.3, size=16, font=MONO)
    text(s, "≠", 6.15, 5.05, 0.7, 0.4, size=23, color=PURPLE, bold=True, align=PP_ALIGN.CENTER)
    text(s, "job.workflow_ref", 7.0, 5.1, 3.0, 0.3, size=16, font=MONO)
    text(s, "Defining identity", 10.1, 5.1, 2.0, 0.3, size=15, color=MUTED, bold=True)
    text(s, "Use cases: provenance · policy · diagnostics · version reporting", 1.0, 6.2, 11.3, 0.4,
         size=18, bold=True, align=PP_ALIGN.CENTER)
    footer(s)
    add_notes(s, "3 minutes", "Four new job context properties expose the ref, SHA, repository, and file path of the workflow that defines a job, including jobs in reusable workflows.", "The github values describe the caller. The job values describe the workflow file defining the current job. They match for an ordinary workflow and diverge for a reusable one. This removes the need to pass source metadata manually. Not available on GitHub Enterprise Server.")

    # 12
    s = new_slide(prs)
    title(s, "Demo setup: one dispatch, two visible results", "Live demo", 12)
    text(s, "Office hours demo", 0.9, 2.0, 3.1, 0.42, size=22, bold=True)
    rect(s, 1.1, 2.75, 3.1, 1.0, PANEL, radius=True, line=BLUE)
    text(s, "workflow_dispatch", 1.1, 3.07, 3.1, 0.3, size=17, color=BLUE, bold=True, font=MONO, align=PP_ALIGN.CENTER)
    text(s, "→", 4.45, 2.96, 0.7, 0.45, size=30, color=PURPLE, bold=True, align=PP_ALIGN.CENTER)
    rect(s, 5.35, 2.25, 3.0, 1.25, PANEL, radius=True, line=GREEN)
    text(s, "Reusable identity", 5.35, 2.66, 3.0, 0.3, size=17, bold=True, align=PP_ALIGN.CENTER)
    rect(s, 5.35, 4.05, 3.0, 1.25, PANEL, radius=True, line=PURPLE)
    text(s, "Alert count", 5.35, 4.46, 3.0, 0.3, size=17, bold=True, align=PP_ALIGN.CENTER)
    text(s, "→", 8.65, 2.96, 0.7, 0.45, size=30, color=PURPLE, bold=True, align=PP_ALIGN.CENTER)
    rect(s, 9.55, 2.75, 2.8, 1.0, PANEL_2, radius=True)
    text(s, "Job summaries", 9.55, 3.07, 2.8, 0.3, size=17, bold=True, align=PP_ALIGN.CENTER)
    text(s, "Fallback: keep one successful run open before the session.", 1.5, 6.08, 10.3, 0.4,
         size=17, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)
    footer(s)
    add_notes(s, "1 minute", "A single manually dispatched workflow runs two independent jobs to demonstrate reusable-workflow provenance and least-privilege Dependabot access.", "Switch to the repository. Show the two small files first, then dispatch the workflow. While it starts, explain that the two jobs are independent. Keep a pre-run result open in case of queue delays.")

    # 13
    s = new_slide(prs)
    title(s, "Demo 1: caller versus defining workflow", "Live demo · 4 minutes", 13)
    code_box(s, "jobs:\n  reusable-identity:\n    uses: ./.github/workflows/reusable-source.yml", 0.85, 2.0, 5.35, 1.7, 18)
    code_box(s, "github.workflow_ref\njob.workflow_ref\njob.workflow_sha\njob.workflow_repository\njob.workflow_file_path", 6.65, 2.0, 5.8, 2.6, 17)
    rect(s, 0.85, 4.25, 5.35, 1.3, PANEL_2, radius=True)
    text(s, "Observe", 1.12, 4.52, 1.3, 0.3, size=15, color=GREEN, bold=True)
    text(s, "The two workflow refs diverge.", 2.25, 4.52, 3.5, 0.3, size=16)
    rect(s, 6.65, 5.05, 5.8, 0.75, PANEL_2, radius=True)
    text(s, "Land the point: provenance is now native.", 6.9, 5.27, 5.3, 0.3,
         size=17, color=BLUE, bold=True, align=PP_ALIGN.CENTER)
    footer(s)
    add_notes(s, "4 minutes", "The reusable-workflow demo shows that github.workflow_ref identifies the caller while job.workflow_ref and related properties identify the workflow defining the job.", "Open both YAML files. Show the caller's uses line. In the run summary, compare github.workflow_ref with job.workflow_ref, then point out SHA, repository, and file path. Ask attendees where they currently pass this metadata as inputs. Close with provenance and diagnostics use cases.")

    # 14
    s = new_slide(prs)
    title(s, "Demo 2: least privilege in three lines", "Live demo · 3 minutes", 14)
    code_box(s, "permissions:\n  contents: read\n  vulnerability-alerts: read", 0.9, 2.0, 4.8, 1.65, 19)
    code_box(s, 'env:\n  GH_TOKEN: ${{ github.token }}\nrun: |\n  gh api "repos/$GITHUB_REPOSITORY/\n    dependabot/alerts?state=open"', 6.15, 2.0, 6.15, 2.8, 15)
    rect(s, 0.9, 4.2, 4.8, 1.15, PANEL_2, radius=True)
    text(s, "No PAT", 1.15, 4.48, 1.2, 0.3, size=17, color=GREEN, bold=True)
    text(s, "No broad token scope", 2.4, 4.48, 2.8, 0.3, size=16)
    text(s, "Expected output: “Open Dependabot alerts: N”", 2.0, 6.0, 9.5, 0.42,
         size=20, bold=True, align=PP_ALIGN.CENTER)
    footer(s)
    add_notes(s, "3 minutes", "The least-privilege demo uses a job-scoped vulnerability-alerts: read grant and the built-in GITHUB_TOKEN to count open Dependabot alerts.", "Highlight that the permission is scoped to only this job. Open the job summary and show the count. Zero is a valid result. If policy prevents access, explain that the explicit failure is preferable to a silent or overly broad fallback.")

    # 15
    s = new_slide(prs)
    title(s, "Five performance moves with outsized returns", "Final five minutes", 15)
    tips = [
        ("1", "Cancel stale work", "Use concurrency groups with cancel-in-progress for PRs.", BLUE),
        ("2", "Cache with intent", "Key on lockfiles; restore broadly; measure hit rate.", PURPLE),
        ("3", "Filter early", "Use paths, branches, and job if conditions before runners start.", GREEN),
        ("4", "Shrink the matrix", "Test representative combinations first; expand on main or nightly.", ORANGE),
        ("5", "Measure the queue", "Separate queue time from execution time before changing code.", RED),
    ]
    for i, (num, head, body, color) in enumerate(tips):
        y = 1.86 + i * 0.95
        rect(s, 0.85, y, 0.58, 0.58, color, radius=True)
        text(s, num, 0.85, y + 0.14, 0.58, 0.24, size=13, color=BG, bold=True, align=PP_ALIGN.CENTER)
        text(s, head, 1.68, y - 0.01, 2.8, 0.34, size=18, bold=True)
        text(s, body, 4.25, y + 0.02, 7.9, 0.32, size=15, color=MUTED)
    rect(s, 0.85, 6.63, 11.55, 0.42, PANEL_2, radius=True)
    text(s, "Optimize total feedback time—not just step duration.", 0.85, 6.73, 11.55, 0.22,
         size=13, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    footer(s)
    add_notes(s, "5 minutes", "The highest-leverage performance improvements eliminate stale runs, reuse expensive deterministic work, prevent unnecessary jobs, control matrix growth, and distinguish queue time from execution time.", "Spend about 45 seconds on each tip. Concurrency avoids paying for obsolete commits. Cache only expensive deterministic inputs. Filters prevent runner allocation entirely. Matrices often grow quadratically. Finally, inspect queue time separately: a faster build cannot fix runner scarcity.")

    # 16
    s = new_slide(prs)
    title(s, "Three actions for next week", "Close", 16)
    card(s, 0.8, 2.0, 3.75, 3.4, "Test", "Arm64 + VS 2026", "Run architecture-specific builds on the explicit image before automatic migration.", GREEN)
    card(s, 4.8, 2.0, 3.75, 3.4, "Review", "Retention + reports", "Confirm the retention window and update Code Quality path or actor filters.", ORANGE)
    card(s, 8.8, 2.0, 3.75, 3.4, "Adopt", "Identity + permission", "Use native reusable-workflow provenance and job-scoped alert access.", BLUE)
    text(s, "Q&A", 0.8, 6.05, 11.75, 0.62, size=34, color=PURPLE, bold=True, align=PP_ALIGN.CENTER)
    text(s, "github.blog/changelog/label/actions", 0.8, 6.72, 11.75, 0.3,
         size=13, color=MUTED, font=MONO, align=PP_ALIGN.CENTER)
    footer(s)
    add_notes(s, "3 minutes, then 15 minutes Q&A", "The immediate next steps are to test the new Arm64 image, review retention and reporting dependencies, and adopt narrower permissions plus native workflow provenance.", "Recap the three action categories. Transition to questions. Useful prompts: Which runner fleet is hardest to maintain? Where do shared workflows lack provenance? What is the longest feedback-time bottleneck today?")

    prs.save(OUT)
    print(f"Wrote {OUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    build()
