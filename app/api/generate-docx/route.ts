import { Document, Packer, Paragraph, HeadingLevel, AlignmentType, PageBreak } from "docx"

interface DocumentInput {
  title: string
  problemStatement: string
  objectives: string[]
  requirements: string[]
  manualSteps: string[]
  automationIdeas: string
}

export async function POST(request: Request) {
  try {
    const body: DocumentInput = await request.json()

    // Validate input
    if (!body.title || typeof body.title !== "string") {
      return Response.json({ error: "Invalid input: title is required and must be a string" }, { status: 400 })
    }

    if (!Array.isArray(body.objectives) || !Array.isArray(body.requirements) || !Array.isArray(body.manualSteps)) {
      return Response.json(
        { error: "Invalid input: objectives, requirements, and manualSteps must be arrays" },
        { status: 400 },
      )
    }

    // Create sections following PDD template structure
    const sections = [
      // Header with footer text
      new Paragraph({
        text: "Created & Managed by Team, for queries write us at ABC@XYZ.COM",
        alignment: AlignmentType.CENTER,
        size: 18,
        spacing: { after: 400 },
      }),

      // Title section
      new Paragraph({
        text: `Process: TMPY HR`,
        size: 20,
        spacing: { after: 100 },
      }),
      new Paragraph({
        text: `Project Name: ${body.title}`,
        size: 20,
        spacing: { after: 100 },
      }),
      new Paragraph({
        text: `Nov | 2024`,
        size: 20,
        spacing: { after: 600 },
      }),

      new Paragraph({
        text: "Process Definition Document",
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        bold: true,
        size: 28,
        spacing: { after: 600 },
      }),

      // Footer line
      new Paragraph({
        text: "Created & Managed by Team, for queries write us at ABC@XYZ.COM",
        alignment: AlignmentType.CENTER,
        size: 18,
        spacing: { after: 400 },
        pageBreakBefore: true,
      }),

      // Table of Contents
      new Paragraph({
        text: "1 Contents",
        heading: HeadingLevel.HEADING_2,
        bold: true,
        size: 22,
        spacing: { after: 200 },
      }),
      createTOCItem("2 INTRODUCTION", 3),
      createTOCItem("2.1 PURPOSE", 3),
      createTOCItem("2.2 OBJECTIVE", 3),
      createTOCItem("2.3 PROCESS KEY CONTACTS", 3),
      createTOCItem("2.4 DOCUMENT CONTROL", 3),
      createTOCItem("3 CHANGE REQUESTS", 3),
      createTOCItem("4 AS IS PROCESS DESCRIPTION", 3),
      createTOCItem("4.1 PROCESS OVERVIEW", 4),
      createTOCItem("4.2 RACI MATRIX", 4),
      createTOCItem("4.3 MINIMUM PRE-REQUISITES FOR THE AUTOMATION", 4),
      createTOCItem("4.4 APPLICATION USED IN THE PROCESS", 4),
      createTOCItem("4.5 AS-IS PROCESS MAP", 4),
      createTOCItem("4.5.1 High Level AS-IS Process Map", 5),
      createTOCItem("4.5.2 Detailed AS-IS Process Map", 5),
      createTOCItem("4.6 VOLUMETRIC", 4),
      createTOCItem("4.7 VOLUME AND HANDLING TIME", 4),
      createTOCItem("4.8 OPERATING WINDOW & STAFFING SCHEDULE", 4),
      createTOCItem("4.9 INPUT DATA DETAILS", 4),
      createTOCItem("5 TO BE PROCESS DESCRIPTION", 3),
      createTOCItem("5.1 TO BE DETAILED PROCESS MAP", 4),
      createTOCItem("5.2 PARALLEL INITIATIVES/ AUTOMATION/ DEVELOPMENT", 4),
      createTOCItem("5.3 IN SCOPE SCENARIOS/CASE TYPES/VOLUME", 4),
      createTOCItem("5.4 OUT OF SCOPE SCENARIOS/CASE TYPES/VOLUME FOR PROJECT", 4),
      createTOCItem("5.5 EXCEPTION HANDLING", 4),
      createTOCItem("5.5.1 Known Business Exception", 5),
      createTOCItem("5.5.2 Unknown Business Exception", 5),
      createTOCItem("5.6 APPLICATIONS ERRORS & EXCEPTIONS HANDLING", 4),
      createTOCItem("5.6.1 Known Applications Errors and Exceptions", 5),
      createTOCItem("5.6.2 Unknown Applications Errors and Exceptions", 5),
      createTOCItem("5.7 REPORTING", 4),
      createTOCItem("6 OTHER", 3),
      createTOCItem("6.1 APPENDIX & OTHER DOCUMENTS", 4),

      new Paragraph({
        text: "",
        spacing: { after: 400 },
      }),

      // Page break before main content
      new PageBreak(),

      // Footer line
      new Paragraph({
        text: "Created & Managed by Team, for queries write us at ABC@XYZ.COM",
        alignment: AlignmentType.CENTER,
        size: 18,
        spacing: { after: 400 },
      }),

      // INTRODUCTION section
      new Paragraph({
        text: "2 INTRODUCTION",
        heading: HeadingLevel.HEADING_1,
        bold: true,
        size: 26,
        spacing: { before: 400, after: 300 },
      }),

      // Problem Statement
      new Paragraph({
        text: "1. Problem Statement",
        heading: HeadingLevel.HEADING_2,
        bold: true,
        size: 22,
        spacing: { before: 200, after: 150 },
      }),
      new Paragraph({
        text: body.problemStatement || "No problem statement provided.",
        size: 22,
        spacing: { after: 300 },
      }),

      // Objectives
      new Paragraph({
        text: "2. Objectives",
        heading: HeadingLevel.HEADING_2,
        bold: true,
        size: 22,
        spacing: { before: 200, after: 150 },
      }),
      ...createBulletList(body.objectives || []),

      // Requirements
      new Paragraph({
        text: "3. Requirements",
        heading: HeadingLevel.HEADING_2,
        bold: true,
        size: 22,
        spacing: { before: 200, after: 150 },
      }),
      ...createBulletList(body.requirements || []),

      // AS-IS Process Map
      new Paragraph({
        text: "4. AS-IS Process Map",
        heading: HeadingLevel.HEADING_2,
        bold: true,
        size: 22,
        spacing: { before: 200, after: 150 },
      }),
      ...createNumberedList(body.manualSteps || []),

      // Automation Ideas
      new Paragraph({
        text: "5. Automation Ideas",
        heading: HeadingLevel.HEADING_2,
        bold: true,
        size: 22,
        spacing: { before: 200, after: 150 },
      }),
      new Paragraph({
        text: body.automationIdeas || "No automation ideas provided.",
        size: 22,
        spacing: { after: 600 },
      }),

      // Footer on last section
      new Paragraph({
        text: "",
        spacing: { after: 400 },
      }),
      new Paragraph({
        text: "Created & Managed by Team, for queries write us at ABC@XYZ.COM",
        alignment: AlignmentType.CENTER,
        size: 18,
      }),
    ]

    // Create document
    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margins: {
                top: 1440, // 1 inch
                right: 1440,
                bottom: 1440,
                left: 1440,
              },
            },
          },
          children: sections,
        },
      ],
    })

    // Generate buffer and convert to base64 (works in Node and edge/browser runtimes)
    const buffer = await Packer.toBuffer(doc)
    const uint8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer as any)

    function uint8ToBase64(u8: Uint8Array) {
      const Global: any = globalThis as any
      if (typeof Global.Buffer !== "undefined") return Global.Buffer.from(u8).toString("base64")
      let binary = ""
      const chunkSize = 0x8000
      for (let i = 0; i < u8.length; i += chunkSize) {
        binary += String.fromCharCode(...u8.subarray(i, i + chunkSize))
      }
      return btoa(binary)
    }

    const base64 = uint8ToBase64(uint8)

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, -5)
    const fileName = `document-${timestamp}.docx`

    return Response.json(
      {
        fileName,
        fileBase64: base64,
      },
      { status: 200 },
    )
  } catch (error) {
    console.error("Error generating DOCX:", error)
    return Response.json(
      { error: "Failed to generate document. Please check your input and try again." },
      { status: 500 },
    )
  }
}

// Helper function to create bullet list items
function createBulletList(items: string[]): any[] {
  return items
    .filter((item) => item && item.trim())
    .map(
      (item) =>
        new Paragraph({
          text: item,
          bullet: {
            level: 0,
          },
          spacing: { after: 100 },
          size: 22,
        }),
    )
}

// Helper function to create numbered list items
function createNumberedList(items: string[]): any[] {
  return items
    .filter((item) => item && item.trim())
    .map(
      (item, index) =>
        new Paragraph({
          text: item,
          numbering: {
            num: 1,
            level: 0,
          },
          spacing: { after: 100 },
          size: 22,
        }),
    )
}

// Helper function for table of contents formatting
function createTOCItem(text: string, indentLevel: number): any {
  return new Paragraph({
    text: text,
    spacing: { after: 100 },
    size: 20,
    indent: { left: (indentLevel - 1) * 400 },
  })
}
