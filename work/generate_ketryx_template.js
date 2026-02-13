const { Document, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, 
        AlignmentType, WidthType, BorderStyle, PageBreak, convertInchesToTwip, Packer } = require('docx');
const fs = require('fs');

// Helper to create bordered cell
function createCell(content, options = {}) {
    const borders = {
        top: { style: BorderStyle.SINGLE, size: 8 },
        bottom: { style: BorderStyle.SINGLE, size: 8 },
        left: { style: BorderStyle.SINGLE, size: 8 },
        right: { style: BorderStyle.SINGLE, size: 8 },
    };
    
    return new TableCell({
        children: Array.isArray(content) ? content : [new Paragraph(content)],
        borders,
        ...options
    });
}

// Helper for header cells
function createHeaderCell(text, options = {}) {
    return createCell([
        new Paragraph({ 
            children: [new TextRun({ text, bold: true })],
            alignment: AlignmentType.CENTER
        })
    ], options);
}

// Create document
const doc = new Document({
    sections: [{
        properties: {
            page: {
                margin: {
                    top: convertInchesToTwip(1),
                    right: convertInchesToTwip(1),
                    bottom: convertInchesToTwip(1),
                    left: convertInchesToTwip(1),
                }
            }
        },
        children: [
            // Title
            new Paragraph({
                text: "Defect Summary Report",
                heading: HeadingLevel.TITLE,
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 }
            }),
            
            new PageBreak(),
            
            // Document History
            new Paragraph({
                text: "Document History",
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 240, after: 240 }
            }),
            
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            createHeaderCell("Version", { width: { size: 25, type: WidthType.PERCENTAGE }}),
                            createHeaderCell("Description", { width: { size: 75, type: WidthType.PERCENTAGE }})
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("01"),
                            createCell("Initial release")
                        ],
                    }),
                ],
            }),
            
            // Table of Contents
            new Paragraph({
                text: "Table of Content",
                heading: HeadingLevel.HEADING_3,
                spacing: { before: 480, after: 240 }
            }),
            
            new Paragraph({
                text: "{@toc}",
                spacing: { after: 240 }
            }),
            
            new PageBreak(),
            
            // 1. Purpose
            new Paragraph({
                text: "Purpose",
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 240, after: 240 }
            }),
            
            new Paragraph({
                text: "This Defect Summary Report summarizes the Defects found in the {project.name} {version.name} in accordance with MQMS-D&D-GSP-05 Defect Management.",
                indent: { left: convertInchesToTwip(0.5) },
                spacing: { before: 80, after: 240 }
            }),
            
            // 2. Scope
            new Paragraph({
                text: "Scope", 
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 480, after: 240 }
            }),
            
            new Paragraph({
                text: "This document covers Accepted, Resolved and Rejected Defects applicable for:",
                indent: { left: convertInchesToTwip(0.5) },
                spacing: { before: 80, after: 80 }
            }),
            
            new Paragraph({ text: "" }),
            new Paragraph({ text: "• {project.name} {version.name}" }),
            new Paragraph({ text: "• MDT – Meeting Management {version.name}" }),
            new Paragraph({ text: "• OSC – Shared Capabilities {version.name}" }),
            new Paragraph({ text: "", spacing: { after: 240 } }),
            
            // 3. Summary
            new Paragraph({
                text: "Summary",
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 480, after: 240 }
            }),
            
            new Paragraph({
                text: "Table 1 below summarizes all defects in states Accepted, Resolved or Rejected, found in the current {project.name} {version.name}.",
                indent: { left: convertInchesToTwip(0.5) },
                spacing: { before: 80, after: 240 }
            }),
            
            new Paragraph({
                children: [new TextRun({ text: "Table 1: Summary of the number of defects", bold: true })],
                indent: { left: convertInchesToTwip(0.5) },
                spacing: { before: 0, after: 200 }
            }),
            
            // Load data with KQL queries
            new Paragraph({ text: "" }),
            new Paragraph({ text: "{$KQL acceptedRiskYes = type:Anomaly state:Accepted field:Risk_Related:yes}" }),
            new Paragraph({ text: "{$KQL acceptedRiskNo = type:Anomaly state:Accepted field:Risk_Related:no}" }),
            new Paragraph({ text: "{$KQL resolved = type:Anomaly state:Resolved}" }),
            new Paragraph({ text: "{$KQL rejected = type:Anomaly state:Rejected}" }),
            new Paragraph({ text: "" }),
            
            // Summary Table
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    // Header rows
                    new TableRow({
                        tableHeader: true,
                        children: [
                            createCell("", { rowSpan: 2 }),
                            createCell([
                                new Paragraph({ 
                                    text: "Accepted", 
                                    alignment: AlignmentType.CENTER, 
                                    children: [new TextRun({ text: "Accepted", bold: true })] 
                                }),
                                new Paragraph({ 
                                    text: "Potential product risk related", 
                                    alignment: AlignmentType.CENTER,
                                    children: [new TextRun({ text: "Potential product risk related", bold: true })]
                                })
                            ], { columnSpan: 2 }),
                            createHeaderCell("Resolved", { rowSpan: 2 }),
                            createHeaderCell("Rejected", { rowSpan: 2 })
                        ],
                    }),
                    new TableRow({
                        children: [
                            createHeaderCell("yes"),
                            createHeaderCell("no")
                        ],
                    }),
                    // Data row
                    new TableRow({
                        children: [
                            createCell([new Paragraph({ 
                                text: "Number of Defects", 
                                children: [new TextRun({ text: "Number of Defects", bold: true })]
                            })]),
                            createCell("{acceptedRiskYes | count}"),
                            createCell("{acceptedRiskNo | count}"),
                            createCell("{resolved | count}"),
                            createCell("{rejected | count}")
                        ],
                    }),
                ],
            }),
            
            new PageBreak(),
            
            // 4. Defects
            new Paragraph({
                text: "Defects",
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 480, after: 240 }
            }),
            
            // 4.1 Accepted Defects
            new Paragraph({
                text: "Accepted Defects",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 480, after: 240 },
                indent: { left: convertInchesToTwip(0.5) }
            }),
            
            new Paragraph({
                text: "Table 2 and Table 3 list all defects found in the current {project.name} {version.name} and previous releases that are accepted to still remain open for the current release.",
                indent: { left: convertInchesToTwip(0.5) },
                spacing: { before: 80, after: 200 }
            }),
            
            // Table 2: Accepted Defects - risk related "yes"
            new Paragraph({
                children: [new TextRun({ text: "Table 2: Accepted Defects – risk related \"yes\"", bold: true })],
                indent: { left: convertInchesToTwip(0.5) },
                spacing: { before: 0, after: 200 }
            }),
            
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        tableHeader: true,
                        children: [
                            createHeaderCell("ID"),
                            createHeaderCell("Summary"),
                            createHeaderCell("Detected in HW SN / Build"),
                            createHeaderCell("Risk Reference"),
                            createHeaderCell("Affected Requirement")
                        ],
                    }),
                    // Template row for defects loop
                    new TableRow({
                        children: [
                            createCell("{#acceptedRiskYes}{docId}"),
                            createCell("{title}"),
                            createCell("{fieldValue.Detected_Version}"),
                            createCell("{relations | where:'type == \"affects\"' | map:'other.docId' | join:', '}"),
                            createCell("{relations | where:'type == \"is related to\"' | map:'other.docId' | join:', '}{/acceptedRiskYes}")
                        ],
                    }),
                    // Summary row
                    new TableRow({
                        children: [
                            createCell("{acceptedRiskYes | count} issue{#acceptedRiskYes | count > 1}s{/acceptedRiskYes | count > 1}"),
                            createCell(""),
                            createCell(""),
                            createCell(""),
                            createCell("")
                        ],
                    }),
                ],
            }),
            
            new Paragraph({ text: "", spacing: { after: 400 } }),
            
            // Table 3: Accepted Defects - risk related "no"
            new Paragraph({
                children: [new TextRun({ text: "Table 3: Accepted Defects – risk related \"no\"", bold: true })],
                indent: { left: convertInchesToTwip(0.5) },
                spacing: { before: 0, after: 200 }
            }),
            
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        tableHeader: true,
                        children: [
                            createHeaderCell("ID"),
                            createHeaderCell("Summary"),
                            createHeaderCell("Detected in HW SN / Build"),
                            createHeaderCell("Affected Requirement")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("{#acceptedRiskNo}{docId}"),
                            createCell("{title}"),
                            createCell("{fieldValue.Detected_Version}"),
                            createCell("{relations | where:'type == \"is related to\"' | map:'other.docId' | join:', '}{/acceptedRiskNo}")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("{acceptedRiskNo | count} issue{#acceptedRiskNo | count > 1}s{/acceptedRiskNo | count > 1}"),
                            createCell(""),
                            createCell(""),
                            createCell("")
                        ],
                    }),
                ],
            }),
            
            new Paragraph({ text: "", spacing: { after: 400 } }),
            
            // 4.2 Resolved Defects
            new Paragraph({
                text: "Resolved Defects",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 480, after: 240 },
                indent: { left: convertInchesToTwip(0.5) }
            }),
            
            new Paragraph({
                text: "Table 4 lists all reported defects that have been successfully verified and confirmed to be resolved.",
                indent: { left: convertInchesToTwip(0.5) },
                spacing: { before: 80, after: 200 }
            }),
            
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        tableHeader: true,
                        children: [
                            createHeaderCell("ID"),
                            createHeaderCell("Summary"),
                            createHeaderCell("Detected in HW SN / Build"),
                            createHeaderCell("Verification/ Validation Reference")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("{#resolved}{docId}"),
                            createCell("{title}"),
                            createCell("{fieldValue.Detected_Version}"),
                            createCell("{relations | where:'type == \"executes\"' | map:'other.docId' | join:', '}{/resolved}")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("{resolved | count} issue{#resolved | count > 1}s{/resolved | count > 1}"),
                            createCell(""),
                            createCell(""),
                            createCell("")
                        ],
                    }),
                ],
            }),
            
            new Paragraph({ text: "", spacing: { after: 400 } }),
            
            // 4.3 Rejected Defects
            new Paragraph({
                text: "Rejected Defects",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 480, after: 240 },
                indent: { left: convertInchesToTwip(0.5) }
            }),
            
            new Paragraph({
                text: "Table 5 list all initially reported defects that have been identified to be invalid.",
                indent: { left: convertInchesToTwip(0.5) },
                spacing: { before: 80, after: 200 }
            }),
            
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        tableHeader: true,
                        children: [
                            createHeaderCell("ID"),
                            createHeaderCell("Summary"),
                            createHeaderCell("Detected in HW SN / Build"),
                            createHeaderCell("Rejection Justification")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("{#rejected}{docId}"),
                            createCell("{title}"),
                            createCell("{fieldValue.Detected_Version}"),
                            createCell("{fieldValue.Rejection_Justification}{/rejected}")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("{rejected | count} issue{#rejected | count > 1}s{/rejected | count > 1}"),
                            createCell(""),
                            createCell(""),
                            createCell("")
                        ],
                    }),
                ],
            }),
            
            new Paragraph({ text: "", spacing: { after: 400 } }),
            
            // 5. Glossary
            new Paragraph({
                text: "Glossary",
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 480, after: 240 }
            }),
            
            new Paragraph({
                text: "Refer to MQMS-LIB-Appendix D for additional definitions and abbreviations.",
                spacing: { before: 80, after: 200 }
            }),
            
            new Paragraph({
                children: [new TextRun({ text: "Table 6: Terms and definitions", bold: true })],
                spacing: { before: 0, after: 200 }
            }),
            
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            createHeaderCell("Abbreviation/ Term"),
                            createHeaderCell("Definition")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("Accepted Defect"),
                            createCell("see MQMS-D&D-GSP-05")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("Resolved Defect"),
                            createCell("see MQMS-D&D-GSP-05")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("Rejected Defect"),
                            createCell("see MQMS-D&D-GSP-05")
                        ],
                    }),
                ],
            }),
            
            new Paragraph({ text: "", spacing: { after: 400 } }),
            
            // 6. References
            new Paragraph({
                text: "References",
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 480, after: 240 }
            }),
            
            new Paragraph({
                children: [new TextRun({ text: "Table 7: References", bold: true })],
                spacing: { before: 0, after: 200 }
            }),
            
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            createHeaderCell("Document ID"),
                            createHeaderCell("Document Title"),
                            createHeaderCell("DMS System")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("MQMS-D&D-GSP-05"),
                            createCell("Defect Management"),
                            createCell("DiaDoc")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("MQMS-LIB-Appendix D"),
                            createCell("Roche Diagnostics Quality System Definitions"),
                            createCell("DiaDoc")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("DD009"),
                            createCell("Design and Development Plan"),
                            createCell("Enzyme")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("HH009"),
                            createCell("Hazard and Harm Review"),
                            createCell("Enzyme")
                        ],
                    }),
                    new TableRow({
                        children: [
                            createCell("{version.name}"),
                            createCell("Design Verification Test Results"),
                            createCell("Ketryx")
                        ],
                    }),
                ],
            }),
        ],
    }],
});

// Generate the document
Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync('/Users/joeycohen/Desktop/Ketryx Templating Wizard Version 3/Template.docx', buffer);
    console.log('Ketryx template document created successfully!');
}).catch((err) => {
    console.error('Error creating document:', err);
});