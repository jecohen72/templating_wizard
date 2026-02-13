const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
        AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel,
        Header, Footer, PageNumber, TableOfContents } = require('docx');
const fs = require('fs');

async function generateTemplate() {
  try {
    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440
            }
          }
        },
        children: [
          // Title
          new Paragraph({
            children: [
              new TextRun({
                text: "Defect Summary Report",
                bold: true,
                size: 32
              })
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
          }),

          // Document History Table
          new Paragraph({
            text: "Document History",
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 }
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: "Version", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Description", bold: true })],
                    shading: { fill: "E0E0E0" }
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: "01" })]
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Initial release" })]
                  })
                ]
              })
            ]
          }),

          // Table of Contents
          new Paragraph({
            text: "Table of Content",
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 }
          }),
          new Paragraph({
            children: [new TextRun({ text: "{@toc}" })]
          }),

          // Purpose
          new Paragraph({
            text: "Purpose",
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 },
            pageBreakBefore: true
          }),
          new Paragraph({
            text: "This Defect Summary Report summarizes the Defects found in the {project.name} {version.name} in accordance with MQMS-D&D-GSP-05 Defect Management.",
            spacing: { after: 200 }
          }),

          // Scope
          new Paragraph({
            text: "Scope",
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 }
          }),
          new Paragraph({
            text: "This document covers Accepted, Resolved and Rejected Defects applicable for:",
            spacing: { after: 100 }
          }),
          new Paragraph({
            text: "{project.name} {version.name}",
            spacing: { after: 50 }
          }),
          new Paragraph({
            text: "MDT – Meeting Management {version.name}",
            spacing: { after: 50 }
          }),
          new Paragraph({
            text: "OSC – Shared Capabilities {version.name}",
            spacing: { after: 200 }
          }),

          // Summary
          new Paragraph({
            text: "Summary",
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 }
          }),
          new Paragraph({
            text: "Table 1 below summarizes all defects in states Accepted, Resolved or Rejected, found in the current {project.name} {version.name}.",
            spacing: { after: 200 }
          }),

          // Summary Table with KQL queries
          new Paragraph({
            text: "Table 1: Summary of the number of defects",
            bold: true,
            spacing: { after: 100 }
          }),
          // KQL queries to get defect counts
          new Paragraph({
            text: "{$KQL acceptedRisk = type:Anomaly status:Accepted fieldValue.Risk_Related:yes}"
          }),
          new Paragraph({
            text: "{$KQL acceptedNoRisk = type:Anomaly status:Accepted fieldValue.Risk_Related:no}"
          }),
          new Paragraph({
            text: "{$KQL resolved = type:Anomaly status:Resolved}"
          }),
          new Paragraph({
            text: "{$KQL rejected = type:Anomaly status:Rejected}"
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: "", bold: true })],
                    shading: { fill: "E0E0E0" },
                    rowSpan: 2
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Accepted", bold: true, alignment: AlignmentType.CENTER })],
                    shading: { fill: "E0E0E0" },
                    columnSpan: 2
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Resolved", bold: true })],
                    shading: { fill: "E0E0E0" },
                    rowSpan: 2
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Rejected", bold: true })],
                    shading: { fill: "E0E0E0" },
                    rowSpan: 2
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: "Potential product risk related", bold: true })],
                    shading: { fill: "E0E0E0" },
                    columnSpan: 2
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: "", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "yes", alignment: AlignmentType.CENTER })]
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "no", alignment: AlignmentType.CENTER })]
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "", alignment: AlignmentType.CENTER })]
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "", alignment: AlignmentType.CENTER })]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: "Number of Defects", bold: true })],
                    shading: { fill: "F0F0F0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "{acceptedRisk | count}", alignment: AlignmentType.CENTER })]
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "{acceptedNoRisk | count}", alignment: AlignmentType.CENTER })]
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "{resolved | count}", alignment: AlignmentType.CENTER })]
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "{rejected | count}", alignment: AlignmentType.CENTER })]
                  })
                ]
              })
            ]
          }),

          // Defects Section
          new Paragraph({
            text: "Defects",
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 },
            pageBreakBefore: true
          }),

          // Accepted Defects
          new Paragraph({
            text: "Accepted Defects",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 300, after: 200 }
          }),
          new Paragraph({
            text: "Table 2 and Table 3 list all defects found in the current {project.name} {version.name} and previous releases that are accepted to still remain open for the current release.",
            spacing: { after: 200 }
          }),

          // Table 2: Accepted Defects - risk related "yes"
          new Paragraph({
            text: "Table 2: Accepted Defects – risk related \"yes\"",
            bold: true,
            spacing: { after: 100 }
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: "ID", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Summary", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Detected in HW SN / Build", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Risk Reference", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Affected Requirement", bold: true })],
                    shading: { fill: "E0E0E0" }
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#acceptedRisk}{docId}{/acceptedRisk}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#acceptedRisk}{title}{/acceptedRisk}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#acceptedRisk}{fieldValue.Detected_Version}{/acceptedRisk}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#acceptedRisk}{relations | where:'type == \"affects\"' | map:'other.docId' | join:', '}{/acceptedRisk}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#acceptedRisk}{relations | where:'type == \"implements\"' | map:'other.docId' | join:', '}{/acceptedRisk}"
                    })]
                  })
                ]
              })
            ]
          }),

          // Show count at the bottom
          new Paragraph({
            text: "{acceptedRisk | count} issue{#acceptedRisk | count > 1}s{/acceptedRisk | count > 1}",
            spacing: { before: 100, after: 200 }
          }),

          // Table 3: Accepted Defects - risk related "no"
          new Paragraph({
            text: "Table 3: Accepted Defects – risk related \"no\"",
            bold: true,
            spacing: { before: 200, after: 100 }
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: "ID", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Summary", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Detected in HW SN / Build", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Affected Requirement", bold: true })],
                    shading: { fill: "E0E0E0" }
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#acceptedNoRisk}{docId}{/acceptedNoRisk}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#acceptedNoRisk}{title}{/acceptedNoRisk}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#acceptedNoRisk}{fieldValue.Detected_Version}{/acceptedNoRisk}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#acceptedNoRisk}{relations | where:'type == \"implements\"' | map:'other.docId' | join:', '}{/acceptedNoRisk}"
                    })]
                  })
                ]
              })
            ]
          }),

          new Paragraph({
            text: "{acceptedNoRisk | count} issue{#acceptedNoRisk | count > 1}s{/acceptedNoRisk | count > 1}",
            spacing: { before: 100, after: 200 }
          }),

          // Resolved Defects
          new Paragraph({
            text: "Resolved Defects",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 300, after: 200 },
            pageBreakBefore: true
          }),
          new Paragraph({
            text: "Table 4 lists all reported defects that have been successfully verified and confirmed to be resolved.",
            spacing: { after: 200 }
          }),

          new Paragraph({
            text: "Table 4: Resolved Defects",
            bold: true,
            spacing: { after: 100 }
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: "ID", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Summary", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Detected in HW SN / Build", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Verification/ Validation Reference", bold: true })],
                    shading: { fill: "E0E0E0" }
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#resolved}{docId}{/resolved}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#resolved}{title}{/resolved}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#resolved}{fieldValue.Detected_Version | default:'None'}{/resolved}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#resolved}{relations | where:'type == \"found anomaly\"' | map:'other.docId' | join:', '}{/resolved}"
                    })]
                  })
                ]
              })
            ]
          }),

          new Paragraph({
            text: "{resolved | count} issue{#resolved | count > 1}s{/resolved | count > 1}",
            spacing: { before: 100, after: 200 }
          }),

          // Rejected Defects
          new Paragraph({
            text: "Rejected Defects",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 300, after: 200 }
          }),
          new Paragraph({
            text: "Table 5 list all initially reported defects that have been identified to be invalid.",
            spacing: { after: 200 }
          }),

          new Paragraph({
            text: "Table 5: Rejected Defects",
            bold: true,
            spacing: { after: 100 }
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: "ID", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Summary", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Detected in HW SN / Build", bold: true })],
                    shading: { fill: "E0E0E0" }
                  }),
                  new TableCell({
                    children: [new Paragraph({ text: "Rejection Justification", bold: true })],
                    shading: { fill: "E0E0E0" }
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#rejected}{docId}{/rejected}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#rejected}{title}{/rejected}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#rejected}{fieldValue.Detected_Version | default:'None'}{/rejected}"
                    })]
                  }),
                  new TableCell({
                    children: [new Paragraph({
                      text: "{#rejected}{fieldValue.Rejection_Justification}{/rejected}"
                    })]
                  })
                ]
              })
            ]
          }),

          new Paragraph({
            text: "{rejected | count} issue{#rejected | count > 1}s{/rejected | count > 1}",
            spacing: { before: 100 }
          })
        ]
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync('/Users/joeycohen/Desktop/Ketryx Templating Wizard Version 3/Template.docx', buffer);
    console.log('Template generated successfully!');
  } catch (err) {
    console.error('Error generating template:', err);
    process.exit(1);
  }
}

generateTemplate();