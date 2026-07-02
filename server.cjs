var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));

// server.ts
var import_express = __toESM(require("express"), 1);
var import_path = __toESM(require("path"), 1);
var import_dotenv = __toESM(require("dotenv"), 1);
var import_mammoth = __toESM(require("mammoth"), 1);
var import_genai = require("@google/genai");
var import_vite = require("vite");
var import_docx = require("docx");
import_dotenv.default.config();
var app = (0, import_express.default)();
var PORT = 3e3;
app.use(import_express.default.json({ limit: "50mb" }));
app.use(import_express.default.urlencoded({ limit: "50mb", extended: true }));
var ai = new import_genai.GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    headers: {
      "User-Agent": "aistudio-build"
    }
  }
});
var rfpSchema = {
  type: import_genai.Type.OBJECT,
  properties: {
    projectName: { type: import_genai.Type.STRING, description: "Name of the project / RFP" },
    scopeSummary: { type: import_genai.Type.STRING, description: "Detailed, professional summary of the scope of work described in the RFP." },
    importantDates: {
      type: import_genai.Type.ARRAY,
      items: {
        type: import_genai.Type.OBJECT,
        properties: {
          date: { type: import_genai.Type.STRING, description: "The specific date or deadline mentioned (e.g. '2026-08-15' or '15th August 2026'). Make sure to parse it clearly." },
          event: { type: import_genai.Type.STRING, description: "The name of the milestone or event (e.g., bid submission, pre-bid meeting, site visit, clarification deadline)." },
          description: { type: import_genai.Type.STRING, description: "Brief details about this date or milestone requirement." }
        },
        required: ["date", "event", "description"]
      }
    },
    requiredCertificates: {
      type: import_genai.Type.ARRAY,
      items: {
        type: import_genai.Type.OBJECT,
        properties: {
          certificate: { type: import_genai.Type.STRING, description: "The name of the certificate, compliance, standard or license required (e.g. ISO 9001, LEED, Security Clearance, Civil Engineering License)." },
          authority: { type: import_genai.Type.STRING, description: "The issuing or verifying authority or standard body if mentioned." },
          description: { type: import_genai.Type.STRING, description: "Why it is required or what it covers." },
          critical: { type: import_genai.Type.BOOLEAN, description: "True if it's a mandatory, pass/fail or critical qualification criteria, false otherwise." }
        },
        required: ["certificate", "description", "critical"]
      }
    },
    keyDeliverables: {
      type: import_genai.Type.ARRAY,
      items: {
        type: import_genai.Type.OBJECT,
        properties: {
          deliverable: { type: import_genai.Type.STRING, description: "The specific deliverable or output requested (e.g., technical proposal, financial proposal, design reports, prototypes, phase reports)." },
          timeline: { type: import_genai.Type.STRING, description: "When this deliverable is due relative to project start or specific dates." },
          recipient: { type: import_genai.Type.STRING, description: "The party or stakeholder who receives this deliverable." },
          format: { type: import_genai.Type.STRING, description: "Format of submission (e.g., PDF, Hardcopy, CAD Drawings, Excel)." }
        },
        required: ["deliverable", "timeline"]
      }
    },
    boqItems: {
      type: import_genai.Type.ARRAY,
      items: {
        type: import_genai.Type.OBJECT,
        properties: {
          itemNo: { type: import_genai.Type.STRING, description: "Item serial number (e.g. '1.1', 'A-2')." },
          description: { type: import_genai.Type.STRING, description: "Detailed item description of quantity, supply, services or work." },
          unit: { type: import_genai.Type.STRING, description: "Unit of measurement (e.g., Sqm, Nos, Ls, Ton, Month, Job)." },
          quantity: { type: import_genai.Type.NUMBER, description: "Quantity of items needed." },
          estimatedRate: { type: import_genai.Type.STRING, description: "Estimated rate or pricing notes if mentioned, otherwise leave empty." }
        },
        required: ["itemNo", "description", "unit", "quantity"]
      }
    },
    mainTowers: {
      type: import_genai.Type.ARRAY,
      items: {
        type: import_genai.Type.OBJECT,
        properties: {
          name: {
            type: import_genai.Type.STRING,
            description: "The tower name. This MUST be one of the following exact tower names: 'ADI', 'Observability', 'Cybersecurity', 'HW/SW Supply', 'EAP', 'Cloud Services', 'DCX', 'Data', 'AI'."
          },
          description: { type: import_genai.Type.STRING, description: "Detailed summary of scope, specifications, or deliverables described in the RFP associated with this particular tower." },
          keyFocusArea: { type: import_genai.Type.STRING, description: "Specific primary focal point, operational standard, or technology stack of this tower in the project." }
        },
        required: ["name", "description"]
      }
    },
    evaluationCriteria: {
      type: import_genai.Type.ARRAY,
      items: {
        type: import_genai.Type.OBJECT,
        properties: {
          criterion: { type: import_genai.Type.STRING, description: "The specific evaluation criterion or scoring category mentioned in the RFP (e.g. Technical Capability, Financial Score, Past Experience, Team Qualifications, Project Methodology)." },
          weight: { type: import_genai.Type.STRING, description: "The percentage weight or points allocated to this criterion (e.g., '30%', '40 points', or 'Not specified')." },
          description: { type: import_genai.Type.STRING, description: "Detailed summary of how the criterion is assessed and what parameters determine a high score." }
        },
        required: ["criterion", "weight", "description"]
      }
    }
  },
  required: ["projectName", "scopeSummary", "importantDates", "requiredCertificates", "keyDeliverables", "boqItems", "mainTowers", "evaluationCriteria"]
};
app.post("/api/analyze-rfp", async (req, res) => {
  try {
    const { fileName, mimeType, fileData, textContent, customApiKey } = req.body;
    const reqCustomKey = req.headers["x-gemini-key"] || customApiKey;
    const activeApiKey = typeof reqCustomKey === "string" && reqCustomKey.trim() ? reqCustomKey.trim() : process.env.GEMINI_API_KEY;
    if (!activeApiKey || activeApiKey === "MY_GEMINI_API_KEY" || activeApiKey.trim() === "") {
      return res.status(401).json({
        error: "Gemini API Key is not configured. Please open Settings and enter a valid Gemini API Key to run the analysis."
      });
    }
    const dynamicAi = new import_genai.GoogleGenAI({
      apiKey: activeApiKey,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build"
        }
      }
    });
    let contents = [];
    const promptText = `You are an expert procurement officer, proposal manager, and bid analyst. Your task is to thoroughly analyze the attached/provided Request for Proposal (RFP) document (which may be written in Arabic, English, or a mix of both).
Extensively extract the following elements and structure them exactly according to the provided schema:
1. Project Name (If in Arabic, provide an accurate English translation or transliteration for title uniformity, but keep the original tone).
2. Detailed summary of the Scope of Work (summarize objectives, deliverables, and expectations professionally. Feel free to explain in professional English, or in Arabic if the original document is Arabic-dominant).
3. Important Dates (bid closing, pre-bid meetings, site visits, clarifications, execution timeline). Ensure you extract these accurately from the text.
4. Required Certificates, Standards, Compliances, or Licenses (ISO, cyber certifications, local regulations, professional credentials, security clearances). Identify which are pass/fail or critical.
5. Key Deliverables (all reports, plans, code, hardware units, configuration manuals, or material handovers).
6. Bill of Quantities (BOQ) Items (if specified or estimated from the work packages, complete with quantities, units, and item identifiers).
7. Main Towers involved. You must map the RFP requirements (in English or Arabic) to one or more of these 9 official main towers. Only include the towers that have relevant scope or deliverables in the RFP:
   - ADI (Application Development & Integration: custom code, API development, system integrations, custom software)
   - Observability (monitoring, system telemetry, SCADA platforms, logging, alerts, APM tools, dashboard tracking)
   - Cybersecurity (OT/IT firewalls, endpoint security, identity management, compliance audits, threat detection, penetration testing)
   - HW/SW Supply (procurement of hardware, servers, switches, smart meters, edge devices, commercial software licenses)
   - EAP (Enterprise Application Platforms: ERP systems, SAP, Oracle, CRMs, corporate core frameworks)
   - Cloud Services (cloud computing, hosting, migration, AWS, Azure, Google Cloud, virtual machines, container orchestration)
   - DCX (Digital Customer Experience: client portals, mobile applications, UI/UX modernizations, consumer touchpoints)
   - Data (databases, pipelines, warehouses, data lakes, analytics, ETL procedures)
   - AI (Machine learning, Predictive modeling, Generative AI, natural language processing, computer vision, smart analytics)
8. Evaluation Criteria (the technical/commercial evaluation breakdown, scoring weights, past experience scores, methodology values, or standard pricing-to-quality ratios).

Make sure your extraction is complete, highly detailed, precise, and professional. If certain details are not explicitly present, you can logically deduce or recommend standard requirements (especially for deliverables, certificates, or evaluation criteria weights) to assist the proposal team in planning.`;
    if (textContent) {
      contents = [promptText, textContent];
    } else if (fileData && mimeType) {
      if (mimeType === "application/pdf") {
        contents = [
          {
            inlineData: {
              mimeType: "application/pdf",
              data: fileData
              // base64 string
            }
          },
          promptText
        ];
      } else if (mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || fileName?.endsWith(".docx")) {
        const buffer = Buffer.from(fileData, "base64");
        const result = await import_mammoth.default.extractRawText({ buffer });
        contents = [promptText, result.value];
      } else if (mimeType.startsWith("text/") || fileName?.endsWith(".txt") || fileName?.endsWith(".csv")) {
        const text = Buffer.from(fileData, "base64").toString("utf-8");
        contents = [promptText, text];
      } else {
        return res.status(400).json({ error: "Unsupported file type. Please upload PDF, DOCX or TXT files." });
      }
    } else {
      return res.status(400).json({ error: "No RFP data or text provided for analysis." });
    }
    const response = await dynamicAi.models.generateContent({
      model: "gemini-3.5-flash",
      contents,
      config: {
        responseMimeType: "application/json",
        responseSchema: rfpSchema,
        temperature: 0.2
        // Lower temperature for high extraction fidelity
      }
    });
    if (!response.text) {
      throw new Error("No response generated from Gemini.");
    }
    const resultData = JSON.parse(response.text.trim());
    return res.json(resultData);
  } catch (error) {
    console.error("Error analyzing RFP:", error);
    return res.status(500).json({
      error: error.message || "An unexpected error occurred during RFP analysis."
    });
  }
});
app.post("/api/export-docx", async (req, res) => {
  try {
    const data = req.body;
    if (!data || !data.projectName) {
      return res.status(400).json({ error: "No RFP analysis data provided for export." });
    }
    const createHeading = (text) => {
      return new import_docx.Paragraph({
        heading: import_docx.HeadingLevel.HEADING_1,
        spacing: { before: 240, after: 120 },
        children: [
          new import_docx.TextRun({
            text,
            bold: true,
            color: "0F172A",
            // Slate-900
            font: "Calibri",
            size: 28
          })
        ]
      });
    };
    const createSubheading = (text) => {
      return new import_docx.Paragraph({
        heading: import_docx.HeadingLevel.HEADING_2,
        spacing: { before: 180, after: 80 },
        children: [
          new import_docx.TextRun({
            text,
            bold: true,
            color: "1E293B",
            // Slate-800
            font: "Calibri",
            size: 24
          })
        ]
      });
    };
    const docChildren = [];
    docChildren.push(
      new import_docx.Paragraph({
        alignment: import_docx.AlignmentType.CENTER,
        spacing: { before: 1200, after: 240 },
        children: [
          new import_docx.TextRun({
            text: "RFP ANALYSIS & PROPOSAL PROFILE",
            bold: true,
            color: "0F172A",
            font: "Calibri",
            size: 40
          })
        ]
      }),
      new import_docx.Paragraph({
        alignment: import_docx.AlignmentType.CENTER,
        spacing: { before: 0, after: 1200 },
        children: [
          new import_docx.TextRun({
            text: data.projectName.toUpperCase(),
            bold: true,
            color: "2563EB",
            // Blue-600
            font: "Calibri",
            size: 32
          })
        ]
      }),
      new import_docx.Paragraph({
        alignment: import_docx.AlignmentType.CENTER,
        spacing: { before: 0, after: 180 },
        children: [
          new import_docx.TextRun({
            text: "Generated on: " + (/* @__PURE__ */ new Date()).toLocaleDateString("en-US", { dateStyle: "long" }),
            italics: true,
            color: "64748B",
            // Slate-500
            font: "Calibri",
            size: 20
          })
        ]
      }),
      new import_docx.Paragraph({
        alignment: import_docx.AlignmentType.CENTER,
        spacing: { before: 0, after: 2400 },
        children: [
          new import_docx.TextRun({
            text: "This document contains intelligent extractions of key milestones, deliverables, requirements, BOQ structures, and scope from the Request for Proposal.",
            italics: true,
            color: "475569",
            font: "Calibri",
            size: 18
          })
        ]
      })
    );
    docChildren.push(createHeading("1. Executive Summary & Scope of Work"));
    docChildren.push(
      new import_docx.Paragraph({
        spacing: { before: 120, after: 240 },
        children: [
          new import_docx.TextRun({
            text: data.scopeSummary || "No scope summary provided.",
            font: "Calibri",
            size: 22
          })
        ]
      })
    );
    docChildren.push(createHeading("2. Important Dates & Timeline"));
    if (data.importantDates && data.importantDates.length > 0) {
      const dateRows = [
        new import_docx.TableRow({
          tableHeader: true,
          children: [
            new import_docx.TableCell({
              width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Milestone Date", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" },
              verticalAlign: import_docx.VerticalAlign.CENTER
            }),
            new import_docx.TableCell({
              width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Event", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" },
              verticalAlign: import_docx.VerticalAlign.CENTER
            }),
            new import_docx.TableCell({
              width: { size: 40, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Description", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" },
              verticalAlign: import_docx.VerticalAlign.CENTER
            })
          ]
        })
      ];
      data.importantDates.forEach((d, i) => {
        dateRows.push(
          new import_docx.TableRow({
            children: [
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: d.date || "" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: d.event || "", bold: true })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: d.description || "" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              })
            ]
          })
        );
      });
      docChildren.push(new import_docx.Table({ rows: dateRows, width: { size: 100, type: import_docx.WidthType.PERCENTAGE } }));
    } else {
      docChildren.push(new import_docx.Paragraph({ text: "No important dates specified." }));
    }
    docChildren.push(new import_docx.Paragraph({ text: "", spacing: { before: 120, after: 120 } }));
    docChildren.push(createHeading("3. Key Deliverables & Timelines"));
    if (data.keyDeliverables && data.keyDeliverables.length > 0) {
      const delRows = [
        new import_docx.TableRow({
          tableHeader: true,
          children: [
            new import_docx.TableCell({
              width: { size: 35, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Deliverable", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            }),
            new import_docx.TableCell({
              width: { size: 25, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Timeline / Due", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            }),
            new import_docx.TableCell({
              width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Recipient", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            }),
            new import_docx.TableCell({
              width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Format", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            })
          ]
        })
      ];
      data.keyDeliverables.forEach((del, i) => {
        delRows.push(
          new import_docx.TableRow({
            children: [
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: del.deliverable || "", bold: true })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: del.timeline || "" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: del.recipient || "Client" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: del.format || "Digital" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              })
            ]
          })
        );
      });
      docChildren.push(new import_docx.Table({ rows: delRows, width: { size: 100, type: import_docx.WidthType.PERCENTAGE } }));
    } else {
      docChildren.push(new import_docx.Paragraph({ text: "No deliverables specified." }));
    }
    docChildren.push(new import_docx.Paragraph({ text: "", spacing: { before: 120, after: 120 } }));
    docChildren.push(createHeading("4. Required Certificates & Compliances"));
    if (data.requiredCertificates && data.requiredCertificates.length > 0) {
      const certRows = [
        new import_docx.TableRow({
          tableHeader: true,
          children: [
            new import_docx.TableCell({
              width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Certificate / Compliance", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            }),
            new import_docx.TableCell({
              width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Authority", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            }),
            new import_docx.TableCell({
              width: { size: 35, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Description", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            }),
            new import_docx.TableCell({
              width: { size: 15, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Criticality", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            })
          ]
        })
      ];
      data.requiredCertificates.forEach((c, i) => {
        certRows.push(
          new import_docx.TableRow({
            children: [
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: c.certificate || "", bold: true })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: c.authority || "N/A" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: c.description || "" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({
                  children: [new import_docx.TextRun({
                    text: c.critical ? "CRITICAL (Pass/Fail)" : "Highly Preferred",
                    bold: true,
                    color: c.critical ? "EF4444" : "10B981"
                    // Red or Green
                  })]
                })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              })
            ]
          })
        );
      });
      docChildren.push(new import_docx.Table({ rows: certRows, width: { size: 100, type: import_docx.WidthType.PERCENTAGE } }));
    } else {
      docChildren.push(new import_docx.Paragraph({ text: "No certificates or compliances specified." }));
    }
    docChildren.push(new import_docx.Paragraph({ text: "", spacing: { before: 120, after: 120 } }));
    docChildren.push(createHeading("5. Bill of Quantities (BOQ) Items"));
    if (data.boqItems && data.boqItems.length > 0) {
      const boqRows = [
        new import_docx.TableRow({
          tableHeader: true,
          children: [
            new import_docx.TableCell({
              width: { size: 15, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Item No.", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            }),
            new import_docx.TableCell({
              width: { size: 50, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Description", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            }),
            new import_docx.TableCell({
              width: { size: 15, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Unit", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            }),
            new import_docx.TableCell({
              width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Quantity", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            })
          ]
        })
      ];
      data.boqItems.forEach((b, i) => {
        boqRows.push(
          new import_docx.TableRow({
            children: [
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: b.itemNo || `${i + 1}` })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: b.description || "" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: b.unit || "Nos" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: b.quantity?.toLocaleString() || "1" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              })
            ]
          })
        );
      });
      docChildren.push(new import_docx.Table({ rows: boqRows, width: { size: 100, type: import_docx.WidthType.PERCENTAGE } }));
    } else {
      docChildren.push(new import_docx.Paragraph({ text: "No BOQ items specified." }));
    }
    docChildren.push(new import_docx.Paragraph({ text: "", spacing: { before: 120, after: 120 } }));
    docChildren.push(createHeading("6. Main Project Towers"));
    if (data.mainTowers && data.mainTowers.length > 0) {
      data.mainTowers.forEach((t) => {
        docChildren.push(createSubheading(t.name));
        docChildren.push(
          new import_docx.Paragraph({
            spacing: { before: 40, after: 80 },
            children: [
              new import_docx.TextRun({ text: "Focus Area: ", bold: true, color: "475569" }),
              new import_docx.TextRun({ text: t.keyFocusArea || "General scope" })
            ]
          }),
          new import_docx.Paragraph({
            spacing: { before: 40, after: 180 },
            children: [
              new import_docx.TextRun({ text: t.description || "" })
            ]
          })
        );
      });
    } else {
      docChildren.push(new import_docx.Paragraph({ text: "No main towers specified." }));
    }
    docChildren.push(new import_docx.Paragraph({ text: "", spacing: { before: 120, after: 120 } }));
    docChildren.push(createHeading("7. RFP Evaluation Criteria"));
    if (data.evaluationCriteria && data.evaluationCriteria.length > 0) {
      const evalRows = [
        new import_docx.TableRow({
          tableHeader: true,
          children: [
            new import_docx.TableCell({
              width: { size: 30, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Criterion", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            }),
            new import_docx.TableCell({
              width: { size: 20, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Weight / Points", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            }),
            new import_docx.TableCell({
              width: { size: 50, type: import_docx.WidthType.PERCENTAGE },
              children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: "Assessment Description", bold: true, color: "FFFFFF" })] })],
              shading: { fill: "1E293B" }
            })
          ]
        })
      ];
      data.evaluationCriteria.forEach((e, i) => {
        evalRows.push(
          new import_docx.TableRow({
            children: [
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: e.criterion || "", bold: true })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: e.weight || "N/A", bold: true, color: "4F46E5" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              }),
              new import_docx.TableCell({
                children: [new import_docx.Paragraph({ children: [new import_docx.TextRun({ text: e.description || "" })] })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" }
              })
            ]
          })
        );
      });
      docChildren.push(new import_docx.Table({ rows: evalRows, width: { size: 100, type: import_docx.WidthType.PERCENTAGE } }));
    } else {
      docChildren.push(new import_docx.Paragraph({ text: "No evaluation criteria specified." }));
    }
    const doc = new import_docx.Document({
      sections: [
        {
          properties: {},
          headers: {
            default: new import_docx.Header({
              children: [
                new import_docx.Paragraph({
                  alignment: import_docx.AlignmentType.RIGHT,
                  children: [
                    new import_docx.TextRun({
                      text: `RFP Extraction Profile | ${data.projectName}`,
                      color: "94A3B8",
                      size: 16,
                      font: "Calibri"
                    })
                  ]
                })
              ]
            })
          },
          footers: {
            default: new import_docx.Footer({
              children: [
                new import_docx.Paragraph({
                  alignment: import_docx.AlignmentType.CENTER,
                  children: [
                    new import_docx.TextRun({
                      text: "Page 1 - Confidential Bid Proposal Resource",
                      color: "94A3B8",
                      size: 16,
                      font: "Calibri"
                    })
                  ]
                })
              ]
            })
          },
          children: docChildren
        }
      ]
    });
    const b64 = await import_docx.Packer.toBase64String(doc);
    res.setHeader("Content-Disposition", `attachment; filename="${data.projectName.replace(/[^a-z0-9]/gi, "_").toLowerCase()}_rfp_analysis.docx"`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    const buffer = Buffer.from(b64, "base64");
    return res.send(buffer);
  } catch (error) {
    console.error("Error exporting DOCX:", error);
    return res.status(500).json({
      error: error.message || "An unexpected error occurred during DOCX generation."
    });
  }
});
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await (0, import_vite.createServer)({
      server: { middlewareMode: true },
      appType: "spa"
    });
    app.use(vite.middlewares);
  } else {
    const distPath = import_path.default.join(process.cwd(), "dist");
    app.use(import_express.default.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(import_path.default.join(distPath, "index.html"));
    });
  }
  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}
startServer();
//# sourceMappingURL=server.cjs.map
