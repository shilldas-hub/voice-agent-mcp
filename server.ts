import express from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { z } from "zod";
import cors from "cors";
import { google } from "googleapis";
import path from "path";
import fs from "fs";
import OpenAI from "openai";
import PDFDocument from "pdfkit";

// --- LIBRARY LOADER ---
import { createRequire } from "module";
const require = createRequire(import.meta.url);

let pdfLib: any;
try { pdfLib = require("pdf-extraction"); } catch (e) { console.error("Warning: Could not load pdf-extraction."); }
let mammoth: any;
try { mammoth = require("mammoth"); } catch (e) { console.error("Warning: Could not load mammoth."); }

// --- CONFIGURATION ---
const CALENDAR_ID = process.env.CALENDAR_ID || ""; 
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";
const EMAIL_WEBHOOK_URL = process.env.EMAIL_WEBHOOK_URL || "";
const TIME_ZONE = "Asia/Kolkata";
const DOCS_DIR = path.join(process.cwd(), "documents");
const PUBLIC_DIR = path.join(process.cwd(), "public");

// --- AUTH SETUP ---
const SCOPES = ["https://www.googleapis.com/auth/calendar", "https://www.googleapis.com/auth/calendar.events"];
let auth: any;

try {
    if (process.env.GOOGLE_JSON) {
        let credentials;
        if (!process.env.GOOGLE_JSON.trim().startsWith('{')) {
            const decoded = Buffer.from(process.env.GOOGLE_JSON, 'base64').toString('utf-8');
            credentials = JSON.parse(decoded);
        } else {
            credentials = JSON.parse(process.env.GOOGLE_JSON);
        }
        auth = new google.auth.GoogleAuth({ credentials, scopes: SCOPES });
    } else {
        const KEY_PATH = path.join(process.cwd(), "service_account.json");
        if (fs.existsSync(KEY_PATH)) auth = new google.auth.GoogleAuth({ keyFile: KEY_PATH, scopes: SCOPES });
    }
} catch (error: any) { console.error("AUTH ERROR:", error.message); }

const calendar = google.calendar({ version: "v3", auth });
const openai = new OpenAI({ apiKey: OPENAI_API_KEY });

// --- HELPER: FORCE IST PARSING ---
function parseIST(dateStr: string): Date {
    let clean = dateStr.replace(/Z$/, '').replace(/([+-]\d{2}:?\d{2})$/, '');
    if (clean.length <= 10) clean += "T00:00:00";
    return new Date(`${clean}+05:30`);
}

// --- HELPER: DOC LOADER ---
if (!fs.existsSync(PUBLIC_DIR)) fs.mkdirSync(PUBLIC_DIR);

let documentKnowledge: { filename: string; content: string }[] = [];
async function loadDocuments() {
  if (!fs.existsSync(DOCS_DIR)) fs.mkdirSync(DOCS_DIR);
  const files = fs.readdirSync(DOCS_DIR);
  documentKnowledge = []; 
  for (const file of files) {
    const filePath = path.join(DOCS_DIR, file);
    try {
      let text = "";
      if (file.toLowerCase().endsWith(".pdf")) {
        const dataBuffer = fs.readFileSync(filePath);
        const data = await pdfLib(dataBuffer); text = data.text;
      } else if (file.toLowerCase().endsWith(".docx")) {
        const result = await mammoth.extractRawText({ path: filePath }); text = result.value;
      } else if (file.toLowerCase().endsWith(".txt")) {
        text = fs.readFileSync(filePath, "utf-8");
      }
      if (text) documentKnowledge.push({ filename: file, content: text.replace(/\s+/g, " ").trim() });
    } catch (err) {}
  }
  console.log(`Loaded ${documentKnowledge.length} docs.`);
}

function calculateFreeSlots(dateStr: string, busyEvents: any[]) {
  const freeSlots = [];
  const startOfDay = parseIST(dateStr); 
  const endOfDay = new Date(startOfDay); 
  endOfDay.setHours(23, 59, 59);

  let candidateTime = new Date(startOfDay.getTime());
  candidateTime.setHours(9, 0, 0, 0); 

  while (candidateTime < endOfDay) {
    const istTimeStr = candidateTime.toLocaleString("en-US", { timeZone: TIME_ZONE, hour12: false, hour: "numeric", minute: "numeric" });
    const [hourStr, minuteStr] = istTimeStr.split(":");
    const hour = parseInt(hourStr || "0");

    if (hour >= 9 && hour < 17) {
        const isBusy = busyEvents.some((event: any) => {
            const eventStart = new Date(event.start.dateTime || event.start.date);
            const eventEnd = new Date(event.end.dateTime || event.end.date);
            if (!event.start.dateTime) return eventStart.toISOString().slice(0,10) === candidateTime.toISOString().slice(0,10);
            return candidateTime >= eventStart && candidateTime < eventEnd;
        });
        if (!isBusy) {
            const slotLabel = candidateTime.toLocaleTimeString("en-US", { timeZone: TIME_ZONE, hour: '2-digit', minute: '2-digit' });
            freeSlots.push(slotLabel);
        }
    }
    candidateTime = new Date(candidateTime.getTime() + 30 * 60000);
  }
  return freeSlots;
}

// --- EMAIL RELAY ---
async function sendEmailViaRelay(to: string, subject: string, body: string, icsContent?: string) {
    if (!EMAIL_WEBHOOK_URL) return;
    try {
        await fetch(EMAIL_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ to, subject, body, ics: icsContent })
        });
    } catch (e) { console.error("Email Relay Failed:", e); }
}

function createICS(title: string, start: Date, end: Date) {
    const formatDate = (date: Date) => date.toISOString().replace(/-|:|\.\d+/g, "");
    return `BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//VoiceAgent//EN\nBEGIN:VEVENT\nUID:${Date.now()}@voiceagent\nDTSTAMP:${formatDate(new Date())}\nDTSTART:${formatDate(start)}\nDTEND:${formatDate(end)}\nSUMMARY:${title}\nEND:VEVENT\nEND:VCALENDAR`;
}

// --- SERVER ---
const app = express();
app.use(cors());
app.use('/files', express.static(PUBLIC_DIR));

const mcp = new McpServer({ name: "VoiceAgent", version: "12.0.0" });

// TOOL 1: CHECK AVAILABILITY
mcp.tool("check_calendar_availability", { date: z.string() }, async ({ date }) => {
    try {
      console.log(`[Check] Checking ${date} (IST)`);
      const start = parseIST(date);
      const end = new Date(start); end.setHours(23, 59, 59);
      const res = await calendar.events.list({
        calendarId: CALENDAR_ID, timeMin: start.toISOString(), timeMax: end.toISOString(), timeZone: TIME_ZONE, singleEvents: true, orderBy: 'startTime'
      });
      const events = res.data.items || [];
      const busyList = events.map((e: any) => {
        const t = e.start.dateTime ? new Date(e.start.dateTime).toLocaleTimeString("en-US", { timeZone: TIME_ZONE, hour: '2-digit', minute:'2-digit'}) : "All Day";
        return `BUSY: ${t} - ${e.summary}`;
      }).join("\n");
      const availableSlots = calculateFreeSlots(date, events);
      return { content: [{ type: "text", text: `STATUS FOR ${date} (${TIME_ZONE}):\n⛔ BUSY:\n${busyList || "None"}\n✅ AVAILABLE:\n${availableSlots.length > 0 ? availableSlots.join(", ") : "None"}` }] };
    } catch (e: any) { return { content: [{ type: "text", text: "Error checking calendar." }] }; }
});

// TOOL 2: BOOKING
mcp.tool("book_appointment", { title: z.string(), dateTime: z.string(), attendeeEmail: z.string() }, async ({ title, dateTime, attendeeEmail }) => {
    try {
      const start = parseIST(dateTime); 
      const end = new Date(start.getTime() + 30 * 60000); 
      console.log(`[Book] ${title} at ${start.toISOString()}`);

      await calendar.events.insert({
        calendarId: CALENDAR_ID,
        requestBody: {
            summary: `${title} (Guest: ${attendeeEmail})`, 
            description: `GUEST CONTACT: ${attendeeEmail}\nBooked via Voice Agent.`,
            start: { dateTime: start.toISOString() },
            end: { dateTime: end.toISOString() },
        }
      });

      const ics = createICS(title, start, end);
      await sendEmailViaRelay(
          attendeeEmail, 
          `Confirmed: ${title}`, 
          `Your appointment is confirmed for ${start.toLocaleString("en-US", { timeZone: TIME_ZONE })}.\n\nSee attached invite.`,
          ics
      );
      
      return { content: [{ type: "text", text: `Success. Booked for ${start.toLocaleTimeString("en-US", { timeZone: TIME_ZONE })} IST.` }] };
    } catch (error: any) { 
        console.error("BOOKING ERROR:", error);
        return { content: [{ type: "text", text: "Error booking slot." }] }; 
    }
});

// TOOL 3: SEARCH
mcp.tool("search_knowledge_base", { query: z.string() }, async ({ query }) => {
  const keywords = query.toLowerCase().split(" ").filter(w => w.length > 3);
  const results = documentKnowledge.map(doc => {
        let score = 0; keywords.forEach(word => { if (doc.content.toLowerCase().includes(word)) score++; });
        return { ...doc, score };
    }).filter(doc => doc.score > 0).sort((a, b) => b.score - a.score).slice(0, 3);
  if (!results.length) return { content: [{ type: "text", text: "No info found." }] };
  const snippets = results.map(r => `[Source: ${r.filename}]\n${r.content.substring(0, 500)}...`).join("\n\n");
  return { content: [{ type: "text", text: `Found details:\n${snippets}` }] };
});

// TOOL 4: GENERATE PDF & EMAIL IT
mcp.tool(
    "generate_collateral", 
    { 
        topic: z.string().describe("Topic (e.g. 'Refund Policy')"), 
        format: z.string().describe("Format (e.g. 'Summary')"),
        recipientEmail: z.string().optional().describe("User's email to send the PDF to") 
    }, 
    async ({ topic, format, recipientEmail }) => {
      console.log(`[AI-Write] Generating PDF: ${format}...`);
      const searchKeyword = topic.toLowerCase().split(" ")[0] || "";
      const relevantDocs = documentKnowledge.filter(d => d.content.toLowerCase().includes(searchKeyword)).slice(0, 2);
      const contextText = relevantDocs.map(d => `[Source: ${d.filename}]\n${d.content}`).join("\n\n").substring(0, 8000); 

      let aiContent = "";
      try {
        const completion = await openai.chat.completions.create({
            model: "gpt-4o",
            messages: [{ role: "system", content: `Write a ${format}. Plain text paragraphs.` }, { role: "user", content: `TOPIC: ${topic}\nCONTEXT:\n${contextText}` }]
        });
        aiContent = completion.choices[0]?.message?.content || "Error generating text.";
      } catch (err) { return { content: [{ type: "text", text: "AI Generation failed." }] }; }

      try {
        const safeName = `${topic.replace(/[^a-z0-9]/gi, '_')}_${format.replace(/[^a-z0-9]/gi, '_')}.pdf`; 
        const filePath = path.join(PUBLIC_DIR, safeName);
        
        const doc = new PDFDocument();
        const stream = fs.createWriteStream(filePath);
        doc.pipe(stream);
        doc.fontSize(20).text(topic.toUpperCase(), { align: 'center' });
        doc.moveDown();
        doc.fontSize(12).text(aiContent, { align: 'justify', indent: 30 });
        doc.end();

        await new Promise<void>((resolve) => stream.on('finish', () => resolve()));
        
        const host = process.env.RENDER_EXTERNAL_HOSTNAME || "your-app.onrender.com";
        const downloadUrl = `https://${host}/files/${safeName}`;

        let message = `I have created the PDF. You can view it here: ${downloadUrl}`;

        // --- EMAIL LOGIC ---
        if (recipientEmail) {
            console.log(`[Email] Sending PDF link to ${recipientEmail}`);
            await sendEmailViaRelay(
                recipientEmail, 
                `Your requested document: ${topic}`, 
                `Here is the ${format} you requested.\n\nDownload Link: ${downloadUrl}`
            );
            message += `\nI have also emailed it to ${recipientEmail}.`;
        } else {
            message += `\n(I didn't send an email because no address was provided).`;
        }

        return { content: [{ type: "text", text: message }] };

      } catch (err: any) {
          console.error("PDF SAVE ERROR:", err);
          return { content: [{ type: "text", text: "Error saving PDF file." }] };
      }
});

let transport: SSEServerTransport;
app.get("/sse", async (req, res) => { transport = new SSEServerTransport("/messages", res); await mcp.connect(transport); });
app.post("/messages", async (req, res) => { if (transport) await transport.handlePostMessage(req, res); });

const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => { console.log(`\n--- Voice MCP Server Running on ${PORT} ---`); await loadDocuments(); });