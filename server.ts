import express from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { z } from "zod";
import cors from "cors";
import { google } from "googleapis";
import path from "path";
import fs from "fs";
import OpenAI from "openai";
import { Readable } from "stream"; // Needed for uploading text to Drive

// --- LIBRARY LOADER ---
import { createRequire } from "module";
const require = createRequire(import.meta.url);

let pdfLib: any;
try { pdfLib = require("pdf-extraction"); } catch (e) { console.error("Warning: Could not load pdf-extraction."); }
let mammoth: any;
try { mammoth = require("mammoth"); } catch (e) { console.error("Warning: Could not load mammoth."); }

// --- CONFIGURATION ---
const CALENDAR_ID = process.env.CALENDAR_ID || ""; 
const EMAIL_USER = process.env.EMAIL_USER || ""; 
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";
const TIME_ZONE = "Asia/Kolkata";
const DOCS_DIR = path.join(process.cwd(), "documents");

// --- AUTH SETUP (Added Drive Scope) ---
const SCOPES = [
    "https://www.googleapis.com/auth/calendar", 
    "https://www.googleapis.com/auth/calendar.events",
    "https://www.googleapis.com/auth/drive" // <--- NEW PERMISSION
];
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

// Initialize Clients
const calendar = google.calendar({ version: "v3", auth });
const drive = google.drive({ version: "v3", auth }); // <--- NEW CLIENT
const openai = new OpenAI({ apiKey: OPENAI_API_KEY });

// --- HELPERS ---
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
  const startOfDay = new Date(dateStr); 
  const endOfDay = new Date(dateStr); 
  endOfDay.setHours(23, 59, 59);

  let candidateTime = new Date(startOfDay.getTime());

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

// --- SERVER ---
const app = express();
app.use(cors());
const mcp = new McpServer({ name: "VoiceAgent", version: "5.0.0" });

// TOOL 1: CHECK AVAILABILITY
mcp.tool("check_calendar_availability", { date: z.string() }, async ({ date }) => {
    try {
      console.log(`[Check] Checking ${date} in ${TIME_ZONE}`);
      const res = await calendar.events.list({
        calendarId: CALENDAR_ID,
        timeMin: new Date(date).toISOString(),
        timeMax: new Date(new Date(date).getTime() + 86400000).toISOString(),
        timeZone: TIME_ZONE,
        singleEvents: true, 
        orderBy: 'startTime'
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

// TOOL 2: BOOKING (Calendar Only - Email Paused)
mcp.tool("book_appointment", { title: z.string(), dateTime: z.string(), attendeeEmail: z.string() }, async ({ title, dateTime, attendeeEmail }) => {
    try {
      console.log(`[Book] ${title} for ${attendeeEmail} at ${dateTime}`);
      const start = new Date(dateTime);
      const end = new Date(start.getTime() + 30 * 60000); 
      await calendar.events.insert({
        calendarId: CALENDAR_ID,
        requestBody: {
            summary: `${title} (Guest: ${attendeeEmail})`, 
            description: `GUEST: ${attendeeEmail}\nBooked via Voice Agent.`,
            start: { dateTime: start.toISOString() },
            end: { dateTime: end.toISOString() },
        }
      });
      return { content: [{ type: "text", text: `Success. Booked on calendar for ${attendeeEmail}.` }] };
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

// TOOL 4: GENERATE GOOGLE DOC (The Upgrade)
mcp.tool(
    "generate_collateral",
    { 
      topic: z.string().describe("The topic (e.g. 'Refund Policy Summary')"),
      format: z.string().describe("Format (e.g. 'One-Pager', 'Memo')") 
    },
    async ({ topic, format }) => {
      console.log(`[AI-Write] Creating Google Doc: ${format} about ${topic}...`);
      
      // 1. Generate Content via OpenAI
      const searchKeyword = topic.toLowerCase().split(" ")[0] || "";
      const relevantDocs = documentKnowledge.filter(d => d.content.toLowerCase().includes(searchKeyword)).slice(0, 2);
      const contextText = relevantDocs.map(d => `[Source: ${d.filename}]\n${d.content}`).join("\n\n").substring(0, 8000); 

      let aiContent = "";
      try {
        const completion = await openai.chat.completions.create({
            model: "gpt-4o",
            messages: [
                { role: "system", content: `You are a professional business writer. Write a ${format}. Use Markdown formatting.` }, 
                { role: "user", content: `TOPIC: ${topic}\nCONTEXT:\n${contextText}` }
            ],
        });
        aiContent = completion.choices[0]?.message?.content || "Error generating text.";
      } catch (err) { return { content: [{ type: "text", text: "AI Generation failed." }] }; }

      // 2. Upload to Google Drive
      try {
        const fileMetadata = {
            name: `${topic} - ${format} (Draft)`,
            mimeType: 'application/vnd.google-apps.document' // Converts text to Google Doc
        };
        const media = {
            mimeType: 'text/plain',
            body: aiContent
        };

        const file = await drive.files.create({
            requestBody: fileMetadata,
            media: media,
            fields: 'id, webViewLink'
        });

        // 3. Share it with YOU (so you can see it)
        if (EMAIL_USER) {
            await drive.permissions.create({
                fileId: file.data.id!,
                requestBody: {
                    role: 'writer',
                    type: 'user',
                    emailAddress: EMAIL_USER
                }
            });
        }

        return { 
            content: [{ 
                type: "text", 
                text: `I have created the ${format} as a Google Doc. You can access it here: ${file.data.webViewLink}` 
            }] 
        };

      } catch (driveErr: any) {
          console.error("DRIVE ERROR:", driveErr);
          return { content: [{ type: "text", text: "I wrote the text, but failed to save it to Google Drive." }] };
      }
    }
);

let transport: SSEServerTransport;
app.get("/sse", async (req, res) => { transport = new SSEServerTransport("/messages", res); await mcp.connect(transport); });
app.post("/messages", async (req, res) => { if (transport) await transport.handlePostMessage(req, res); });

const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => { console.log(`\n--- Voice MCP Server Running on ${PORT} ---`); await loadDocuments(); });