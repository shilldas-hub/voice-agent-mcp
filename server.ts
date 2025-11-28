import express from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { z } from "zod";
import cors from "cors";
import { google } from "googleapis";
import path from "path";
import fs from "fs";
import nodemailer from "nodemailer";
import OpenAI from "openai";

// --- LIBRARY LOADER ---
import { createRequire } from "module";
const require = createRequire(import.meta.url);

// Load Libraries safely
let pdfLib: any;
try { pdfLib = require("pdf-extraction"); } catch (e) { console.error("Warning: Could not load pdf-extraction."); }

let mammoth: any;
try { mammoth = require("mammoth"); } catch (e) { console.error("Warning: Could not load mammoth."); }

// --- CONFIGURATION ---
// FIX: We add '|| ""' to force these to be strings, satisfying TypeScript
const CALENDAR_ID = process.env.CALENDAR_ID || ""; 
const EMAIL_USER = process.env.EMAIL_USER || ""; 
const EMAIL_PASS = process.env.EMAIL_PASS || ""; 
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";
const DOCS_DIR = path.join(process.cwd(), "documents");

// --- GOOGLE AUTH SETUP ---
const SCOPES = ["https://www.googleapis.com/auth/calendar", "https://www.googleapis.com/auth/calendar.events"];
let auth: any;

try {
    if (process.env.GOOGLE_JSON) {
        console.log("Reading Google Creds...");
        let credentials;
        
        // BASE64 CHECK: Decodes the string if you pasted the Base64 version in Render
        if (!process.env.GOOGLE_JSON.trim().startsWith('{')) {
            console.log("-> Detected Base64 Encoded Credentials.");
            const decoded = Buffer.from(process.env.GOOGLE_JSON, 'base64').toString('utf-8');
            credentials = JSON.parse(decoded);
        } else {
            console.log("-> Detected Raw JSON.");
            credentials = JSON.parse(process.env.GOOGLE_JSON);
        }

        auth = new google.auth.GoogleAuth({ credentials, scopes: SCOPES });
        console.log("-> Google Creds loaded successfully.");
    } else {
        // Fallback for local testing if file exists
        console.log("Loading Google Creds from local file...");
        const KEY_PATH = path.join(process.cwd(), "service_account.json");
        if (fs.existsSync(KEY_PATH)) {
            auth = new google.auth.GoogleAuth({ keyFile: KEY_PATH, scopes: SCOPES });
        } else {
            console.error("WARNING: No Google Credentials found (Env Var or File).");
        }
    }
} catch (error: any) {
    console.error("CRITICAL AUTH ERROR:", error.message);
}

// --- INITIALIZE SERVICES ---
const calendar = google.calendar({ version: "v3", auth });
const openai = new OpenAI({ apiKey: OPENAI_API_KEY });

// --- HELPER FUNCTIONS ---
let documentKnowledge: { filename: string; content: string }[] = [];

async function loadDocuments() {
  console.log("Loading documents from:", DOCS_DIR);
  if (!fs.existsSync(DOCS_DIR)) fs.mkdirSync(DOCS_DIR);
  const files = fs.readdirSync(DOCS_DIR);
  documentKnowledge = []; 
  for (const file of files) {
    const filePath = path.join(DOCS_DIR, file);
    try {
      let text = "";
      if (file.toLowerCase().endsWith(".pdf")) {
        if (!pdfLib) throw new Error("PDF Library missing");
        const dataBuffer = fs.readFileSync(filePath);
        const data = await pdfLib(dataBuffer); text = data.text;
      } else if (file.toLowerCase().endsWith(".docx")) {
        if (!mammoth) throw new Error("Mammoth Library missing");
        let extractor = mammoth;
        if (!extractor.extractRawText && extractor.default) extractor = extractor.default;
        const result = await extractor.extractRawText({ path: filePath }); text = result.value;
      } else if (file.toLowerCase().endsWith(".txt")) {
        text = fs.readFileSync(filePath, "utf-8");
      }
      if (text) {
        text = text.replace(/\s+/g, " ").trim(); 
        documentKnowledge.push({ filename: file, content: text });
      }
    } catch (err: any) { console.error(` -> Failed to read ${file}: ${err.message}`); }
  }
  console.log(`Total documents available: ${documentKnowledge.length}`);
}

function calculateFreeSlots(dateStr: string, busyEvents: any[]) {
  const freeSlots = [];
  const workStartHour = 9; const workEndHour = 17;
  const candidateTime = new Date(dateStr); candidateTime.setHours(workStartHour, 0, 0, 0);
  const endTime = new Date(dateStr); endTime.setHours(workEndHour, 0, 0, 0);

  while (candidateTime < endTime) {
    const isBusy = busyEvents.some((event: any) => {
      const startStr = event.start.dateTime || event.start.date;
      const endStr = event.end.dateTime || event.end.date;
      const eventStart = new Date(startStr);
      const eventEnd = endStr ? new Date(endStr) : new Date(eventStart.getTime() + 30*60000);
      if (!event.start.dateTime) {
         const targetDate = new Date(dateStr); return eventStart.getDate() === targetDate.getDate();
      }
      return candidateTime >= eventStart && candidateTime < eventEnd;
    });
    if (!isBusy) freeSlots.push(candidateTime.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' }));
    candidateTime.setMinutes(candidateTime.getMinutes() + 30);
  }
  return freeSlots;
}

function createICS(title: string, description: string, start: Date, end: Date, location: string = "Online") {
    const formatDate = (date: Date) => date.toISOString().replace(/-|:|\.\d+/g, "");
    return `BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//VoiceAgent//NONSGML v1.0//EN\nBEGIN:VEVENT\nUID:${Date.now()}@voiceagent.com\nDTSTAMP:${formatDate(new Date())}\nDTSTART:${formatDate(start)}\nDTEND:${formatDate(end)}\nSUMMARY:${title}\nDESCRIPTION:${description}\nLOCATION:${location}\nEND:VEVENT\nEND:VCALENDAR`;
}

// --- SERVER SETUP ---
const app = express();
app.use(cors());
const mcp = new McpServer({ name: "VoiceAgent", version: "3.6.0" });

// --- TOOLS ---

mcp.tool("check_calendar_availability", { date: z.string() }, async ({ date }) => {
    console.log(`[Check] Checking ${date}`);
    try {
      const start = new Date(date); start.setHours(0,0,0,0);
      const end = new Date(date); end.setHours(23,59,59,999);
      const res = await calendar.events.list({
        calendarId: CALENDAR_ID, timeMin: start.toISOString(), timeMax: end.toISOString(), singleEvents: true, orderBy: 'startTime'
      });
      const events = res.data.items || [];
      const busyList = events.map((e: any) => {
        const t = e.start.dateTime ? new Date(e.start.dateTime).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) : "All Day";
        return `BUSY: ${t} - ${e.summary}`;
      }).join("\n");
      const availableSlots = calculateFreeSlots(date, events);
      return { content: [{ type: "text", text: `STATUS FOR ${date}:\n⛔ BUSY:\n${busyList || "None"}\n✅ AVAILABLE:\n${availableSlots.length > 0 ? availableSlots.join(", ") : "None"}` }] };
    } catch (e: any) { 
        console.error("CALENDAR CHECK ERROR:", e);
        return { content: [{ type: "text", text: "Error checking calendar." }] }; 
    }
});

mcp.tool("book_appointment", { title: z.string(), dateTime: z.string(), attendeeEmail: z.string(), durationMinutes: z.number().default(30) }, async ({ title, dateTime, attendeeEmail, durationMinutes }) => {
    console.log(`[Book] '${title}' for ${attendeeEmail} at ${dateTime}`);
    try {
      const start = new Date(dateTime);
      const end = new Date(start.getTime() + durationMinutes * 60000);
      
      console.log("-> Inserting into Calendar...");
      await calendar.events.insert({
        calendarId: CALENDAR_ID,
        requestBody: {
            summary: `${title} (Guest: ${attendeeEmail})`, 
            description: `GUEST EMAIL: ${attendeeEmail}\nBooked via Voice Agent.`,
            start: { dateTime: start.toISOString() },
            end: { dateTime: end.toISOString() },
        }
      });

      console.log("-> Sending Email (Port 587)...");
      // FIX: Use Port 587 (STARTTLS) for reliable Cloud delivery
      const transporter = nodemailer.createTransport({
          host: "smtp.gmail.com",
          port: 587,
          secure: false, // Must be false for 587
          auth: { user: EMAIL_USER, pass: EMAIL_PASS }
      });

      const icsContent = createICS(title, `Meeting with ${attendeeEmail}`, start, end);
      await transporter.sendMail({
          from: `"Voice Agent" <${EMAIL_USER}>`, to: attendeeEmail, subject: `Confirmed: ${title}`,
          text: `Your appointment is confirmed for ${start.toLocaleString()}.`,
          attachments: [{ filename: 'invite.ics', content: icsContent, contentType: 'text/calendar' }]
      });

      return { content: [{ type: "text", text: `Success! Booked and emailed invite to ${attendeeEmail}.` }] };
    } catch (error: any) { 
        console.error("BOOKING/EMAIL ERROR:", error);
        return { content: [{ type: "text", text: "I tried to book it, but a system error occurred." }] }; 
    }
});

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

mcp.tool("send_email", { to: z.string(), subject: z.string(), body: z.string() }, async ({ to, subject, body }) => {
      try {
        // FIX: Use Port 587 (STARTTLS) here too
        const transporter = nodemailer.createTransport({
            host: "smtp.gmail.com",
            port: 587,
            secure: false,
            auth: { user: EMAIL_USER, pass: EMAIL_PASS }
        });
        await transporter.sendMail({ from: `"Voice Agent" <${EMAIL_USER}>`, to, subject, text: body });
        return { content: [{ type: "text", text: `Email sent.` }] };
      } catch (error) { 
          console.error("EMAIL ERROR:", error); 
          return { content: [{ type: "text", text: "Failed to send email." }] }; 
      }
});

mcp.tool("generate_collateral", { topic: z.string(), format: z.string() }, async ({ topic, format }) => {
      console.log(`[AI-Write] Generating ${format}...`);
      const searchKeyword = topic.toLowerCase().split(" ")[0] || "";
      const relevantDocs = documentKnowledge.filter(d => d.content.toLowerCase().includes(searchKeyword)).slice(0, 2);
      const contextText = relevantDocs.map(d => `[Source: ${d.filename}]\n${d.content}`).join("\n\n").substring(0, 8000); 
      
      try {
        const completion = await openai.chat.completions.create({
            model: "gpt-4o",
            messages: [
                { role: "system", content: `Write a ${format} about '${topic}' based ONLY on the context.` },
                { role: "user", content: `CONTEXT:\n${contextText}` }
            ],
        });
        const generatedContent = completion.choices[0]?.message?.content || "Error";
        return { content: [{ type: "text", text: `I have generated the document. Here is the content:\n\n${generatedContent}\n\nWould you like me to email this to you?` }] };
      } catch (error: any) { 
          console.error("OPENAI ERROR:", error);
          return { content: [{ type: "text", text: "AI Generation failed." }] }; 
      }
});

let transport: SSEServerTransport;
app.get("/sse", async (req, res) => { transport = new SSEServerTransport("/messages", res); await mcp.connect(transport); });
app.post("/messages", async (req, res) => { if (transport) await transport.handlePostMessage(req, res); });

const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => { console.log(`\n--- Voice MCP Server Running on ${PORT} ---`); await loadDocuments(); });