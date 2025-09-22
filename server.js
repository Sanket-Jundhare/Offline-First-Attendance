const express = require("express");
const bodyParser = require("body-parser");
const fs = require("fs");
const os = require("os");
const { exec } = require("child_process");
const XLSX = require("xlsx");
const QRCode = require("qrcode");
const dgram = require("dgram");
const axios = require("axios"); // Added axios

require('dotenv').config();

const app = express();
const PORT = 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static("."));

// --- Helper functions ---
function timeToMinutes(t) {
  if (typeof t === "number") return Math.round(t * 24 * 60);
  if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
  if (typeof t === "string") {
    let match = t.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
    if (match) {
      let hour = parseInt(match[1], 10);
      const min = parseInt(match[2], 10);
      const ampm = match[3] ? match[3].toUpperCase() : null;
      if (ampm) {
        if (ampm === "PM" && hour !== 12) hour += 12;
        if (ampm === "AM" && hour === 12) hour = 0;
      }
      return hour * 60 + min;
    }
    match = t.match(/^(\d{1,2}):(\d{2})$/);
    if (match) return parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
  }
  return 0;
}

function loadStudents() {
  if (!fs.existsSync("students.xlsx")) return {};
  const workbook = XLSX.readFile("students.xlsx");
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet);
  const students = {};
  data.forEach((row) => (students[row.ID] = { name: row.Name, class: row.Class }));
  return students;
}

function loadSchedule() {
  if (!fs.existsSync("schedule.xlsx")) return {};
  const workbook = XLSX.readFile("schedule.xlsx", { cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet);
  const schedule = {};
  data.forEach((row) => {
    if (!row.Class || !row.Day) return;
    const cls = String(row.Class).trim().toUpperCase();
    const day = String(row.Day).trim().toLowerCase();
    if (!schedule[cls]) schedule[cls] = {};
    if (!schedule[cls][day]) schedule[cls][day] = [];
    const start = timeToMinutes(row.Start);
    const end = timeToMinutes(row.End);
    if (start <= end) schedule[cls][day].push({ subject: row.Subject || "", start, end });
  });
  Object.keys(schedule).forEach(cls =>
    Object.keys(schedule[cls]).forEach(day =>
      schedule[cls][day].sort((a,b)=>a.start-b.start)
    )
  );
  return schedule;
}

function getLocalMac() {
  const interfaces = os.networkInterfaces();
  for (const name in interfaces) {
    for (const net of interfaces[name]) {
      if (net.family === "IPv4" && !net.internal && net.mac) return net.mac;
    }
  }
  return null;
}

function getMac(ip, callback) {
  exec(`arp -a ${ip}`, (err, stdout) => {
    if (err) return callback(null);
    const match = stdout.match(/([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})/);
    callback(match ? match[0].toUpperCase().replace(/-/g, ":") : null);
  });
}

function getCurrentLecture(studentClass, schedule, dateObj) {
  if (!studentClass) return null;
  const now = dateObj.getHours()*60 + dateObj.getMinutes();
  const day = dateObj.toLocaleDateString("en-US",{weekday:"long"}).toLowerCase();
  const lectures = schedule[studentClass?.trim().toUpperCase()]?.[day] || [];
  return lectures.find(l => now >= l.start && now <= l.end) || null;
}

// --- AI Function (Directly in server.js) ---
async function askAttendanceAI(records, question) {
  try {
    // Check if API key exists
    if (!process.env.GEMINI_API_KEY) {
      return "❌ Gemini API key not configured. Please set GEMINI_API_KEY in your .env file.";
    }

    const prompt = `You are an attendance assistant. Analyze the following attendance records and answer the question clearly and concisely.

ATTENDANCE RECORDS:
${JSON.stringify(records, null, 2)}

QUESTION: ${question}

Please provide a helpful answer based on the attendance data:`;

    // Try multiple model endpoints
    const modelsToTry = [
      'gemini-1.5-flash-latest',
      'gemini-1.5-pro-latest', 
      'gemini-pro',
      'gemini-1.0-pro'
    ];

    for (const model of modelsToTry) {
      try {
        console.log(`Trying model: ${model}`);
        
        const response = await axios.post(
          `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${process.env.GEMINI_API_KEY}`,
          {
            contents: [
              {
                parts: [
                  {
                    text: prompt
                  }
                ]
              }
            ],
            generationConfig: {
              temperature: 0.3,
              maxOutputTokens: 1000,
              topP: 0.8,
              topK: 40
            }
          },
          {
            headers: {
              "Content-Type": "application/json"
            },
            timeout: 15000
          }
        );

        console.log(`Success with model: ${model}`);
        
        // Extract the response text
        if (response.data && response.data.candidates && response.data.candidates.length > 0) {
          return response.data.candidates[0].content.parts[0].text;
        }
      } catch (err) {
        console.log(`Model ${model} failed: ${err.response?.data?.error?.message || err.message}`);
        // Continue to next model
      }
    }

    // If all models failed, provide basic analysis
    return provideBasicAnalysis(records, question);

  } catch (err) {
    console.error("AI Error:", err.message);
    return provideBasicAnalysis(records, question);
  }
}

// Fallback analysis function
function provideBasicAnalysis(records, question) {
  if (!records || records.length === 0) {
    return "No attendance records found to analyze.";
  }

  const lowerQuestion = question.toLowerCase();
  
  // Basic analysis based on common questions
  if (lowerQuestion.includes('total') || lowerQuestion.includes('how many')) {
    const totalRecords = records.length;
    const uniqueStudents = new Set(records.map(r => r.ID)).size;
    return `Based on available records:\n- Total attendance entries: ${totalRecords}\n- Unique students: ${uniqueStudents}`;
  }
  
  if (lowerQuestion.includes('present') || lowerQuestion.includes('attended')) {
    const studentCount = {};
    records.forEach(record => {
      studentCount[record.ID] = (studentCount[record.ID] || 0) + 1;
    });
    
    const mostFrequent = Object.entries(studentCount).sort((a, b) => b[1] - a[1])[0];
    return `Attendance analysis:\n- Most frequent attendee: ${mostFrequent[0]} (${mostFrequent[1]} times)`;
  }
  
  if (lowerQuestion.includes('subject') || lowerQuestion.includes('class')) {
    const subjectCount = {};
    records.forEach(record => {
      if (record.Subject) {
        subjectCount[record.Subject] = (subjectCount[record.Subject] || 0) + 1;
      }
    });
    
    const subjects = Object.entries(subjectCount).map(([subject, count]) => `${subject}: ${count} records`).join('\n- ');
    return `Subjects with attendance records:\n- ${subjects || 'No subject data available'}`;
  }

  return `I analyzed ${records.length} attendance records. For detailed AI analysis, please check your Gemini API configuration. Question: "${question}"`;
}

// --- Routes ---

// Home page
app.get("/", (req,res) => {
  res.send(`<!DOCTYPE html>
<html>
<head>
  <title>Attendance System</title>
  <style>
    body{font-family:Arial,sans-serif;background:#f4f6f9;margin:0;padding:0;color:#333;}
    header{background:#007bff;color:white;padding:1rem;text-align:center;}
    .container{max-width:500px;margin:2rem auto;background:white;padding:2rem;border-radius:10px;box-shadow:0 4px 8px rgba(0,0,0,0.1);}
    label{font-weight:bold;display:block;margin-bottom:0.5rem;}
    input{width:100%;padding:10px;margin-bottom:1.5rem;border:1px solid #ccc;border-radius:5px;}
    button{width:100%;padding:12px;background:#007bff;color:white;border:none;border-radius:5px;font-size:16px;cursor:pointer;transition:background 0.3s;}
    button:hover{background:#0056b3;}
  </style>
</head>
<body>
  <header><h1>📌 Attendance System</h1></header>
  <div class="container">
    <h2>Mark Attendance</h2>
    <form action="/submit" method="POST">
      <label for="studentId">Student ID:</label>
      <input type="text" id="studentId" name="studentId" required>
      <button type="submit">Mark Attendance</button>
    </form>

    <h2>View Attendance by Student ID</h2>
    <form action="/view" method="GET">
      <label for="viewId">Student ID:</label>
      <input type="text" id="viewId" name="studentId" required>
      <button type="submit">View Attendance</button>
    </form>

    <h2>Datewise Attendance</h2>
    <form action="/list" method="GET">
      <button type="submit">View Datewise Attendance</button>
    </form>

    <h2>Attendance AI Chat</h2>
    <form action="/chat" method="GET">
      <button type="submit">Open Chat Bot</button>
    </form>

    <h2>API Status</h2>
    <p>Gemini API: ${process.env.GEMINI_API_KEY ? '✅ Configured' : '❌ Not configured'}</p>
    <a href="/test-api">Test API Connection</a>
  </div>
</body>
</html>`);
});

// Test API route
app.get("/test-api", async (req, res) => {
  try {
    if (!process.env.GEMINI_API_KEY) {
      return res.send("❌ GEMINI_API_KEY is not set in environment variables");
    }
    
    const response = await axios.get(
      `https://generativelanguage.googleapis.com/v1beta/models?key=${process.env.GEMINI_API_KEY}`,
      {
        headers: {
          "Content-Type": "application/json"
        },
        timeout: 10000
      }
    );

    const models = response.data.models || [];
    let html = `<h1>Available Models</h1>`;
    html += `<ul>`;
    models.forEach(model => {
      html += `<li><strong>${model.name}</strong> - ${model.displayName || 'No name'} (${model.version})</li>`;
    });
    html += `</ul>`;
    html += `<a href="/">⬅️ Back to Home</a>`;
    
    res.send(html);
  } catch (err) {
    res.send(`❌ API Test Failed: ${err.message}<br>Error: ${JSON.stringify(err.response?.data)}<br><a href="/">⬅️ Back to Home</a>`);
  }
});

// --- Mark attendance ---
app.post("/submit", (req,res)=>{
  const students = loadStudents();
  const schedule = loadSchedule();
  const { studentId } = req.body;
  if(!studentId) return res.status(400).send("<h2>❌ Missing Student ID!</h2><a href='/'>Go back</a>");
  const student = students[studentId];
  if(!student) return res.send("<h2>❌ Student ID not found!</h2><a href='/'>Go back</a>");

  const dateObj = new Date();
  const date = dateObj.toISOString().slice(0,10);
  let h = dateObj.getHours(); const m = dateObj.getMinutes().toString().padStart(2,"0");
  const ampm = h>=12?"PM":"AM"; h = h%12||12; const time = `${h}:${m} ${ampm}`;
  const day = dateObj.toLocaleDateString("en-US",{weekday:"long"});
  const lecture = getCurrentLecture(student.class, schedule, dateObj);
  if(!lecture) return res.send(`<h2>❌ No lecture at this time.</h2><a href='/'>Go back</a>`);

  let ip = req.headers["x-forwarded-for"]||req.connection.remoteAddress;
  if(ip.startsWith("::ffff:")) ip=ip.replace("::ffff:","");
  const processEntry = (device)=>{
    const fileName = `attendance-${date}.xlsx`;
    let workbook, worksheet;
    if(fs.existsSync(fileName)){
      workbook = XLSX.readFile(fileName);
      worksheet = workbook.Sheets[workbook.SheetNames[0]];
    }else{
      workbook = XLSX.utils.book_new();
      worksheet = XLSX.utils.json_to_sheet([]);
      XLSX.utils.book_append_sheet(workbook, worksheet, "Attendance");
    }
    let data = XLSX.utils.sheet_to_json(worksheet);
    const exists = data.some(r=>r.Device===device && r.Date===date && r.Subject===lecture.subject);
    if(exists) return res.send(`<h2>❌ Attendance already marked for ${lecture.subject} from device (${device})</h2><a href='/'>Go back</a>`);
    data.push({ Name:student.name, ID:studentId, Class:student.class, Day:day, Time:time, Subject:lecture.subject, Device:device, Date:date });
    worksheet = XLSX.utils.json_to_sheet(data); workbook.Sheets["Attendance"]=worksheet; XLSX.writeFile(workbook,fileName);
    res.send(`<h2>✅ Attendance marked for ${lecture.subject} on ${date}</h2><a href='/'>Go back</a>`);
  };
  if(ip==="::1"||ip==="127.0.0.1") processEntry(getLocalMac()||"LOCAL");
  else getMac(ip, mac=>processEntry(mac||`IP:${ip}`));
});

// --- View attendance by Student ID ---
app.get("/view", (req,res)=>{
  const studentId = req.query.studentId;
  if(!studentId) return res.send(`<h2>❌ Please enter a Student ID.</h2><a href="/">Go back</a>`);

  const files = fs.readdirSync(".").filter(f=>f.startsWith("attendance-") && f.endsWith(".xlsx"));
  let allRecords = [];
  files.forEach(file=>{
    const workbook = XLSX.readFile(file);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);
    allRecords.push(...rows.filter(r=>r.ID===studentId.trim()));
  });

  if(allRecords.length===0) return res.send(`<h2>No attendance records found for Student ID: ${studentId}</h2><a href="/">Go back</a>`);

  const grouped = {};
  allRecords.forEach(r=>{
    if(!grouped[r.Date]) grouped[r.Date]={};
    if(!grouped[r.Date][r.Subject]) grouped[r.Date][r.Subject]=[];
    grouped[r.Date][r.Subject].push(r);
  });

  let htmlTables="";
  for(const date of Object.keys(grouped).sort()){
    htmlTables+=`<h2 style="margin-top:2rem;color:#007bff;">Date: ${date}</h2>`;
    for(const subject of Object.keys(grouped[date])) {
      htmlTables+=`<h3 style="color:#28a745;">Subject: ${subject}</h3>
        <table>
          <tr><th>Name</th><th>ID</th><th>Class</th><th>Day</th><th>Time</th><th>Device</th></tr>
          ${grouped[date][subject].map(r=>`<tr><td>${r.Name}</td><td>${r.ID}</td><td>${r.Class}</td><td>${r.Day}</td><td>${r.Time}</td><td>${r.Device}</td></tr>`).join("")}
        </table>`;
    }
  }

  res.send(`<!DOCTYPE html>
<html>
<head>
<title>Attendance Records for ${studentId}</title>
<style>
body{font-family:Arial,sans-serif;margin:2rem;}
table{border-collapse:collapse;margin-bottom:1.5rem;width:100%;}
th,td{border:1px solid #ddd;padding:8px;text-align:left;}
th{background:#007bff;color:white;}
tr:nth-child(even){background:#f9f9f9;}
a{display:block;margin-top:1rem;color:#007bff;text-decoration:none;}
</style>
</head>
<body>
<h1>📋 Attendance for Student ID: ${studentId}</h1>
${htmlTables}
<a href="/">⬅️ Back to Home</a>
</body>
</html>`);
});

// --- Datewise attendance with date picker ---
app.get("/list", (req,res)=>{
  const date = req.query.date || new Date().toISOString().slice(0,10);
  const fileName = `attendance-${date}.xlsx`;

  if (!fs.existsSync(fileName)) {
    return res.send(`<h2>No attendance records for ${date}.</h2><a href="/">Go back</a>`);
  }

  const workbook = XLSX.readFile(fileName);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  let rows = XLSX.utils.sheet_to_json(sheet);

  if (req.query.studentId) rows = rows.filter(r=>r.ID===req.query.studentId.trim());

  if (rows.length===0) return res.send(`<h2>No records found for ${req.query.studentId || "all students"} on ${date}</h2><a href="/">Go back</a>`);

  const subjects = {};
  rows.forEach(r=>{ if(!subjects[r.Subject]) subjects[r.Subject]=[]; subjects[r.Subject].push(r); });

  let tablesHtml="";
  for(const subject of Object.keys(subjects)){
    tablesHtml+=`<h2 style="margin-top:2rem;color:#007bff;">Subject: ${subject}</h2>
      <table>
        <thead><tr><th>Name</th><th>ID</th><th>Class</th><th>Day</th><th>Time</th><th>Device</th></tr></thead>
        <tbody>${subjects[subject].map(r=>`<tr><td>${r.Name}</td><td>${r.ID}</td><td>${r.Class}</td><td>${r.Day}</td><td>${r.Time}</td><td>${r.Device}</td></tr>`).join("")}</tbody>
      </table>`;
  }

  res.send(`<!DOCTYPE html>
<html>
<head>
<title>Datewise Attendance (${date})</title>
<style>
body{font-family:Arial,sans-serif;margin:2rem;}
table{border-collapse:collapse;margin-bottom:1.5rem;width:100%;}
th,td{border:1px solid #ddd;padding:8px;text-align:left;}
th{background:#007bff;color:white;}
tr:nth-child(even){background:#f9f9f9;}
a, button{display:inline-block;margin-top:1rem;color:#007bff;text-decoration:none;padding:10px 15px;background:#007bff;color:white;border:none;border-radius:5px;cursor:pointer;}
button:hover, a:hover{background:#0056b3;color:white;}
</style>
</head>
<body>
<h1>📋 Attendance Records for ${date}</h1>
<form method="GET" action="/list">
<label for="date">Select Date:</label>
<input type="date" name="date" id="date" value="${date}">
<button type="submit">View</button>
</form>
<form action="/download" method="GET" style="margin-top:10px;">
<input type="hidden" name="date" value="${date}">
<button type="submit">⬇️ Download Excel</button>
</form>
${tablesHtml}
<a href="/">⬅️ Back to Home</a>
</body>
</html>`);
});

// --- Download attendance file ---
app.get("/download",(req,res)=>{
  const date = req.query.date || new Date().toISOString().slice(0,10);
  const fileName = `attendance-${date}.xlsx`;
  if(!fs.existsSync(fileName)) return res.send(`<h2>No attendance file found for ${date}.</h2><a href="/">Go back</a>`);
  res.download(fileName, fileName);
});

// --- Attendance AI Chat ---
app.get("/chat",(req,res)=>{
  res.send(`
<html>
<head>
<title>Attendance AI Chat</title>
<style>
body{font-family:Arial;background:#f4f6f9;padding:2rem;}
h1{color:#007bff;}
textarea{width:100%;height:100px;margin-bottom:1rem;padding:10px;}
button{padding:10px 20px;background:#007bff;color:white;border:none;border-radius:5px;}
.response{margin-top:2rem;padding:1rem;background:white;border-radius:5px;box-shadow:0 2px 4px rgba(0,0,0,0.1);}
</style>
</head>
<body>
<h1>💬 Attendance AI Chat</h1>
<form action="/chat/query" method="POST">
<textarea name="query" placeholder="Ask about attendance..."></textarea><br>
<button type="submit">Ask AI</button>
</form>
<a href="/">⬅️ Back to Home</a>
</body>
</html>`);
});

app.post("/chat/query", async (req, res) => {
  const query = req.body.query;

  // Load all attendance records
  const files = fs.readdirSync(".").filter(f => f.startsWith("attendance-") && f.endsWith(".xlsx"));
  let allRecords = [];
  files.forEach(file => {
    const workbook = XLSX.readFile(file);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    allRecords.push(...XLSX.utils.sheet_to_json(sheet));
  });

  const answer = await askAttendanceAI(allRecords, query);

  res.send(`
    <html>
      <head>
        <title>AI Response</title>
        <style>
          body{font-family:Arial;padding:2rem;}
          h1{color:#007bff;}
          pre{white-space:pre-wrap;background:#f8f9fa;padding:1rem;border-radius:5px;}
          a{display:block;margin-top:1rem;color:#007bff;text-decoration:none;}
        </style>
      </head>
      <body>
        <h1>AI Response</h1>
        <pre>${answer}</pre>
        <a href="/chat">⬅️ Ask Again</a>
        <a href="/">🏠 Home</a>
      </body>
    </html>
  `);
});

// --- Start server ---
app.listen(PORT,"0.0.0.0",()=>{
  const socket = dgram.createSocket("udp4");
  socket.connect(80,"8.8.8.8",()=>{
    const ip = socket.address().address;
    socket.close();
    const url = `http://${ip}:${PORT}`;
    console.log(`📡 Attendance server running at: ${url}`);
    QRCode.toString(url,{type:"terminal",small:true},(err,qr)=>{if(!err)console.log("\n📱 Scan QR to open site:\n",qr);});
  });
});