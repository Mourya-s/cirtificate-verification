const express = require("express");
const { MongoClient } = require("mongodb");
const multer = require("multer");
const path = require("path");
const xlsx = require("xlsx");
const fs = require("fs");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");

const app = express();

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static("public"));

// JWT Secret Key - CHANGE THIS IN PRODUCTION!
const JWT_SECRET = "your-secret-key-change-this-in-production";

// Multer storage for Excel
const excelStorage = multer.diskStorage({
  destination: "uploads/excel",
  filename: (req, file, cb) => cb(null, "students.xlsx"),
});
const uploadExcel = multer({ storage: excelStorage });

let db;
let studentsCollection;
let usersCollection;
let settingsCollection;

// JWT Middleware to verify token
function authenticateToken(req, res, next) {
  const token = req.headers['authorization']?.split(' ')[1];
  
  if (!token) {
    return res.status(401).json({ message: "Access denied. No token provided." });
  }

  try {
    const verified = jwt.verify(token, JWT_SECRET);
    req.user = verified;
    next();
  } catch (err) {
    res.status(403).json({ message: "Invalid token" });
  }
}

// Middleware to check if user is admin
function isAdmin(req, res, next) {
  if (req.user.role !== 'admin') {
    return res.status(403).json({ message: "Access denied. Admin only." });
  }
  next();
}

// Connect to MongoDB
async function connectDB() {
  try {
    const client = new MongoClient("mongodb://127.0.0.1:27017");
    await client.connect();
    console.log("✅ MongoDB Connected Successfully");
    
    db = client.db("certificatesDB");
    studentsCollection = db.collection("students");
    usersCollection = db.collection("users");
    settingsCollection = db.collection("settings");
    
    // Create unique index on username
    await usersCollection.createIndex({ username: 1 }, { unique: true });
    
    // Initialize default template if not exists
    const settings = await settingsCollection.findOne({ type: 'certificate_template' });
    if (!settings) {
      await settingsCollection.insertOne({
        type: 'certificate_template',
        template: 'classic',
        updatedAt: new Date()
      });
    }
    
    // Check if data already exists
    const count = await studentsCollection.countDocuments();
    console.log(`📊 Found ${count} existing records in database`);
    
    if (count === 0) {
      console.log("⚠️  No data found. Please upload Excel file.");
      await loadExcelIfExists();
    } else {
      console.log("✅ Database already has data - ready to search!");
    }
    
    startServer();
  } catch (err) {
    console.error("❌ MongoDB Connection Error:", err);
  }
}

// Load Excel file automatically if it exists
async function loadExcelIfExists() {
  const excelPath = "./uploads/excel/students.xlsx";
  
  if (fs.existsSync(excelPath)) {
    try {
      console.log("📂 Found existing Excel file, loading data...");
      const workbook = xlsx.readFile(excelPath);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = xlsx.utils.sheet_to_json(sheet);

      // Debug: Check what columns are available
      if (data.length > 0) {
        console.log("📋 Excel columns found:", Object.keys(data[0]));
        console.log("📝 First row sample:", data[0]);
      }

      const mappedData = data.map((row) => ({
        name: row.NAME,
        certificate: row.CIRTIFICATES,
        link: row.links,
        college: row.college
      }));

      await studentsCollection.insertMany(mappedData);
      console.log(`✅ Auto-loaded ${mappedData.length} records from saved Excel file`);
    } catch (err) {
      console.error("❌ Error auto-loading Excel:", err.message);
    }
  }
}

// Certificate Template 1 - Classic Design
function generateClassicCertificate(name, certificateType) {
  const currentDate = new Date().toLocaleDateString('en-US', { 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  });
  
  return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Certificate - ${name}</title>
    <style>
        @page {
            size: A4 landscape;
            margin: 0;
        }
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Georgia', 'Times New Roman', serif;
            background: #f5f5f5;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            padding: 20px;
        }
        
        .certificate {
            width: 297mm;
            height: 210mm;
            background: white;
            padding: 30px 50px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            position: relative;
            border: 15px solid transparent;
            border-image: linear-gradient(135deg, #667eea 0%, #764ba2 100%) 1;
        }
        
        .certificate::before {
            content: '';
            position: absolute;
            top: 20px;
            left: 20px;
            right: 20px;
            bottom: 20px;
            border: 2px solid #667eea;
            pointer-events: none;
        }
        
        .header {
            text-align: center;
            margin-bottom: 25px;
        }
        
        .logo {
            width: 60px;
            height: 60px;
            margin: 0 auto 15px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 32px;
            font-weight: bold;
        }
        
        .title {
            font-size: 48px;
            font-weight: bold;
            color: #667eea;
            text-transform: uppercase;
            letter-spacing: 6px;
            margin-bottom: 8px;
        }
        
        .subtitle {
            font-size: 18px;
            color: #666;
            font-style: italic;
            letter-spacing: 2px;
        }
        
        .content {
            text-align: center;
            margin: 30px 0;
        }
        
        .awarded-to {
            font-size: 20px;
            color: #666;
            margin-bottom: 15px;
            font-style: italic;
        }
        
        .student-name {
            font-size: 52px;
            font-weight: bold;
            color: #333;
            margin: 20px 0;
            padding: 15px;
            border-bottom: 3px solid #667eea;
            display: inline-block;
            min-width: 400px;
        }
        
        .certificate-text {
            font-size: 18px;
            color: #555;
            line-height: 1.6;
            margin: 25px auto;
            max-width: 700px;
        }
        
        .certificate-type {
            font-size: 24px;
            font-weight: bold;
            color: #764ba2;
            margin: 15px 0;
        }
        
        .date {
            font-size: 16px;
            color: #666;
            margin-top: 15px;
        }
        
        .footer {
            display: flex;
            justify-content: space-around;
            align-items: center;
            margin-top: 40px;
            padding: 0 40px;
        }
        
        .signature-block {
            text-align: center;
        }
        
        .signature-line {
            width: 200px;
            border-top: 2px solid #333;
            margin-bottom: 8px;
        }
        
        .signature-label {
            font-size: 14px;
            color: #666;
            font-weight: 600;
        }
        
        .seal {
            width: 80px;
            height: 80px;
            border: 3px solid #667eea;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #667eea;
            font-weight: bold;
            font-size: 12px;
            text-align: center;
            line-height: 1.2;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            
            .certificate {
                box-shadow: none;
            }
        }
    </style>
</head>
<body>
    <div class="certificate">
        <div class="header">
            <div class="logo">🎓</div>
            <div class="title">Certificate</div>
            <div class="subtitle">of Achievement</div>
        </div>
        
        <div class="content">
            <div class="awarded-to">This is proudly presented to</div>
            <div class="student-name">${name}</div>
            <div class="certificate-text">
                For successfully completing the course and demonstrating 
                outstanding performance in course ML
                from dates 12/7/27 to 13/8/27
            </div>
            <div class="certificate-type">${certificateType}</div>
            <div class="date">Awarded on ${currentDate}</div>
        </div>
        
        <div class="footer">
            <div class="signature-block">
                <div class="signature-line"></div>
                <div class="signature-label">Director</div>
            </div>
            
            <div class="seal">
                OFFICIAL<br>SEAL
            </div>
            
            <div class="signature-block">
                <div class="signature-line"></div>
                <div class="signature-label">Instructor</div>
            </div>
        </div>
    </div>
</body>
</html>
  `;
}

// Certificate Template 2 - Modern Design
function generateModernCertificate(name, certificateType) {
  const currentDate = new Date().toLocaleDateString('en-US', { 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  });
  
  return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Certificate - ${name}</title>
    <style>
        @page {
            size: A4 landscape;
            margin: 0;
        }
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Arial', 'Helvetica', sans-serif;
            background: #f5f5f5;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            padding: 20px;
        }
        
        .certificate {
            width: 297mm;
            height: 210mm;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            position: relative;
            display: flex;
            flex-direction: column;
            overflow: hidden;
        }
        
        .certificate::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 150px;
            background: rgba(255,255,255,0.1);
            transform: skewY(-3deg);
            transform-origin: top left;
        }
        
        .content-wrapper {
            flex: 1;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            padding: 30px 50px;
            position: relative;
            z-index: 1;
        }
        
        .badge {
            width: 70px;
            height: 70px;
            background: white;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 36px;
            margin-bottom: 20px;
            box-shadow: 0 5px 20px rgba(0,0,0,0.2);
        }
        
        .title {
            font-size: 56px;
            font-weight: bold;
            color: white;
            text-transform: uppercase;
            letter-spacing: 8px;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }
        
        .subtitle {
            font-size: 20px;
            color: rgba(255,255,255,0.9);
            letter-spacing: 3px;
            margin-bottom: 30px;
        }
        
        .white-box {
            background: white;
            padding: 40px 60px;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            text-align: center;
            max-width: 800px;
        }
        
        .awarded-text {
            font-size: 18px;
            color: #666;
            margin-bottom: 15px;
        }
        
        .student-name {
            font-size: 48px;
            font-weight: bold;
            color: #667eea;
            margin: 20px 0;
            padding-bottom: 15px;
            border-bottom: 3px solid #667eea;
        }
        
        .certificate-text {
            font-size: 16px;
            color: #555;
            margin: 20px 0;
            line-height: 1.5;
        }
        
        .certificate-type {
            font-size: 22px;
            font-weight: bold;
            color: #764ba2;
            margin: 15px 0;
        }
        
        .date {
            font-size: 14px;
            color: #999;
            margin-top: 15px;
        }
        
        .footer {
            background: rgba(255,255,255,0.15);
            padding: 15px 50px;
            display: flex;
            justify-content: space-around;
            align-items: center;
        }
        
        .signature-block {
            text-align: center;
            color: white;
        }
        
        .signature-line {
            width: 180px;
            border-top: 2px solid white;
            margin-bottom: 8px;
        }
        
        .signature-label {
            font-size: 13px;
            font-weight: 600;
            opacity: 0.9;
        }
        
        .seal {
            width: 70px;
            height: 70px;
            border: 3px solid white;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
            font-size: 11px;
            text-align: center;
            line-height: 1.2;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            
            .certificate {
                box-shadow: none;
            }
        }
    </style>
</head>
<body>
    <div class="certificate">
        <div class="content-wrapper">
            <div class="badge">🏆</div>
            <div class="title">Certificate</div>
            <div class="subtitle">OF EXCELLENCE</div>
            
            <div class="white-box">
                <div class="awarded-text">This certificate is awarded to</div>
                <div class="student-name">${name}</div>
                <div class="certificate-text">
                    For outstanding achievement and successful completion of
                </div>
                <div class="certificate-type">${certificateType}</div>
                <div class="date">${currentDate}</div>
            </div>
        </div>
        
        <div class="footer">
            <div class="signature-block">
                <div class="signature-line"></div>
                <div class="signature-label">Director Signature</div>
            </div>
            
            <div class="seal">
                OFFICIAL<br>SEAL
            </div>
            
            <div class="signature-block">
                <div class="signature-line"></div>
                <div class="signature-label">Instructor Signature</div>
            </div>
        </div>
    </div>
</body>
</html>
  `;
}

function startServer() {
  // ========== AUTHENTICATION ROUTES ==========
  
  // Register route
  app.post("/api/register", async (req, res) => {
    try {
      const { username, password, role } = req.body;
      
      if (!username || !password || !role) {
        return res.status(400).json({ message: "All fields are required" });
      }
      
      if (role !== 'admin' && role !== 'student') {
        return res.status(400).json({ message: "Invalid role. Must be 'admin' or 'student'" });
      }
      
      if (password.length < 6) {
        return res.status(400).json({ message: "Password must be at least 6 characters" });
      }
      
      const existingUser = await usersCollection.findOne({ username });
      if (existingUser) {
        return res.status(400).json({ message: "Username already exists" });
      }
      
      const hashedPassword = await bcrypt.hash(password, 10);
      
      const newUser = {
        username,
        password: hashedPassword,
        role,
        createdAt: new Date()
      };
      
      await usersCollection.insertOne(newUser);
      
      console.log(`✅ New ${role} registered: ${username}`);
      res.status(201).json({ 
        message: "Registration successful!",
        username,
        role 
      });
      
    } catch (err) {
      console.error("❌ Registration error:", err);
      res.status(500).json({ message: "Error during registration" });
    }
  });
  
  // Login route
  app.post("/api/login", async (req, res) => {
    try {
      const { username, password } = req.body;
      
      if (!username || !password) {
        return res.status(400).json({ message: "Username and password are required" });
      }
      
      const user = await usersCollection.findOne({ username });
      if (!user) {
        return res.status(400).json({ message: "Invalid username or password" });
      }
      
      const validPassword = await bcrypt.compare(password, user.password);
      if (!validPassword) {
        return res.status(400).json({ message: "Invalid username or password" });
      }
      
      const token = jwt.sign(
        { 
          userId: user._id,
          username: user.username,
          role: user.role 
        },
        JWT_SECRET,
        { expiresIn: '24h' }
      );
      
      console.log(`✅ ${user.role} logged in: ${username}`);
      res.json({
        message: "Login successful!",
        token,
        role: user.role,
        username: user.username
      });
      
    } catch (err) {
      console.error("❌ Login error:", err);
      res.status(500).json({ message: "Error during login" });
    }
  });
  
  // Verify token route
  app.get("/api/verify", authenticateToken, (req, res) => {
    res.json({ 
      valid: true,
      user: {
        username: req.user.username,
        role: req.user.role
      }
    });
  });

  // ========== TEMPLATE ROUTES ==========
  
  // Get current template
  app.get("/api/template", authenticateToken, isAdmin, async (req, res) => {
    try {
      const settings = await settingsCollection.findOne({ type: 'certificate_template' });
      res.json({ template: settings?.template || 'classic' });
    } catch (err) {
      res.status(500).json({ message: "Error fetching template" });
    }
  });
  
  // Set template
  app.post("/api/template", authenticateToken, isAdmin, async (req, res) => {
    try {
      const { template } = req.body;
      
      if (!['classic', 'modern'].includes(template)) {
        return res.status(400).json({ message: "Invalid template" });
      }
      
      await settingsCollection.updateOne(
        { type: 'certificate_template' },
        { $set: { template, updatedAt: new Date() } },
        { upsert: true }
      );
      
      console.log(`✅ Template changed to: ${template}`);
      res.json({ message: "Template updated successfully", template });
      
    } catch (err) {
      console.error("❌ Template update error:", err);
      res.status(500).json({ message: "Error updating template" });
    }
  });

  // ========== PROTECTED ROUTES ==========
  
 // Upload Excel route
  app.post("/upload-excel", authenticateToken, isAdmin, uploadExcel.single("excel"), async (req, res) => {
    try {
      console.log("📂 Reading Excel file...");
      const workbook = xlsx.readFile("./uploads/excel/students.xlsx");
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = xlsx.utils.sheet_to_json(sheet);

      console.log("🔄 Processing", data.length, "records");
      
      // Debug: Check what columns are available
      if (data.length > 0) {
        console.log("📋 Excel columns found:", Object.keys(data[0]));
        console.log("📝 First row sample:", data[0]);
      }
      
      const mappedData = data.map((row) => ({
        name: row.NAME,
        certificate: row.CIRTIFICATES,
        link: row.links,
        college: row.college
      }));

      console.log("🗑️ Deleting old records...");
      await studentsCollection.deleteMany({});
      
      console.log("💾 Inserting new records...");
      await studentsCollection.insertMany(mappedData);

      console.log("✅ Success! Uploaded", mappedData.length, "records");
      res.json({
        message: "Excel uploaded successfully!",
        recordCount: mappedData.length
      });
    } catch (err) {
      console.error("❌ Upload error:", err);
      res.status(500).json({ message: "Error saving data", error: err.message });
    }
  });
  // Search route
app.get("/api/search", authenticateToken, async (req, res) => {
  const name = req.query.name;
  try {
    console.log(`🔍 ${req.user.username} (${req.user.role}) searching for: ${name}`);
    
    const count = await studentsCollection.countDocuments();
    if (count === 0) {
      return res.status(404).json({
        message: "No data available. Please contact admin to upload student data."
      });
    }
    
    const student = await studentsCollection.findOne({ name: name });
    
    if (!student) {
      return res.status(404).json({
        message: `No record found for "${name}"`
      });
    }

    res.json({
      success: true,
      student: {
        name: student.name,
        certificate: student.certificate,
        college: student.college,
        link: student.link
      }
    });
  } catch (err) {
    console.error("❌ Search error:", err);
    res.status(500).json({ message: "Error searching", error: err.message });
  }
});

 // Generate certificate route - UPDATED to handle token from query param
  app.get("/api/generate-certificate", async (req, res) => {
    const name = req.query.name;
    const token = req.query.token || req.headers['authorization']?.split(' ')[1];
    
    // Verify token
    if (!token) {
      return res.status(401).send('<h1>Access Denied</h1><p>Please login first</p>');
    }
    
    try {
      const verified = jwt.verify(token, JWT_SECRET);
      console.log(`📜 ${verified.username} generating certificate for: ${name}`);
      
      const student = await studentsCollection.findOne({ name: name });
      
      if (!student) {
        return res.status(404).send(`<h1>Not Found</h1><p>No record found for "${name}"</p>`);
      }

      // Get current template
      const settings = await settingsCollection.findOne({ type: 'certificate_template' });
      const template = settings?.template || 'classic';
      
      // Generate certificate based on template
      let certificateHTML;
      if (template === 'modern') {
        certificateHTML = generateModernCertificate(student.name, student.certificate);
      } else {
        certificateHTML = generateClassicCertificate(student.name, student.certificate);
      }
      
      // Add auto-print script to the certificate HTML
      certificateHTML = certificateHTML.replace('</body>', `
        <script>
          window.onload = function() {
            window.print();
          };
        </script>
      </body>`);
      
      res.setHeader('Content-Type', 'text/html');
      res.send(certificateHTML);
      
      console.log(`✅ Certificate generated for ${name} using ${template} template`);
      
    } catch (err) {
      console.error("❌ Certificate generation error:", err);
      res.status(500).send('<h1>Error</h1><p>Error generating certificate</p>');
    }
  });
  // Check database status route
  app.get("/api/status", authenticateToken, async (req, res) => {
    try {
      const count = await studentsCollection.countDocuments();
      const settings = await settingsCollection.findOne({ type: 'certificate_template' });
      
      res.json({
        status: "connected",
        recordCount: count,
        template: settings?.template || 'classic',
        message: count > 0 ? "Database ready" : "No data uploaded yet",
        user: {
          username: req.user.username,
          role: req.user.role
        }
      });
    } catch (err) {
      res.status(500).json({ status: "error", message: err.message });
    }
  });

  // Reload data from Excel file (for debugging/fixing data issues)
  app.get("/api/reload-data", authenticateToken, isAdmin, async (req, res) => {
    try {
      console.log("🔄 Reloading data from Excel file...");
      
      // Clear existing data
      await studentsCollection.deleteMany({});
      console.log("🗑️ Cleared old data");
      
      // Reload from Excel
      await loadExcelIfExists();
      
      const count = await studentsCollection.countDocuments();
      
      res.json({
        message: "Data reloaded successfully!",
        recordCount: count
      });
      
    } catch (err) {
      console.error("❌ Reload error:", err);
      res.status(500).json({ message: "Error reloading data", error: err.message });
    }
  });

  // ========== PUBLIC ROUTES (HTML Pages) ==========
  
  app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "login.html"));
  });
  
  app.get("/admin.html", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "admin.html"));
  });
  
  app.get("/search.html", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "search.html"));
  });

  app.listen(5000, () => {
    console.log("🚀 Server running on http://localhost:5000");
   
  });
}

connectDB();