const express = require('express');
const multer = require('multer');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
const port = 3000;

// Enable CORS for all routes
app.use(cors());

// Add route handler for root path to redirect to extractus_demo.html
app.get('/', (req, res) => {
  res.redirect('/extractus_demo.html');
});

// Serve static files from the current directory
app.use(express.static('./'));

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    const uploadDir = './uploads';
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir);
    }
    cb(null, uploadDir);
  },
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  }
});

const upload = multer({ storage: storage });

// Store progress and results for each extraction
const progressos = {};
const resultados = {};

// Endpoint for starting extraction
app.post('/api/extractus', upload.single('file'), (req, res) => {
  const excelPath = req.file.path;
  const saveDir = path.join(__dirname, 'downloads');
  
  // Create downloads directory if it doesn't exist
  if (!fs.existsSync(saveDir)) {
    fs.mkdirSync(saveDir);
  }
  
  const dueDay = req.body.dueDay;
  const monthNumber = req.body.monthNumber;
  const tipos = JSON.parse(req.body.tipos);
  
  // Generate unique ID for this extraction
  const extractionId = Date.now().toString();
  
  // Initialize progress and results
  progressos[extractionId] = {};
  resultados[extractionId] = {};
  
  tipos.forEach(tipo => {
    progressos[extractionId][tipo] = 0;
    resultados[extractionId][tipo] = 0;
  });
  
  console.log(`Starting extraction ${extractionId} with parameters:
    Excel: ${excelPath}
    Save Dir: ${saveDir}
    Due Day: ${dueDay}
    Month: ${monthNumber}
    Types: ${tipos.join(', ')}
  `);
  
  // Execute Python bridge script
  const pythonProcess = spawn('python', [
    'extractus_bridge.py',
    excelPath,
    saveDir,
    dueDay,
    monthNumber,
    tipos.join(',')
  ]);
  
  let stdoutData = '';
  
  pythonProcess.stdout.on('data', (data) => {
    const dataStr = data.toString();
    stdoutData += dataStr;
    
    try {
      // Try to parse JSON from the output
      // The Python script may output multiple JSON objects, so we need to handle each one
      const jsonLines = dataStr.split('\n').filter(line => line.trim().startsWith('{'));
      
      jsonLines.forEach(line => {
        try {
          const jsonData = JSON.parse(line);
          
          if (jsonData.progresso) {
            const { tipo, porcentagem } = jsonData.progresso;
            progressos[extractionId][tipo] = porcentagem;
            console.log(`Progress update for ${extractionId}/${tipo}: ${porcentagem}%`);
          }
          
          if (jsonData.resultados) {
            resultados[extractionId] = jsonData.resultados;
            console.log(`Results for ${extractionId}:`, jsonData.resultados);
          }
        } catch (e) {
          // Ignore parsing errors for incomplete JSON
        }
      });
    } catch (error) {
      console.error('Error processing Python script output:', error);
    }
  });
  
  pythonProcess.stderr.on('data', (data) => {
    console.error(`Python script error: ${data}`);
  });
  
  pythonProcess.on('close', (code) => {
    console.log(`Python script exited with code ${code}`);
    
    // If we haven't parsed any results yet, try to parse the complete stdout
    if (!resultados[extractionId].boleto && !resultados[extractionId].nota && !resultados[extractionId].relatorio) {
      try {
        const jsonMatch = stdoutData.match(/\{.*\}/s);
        if (jsonMatch) {
          const jsonData = JSON.parse(jsonMatch[0]);
          if (jsonData.resultados) {
            resultados[extractionId] = jsonData.resultados;
          }
        }
      } catch (error) {
        console.error('Error parsing final output:', error);
      }
    }
    
    // Mark all processes as complete
    tipos.forEach(tipo => {
      progressos[extractionId][tipo] = 100;
    });
  });
  
  res.json({ extractionId });
});

// Endpoint for tracking progress
app.get('/api/extractus/progress', (req, res) => {
  const extractionId = req.query.extractionId;
  const tipo = req.query.tipo;
  
  if (!extractionId || !tipo || !progressos[extractionId]) {
    return res.status(400).json({ error: 'Invalid parameters' });
  }
  
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  
  // Send current progress
  res.write(`data: ${JSON.stringify({
    progresso: { tipo, porcentagem: progressos[extractionId][tipo] }
  })}\n\n`);
  
  // Send results if available
  if (resultados[extractionId][tipo]) {
    res.write(`data: ${JSON.stringify({
      resultados: { [tipo]: resultados[extractionId][tipo] }
    })}\n\n`);
  }
  
  // Set up interval to send updates
  const interval = setInterval(() => {
    res.write(`data: ${JSON.stringify({
      progresso: { tipo, porcentagem: progressos[extractionId][tipo] }
    })}\n\n`);
    
    if (progressos[extractionId][tipo] >= 100) {
      res.write(`data: ${JSON.stringify({
        resultados: { [tipo]: resultados[extractionId][tipo] || 0 }
      })}\n\n`);
      
      clearInterval(interval);
      res.end();
    }
  }, 1000);
  
  // Clean up interval if client disconnects
  req.on('close', () => {
    clearInterval(interval);
  });
});

// Endpoint for downloading results
app.get('/api/extractus/download', (req, res) => {
  const downloadDir = path.join(__dirname, 'downloads');
  
  // In a real implementation, you would zip the files and send them
  // For now, just return a success message
  res.json({ message: 'Download endpoint. In a real implementation, this would provide a ZIP file of the results.' });
});

// Start the server
app.listen(port, () => {
  console.log(`Extractus backend server running at http://localhost:${port}`);
  console.log(`Access the demo at http://localhost:${port}/extractus_demo.html`);
});