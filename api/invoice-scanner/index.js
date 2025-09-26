const Busboy = require("busboy");  
const fetch = require("node-fetch"); // For Node <18; otherwise use global fetch  
  
const POLL_INTERVAL_MS = 2000;  
const MAX_POLL_TRIES = 15;  
const SUPPORTED_TYPES = [  
  "application/pdf",  
  "application/octet-stream", // fallback for PDFs  
  "image/jpeg",  
  "image/png",  
  "image/tiff",  
  "image/bmp",  
  "image/heif",  
  "image/heic"  
];  
  
module.exports = async function (context, req) {  
  // Only allow POST  
  if (req.method !== "POST") {  
    context.res = { status: 405, body: "Method Not Allowed" };  
    return;  
  }  
  
  // Validate environment variables early  
  const endpoint = process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT;  
  const apiKey = process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY;  
  if (!endpoint || !apiKey) {  
    context.res = { status: 500, body: "Missing Document Intelligence config" };  
    return;  
  }  
  
  let fileBuffer, mimeType, fileFound = false;  
  
  try {  
    // Check Content-Type  
    const contentType = req.headers['content-type'] || '';  
  
    if (contentType.startsWith('multipart/form-data')) {  
      // Use Busboy for multipart uploads (your old logic)  
      let parts = [];  
      await new Promise((resolve, reject) => {  
        const busboy = Busboy({ headers: req.headers });  
        busboy.on("file", (fieldname, file, filename, encoding, mimetype) => {  
          fileFound = true;  
          mimeType = mimetype;  
          file.on("data", (data) => parts.push(data));  
        });  
        busboy.on("finish", resolve);  
        busboy.on("error", reject);  
        busboy.end(req.rawBody || req.body);  
      });  
  
      if (!fileFound) {  
        context.res = { status: 400, body: "No file uploaded" };  
        return;  
      }  
      fileBuffer = Buffer.concat(parts);  
  
    } else if (  
      // Accept raw uploads of supported types  
      SUPPORTED_TYPES.some(type => contentType.startsWith(type))  
    ) {  
        if (Buffer.isBuffer(req.body)) {  
        fileBuffer = req.body;  
        } else if (typeof req.body === 'string') {  
        // Is it base64? If so, decode!  
        if (/^[A-Za-z0-9+/=]+$/.test(req.body.trim())) {  
            fileBuffer = Buffer.from(req.body, 'base64');  
        } else {  
            // This is probably a bug - warn or throw  
            throw new Error('req.body is a string, but not base64. Uploads may be broken.');  
        }  
        } else {  
        throw new Error('req.body is not a Buffer or string');  
        }  
        
      mimeType = contentType;  
      fileFound = true;  
  
    } else {  
      context.res = { status: 415, body: `Unsupported content type: ${contentType}` };  
      return;  
    }  
  
    // Accept octet-stream if the file extension is .pdf  
    const isPdf =  
      mimeType === "application/pdf" ||  
      (mimeType === "application/octet-stream" &&  
        req.headers["x-file-name"] &&  
        req.headers["x-file-name"].toLowerCase().endsWith(".pdf"));  
  
    if (!SUPPORTED_TYPES.includes(mimeType) && !isPdf) {  
      context.res = { status: 415, body: `Unsupported file type: ${mimeType}` };  
      return;  
    }  
  
    // Prepare Document Intelligence Analyze API URL  
    const apiUrl = `${endpoint}/formrecognizer/documentModels/prebuilt-invoice:analyze?api-version=2023-07-31`;  
    
    console.log('fileBuffer.length:', fileBuffer.length);  
    console.log('First 20 bytes:', fileBuffer.slice(0, 20).toString('hex'));  
    // POST the raw file buffer with the correct Content-Type  
    const analyzeRes = await fetch(apiUrl, {  
      method: "POST",  
      headers: {  
        "Content-Type": mimeType,  
        "Ocp-Apim-Subscription-Key": apiKey  
      },  
      body: fileBuffer  
    });  
  
    if (!analyzeRes.ok) {  
      const errorText = await analyzeRes.text();  
      context.res = { status: analyzeRes.status, body: errorText };  
      return;  
    }  
  
    const operationLocation = analyzeRes.headers.get("operation-location");  
    if (!operationLocation) {  
      context.res = { status: 500, body: "No operation-location header in response" };  
      return;  
    }  
  
    // Poll until the analysis is complete  
    let pollRes, pollData;  
    let tries = 0;  
    do {  
      await new Promise((r) => setTimeout(r, POLL_INTERVAL_MS));  
      pollRes = await fetch(operationLocation, {  
        headers: { "Ocp-Apim-Subscription-Key": apiKey }  
      });  
      pollData = await pollRes.json();  
      tries++;  
    } while (  
      pollData.status &&  
      pollData.status.toLowerCase() === "running" &&  
      tries < MAX_POLL_TRIES  
    );  
  
    context.res = {  
      status: 200,  
      body: pollData,  
      headers: { "Content-Type": "application/json" }  
    };  
  } catch (err) {  
    context.res = { status: 500, body: "Server error: " + err.message };  
  }  
};  