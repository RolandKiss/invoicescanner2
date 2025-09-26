// api/invoice-scanner/index.js  
  
const Busboy = require("busboy");  
const fetch = require("node-fetch"); // If on Node 18+, you can use global fetch  
  
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
  
  let fileBuffer = [];  
  let mimeType = "";  
  let fileFound = false;  
  
  try {  
    // Parse the multipart form-data using Busboy  
    await new Promise((resolve, reject) => {  
      const busboy = Busboy({ headers: req.headers });  
      busboy.on("file", (fieldname, file, filename, encoding, mimetype) => {  
        fileFound = true;  
        mimeType = mimetype;  
        file.on("data", (data) => fileBuffer.push(data));  
      });  
      busboy.on("finish", resolve);  
      busboy.on("error", reject);  
      busboy.end(req.rawBody || req.body);  
    });  
  
    if (!fileFound) {  
      context.res = { status: 400, body: "No file uploaded" };  
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
  
    const buffer = Buffer.concat(fileBuffer);  
  
    // Prepare Document Intelligence Analyze API URL  
    const apiUrl = `${endpoint}/formrecognizer/documentModels/prebuilt-invoice:analyze?api-version=2023-07-31`;  
  
    // POST the raw file buffer with the correct Content-Type  
    const analyzeRes = await fetch(apiUrl, {  
      method: "POST",  
      headers: {  
        "Content-Type": mimeType,  
        "Ocp-Apim-Subscription-Key": apiKey  
      },  
      body: buffer  
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