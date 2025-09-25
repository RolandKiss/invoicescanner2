// api/invoice-scanner/index.js  
const Busboy = require("busboy");  
const fetch = require("node-fetch");  
  
module.exports = async function (context, req) {  
  if (req.method !== "POST") {  
    context.res = { status: 405, body: "Method Not Allowed" };  
    return;  
  }  
  
  let fileBuffer = [];  
  let mimeType = "";  
  let fileFound = false;  
  
  try {  
    await new Promise((resolve, reject) => {  
      const busboy = new Busboy({ headers: req.headers });  
      busboy.on("file", (fieldname, file, filename, encoding, mimetype) => {  
        fileFound = true;  
        mimeType = mimetype;  
        file.on("data", (data) => fileBuffer.push(data));  
      });  
      busboy.on("finish", resolve);  
      busboy.on("error", reject);  
  
      busboy.end(req.rawBody);  
    });  
  
    if (!fileFound) {  
      context.res = { status: 400, body: "No file uploaded" };  
      return;  
    }  
  
    const buffer = Buffer.concat(fileBuffer);  
  
    const endpoint = process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT;  
    const apiKey = process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY;  
  
    if (!endpoint || !apiKey) {  
      context.res = { status: 500, body: "Missing Document Intelligence config" };  
      return;  
    }  
  
    const apiUrl = `${endpoint}/formrecognizer/documentModels/prebuilt-invoice:analyze?api-version=2023-07-31`;  
  
    // Submit document for analysis  
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
  
    // The Analyze API is async; you must poll the "operation-location" header  
    const operationLocation = analyzeRes.headers.get("operation-location");  
    if (!operationLocation) {  
      context.res = { status: 500, body: "No operation-location header in response" };  
      return;  
    }  
  
    // Poll until done  
    let pollRes, pollData;  
    let tries = 0;  
    do {  
      await new Promise((r) => setTimeout(r, 2000));  
      pollRes = await fetch(operationLocation, {  
        headers: { "Ocp-Apim-Subscription-Key": apiKey }  
      });  
      pollData = await pollRes.json();  
      tries++;  
    } while (  
      pollData.status &&  
      pollData.status.toLowerCase() === "running" &&  
      tries < 15  
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