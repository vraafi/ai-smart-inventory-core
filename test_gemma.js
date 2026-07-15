const https = require('https');

const apiKey = "AIzaSyDI6s__DL5lkmWfy-7aGiWxC8WmbZuzG_M";
const model = "gemma-4-31b-it";
const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

const payload = {
  systemInstruction: {
    parts: [{ text: "You are a test agent." }]
  },
  contents: [
    { role: "user", parts: [{ text: "Hello, reply with JSON." }] }
  ],
  generationConfig: {
    temperature: 0.0
  }
};

const req = https.request(url, {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json'
  }
}, (res) => {
  let data = '';
  res.on('data', chunk => data += chunk);
  res.on('end', () => {
    console.log("Status Code:", res.statusCode);
    console.log("Response Body:", data);
  });
});

req.on('error', console.error);
req.write(JSON.stringify(payload));
req.end();
