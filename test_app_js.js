const fs = require('fs');
const path = require('path');

// Simulate the file that app.js sends
const fileBuffer = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000a49444154789c63000100000500010d0a2db40000000049454e44ae426082', 'hex');

async function run() {
  const file = new Blob([fileBuffer], { type: 'image/png' });
  const formData = new FormData();
  formData.append('image', file, 'test.png');
  
  try {
    const uploadResponse = await fetch('http://localhost:8000/api/gambarsoal', { method: 'POST', body: formData });
    const text = await uploadResponse.text();
    console.log("Status:", uploadResponse.status);
    console.log("Response:", text);
  } catch (e) {
    console.error("Fetch failed:", e);
  }
}
run();
