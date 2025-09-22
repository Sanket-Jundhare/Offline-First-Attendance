const axios = require("axios");

async function askAttendanceAI(records, question) {
  try {
    const prompt = `You are an attendance assistant. Use the following attendance records to answer the question.\n\nAttendance Records:\n${JSON.stringify(records, null, 2)}\n\nQuestion: ${question}\nAnswer:`;

    console.log("Making request to Gemini API...");
    
    // Use the correct endpoint for Gemini 1.0 Pro
    const response = await axios.post(
      `https://generativelanguage.googleapis.com/v1/models/gemini-1.0-pro:generateContent?key=${process.env.GEMINI_API_KEY}`,
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
          temperature: 0.5,
          maxOutputTokens: 500,
          topP: 0.8,
          topK: 40
        },
        safetySettings: [
          {
            category: "HARM_CATEGORY_HARASSMENT",
            threshold: "BLOCK_MEDIUM_AND_ABOVE"
          },
          {
            category: "HARM_CATEGORY_HATE_SPEECH",
            threshold: "BLOCK_MEDIUM_AND_ABOVE"
          }
        ]
      },
      {
        headers: {
          "Content-Type": "application/json"
        },
        timeout: 30000
      }
    );

    console.log("API Response received");
    
    // Extract the response text
    if (response.data && response.data.candidates && response.data.candidates.length > 0) {
      const responseText = response.data.candidates[0].content.parts[0].text;
      return responseText;
    } else {
      return "❌ No response generated from AI. Please try again.";
    }
  } catch (err) {
    console.error("Full Gemini AI Error:", err.message);
    console.error("Error response data:", err.response?.data);
    console.error("Error status:", err.response?.status);
    
    if (err.response?.status === 404) {
      return "❌ Model not found. Please check if your API key has access to the Gemini models.";
    } else if (err.response?.status === 400) {
      return "❌ Bad request. Please check your API key and request format.";
    } else if (err.response?.status === 403) {
      return "❌ Permission denied. Please check your API key and billing status.";
    } else if (err.code === 'ENOTFOUND') {
      return "❌ Network error. Please check your internet connection.";
    }
    
    return "❌ Error contacting Gemini AI. Please check your API key and try again.";
  }
}

module.exports = { askAttendanceAI };