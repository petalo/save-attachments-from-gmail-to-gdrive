require("dotenv").config();
const { GoogleGenerativeAI } = require("@google/generative-ai");

const API_KEY = process.env.GEMINI_API_KEY;

async function listGeminiModels() {
  if (!API_KEY) {
    console.error("Error: GEMINI_API_KEY not found in .env");
    return;
  }

  const genAI = new GoogleGenerativeAI(API_KEY);

  console.log("genAI object:", genAI);
  console.log("genAI keys:", Object.keys(genAI));

  try {
    const models = await genAI.getAvailableModels();
    console.log("Models:", models);
  } catch (error) {
    console.error("Error:", error);
  }
}

listGeminiModels();
