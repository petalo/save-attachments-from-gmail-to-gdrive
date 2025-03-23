require("dotenv").config();

const API_KEY = process.env.GEMINI_API_KEY;
const MODELS_URL = `https://generativelanguage.googleapis.com/v1/models?key=${API_KEY}`;

async function listAvailableModels() {
  try {
    const response = await fetch(MODELS_URL, {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
      },
    });

    const data = await response.json();

    if (!response.ok) {
      console.error("❌ Error fetching models:", data);
      return;
    }

    console.log("✅ Available models:");
    data.models.forEach((model, index) => {
      console.log(`\n${index + 1}. ${model.name}`);
      if (model.description) console.log(`   📘 ${model.description}`);
      if (model.supportedGenerationMethods) {
        console.log(
          `   ⚙️ Supported methods: ${model.supportedGenerationMethods.join(
            ", "
          )}`
        );
      }
    });
  } catch (error) {
    console.error("❌ Network or request error:", error);
  }
}

listAvailableModels();
