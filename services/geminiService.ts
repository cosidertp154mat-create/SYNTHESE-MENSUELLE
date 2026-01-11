
import { GoogleGenAI } from "@google/genai";
import { DataPoint } from "../types";

/**
 * Generates an AI synthesis of performance data using the Gemini API.
 */
export const generateAiSynthesis = async (points: DataPoint[], periodLabel: string) => {
  // Always use a named parameter for the API key from process.env.API_KEY directly
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  const pointsSummary = points
    .sort((a, b) => a.date.localeCompare(b.date))
    .map(p => `- Le ${p.date}: ${p.label} (${p.category}): ${p.value}/${p.maxValue}`)
    .join('\n');
  
  const prompt = `
    J'ai une liste de points obtenus pour la période du ${periodLabel}.
    Note : Dans ce contexte, les week-ends sont le Vendredi et le Samedi.
    
    Voici les données chronologiques :
    ${pointsSummary}

    Peux-tu me fournir une synthèse professionnelle en français comprenant :
    1. Une analyse de la régularité sur la période du 21 au 20.
    2. Les points forts identifiés.
    3. Une observation sur la répartition des scores (début vs fin de période).
    4. Une conclusion stratégique.

    Réponds en format Markdown structuré.
  `;

  try {
    // Using gemini-3-flash-preview for basic text task
    const response = await ai.models.generateContent({
      model: 'gemini-3-flash-preview',
      contents: prompt,
      config: {
        systemInstruction: "Agis en tant qu'analyste de performance expert.",
      }
    });
    // Direct access to response.text property (not a method)
    return response.text;
  } catch (error) {
    console.error("Erreur lors de la génération de la synthèse IA:", error);
    return "Désolé, une erreur est survenue lors de la génération de la synthèse par l'IA.";
  }
};
