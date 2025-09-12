/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */
import { GoogleGenAI, Type, GenerateContentConfig } from "@google/genai";
import * as XLSX from 'xlsx';

const fileInput = document.getElementById('file-upload') as HTMLInputElement;
const fileNameSpan = document.getElementById('file-name') as HTMLSpanElement;
const feedbackInput = document.getElementById('feedback-input') as HTMLTextAreaElement;
const systemPromptInput = document.getElementById('system-prompt-input') as HTMLTextAreaElement;
const customLabelsInput = document.getElementById('custom-labels-input') as HTMLInputElement;
const analyzeButton = document.getElementById('analyze-button') as HTMLButtonElement;
const loadingSpinner = document.getElementById('loading-spinner') as HTMLDivElement;
const resultsSection = document.getElementById('results-section') as HTMLElement;

const API_KEY = process.env.API_KEY;
if (!API_KEY) {
  resultsSection.innerHTML = '<p class="error">API_KEY environment variable not set. Please configure it to use the application.</p>';
  throw new Error("API_KEY not set");
}
const ai = new GoogleGenAI({ apiKey: API_KEY });

const checkInputs = () => {
  analyzeButton.disabled = !feedbackInput.value.trim();
};

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (file) {
    fileNameSpan.textContent = file.name;
    const reader = new FileReader();

    if (file.name.endsWith('.csv')) {
        reader.onload = (e) => {
          feedbackInput.value = e.target?.result as string;
          checkInputs();
        };
        reader.onerror = () => {
            resultsSection.innerHTML = `<p class="error">Error reading CSV file.</p>`;
        };
        reader.readAsText(file);
    } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target?.result as ArrayBuffer);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const json: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                const feedbackColumn = json
                    .map(row => row[0]) // Get first column
                    .filter(cell => cell !== undefined && cell !== null && String(cell).trim() !== '') // Filter out empty cells
                    .join('\n');
                feedbackInput.value = feedbackColumn;
                checkInputs();
            } catch (err) {
                 resultsSection.innerHTML = `<p class="error">Error parsing Excel file.</p>`;
                 console.error(err);
            }
        };
        reader.onerror = () => {
            resultsSection.innerHTML = `<p class="error">Error reading Excel file.</p>`;
        };
        reader.readAsArrayBuffer(file);
    } else {
        resultsSection.innerHTML = `<p class="error">Unsupported file type. Please upload a CSV or Excel file.</p>`;
        fileNameSpan.textContent = 'No file chosen';
        fileInput.value = ''; // Clear the input
    }
  } else {
    fileNameSpan.textContent = 'No file chosen';
  }
});

feedbackInput.addEventListener('input', checkInputs);

analyzeButton.addEventListener('click', async () => {
  const feedbackText = feedbackInput.value.trim();
  if (!feedbackText) return;

  resultsSection.innerHTML = '';
  loadingSpinner.hidden = false;
  analyzeButton.disabled = true;

  const systemPrompt = systemPromptInput.value.trim();
  const customLabelsValue = customLabelsInput.value.trim();
  const customLabels = customLabelsValue ? customLabelsValue.split(',').map(label => label.trim()).filter(label => label) : [];

  let prompt = `Please analyze the following customer feedback. For each distinct piece of feedback, provide the original feedback text and assign a few relevant labels (e.g., 'Bug Report', 'Feature Request', 'Positive Feedback', 'UI/UX').`;

  if (customLabels.length > 0) {
    prompt += ` In addition to any other relevant labels, please consider using the following custom labels if they are applicable: ${customLabels.join(', ')}.`;
  }

  prompt += `\n\n---FEEDBACK---\n${feedbackText}`;

  const modelConfig: GenerateContentConfig = {
    responseMimeType: "application/json",
    responseSchema: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          feedback: {
            type: Type.STRING,
            description: "The original piece of customer feedback.",
          },
          labels: {
            type: Type.ARRAY,
            items: {
              type: Type.STRING
            },
            description: "A list of concise labels for the feedback.",
          },
        },
        required: ["feedback", "labels"],
      },
    },
  };

  if (systemPrompt) {
    modelConfig.systemInstruction = systemPrompt;
  }

  try {
    const response = await ai.models.generateContent({
      model: "gemini-2.5-flash",
      contents: prompt,
      config: modelConfig,
    });

    const resultText = response.text;
    const labeledFeedback = JSON.parse(resultText);

    if (labeledFeedback.length > 0) {
        generateAndDownloadExcel(labeledFeedback);
        resultsSection.innerHTML = '<p class="success">Analysis complete! Your Excel file has been downloaded.</p>';
    } else {
        resultsSection.innerHTML = '<p>No feedback could be analyzed. Please check your input.</p>';
    }

  } catch (error) {
    console.error(error);
    resultsSection.innerHTML = `<p class="error">An error occurred while analyzing the feedback. Please try again.</p>`;
  } finally {
    loadingSpinner.hidden = true;
    checkInputs();
  }
});

function generateAndDownloadExcel(labeledFeedback: { feedback: string; labels: string[] }[]) {
    const dataForSheet = labeledFeedback.map(item => ({
        Feedback: item.feedback,
        Labels: item.labels.join(', '),
    }));

    const ws = XLSX.utils.json_to_sheet(dataForSheet);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Feedback Analysis");
    XLSX.writeFile(wb, "feedback_analysis_results.xlsx");
}