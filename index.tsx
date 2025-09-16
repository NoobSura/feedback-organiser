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
const analyzeButton = document.getElementById('analyze-button') as HTMLButtonElement;
const exportButton = document.getElementById('export-button') as HTMLButtonElement;
const exportExplodedButton = document.getElementById('export-exploded-button') as HTMLButtonElement;
const loadingSpinner = document.getElementById('loading-spinner') as HTMLDivElement;
const resultsSection = document.getElementById('results-section') as HTMLElement;
const resultsHeader = document.getElementById('results-header') as HTMLElement;
const resultsTableContainer = document.getElementById('results-table-container') as HTMLElement;
const statusMessage = document.getElementById('status-message') as HTMLParagraphElement;
const labelCountsSummary = document.getElementById('label-counts-summary') as HTMLDivElement;


const API_KEY = process.env.API_KEY;
if (!API_KEY) {
  resultsSection.innerHTML = '<p class="error">API_KEY environment variable not set. Please configure it to use the application.</p>';
  throw new Error("API_KEY not set");
}
const ai = new GoogleGenAI({ apiKey: API_KEY });

let analysisResults: { feedback: string; labels: string[]; isIncorrect?: boolean }[] = [];

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

  // UI Reset
  resultsTableContainer.innerHTML = '';
  resultsHeader.hidden = true;
  exportButton.disabled = true;
  exportExplodedButton.disabled = true;
  loadingSpinner.hidden = false;
  analyzeButton.disabled = true;
  statusMessage.textContent = '';
  statusMessage.classList.remove('error');
  labelCountsSummary.innerHTML = '';
  labelCountsSummary.hidden = true;
  analysisResults = [];

  const feedbackLines = feedbackText.split('\n').filter(line => line.trim() !== '');
  const BATCH_SIZE = 250;
  const totalBatches = Math.ceil(feedbackLines.length / BATCH_SIZE);

  const systemPrompt = systemPromptInput.value.trim();
  
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
    for (let i = 0; i < totalBatches; i++) {
        const batchStart = i * BATCH_SIZE;
        const batchEnd = batchStart + BATCH_SIZE;
        const batchLines = feedbackLines.slice(batchStart, batchEnd);
        const batchText = batchLines.join('\n');

        statusMessage.textContent = `Processing batch ${i + 1} of ${totalBatches}...`;

        const prompt = `Analyze each line of the following customer feedback individually. Return a JSON array where each object corresponds to a single line of feedback from the input. Do not skip, merge, or alter any lines. Each object must contain two keys: 'feedback' (the original, unmodified text of the feedback line) and 'labels' (an array of relevant string labels, e.g., 'Bug Report', 'Feature Request').\n\n---FEEDBACK---\n${batchText}`;

        const response = await ai.models.generateContent({
            model: "gemini-2.5-flash",
            contents: prompt,
            config: modelConfig,
        });

        const resultText = response.text;
        const labeledFeedbackBatch = JSON.parse(resultText);

        if (labeledFeedbackBatch && labeledFeedbackBatch.length > 0) {
            analysisResults.push(...labeledFeedbackBatch.map((item: any) => ({ ...item, isIncorrect: false })));
        }
    }

    if (analysisResults.length > 0) {
        statusMessage.textContent = `Analysis complete! ${analysisResults.length} items processed.`;
        renderResultsTable();
        updateAndRenderStats();
        resultsHeader.hidden = false;
        exportButton.disabled = false;
        exportExplodedButton.disabled = false;
    } else {
        statusMessage.textContent = 'No feedback could be analyzed. Please check your input.';
    }

  } catch (error) {
    console.error(error);
    statusMessage.textContent = `An error occurred during analysis. Please try again.`;
    statusMessage.classList.add('error');
    resultsTableContainer.innerHTML = `<p class="error">An error occurred while analyzing the feedback. Please try again.</p>`;
  } finally {
    loadingSpinner.hidden = true;
    checkInputs();
  }
});

function normalizeLabel(label: string): string {
    return label.trim().toLowerCase().replace(/^#+/, '');
}

function updateAndRenderStats() {
    const labelCounts = new Map<string, number>();
    analysisResults.forEach(item => {
        item.labels.forEach(label => {
            const normalizedLabel = normalizeLabel(label);
            if(normalizedLabel) {
                labelCounts.set(normalizedLabel, (labelCounts.get(normalizedLabel) || 0) + 1);
            }
        });
    });

    labelCountsSummary.innerHTML = '';
    if (labelCounts.size === 0) {
        labelCountsSummary.hidden = true;
        return;
    }
    labelCountsSummary.hidden = false;

    const sortedLabels = [...labelCounts.entries()].sort((a, b) => b[1] - a[1]);

    const summaryHtml = sortedLabels.map(([label, count]) => 
        `<span class="label-count-item">
            <span class="label-name">${label}</span>
            <span class="label-count">${count}</span>
        </span>`
    ).join('');

    labelCountsSummary.innerHTML = `<strong>Label Counts:</strong> ${summaryHtml}`;
}

function renderResultsTable() {
  resultsTableContainer.innerHTML = '';

  if (analysisResults.length === 0) return;

  const table = document.createElement('table');
  table.className = 'results-table';

  const thead = document.createElement('thead');
  thead.innerHTML = `
    <tr>
      <th>Feedback</th>
      <th>Labels (Editable)</th>
      <th class="incorrect-col">Incorrect?</th>
    </tr>
  `;
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  analysisResults.forEach((item, index) => {
    const row = document.createElement('tr');
    if (item.isIncorrect) {
      row.classList.add('incorrect-row');
    }
    const feedbackCell = document.createElement('td');
    feedbackCell.textContent = item.feedback;
    
    const labelsCell = document.createElement('td');
    const labelsInput = document.createElement('input');
    labelsInput.type = 'text';
    labelsInput.className = 'labels-input';
    labelsInput.value = item.labels.join(', ');
    labelsInput.dataset.index = index.toString();

    labelsInput.addEventListener('change', (e) => {
      const target = e.target as HTMLInputElement;
      const updatedIndex = parseInt(target.dataset.index!, 10);
      const updatedLabels = target.value.split(',').map(l => l.trim()).filter(Boolean);
      analysisResults[updatedIndex].labels = updatedLabels;
      updateAndRenderStats(); // Update stats on label change
    });
    
    const incorrectCell = document.createElement('td');
    incorrectCell.className = 'incorrect-col';
    const incorrectCheckbox = document.createElement('input');
    incorrectCheckbox.type = 'checkbox';
    incorrectCheckbox.checked = !!item.isIncorrect;
    incorrectCheckbox.dataset.index = index.toString();
    incorrectCheckbox.title = "Mark this analysis as incorrect";

    incorrectCheckbox.addEventListener('change', (e) => {
        const target = e.target as HTMLInputElement;
        const updatedIndex = parseInt(target.dataset.index!, 10);
        analysisResults[updatedIndex].isIncorrect = target.checked;
        if (target.checked) {
            row.classList.add('incorrect-row');
        } else {
            row.classList.remove('incorrect-row');
        }
    });

    labelsCell.appendChild(labelsInput);
    incorrectCell.appendChild(incorrectCheckbox);

    row.appendChild(feedbackCell);
    row.appendChild(labelsCell);
    row.appendChild(incorrectCell);
    tbody.appendChild(row);
  });

  table.appendChild(tbody);
  resultsTableContainer.appendChild(table);
}


exportButton.addEventListener('click', () => {
    if (analysisResults.length > 0) {
        generateAndDownloadExcel(analysisResults);
    }
});

exportExplodedButton.addEventListener('click', () => {
    if (analysisResults.length > 0) {
        generateAndDownloadExplodedExcel(analysisResults);
    }
});

function generateAndDownloadExcel(labeledFeedback: { feedback: string; labels: string[], isIncorrect?: boolean }[]) {
    const dataForSheet = labeledFeedback.map(item => ({
        Feedback: item.feedback,
        Labels: item.labels.join(', '),
        'Incorrect Analysis': item.isIncorrect ? 'Yes' : 'No',
    }));

    const ws = XLSX.utils.json_to_sheet(dataForSheet);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Feedback Analysis");
    XLSX.writeFile(wb, "feedback_analysis_results.xlsx");
}

function generateAndDownloadExplodedExcel(labeledFeedback: { feedback: string; labels: string[], isIncorrect?: boolean }[]) {
    const wb = XLSX.utils.book_new();

    // --- Sheet 1: Exploded by Label ---
    const explodedData: { Feedback: string; 'Single Label': string; 'Incorrect Analysis': string }[] = [];
    labeledFeedback.forEach(item => {
        if (item.labels.length === 0) {
            explodedData.push({
                Feedback: item.feedback,
                'Single Label': '',
                'Incorrect Analysis': item.isIncorrect ? 'Yes' : 'No',
            });
        } else {
            item.labels.forEach(label => {
                explodedData.push({
                    Feedback: item.feedback,
                    'Single Label': label,
                    'Incorrect Analysis': item.isIncorrect ? 'Yes' : 'No',
                });
            });
        }
    });
    const wsExploded = XLSX.utils.json_to_sheet(explodedData);
    XLSX.utils.book_append_sheet(wb, wsExploded, "Exploded by Label");


    // --- Sheet 2: Compiled View ---
    const compiledData = labeledFeedback.map(item => ({
        Feedback: item.feedback,
        Labels: item.labels.join(', '),
        'Incorrect Analysis': item.isIncorrect ? 'Yes' : 'No',
    }));
    const wsCompiled = XLSX.utils.json_to_sheet(compiledData);
    XLSX.utils.book_append_sheet(wb, wsCompiled, "Compiled View");


    // --- Sheet 3: Label Counts ---
    const labelCounts = new Map<string, number>();
    labeledFeedback.forEach(item => {
        item.labels.forEach(label => {
            const normalizedLabel = normalizeLabel(label);
            if(normalizedLabel) {
                labelCounts.set(normalizedLabel, (labelCounts.get(normalizedLabel) || 0) + 1);
            }
        });
    });
    const countsData = Array.from(labelCounts.entries())
        .map(([label, count]) => ({ Label: label, Count: count }))
        .sort((a, b) => b.Count - a.Count);
        
    const wsCounts = XLSX.utils.json_to_sheet(countsData);
    XLSX.utils.book_append_sheet(wb, wsCounts, "Label Counts");

    // --- Download the file ---
    XLSX.writeFile(wb, "feedback_analysis_detailed_export.xlsx");
}