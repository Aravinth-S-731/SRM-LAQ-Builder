let weeksData = [
    { week: 7, question: "", answer: "", exported: false, files: null }
];
let activeIndex = 0;

// Elements
const weekListEl = document.getElementById("weekList");
const weekNumberEl = document.getElementById("weekNumber");
const questionInputEl = document.getElementById("questionInput");
const answerInputEl = document.getElementById("answerInput");
const rawPortalInputEl = document.getElementById("rawPortalInput");
const wordCountBadge = document.getElementById("wordCountBadge");
const pageEstBadge = document.getElementById("pageEstBadge");
const statusMsg = document.getElementById("statusMsg");
const downloadLinks = document.getElementById("downloadLinks");
const apiKeyEl = document.getElementById("apiKey");
const courseCodeEl = document.getElementById("courseCode");

// Load saved API key & course code from localStorage
apiKeyEl.value = localStorage.getItem("srm_gemini_api_key") || "";
courseCodeEl.value = localStorage.getItem("srm_course_code") || "ENT";

apiKeyEl.addEventListener("input", () => localStorage.setItem("srm_gemini_api_key", apiKeyEl.value));
courseCodeEl.addEventListener("input", () => localStorage.setItem("srm_course_code", courseCodeEl.value));

// Smart Parser for SRM Portal Copy-Paste
function parseSrmPortalText(rawText) {
    if (!rawText.trim()) return;

    // 1. Extract "Week X"
    const weekMatch = rawText.match(/Week\s*(\d+)/i);
    if (weekMatch) {
        weekNumberEl.value = parseInt(weekMatch[1], 10);
        weeksData[activeIndex].week = parseInt(weekMatch[1], 10);
    }

    // 2. Remove "Week X - LAQ" and "Instructions" headers to isolate the question
    let cleaned = rawText
        .replace(/Week\s*\d+\s*[-–—]?\s*LAQ/gi, "")
        .replace(/^\s*Instructions\s*$/gim, "")
        .trim();

    // Collapse multiple blank lines into clean spacing
    cleaned = cleaned.replace(/\n{2,}/g, " ").replace(/\s{2,}/g, " ").trim();

    if (cleaned) {
        questionInputEl.value = cleaned;
        weeksData[activeIndex].question = cleaned;
    }
    renderSidebar();
}

// Automatically parse as soon as user pastes into the Smart Portal box
rawPortalInputEl.addEventListener("input", (e) => parseSrmPortalText(e.target.value));
document.getElementById("parsePortalBtn").addEventListener("click", () => parseSrmPortalText(rawPortalInputEl.value));

// Update Word & Page Counter (Week7_ECS.pdf is ~1,400-1,700 words across 4 pages with tables/headings)
// Replace updateMetrics() in static/script.js:
function updateMetrics() {
    const text = answerInputEl.value.trim();
    const words = text ? text.split(/\s+/).length : 0;
    const estPages = (words / 310).toFixed(1);

    wordCountBadge.textContent = `${words} words`;
    pageEstBadge.textContent = `~${estPages} pages (Target: 4–5 pages)`;

    if (estPages >= 3.8 && estPages <= 5.5) {
        pageEstBadge.classList.add("ready");
    } else {
        pageEstBadge.classList.remove("ready");
    }
}

function renderSidebar() {
    weekListEl.innerHTML = "";
    weeksData.forEach((item, idx) => {
        const li = document.createElement("li");
        li.className = `week-item ${idx === activeIndex ? "active" : ""}`;
        li.innerHTML = `
            <span>Week ${item.week}</span>
            <span class="done-dot">${item.exported ? "✔ PDF Ready" : ""}</span>
        `;
        li.addEventListener("click", () => {
            activeIndex = idx;
            loadActiveWeek();
            renderSidebar();
        });
        weekListEl.appendChild(li);
    });
}

function loadActiveWeek() {
    const current = weeksData[activeIndex];
    weekNumberEl.value = current.week;
    questionInputEl.value = current.question;
    answerInputEl.value = current.answer;
    rawPortalInputEl.value = "";
    updateMetrics();
    renderDownloadLinks(current.files);
}

function renderDownloadLinks(files) {
    if (!files) {
        downloadLinks.innerHTML = "";
        return;
    }
    downloadLinks.innerHTML = `
        <a class="download-btn" href="${files.docx_url}" download>⬇ ${files.docx_filename}</a>
        <a class="download-btn" href="${files.pdf_url}" download>⬇ ${files.pdf_filename}</a>
    `;
}

// Sync input changes
weekNumberEl.addEventListener("input", () => {
    weeksData[activeIndex].week = parseInt(weekNumberEl.value || 1, 10);
    renderSidebar();
});
questionInputEl.addEventListener("input", () => {
    weeksData[activeIndex].question = questionInputEl.value;
});
answerInputEl.addEventListener("input", () => {
    weeksData[activeIndex].answer = answerInputEl.value;
    updateMetrics();
});

// Add Single Week
document.getElementById("addWeekBtn").addEventListener("click", () => {
    const lastWeek = weeksData.length ? weeksData[weeksData.length - 1].week : 0;
    weeksData.push({ week: lastWeek + 1, question: "", answer: "", exported: false, files: null });
    activeIndex = weeksData.length - 1;
    loadActiveWeek();
    renderSidebar();
});

// Bulk Create N Weeks
document.getElementById("bulkCreateBtn").addEventListener("click", () => {
    const count = parseInt(document.getElementById("bulkCount").value, 10);
    if (!count || count < 1) return;
    weeksData = [];
    for (let i = 1; i <= count; i++) {
        weeksData.push({ week: i, question: "", answer: "", exported: false, files: null });
    }
    activeIndex = 0;
    loadActiveWeek();
    renderSidebar();
});

// Copy Master Prompt for Gemini Pro Web Chat
// Copy Master Prompt for Gemini Pro Web Chat
document.getElementById("copyPromptBtn").addEventListener("click", () => {
    const q = questionInputEl.value.trim();
    if (!q) {
        statusMsg.textContent = "Please paste or enter a question first!";
        return;
    }
    const masterPrompt = `Write a comprehensive, formal, human-like academic response for the following university Long Answer Question (LAQ):

"${q}"

STRICT STRUCTURAL & LENGTH REQUIREMENTS (TARGET: 1,200 TO 1,350 WORDS / EXACTLY 4 PAGES):
1. Opening Definition & Core Framing (2 paragraphs) with scholarly (Author, Year) citation.
2. Core Pillars / Foundational Dimensions (### Heading + 2-3 numbered items on single paragraphs: 1. **Title:** Explanation... Do NOT use subheadings for pillars).
3. Strategic Importance in Contemporary Business (### Heading + 1 intro sentence + 4 bullet points with **Bold Lead-ins:** and concrete corporate examples).
4. Comparative Analysis Table (### Heading + 5-column Markdown table with 1 header row and 3-4 rows. Keep cell text concise: 8-14 words per cell).
5. How Leaders/Entrepreneurs Effectively Harness/Implement (### Heading + 5 numbered subheadings #### 1..5, each with 1 intro line and 3 concise bullet points with **Bold Lead-ins:**).
6. Conclusion (### Heading + 2 concise synthesizing paragraphs).

STYLE RULES: Write in a natural, authoritative academic tone. Avoid AI clichés ("delve", "tapestry", "in today's fast-paced world"). Do NOT include "Week #" or repeat the question at the start. Output clean Markdown.`;

    navigator.clipboard.writeText(masterPrompt);
    statusMsg.textContent = "✔ Master 4-Page Prompt copied! Paste it into Gemini Pro, then paste the Markdown response below.";
});

// Auto-Generate via Gemini API
document.getElementById("generateApiBtn").addEventListener("click", async () => {
    const q = questionInputEl.value.trim();
    const apiKey = apiKeyEl.value.trim();
    const model = document.getElementById("geminiModel").value;

    if (!q) {
        statusMsg.textContent = "Please enter or extract a question first.";
        return;
    }

    const btn = document.getElementById("generateApiBtn");
    btn.disabled = true;
    statusMsg.textContent = "⏳ Generating 4–5 page formal response with Gemini...";

    try {
        const res = await fetch("/api/generate-answer", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ api_key: apiKey, question: q, model })
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || "Generation failed");

        answerInputEl.value = data.answer;
        weeksData[activeIndex].answer = data.answer;
        updateMetrics();
        statusMsg.textContent = "✔ Answer generated! Click 'Generate Word & PDF' below.";
    } catch (err) {
        statusMsg.textContent = `❌ ${err.message}`;
    } finally {
        btn.disabled = false;
    }
});

// Export Word & PDF
document.getElementById("exportBtn").addEventListener("click", async () => {
    const current = weeksData[activeIndex];
    const courseCode = courseCodeEl.value.trim();

    if (!current.question.trim() || !current.answer.trim()) {
        statusMsg.textContent = "Please make sure both Question and Answer are filled in.";
        return;
    }

    const btn = document.getElementById("exportBtn");
    btn.disabled = true;
    statusMsg.textContent = "⏳ Building Word (.docx) and converting to PDF...";

    try {
        const res = await fetch("/api/export", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                week: current.week,
                course_code: courseCode,
                question: current.question,
                answer: current.answer
            })
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || "Export failed");

        current.exported = true;
        current.files = data;
        renderSidebar();
        renderDownloadLinks(data);
        statusMsg.textContent = `✔ Saved in /generated_files/${data.pdf_filename}!`;
    } catch (err) {
        statusMsg.textContent = `❌ ${err.message}`;
    } finally {
        btn.disabled = false;
    }
});

// Initialize UI
loadActiveWeek();
renderSidebar();