/* --- TAB SWITCHING --- */
function switchTab(tabName) {
    document.getElementById('ocrSection').style.display = 'none';
    document.getElementById('splitSection').style.display = 'none';
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));

    if (tabName === 'ocr') {
        document.getElementById('ocrSection').style.display = 'block';
        document.querySelectorAll('.tab-btn')[0].classList.add('active');
        document.getElementById('ocrSection').classList.add('active'); // for state check
    } else {
        document.getElementById('splitSection').style.display = 'block';
        document.querySelectorAll('.tab-btn')[1].classList.add('active');
        document.getElementById('splitSection').classList.add('active'); // for state check
    }
}

/* ============================
   SECTION 1: OCR LOGIC
   ============================ */
const imageInput = document.getElementById('imageInput');
const imagePreview = document.getElementById('imagePreview');
const convertBtn = document.getElementById('convertBtn');
const ocrResultContainer = document.getElementById('ocrResultContainer');
const ocrResultText = document.getElementById('ocrResultText');
const copyOcrBtn = document.getElementById('copyOcrBtn');
const statusText = document.getElementById('statusText');
const loadingBar = document.getElementById('loadingBar');
const progressBar = document.getElementById('progressBar');

// Global variable to store the image file
let currentOcrFile = null;

// 1. Handle File Upload (Click)
if(imageInput) {
    imageInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            loadFile(file);
        }
    });
}

// 2. Handle Paste (Ctrl+V)
document.addEventListener('paste', function(e) {
    // Only work if we are on the OCR tab
    const ocrSection = document.getElementById('ocrSection');
    if (ocrSection.style.display === 'none') return;

    const items = e.clipboardData.items;
    for (let i = 0; i < items.length; i++) {
        if (items[i].type.indexOf('image') !== -1) {
            const blob = items[i].getAsFile();
            loadFile(blob);
            statusText.innerText = "📸 Image pasted from clipboard!";
            e.preventDefault(); // Stop browser from pasting into other fields
            break;
        }
    }
});

// Helper: Load File into Preview
function loadFile(file) {
    currentOcrFile = file;
    imagePreview.src = URL.createObjectURL(file);
    imagePreview.style.display = 'block';
    convertBtn.disabled = false;
    
    // Reset Results
    ocrResultContainer.style.display = 'none';
    statusText.innerText = "Ready to extract.";
    progressBar.style.width = '0%';
    loadingBar.style.display = 'none';
}

// 3. Run OCR (Extract Text)
if(convertBtn) {
    convertBtn.addEventListener('click', function() {
        if (!currentOcrFile) {
            statusText.innerText = "⚠️ Please select or paste an image first.";
            return;
        }

        convertBtn.disabled = true;
        convertBtn.innerText = "Processing...";
        loadingBar.style.display = 'block';
        statusText.innerText = "Initializing AI...";
        
        Tesseract.recognize(
            currentOcrFile, 
            'eng',
            { logger: m => {
                if (m.status === 'recognizing text') {
                    const progress = Math.round(m.progress * 100);
                    progressBar.style.width = `${progress}%`;
                    statusText.innerText = `Scanning: ${progress}%`;
                }
            }}
        ).then(({ data: { text } }) => {
            statusText.innerText = "✅ Done!";
            convertBtn.innerText = "Extract Text";
            convertBtn.disabled = false;
            ocrResultContainer.style.display = 'block';
            ocrResultText.value = text;
        }).catch(err => {
            statusText.innerText = "Error: " + err.message;
            convertBtn.disabled = false;
        });
    });
}

// Copy Button for OCR
if(copyOcrBtn) {
    copyOcrBtn.addEventListener('click', function() {
        copyToClipboard('ocrResultText', this);
    });
}


/* ============================
   SECTION 2: SPLIT COLUMNS LOGIC
   ============================ */
const splitBtn = document.getElementById('splitBtn');
const splitInput = document.getElementById('splitInput');
const col1Text = document.getElementById('col1Text');
const col2Text = document.getElementById('col2Text');
const splitResultContainer = document.getElementById('splitResultContainer');

if(splitBtn) {
    splitBtn.addEventListener('click', () => {
        const rawText = splitInput.value;
        if (!rawText.trim()) return;

        let column1 = [];
        let column2 = [];

        // Split by lines
        const lines = rawText.trim().split('\n');

        lines.forEach(line => {
            // Remove empty space around line
            const cleanLine = line.trim();
            if (cleanLine.length === 0) return; // Skip empty lines

            // Split by whitespace (tabs or spaces)
            const parts = cleanLine.split(/\s+/);
            
            if (parts.length >= 1) column1.push(parts[0]); 
            
            // Logic for 2nd column: Take the LAST item to avoid issues with middle spaces
            if (parts.length >= 2) column2.push(parts[parts.length - 1]); 
            else column2.push(""); // Push empty string if no 2nd column exists
        });

        // Join with new lines
        col1Text.value = column1.join('\n');
        col2Text.value = column2.join('\n');

        splitResultContainer.style.display = 'block';
    });
}

/* --- COPY HELPER FUNCTION --- */
function copyToClipboard(elementId, btnElement) {
    const textArea = document.getElementById(elementId);
    if (!textArea) return;

    textArea.select();
    navigator.clipboard.writeText(textArea.value).then(() => {
        const originalText = btnElement.innerText;
        btnElement.innerText = "Copied! ✅";
        setTimeout(() => { btnElement.innerText = originalText; }, 1500);
    });
}
