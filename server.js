require('dotenv').config();
const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { GoogleGenerativeAI } = require('@google/generative-ai');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// מצב עיבוד גלובלי
let processingState = {
  status: 'idle', // idle / processing / done / error
  error: null,
  wordCount: 0,
  storyCount: 0,
  filename: null
};

// ==================== עיבוד הנתונים ====================

async function completeCategoryWords(categoryName, existingWords, targetCount = 100) {
  const missingCount = targetCount - existingWords.length;

  if (missingCount <= 0) {
    return existingWords.slice(0, targetCount);
  }

  const wordsStr = existingWords.slice(0, 30).join(', ');

  const prompt = `אתה עוזר ליצור מילים למשחק קופסה גדודי.

הקטגוריה: ${categoryName}

המילים שכבר יש: ${wordsStr}${existingWords.length > 30 ? '... (ועוד)' : ''}

צור לי ${missingCount} מילים חדשות שלא חוזרות על עצמן, קשורות לקטגוריה "${categoryName}".
המילים צריכות להיות:
- בעברית
- קצרות (מילה או שתיים)
- רלוונטיות לקטגוריה
- מגוונות

החזר רק רשימה של מילים, מופרדות בשורה חדשה. ללא מספורים, ללא הסברים.`;

  const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });
  const result = await model.generateContent(prompt);
  const generatedText = result.response.text();

  const newWords = generatedText.split('\n')
    .map(w => w.trim())
    .filter(w => w.length > 0 && !existingWords.includes(w))
    .slice(0, missingCount);

  return [...existingWords, ...newWords].slice(0, targetCount);
}

async function extractQuestionsFromStories(stories, numQuestions = 48) {
  const storiesText = stories
    .map(s => `[${s.company}] ${s.story}`)
    .join('\n\n');

  const prompt = `אתה עוזר לחלץ שאלות משחק מסיפורים גדודיים.

הסיפורים:
${storiesText}

צור לי ${numQuestions} שאלות/משימות קצרות בנוסח "ספרו לנו..." או "תסביר לי..." שמתוך הסיפורים האלה.
השאלות צריכות להיות:
- קצרות (עד 10 מילים כל אחת)
- מעוררות סיפור וזיכרון
- קשורות לתוכן של הסיפורים
- מגוונות

החזר רק רשימה של שאלות, מופרדות בשורה חדשה. ללא מספורים.`;

  const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });
  const result = await model.generateContent(prompt);
  const generatedText = result.response.text();

  return generatedText.split('\n')
    .map(q => q.trim())
    .filter(q => q.length > 0)
    .slice(0, numQuestions);
}

// ==================== קריאת קבצים ====================

function readCellsFromSheet(sheet) {
  const ref = sheet['!ref'];
  if (!ref) return [];
  const range = xlsx.utils.decode_range(ref);
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[xlsx.utils.encode_cell({ r, c })];
      row.push(cell && cell.v !== undefined ? String(cell.v).trim() : '');
    }
    rows.push(row);
  }
  return rows;
}

// קורא קובץ מילים לפי קטגוריות — שורה ראשונה = שמות קטגוריות, שאר השורות = מילים
function readWordsFileByCategory(buffer) {
  const workbook = xlsx.read(buffer, { type: 'buffer', cellFormula: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = readCellsFromSheet(sheet);

  if (rows.length < 1) return {};

  const headers = rows[0]; // שורה ראשונה = שמות קטגוריות
  const result = {};

  headers.forEach((header, colIdx) => {
    if (!header.trim()) return;
    const words = rows.slice(1)
      .map(row => (row[colIdx] || '').trim())
      .filter(w => w.length > 0);
    result[header] = words;
  });

  return result;
}

function readStoriesFile(buffer) {
  const workbook = xlsx.read(buffer, { type: 'buffer', cellFormula: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = readCellsFromSheet(sheet);

  if (rows.length < 2) return [];

  const headers = rows[0];
  const companyIdx = headers.findIndex(h => h === 'פלוגה' || h === 'Company');
  const storyIdx = headers.findIndex(h => h === 'סיפור' || h === 'Story' || h === 'סיפור/זיכרון');

  return rows.slice(1).map(row => ({
    company: companyIdx >= 0 ? (row[companyIdx] || 'לא מוגדר') : 'לא מוגדר',
    story: storyIdx >= 0 ? (row[storyIdx] || '') : ''
  })).filter(s => s.story.trim().length > 0);
}

// ==================== יצוא Excel ====================

function createExcelFile(categoryWords, questions, metadata) {
  const wb = xlsx.utils.book_new();

  // גיליון קלפים בסיסיים — עמודה לכל קטגוריה
  const categoryNames = Object.keys(categoryWords);
  const maxWords = Math.max(...categoryNames.map(cat => categoryWords[cat].length), 0);

  const basicAOA = [categoryNames];
  for (let r = 0; r < maxWords; r++) {
    const row = categoryNames.map(cat => categoryWords[cat][r] || '');
    basicAOA.push(row);
  }
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(basicAOA), 'קלפים בסיסיים');

  // גיליון קלפי סיפור
  const storyAOA = [
    ['קלף #', 'שאלה', 'פלוגה'],
    ...questions.map((q, idx) => [idx + 1, q, metadata.companies[idx % metadata.companies.length] || 'מחולק'])
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(storyAOA), 'קלפי סיפור');

  const totalWords = categoryNames.reduce((sum, cat) => sum + categoryWords[cat].length, 0);

  const metaAOA = [
    ['שם המשחק:', metadata.gameName],
    ['סלוגן:', metadata.slogan],
    ['שם קלפי סיפור:', metadata.storyCardName],
    ['שם קלפי אזרחות:', metadata.civilianCardName],
    ['כמות מילים סה״כ:', totalWords],
    ['כמות שאלות:', questions.length],
    ['כמות פלוגות:', metadata.companies.length],
    ['תאריך יצוא:', new Date().toLocaleDateString('he-IL')]
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(metaAOA), 'מטא-נתונים');

  return xlsx.write(wb, { bookType: 'xlsx', type: 'buffer' });
}

// ==================== API Routes ====================

// התחלת עיבוד - מחזיר מיד ועובד ברקע
app.post('/api/process', upload.fields([
  { name: 'wordsFile', maxCount: 1 },
  { name: 'storiesFile', maxCount: 1 }
]), (req, res) => {
  try {
    if (processingState.status === 'processing') {
      return res.json({ status: 'processing' });
    }

    const { gameName, slogan, storyCardName, civilianCardName, categories, companies } = req.body;
    const categoriesList = JSON.parse(categories || '[]');
    const companiesList = JSON.parse(companies || '[]');

    let wordsByCategory = {};
    let stories = [];

    try {
      if (req.files.wordsFile) {
        wordsByCategory = readWordsFileByCategory(req.files.wordsFile[0].buffer);
      }
    } catch (e) {
      console.error('Error reading words file:', e.message);
    }

    try {
      if (req.files.storiesFile) {
        stories = readStoriesFile(req.files.storiesFile[0].buffer);
      }
    } catch (e) {
      console.error('Error reading stories file:', e.message);
    }

    // אם אין קטגוריות מהקובץ — משתמשים בקטגוריות מהטופס
    if (Object.keys(wordsByCategory).length === 0) {
      categoriesList.forEach(cat => {
        if (cat) wordsByCategory[cat] = [];
      });
    }

    processingState = { status: 'processing', error: null, wordCount: 0, storyCount: 0, filename: null };

    // מחזיר מיד — לא מחכה לסיום
    res.json({ status: 'processing' });

    // עיבוד ברקע
    (async () => {
      try {
        console.log('🤖 Gemini עובד על המילים לפי קטגוריות...');
        const categoryWords = {};
        for (const [catName, existingWords] of Object.entries(wordsByCategory)) {
          console.log(`  קטגוריה: ${catName} (${existingWords.length} מילים קיימות)`);
          categoryWords[catName] = await completeCategoryWords(catName, existingWords, 100);
        }

        console.log('🤖 Gemini עובד על השאלות...');
        const allQuestions = await extractQuestionsFromStories(stories, 48);

        console.log('📊 יוצר Excel...');
        const excelBuffer = createExcelFile(categoryWords, allQuestions, {
          gameName, slogan, storyCardName, civilianCardName,
          categories: categoriesList, companies: companiesList
        });

        const filename = `game-${Date.now()}.xlsx`;
        fs.writeFileSync(`/tmp/${filename}`, excelBuffer);

        const totalWords = Object.values(categoryWords).reduce((sum, arr) => sum + arr.length, 0);

        processingState = {
          status: 'done',
          error: null,
          wordCount: totalWords,
          storyCount: stories.length,
          filename
        };

        console.log(`✅ קובץ מוכן: ${filename}`);
      } catch (error) {
        console.error('Processing error:', error.message);
        processingState = { status: 'error', error: error.message, wordCount: 0, storyCount: 0, filename: null };
      }
    })();

  } catch (error) {
    console.error('API error:', error);
    res.status(500).json({ error: error.message });
  }
});

// בדיקת סטטוס עיבוד
app.get('/api/status', (req, res) => {
  res.json(processingState);
});

// הורדת קובץ
app.get('/api/download/latest', (_req, res) => {
  if (!processingState.filename) {
    return res.status(404).json({ error: 'קובץ לא נמצא.' });
  }

  const filepath = path.join('/tmp', processingState.filename);

  if (!fs.existsSync(filepath)) {
    return res.status(404).json({ error: 'קובץ לא זמין.' });
  }

  const downloadName = `game-${Date.now()}.xlsx`;
  res.download(filepath, downloadName, () => {
    try { fs.unlinkSync(filepath); } catch (e) {}
    processingState = { status: 'idle', error: null, wordCount: 0, storyCount: 0, filename: null };
  });
});

// ==================== Start Server ====================

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🎲 משחק קופסה גדודי - מחולל תוכן`);
  console.log(`🌐 שרת פועל על: http://localhost:${PORT}`);
  console.log(`🤖 Gemini API: ${process.env.GEMINI_API_KEY ? '✅ מוגדר' : '❌ לא מוגדר'}`);
});
