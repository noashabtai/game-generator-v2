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

// משתנה גלובלי לשמירת שם הקובץ האחרון
let lastFilename = null;

// ==================== עיבוד הנתונים ====================

async function generateMissingWords(categories, existingWords, totalWords = 800) {
  try {
    const missingCount = totalWords - existingWords.length;
    
    if (missingCount <= 0) {
      return existingWords.slice(0, totalWords);
    }

    const categoriesStr = categories.join(', ');
    const wordsStr = existingWords.slice(0, 50).join(', ');

    const prompt = `אתה עוזר ליצור מילים למשחק קופסה גדודי.

הקטגוריות הן: ${categoriesStr}

המילים שכבר יש: ${wordsStr}... (ועוד)

צור לי ${missingCount} מילים חדשות שלא חוזרות על עצמן, קשורות לקטגוריות האלה.
המילים צריכות להיות:
- בעברית
- קצרות (מילה או שתיים)
- רלוונטיות לגדוד/חיים/משפחה
- מגוונות

החזר רק רשימה של מילים, מופרדות בשורה חדשה. ללא מספורים, ללא הסברים.`;

    const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' }, { apiVersion: 'v1beta' });
    const result = await model.generateContent(prompt);
    const generatedText = result.response.text();

    const newWords = generatedText.split('\n')
      .map(w => w.trim())
      .filter(w => w.length > 0 && !existingWords.includes(w))
      .slice(0, missingCount);

    return [...existingWords, ...newWords].slice(0, totalWords);
  } catch (error) {
    console.error('Error generating words:', error.message);
    throw error;
  }
}

async function extractQuestionsFromStories(stories, numQuestions = 48) {
  try {
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

    const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' }, { apiVersion: 'v1beta' });
    const result = await model.generateContent(prompt);
    const generatedText = result.response.text();

    const questions = generatedText.split('\n')
      .map(q => q.trim())
      .filter(q => q.length > 0)
      .slice(0, numQuestions);

    return questions;
  } catch (error) {
    console.error('Error extracting questions:', error.message);
    throw error;
  }
}

// ==================== קריאת קבצים ====================

function readExcelFile(buffer, sheetName = 0) {
  try {
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[sheetName]];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    
    return data.flat().filter(cell => cell && String(cell).trim().length > 0);
  } catch (error) {
    console.error('Error reading Excel:', error.message);
    throw error;
  }
}

function readStoriesFile(buffer) {
  try {
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });

    if (rows.length < 2) return [];

    const headers = rows[0].map(h => String(h || '').trim());
    const companyIdx = headers.findIndex(h => h === 'פלוגה' || h === 'Company');
    const storyIdx = headers.findIndex(h => h === 'סיפור' || h === 'Story' || h === 'סיפור/זיכרון');

    return rows.slice(1).map(row => ({
      company: companyIdx >= 0 ? String(row[companyIdx] || 'לא מוגדר') : 'לא מוגדר',
      story: storyIdx >= 0 ? String(row[storyIdx] || '') : ''
    })).filter(s => s.story.trim().length > 0);
  } catch (error) {
    console.error('Error reading stories:', error.message);
    throw error;
  }
}

// ==================== יצוא Excel ====================

function createExcelFile(words, questions, metadata) {
  try {
    const wb = xlsx.utils.book_new();

    // גיליון 1: קלפים בסיסיים
    const basicCards = words.map((word, idx) => ({
      'קלף #': idx + 1,
      'מילה': word,
      'קטגוריה': '(עיצוב בעיצובנית)'
    }));
    const ws1 = xlsx.utils.json_to_sheet(basicCards);
    xlsx.utils.book_append_sheet(wb, ws1, 'קלפים בסיסיים');

    // גיליון 2: קלפי סיפור
    const storyCards = questions.map((q, idx) => ({
      'קלף #': idx + 1,
      'שאלה': q,
      'פלוגה': metadata.companies[idx % metadata.companies.length] || 'מחולק'
    }));
    const ws2 = xlsx.utils.json_to_sheet(storyCards);
    xlsx.utils.book_append_sheet(wb, ws2, 'קלפי סיפור');

    // גיליון 3: מטא-נתונים
    const metaData = [
      ['שם המשחק:', metadata.gameName],
      ['סלוגן:', metadata.slogan],
      ['שם קלפי סיפור:', metadata.storyCardName],
      ['שם קלפי אזרחות:', metadata.civilianCardName],
      ['כמות מילים סה״כ:', words.length],
      ['כמות שאלות:', questions.length],
      ['כמות פלוגות:', metadata.companies.length],
      ['תאריך יצוא:', new Date().toLocaleDateString('he-IL')],
      [''],
      ['קטגוריות (למידע בלבד):', metadata.categories.join(', ')]
    ];
    const ws3 = xlsx.utils.json_to_sheet(metaData, { header: 1 });
    xlsx.utils.book_append_sheet(wb, ws3, 'מטא-נתונים');

    return xlsx.write(wb, { bookType: 'xlsx', type: 'buffer' });
  } catch (error) {
    console.error('Error creating Excel:', error.message);
    throw error;
  }
}

// ==================== API Routes ====================

// עיבוד קבצים
app.post('/api/process', upload.fields([
  { name: 'wordsFile', maxCount: 1 },
  { name: 'storiesFile', maxCount: 1 }
]), async (req, res) => {
  try {
    const { gameName, slogan, storyCardName, civilianCardName, categories, companies } = req.body;

    const categoriesList = JSON.parse(categories || '[]');
    const companiesList = JSON.parse(companies || '[]');

    let words = [];
    let stories = [];

    if (req.files.wordsFile) {
      words = readExcelFile(req.files.wordsFile[0].buffer);
    }

    if (req.files.storiesFile) {
      stories = readStoriesFile(req.files.storiesFile[0].buffer);
    }

    console.log('🤖 Gemini עובד על המילים...');
    const allWords = await generateMissingWords(categoriesList, words, 800);

    console.log('🤖 Gemini עובד על השאלות...');
    const allQuestions = await extractQuestionsFromStories(stories, 48);

    console.log('📊 יוצר Excel...');
    const excelBuffer = createExcelFile(allWords, allQuestions, {
      gameName,
      slogan,
      storyCardName,
      civilianCardName,
      categories: categoriesList,
      companies: companiesList
    });

    const filename = `game-${Date.now()}.xlsx`;
    fs.writeFileSync(`/tmp/${filename}`, excelBuffer);
    lastFilename = filename;

    console.log(`✅ קובץ יצוא: ${filename}`);

    res.json({
      status: 'done',
      wordCount: allWords.length,
      storyCount: stories.length
    });

  } catch (error) {
    console.error('API error:', error);
    res.status(500).json({ error: error.message });
  }
});

// הורדת קובץ - גרסה מתוקנת
app.get('/api/download/latest', (req, res) => {
  try {
    if (!lastFilename) {
      return res.status(404).json({ error: 'קובץ לא נמצא. אנא העלו קבצים תחילה.' });
    }

    const filepath = path.join('/tmp', lastFilename);

    if (!fs.existsSync(filepath)) {
      return res.status(404).json({ error: 'קובץ לא זמין. אנא נסו שוב.' });
    }

    res.download(filepath, `game-${Date.now()}.xlsx`, () => {
      try {
        fs.unlinkSync(filepath);
        lastFilename = null;
      } catch (e) {
        console.error('Error deleting file:', e);
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ==================== Start Server ====================

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🎲 משחק קופסה גדודי - מחולל תוכן`);
  console.log(`🌐 שרת פועל על: http://localhost:${PORT}`);
  console.log(`🤖 Gemini API: ${process.env.GEMINI_API_KEY ? '✅ מוגדר' : '❌ לא מוגדר'}`);
});