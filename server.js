import express from 'express';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

const TEMPLATES_DIR = path.join(__dirname, 'Шаблоны');
const PUBLIC_TEMPLATES_DIR = path.join(__dirname, 'public', 'Шаблоны');

if(!fs.existsSync(TEMPLATES_DIR)) fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
if(!fs.existsSync(PUBLIC_TEMPLATES_DIR)) fs.mkdirSync(PUBLIC_TEMPLATES_DIR, { recursive: true });

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

// Получить информацию о файле шаблонов
app.get('/api/templates/info', (req, res) => {
  try {
    const candidates = [
      path.join(TEMPLATES_DIR, 'Шаблоны документов.xlsx'),
      path.join(PUBLIC_TEMPLATES_DIR, 'Шаблоны документов.xlsx'),
      path.join(TEMPLATES_DIR, 'Шаблоны документов.xlsm'),
      path.join(PUBLIC_TEMPLATES_DIR, 'Шаблоны документов.xlsm')
    ];
    let filePath = candidates.find(p => fs.existsSync(p));
    if(!filePath){
      // Проверяем любой .xlsx / .xlsm в папке Шаблоны
      const files = fs.readdirSync(TEMPLATES_DIR).filter(f => f.endsWith('.xlsx') || f.endsWith('.xlsm'));
      if(files.length > 0) filePath = path.join(TEMPLATES_DIR, files[0]);
    }

    if(!filePath){
      return res.json({ exists: false, message: 'Файл шаблона не найден в папке Шаблоны' });
    }

    const stat = fs.statSync(filePath);
    const wb = XLSX.readFile(filePath, { bookSheets: true });
    return res.json({
      exists: true,
      fileName: path.basename(filePath),
      size: stat.size,
      updatedAt: stat.mtime.toISOString(),
      sheets: wb.SheetNames || [],
      hasKs2: (wb.SheetNames || []).some(s => s.toLowerCase().includes('кс-2') || s.toLowerCase().includes('кс2')),
      hasKs3: (wb.SheetNames || []).some(s => s.toLowerCase().includes('кс-3') || s.toLowerCase().includes('кс3')),
      hasKs6: (wb.SheetNames || []).some(s => s.toLowerCase().includes('кс-6') || s.toLowerCase().includes('кс6'))
    });
  } catch (err) {
    return res.status(500).json({ exists: false, error: err.message });
  }
});

// Загрузить / обновить файл шаблонов
app.post('/api/templates/upload', express.raw({ type: '*/*', limit: '50mb' }), (req, res) => {
  try {
    const origName = req.query.name || 'Шаблоны документов.xlsx';
    const cleanName = path.basename(origName);
    const targetPath1 = path.join(TEMPLATES_DIR, cleanName);
    const targetPath2 = path.join(PUBLIC_TEMPLATES_DIR, cleanName);

    fs.writeFileSync(targetPath1, req.body);
    fs.writeFileSync(targetPath2, req.body);

    // Если имя отличается, также дублируем в Шаблоны документов.xlsx
    if(cleanName !== 'Шаблоны документов.xlsx'){
      fs.writeFileSync(path.join(TEMPLATES_DIR, 'Шаблоны документов.xlsx'), req.body);
      fs.writeFileSync(path.join(PUBLIC_TEMPLATES_DIR, 'Шаблоны документов.xlsx'), req.body);
    }

    const wb = XLSX.readFile(targetPath1, { bookSheets: true });
    return res.json({
      success: true,
      fileName: cleanName,
      sheets: wb.SheetNames || [],
      message: 'Файл шаблона сохранен в папку Шаблоны'
    });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

// Fallback to index.html
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server is running at http://0.0.0.0:${PORT}`);
});
