import express from 'express';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 3000;

// Serve static assets (CSS, fonts cached by browser)
app.use(express.static(join(__dirname, 'public')));

// Clean URL routing
app.get('/',      (_req, res) => res.sendFile(join(__dirname, 'public/index.html')));
app.get('/short', (_req, res) => res.sendFile(join(__dirname, 'public/short.html')));
app.get('/full',  (_req, res) => res.sendFile(join(__dirname, 'public/full.html')));

// Fallback
app.use((_req, res) => res.redirect('/'));

app.listen(PORT, () => {
  console.log(`PSU Darpan deck running at http://localhost:${PORT}`);
});
