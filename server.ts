import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { GoogleGenAI } from '@google/genai';
import { createServer as createViteServer } from 'vite';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: '5mb' }));

  // Helper to get GenAI instance
  const getAI = () => {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return null;
    }
    return new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          'User-Agent': 'aistudio-build',
        },
      },
    });
  };

  // Health check
  app.get('/api/health', (req, res) => {
    res.json({ status: 'ok', hasGeminiKey: !!process.env.GEMINI_API_KEY });
  });

  // AI Project & Web Consultant for Web Pro BD
  app.post('/api/ai/consult', async (req, res) => {
    try {
      const { message, history, language = 'bn', projectType, budget } = req.body;
      const ai = getAI();
      if (!ai) {
        return res.json({
          reply: language === 'bn' 
            ? 'ধন্যবাদ Web Pro BD তে যোগাযোগ করার জন্য! আমাদের রেডিমেড ওয়েবসাইট মাত্র ৩,৯৯৯৳ থেকে শুরু এবং কাস্টম ওয়েব অ্যাপ ৯,৯৯৯৳ থেকে (লাইফটাইম ফ্রি ডোমেন ও হোস্টিং সহ)। সরাসরি প্রতিষ্ঠাতা আবু তালহা খানের WhatsApp-এ (01516513987) কথা বলুন।'
            : 'Thank you for reaching out to Web Pro BD! Our ready-made websites start at 3,999 BDT with lifetime free domain & hosting. Connect with founder Abu Talha Khan directly on WhatsApp at 01516513987.',
        });
      }

      const systemInstruction = `You are "Web Pro BD Assistant" (ওয়েব প্রো বিডি কনসালট্যান্ট), an expert digital solutions advisor for Web Pro BD.
Web Pro BD sells high-quality ready-made websites (E-commerce, News Portals, POS, School/Hospital Management, Portfolio, Restaurant) and builds custom full-stack Web Applications & Backend solutions (Node.js, React, Express, MongoDB, MySQL, Laravel, Python).
Pricing guidelines:
- Ready Landing Pages: 2,999৳ - 4,999৳
- Full E-commerce (bKash/Nagad gateway + Admin): 5,999৳ - 12,999৳
- School / Hospital / POS Management Systems: 8,999৳ - 19,999৳
- Custom Full-Stack Web Apps & SaaS: 14,999৳+
- Backend API & Custom Code (\`backendcode.js\` services): 4,999৳+

Respond kindly in natural, business-friendly ${language === 'bn' ? 'Bengali (বাংলা)' : 'English'}.
Help the customer understand features, suggest the best website/web app package for their business, estimate timeframe (1 to 7 days), and guide them on placing orders with cash memos/invoices. Keep responses concise, clear, and professional.`;

      const contents = [];
      if (Array.isArray(history)) {
        for (const item of history.slice(-6)) {
          contents.push({
            role: item.role === 'user' ? 'user' : 'model',
            parts: [{ text: item.text }],
          });
        }
      }
      contents.push({
        role: 'user',
        parts: [{ text: `Customer asked: ${message}${projectType ? ` | Interested in: ${projectType}` : ''}${budget ? ` | Budget: ${budget}` : ''}` }],
      });

      const response = await ai.models.generateContent({
        model: 'gemini-3.7-flash',
        contents,
        config: {
          systemInstruction,
          temperature: 0.7,
        },
      });

      res.json({ reply: response.text || 'ধন্যবাদ! আমাদের প্রতিনিধি শীঘ্রই আপনার সাথে যোগাযোগ করবেন।' });
    } catch (err: any) {
      console.error('Gemini consult error:', err);
      res.status(500).json({ error: err.message || 'AI consultation failed' });
    }
  });

  // AI Instant Quote / Scope Generator
  app.post('/api/ai/quote', async (req, res) => {
    try {
      const { requirements, businessType, language = 'bn' } = req.body;
      const ai = getAI();
      if (!ai) {
        return res.json({
          plan: 'স্ট্যান্ডার্ড ই-কমার্স / ওয়েব অ্যাপ প্যাকেজ',
          estimatedDays: '২-৪ দিন',
          estimatedPriceBDT: '৭,৫০০ ৳',
          features: ['রেস্পন্সিভ ডিজাইন', 'অ্যাডমিন প্যানেল', 'বিকাশ/নগদ পেমেন্ট গেটওয়ে', '১ বছর ফ্রি সাপোর্ট'],
        });
      }

      const prompt = `Based on the following client project requirements for a website or web app, generate a structured project proposal in ${language === 'bn' ? 'Bengali (বাংলা)' : 'English'}:
Business Type: ${businessType || 'General Business / Startup'}
Requirements: ${requirements}

Provide in clear formatted markdown:
1. Recommended Solution (Ready-made site vs Custom Web App)
2. Core Features List
3. Tech Stack Recommendation (Frontend, Backend, Database)
4. Estimated Timeline (in days)
5. Estimated Cost in BDT (Bangladeshi Taka)`;

      const response = await ai.models.generateContent({
        model: 'gemini-3.7-flash',
        contents: prompt,
      });

      res.json({ quote: response.text });
    } catch (err: any) {
      console.error('Gemini quote error:', err);
      res.status(500).json({ error: err.message || 'Quote generation failed' });
    }
  });

  // AI Explain Code endpoint
  app.post('/api/ai/explain', async (req, res) => {
    try {
      const { code, language = 'bn', codeLanguage = 'javascript' } = req.body;
      if (!code) {
        return res.status(400).json({ error: 'Code is required' });
      }

      const ai = getAI();
      if (!ai) {
        return res.status(503).json({
          error: 'Gemini API key not configured.',
        });
      }

      const prompt = `Please explain the following ${codeLanguage} code step-by-step for a beginner learner in ${language === 'bn' ? 'Bengali (বাংলা)' : 'English'}:
- Explain what the code does as a whole
- Explain line by line in bullet points
- Mention any good practices or potential edge cases

Code:
\`\`\`${codeLanguage}
${code}
\`\`\``;

      const response = await ai.models.generateContent({
        model: 'gemini-3.7-flash',
        contents: prompt,
      });

      res.json({ explanation: response.text });
    } catch (err: any) {
      console.error('Gemini explain error:', err);
      res.status(500).json({ error: err.message || 'AI explanation failed' });
    }
  });

  // AI Debug / Fix endpoint
  app.post('/api/ai/debug', async (req, res) => {
    try {
      const { code, errorOutput, language = 'bn', codeLanguage = 'javascript' } = req.body;
      if (!code) {
        return res.status(400).json({ error: 'Code is required' });
      }

      const ai = getAI();
      if (!ai) {
        return res.status(503).json({
          error: 'Gemini API key not configured.',
        });
      }

      const prompt = `You are a code debugging expert. The user is writing ${codeLanguage} code and encountered an error or needs help fixing their code.
Language for explanation: ${language === 'bn' ? 'Bengali (বাংলা)' : 'English'}.

User Code:
\`\`\`${codeLanguage}
${code}
\`\`\`

${errorOutput ? `Error Message / Output:\n${errorOutput}\n` : ''}

Please provide:
1. What the bug/problem is
2. Fixed and corrected working code block (enclosed in \`\`\`${codeLanguage})
3. Quick tip to avoid this bug in the future`;

      const response = await ai.models.generateContent({
        model: 'gemini-3.7-flash',
        contents: prompt,
      });

      res.json({ debugInfo: response.text });
    } catch (err: any) {
      console.error('Gemini debug error:', err);
      res.status(500).json({ error: err.message || 'AI debugging failed' });
    }
  });

  // AI Generate Code endpoint
  app.post('/api/ai/generate', async (req, res) => {
    try {
      const { prompt, codeLanguage = 'javascript', language = 'bn' } = req.body;
      if (!prompt) {
        return res.status(400).json({ error: 'Prompt is required' });
      }

      const ai = getAI();
      if (!ai) {
        return res.status(503).json({
          error: 'Gemini API key not configured.',
        });
      }

      const systemPrompt = `Write clean, well-commented ${codeLanguage} code for the user prompt.
Provide a short explanation in ${language === 'bn' ? 'Bengali (বাংলা)' : 'English'}, followed by the complete code enclosed in triple backticks.`;

      const response = await ai.models.generateContent({
        model: 'gemini-3.7-flash',
        contents: prompt,
        config: {
          systemInstruction: systemPrompt,
        },
      });

      res.json({ result: response.text });
    } catch (err: any) {
      console.error('Gemini generate error:', err);
      res.status(500).json({ error: err.message || 'AI generation failed' });
    }
  });

  // Vite middleware for development vs static build in production
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
