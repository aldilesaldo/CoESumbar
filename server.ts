import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Use JSON middleware
  app.use(express.json());

  // API routes
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  const SHEET_URL = "https://script.google.com/macros/s/AKfycbwtvb9Ffd_W0ErZbTvdtVV-z-1jlwLCPK38O0FJ79z9ZQczaUH6W1yx-ofSunGcJtwGwA/exec";

  // Proxy for Google Sheets GET
  app.get("/api/events", async (req, res) => {
    try {
      const response = await fetch(`${SHEET_URL}?action=getAll`, {
        headers: {
          "Accept": "application/json",
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
      });
      
      const text = await response.text();
      
      if (!response.ok) {
        console.error(`Google Sheets responded with status ${response.status}. Body:`, text.substring(0, 200));
        return res.status(response.status).json({ error: "Google Sheets returned an error" });
      }

      try {
        const data = JSON.parse(text);
        res.json(data);
      } catch (parseError) {
        console.error("Failed to parse JSON from Google Sheets. First 500 chars:", text.substring(0, 500));
        res.status(500).json({ error: "Invalid response format from Google Sheets" });
      }
    } catch (error) {
      console.error("Error fetching from sheets:", error);
      res.status(500).json({ error: "Failed to fetch events from Google Sheets" });
    }
  });

  // Proxy for Google Sheets POST
  app.post("/api/events", async (req, res) => {
    try {
      const action = req.body.action || "save";
      const targetUrl = `${SHEET_URL}${SHEET_URL.includes("?") ? "&" : "?"}action=${action}`;
      
      console.log(`Saving to Google Sheets (${action}):`, targetUrl);
      const response = await fetch(targetUrl, {
        method: "POST",
        headers: { 
          "Content-Type": "application/json",
          "Accept": "application/json",
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        },
        body: JSON.stringify(req.body),
        redirect: "follow"
      });
      
      const text = await response.text();
      
      if (!response.ok) {
        console.error(`Google Sheets POST failed. Status: ${response.status}, Body: ${text.substring(0, 500)}`);
        return res.status(response.status).json({ 
          error: "Failed to save to Google Sheets", 
          status: response.status,
          detail: text.substring(0, 200)
        });
      }

      console.log("Google Sheets POST success:", text.substring(0, 100));
      res.json({ status: "ok", detail: text.substring(0, 500) });
    } catch (error: any) {
      console.error("Error saving to sheets:", error);
      res.status(500).json({ 
        error: "Failed to save event to Google Sheets", 
        detail: error.message 
      });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
