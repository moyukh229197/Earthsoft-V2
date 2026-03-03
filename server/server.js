const http = require('http');

const PORT = 3000;
const GROQ_API_KEY = process.env.GROQ_API_KEY; // Loaded from .env file

// The system prompt that forces Groq to act as a strict data parser
const SYSTEM_PROMPT = `
You are a highly precise civil engineering data extraction assistant.
The human will provide messy, unstructured text (surveyor notes, emails, messy tables).
Your ONLY job is to extract Bridge/Structure information and return a strictly formatted JSON array.
If the text contains any structures explicitly labeled as 'ROB' or 'Road Over Bridge', IGNORE THEM completely.

For every valid bridge found, return an object in this EXACT JSON structure:
{
  "bridgeNo": "string (e.g., BR-1, 10/2, etc.)",
  "bridgeCategory": "string (Major, Minor, Viaduct, Important)",
  "bridgeType": "string (Box, PSC Slab, Composite Girder, OWG, etc.)",
  "bridgeSize": "string (e.g., 1x7.5x5.5, 10x9.15)",
  "bridgeSpans": "string (number of spans, e.g., '1', '10')",
  "startChainage": "number (in meters, e.g. 12500)",
  "endChainage": "number (in meters)",
  "length": "number (in meters)"
}

Rules:
1. Try your best to extract the Start and End chainage. If you only see a single Chainage (like 'Ch: 12+500'), treat that as the center. 
2. If only center chainage and size are given, calculate start/end mathematically (Length = spans * span_size).
3. Chainages of '12+500' mean 12500 meters.
4. ONLY return a beautifully formatted JSON Array of these objects. No markdown formatting, no conversational text, no backticks. JUST the literal JSON array starting with '[' and ending with ']'.
`;

const server = http.createServer(async (req, res) => {
    // CORS Headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    // Handle preflight
    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    if (req.method === 'POST' && req.url === '/api/extract-bridges') {
        let body = '';
        req.on('data', chunk => {
            body += chunk.toString();
        });

        req.on('end', async () => {
            try {
                const payload = JSON.parse(body);
                const { text } = payload;

                if (!text) {
                    res.writeHead(400, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: 'Text input is required' }));
                    return;
                }

                console.log(`Received request to parse ${text.length} characters of unstructured text.`);

                // Direct REST API call to Groq via native Fetch
                const groqResponse = await fetch('https://api.groq.com/openai/v1/chat/completions', {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${GROQ_API_KEY}`,
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        model: 'llama3-70b-8192',
                        messages: [
                            { role: 'system', content: SYSTEM_PROMPT },
                            { role: 'user', content: text }
                        ],
                        temperature: 0.1, // Low temperature for high precision data extraction
                        max_tokens: 4000,
                    })
                });

                if (!groqResponse.ok) {
                    const errText = await groqResponse.text();
                    console.error("Groq API Error:", errText);
                    throw new Error(`Groq API returned status ${groqResponse.status}`);
                }

                const groqData = await groqResponse.json();
                let assistantMessage = groqData.choices[0].message.content.trim();

                // Safety cleanup if the LLM hallucinated markdown backticks despite instructions
                if (assistantMessage.startsWith('```json')) {
                    assistantMessage = assistantMessage.replace(/^```json\n/, '').replace(/\n```$/, '');
                } else if (assistantMessage.startsWith('```')) {
                    assistantMessage = assistantMessage.replace(/^```\n?/, '').replace(/\n?```$/, '');
                }

                const extractedBridges = JSON.parse(assistantMessage);

                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ success: true, rows: extractedBridges }));

            } catch (err) {
                console.error('Server error:', err);
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: err.message || 'Failed to process AI extraction' }));
            }
        });
    } else {
        res.writeHead(404);
        res.end();
    }
});

server.listen(PORT, () => {
    console.log(`Earthsoft AI Backend running silently at http://localhost:${PORT}`);
    console.log(`Listening for incoming unstructured text to parse using Llama 3 on Groq...`);
});
