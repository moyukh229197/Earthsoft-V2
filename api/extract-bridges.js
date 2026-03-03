export default async function handler(req, res) {
    // CORS setup if accessed externally
    res.setHeader('Access-Control-Allow-Credentials', true)
    res.setHeader('Access-Control-Allow-Origin', '*')
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT')
    res.setHeader(
        'Access-Control-Allow-Headers',
        'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
    )

    if (req.method === 'OPTIONS') {
        res.status(200).end()
        return
    }

    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed' })
    }

    const GROQ_API_KEY = process.env.GROQ_API_KEY;

    if (!GROQ_API_KEY) {
        return res.status(500).json({ error: 'GROQ_API_KEY is not configured in Vercel Environment Variables.' })
    }

    const { text } = req.body;
    if (!text) {
        return res.status(400).json({ error: 'Text input is required' });
    }

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
1. Try your best to extract Start and End chainage. If you only see a single Chainage (like 'Ch: 12+500'), treat that as the center. 
2. If only center chainage and size are given, calculate start/end mathematically (Length = spans * span_size).
3. Chainages of '12+500' mean 12500 meters.
4. ONLY return a beautifully formatted JSON Array of these objects. No markdown formatting, no conversational text, no backticks. JUST the literal JSON array starting with '[' and ending with ']'.
`;

    try {
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
                temperature: 0.1,
                max_tokens: 4000,
            })
        });

        if (!groqResponse.ok) {
            const errData = await groqResponse.json().catch(() => ({}));
            throw new Error(errData.error?.message || `Groq API Error: ${groqResponse.status}`);
        }

        const groqData = await groqResponse.json();
        let aiOutput = groqData.choices && groqData.choices[0]?.message?.content?.trim();

        if (!aiOutput) throw new Error("No readable response from AI.");

        if (aiOutput.startsWith('```json')) {
            aiOutput = aiOutput.replace(/^```json\n/, '').replace(/\n```$/, '');
        } else if (aiOutput.startsWith('```')) {
            aiOutput = aiOutput.replace(/^```\n?/, '').replace(/\n?```$/, '');
        }

        const extractedBridges = JSON.parse(aiOutput);

        return res.status(200).json({ success: true, rows: extractedBridges });

    } catch (err) {
        console.error('API Route Error:', err);
        return res.status(500).json({ error: err.message || 'Failed to process AI extraction' });
    }
}
