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

    const { text, dataType = 'bridges' } = req.body;

    if (!text) {
        return res.status(400).json({ error: 'Text input is required' });
    }

    let SYSTEM_PROMPT = "";
    let IS_MAPPER = false;

    // Branch prompts based on data type
    if (dataType === 'bridges') {
        SYSTEM_PROMPT = `
You are a highly precise civil engineering data extraction assistant.
The human will provide messy CSV/text of Bridge/Structure data.
Your ONLY job is to extract Bridge information and return a strictly formatted JSON array.
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
4. Bridge identifiers/names should be taken primarily from columns like 'Bridge No.', 'Bridge Number', 'Structure', or 'Sl. No.'.
5. Tunnel identifiers/names MUST be explicitly parsed from columns titled 'Tunnel No.', 'Tunnel Name', or 'Tunnel'.
6. ONLY return a beautifully formatted JSON Array of these objects. No markdown formatting, no conversational text. JUST the literal JSON array starting with '[' and ending with ']'.
`;
    } else if (dataType === 'curves') {
        SYSTEM_PROMPT = `
You are a highly precise civil engineering data extraction assistant.
The human will provide a raw CSV of horizontal Curve data.
Your ONLY job is to return a beautiful JSON array of objects representing the curves.

For each curve, extract:
        {
            "chainage": "number (in meters, e.g. 15300)",
                "curve": "string (e.g. C-1, Curve 2. etc. Generate one if missing.)",
                    "radius": "number (in meters)",
                        "length": "number (in meters)"
        }

        Rules: ONLY return the literal JSON Array.No markdown formatting, no conversational text.
`;
    } else if (dataType === 'loops') {
        SYSTEM_PROMPT = `
You are a highly precise civil engineering data extraction assistant.
The human will provide a raw CSV of Railway Station Loops and Platforms data.
Your ONLY job is to return a perfect JSON array representing this data.

For each station row, extract:
        {
            "station": "string (Name of station)",
                "csb": "number (Center of station board chainage in meters, or null)",
                    "loopStartCh": "number (in meters)",
                        "loopEndCh": "number (in meters)",
                            "pfWidth": "number (platform width in meters)",
                                "pfStartCh": "number (in meters)",
                                    "pfEndCh": "number (in meters)",
                                        "tc": "number (Track center displacement in meters)",
                                            "remarks": "string (Any extra notes)"
        }
        Rules:
        1. "12+500" means 12500.
        2. ONLY return the literal JSON Array.No markdown formatting, no conversational text.
`;
    } else if (dataType === 'levels') {
        IS_MAPPER = true;
        SYSTEM_PROMPT = `
You are an intelligent data - mapping assistant.
The human will provide the first 50 rows of a massive Topographical Levels CSV. 
We CANNOT process all 10,000 rows right now. 

Your ONLY job is to figure out WHICH array index(0 - indexed) holds the critical data columns.
Look at the headers and the sample data rows provided.The rows are comma - separated.

Return a SINGLE JSON object exactly like this:
        {
            "chainageIndex": number(The 0 - based column index containing distance / chainage, e.g. 10 + 500, or - 1 if not found),
            "groundLevelIndex": number(The column containing Ground RL / OGL / Ground Elevation, or - 1 if not found),
            "proposedLevelIndex": number(The column containing Proposed RL / PRL / Formation Level, or - 1 if not found),
            "stationIndex": number(The column containing Station Name, or - 1 if none),
                "structureNoIndex": number(The column containing Structure No, or - 1 if none),
                    "dataStartRowIndex": number(The 0 - based row index where the actual numeric data begins, skipping headers)
        }

        Rules:
ONLY return the literal JSON Object.No markdown formatting, no arrays, no conversational text.JUST the JSON object starting with '{' and ending with '}'.
`;
    } else {
        return res.status(400).json({ error: 'Unsupported dataType' });
    }

    try {
        const groqResponse = await fetch('https://api.groq.com/openai/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${GROQ_API_KEY} `,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                model: 'llama-3.3-70b-versatile',
                messages: [
                    { role: 'system', content: SYSTEM_PROMPT },
                    { role: 'user', content: text }
                ],
                temperature: 0.1,
                max_tokens: IS_MAPPER ? 500 : 4000,
            })
        });

        if (!groqResponse.ok) {
            const errData = await groqResponse.json().catch(() => ({}));
            throw new Error(errData.error?.message || `Groq API Error: ${groqResponse.status} `);
        }

        const groqData = await groqResponse.json();
        let aiOutput = groqData.choices && groqData.choices[0]?.message?.content?.trim();

        if (!aiOutput) throw new Error("No readable response from AI.");

        if (aiOutput.startsWith('\`\`\`json')) {
            aiOutput = aiOutput.replace(/^\`\`\`json\n/, '').replace(/\n\`\`\`$/, '');
        } else if (aiOutput.startsWith('\`\`\`')) {
            aiOutput = aiOutput.replace(/^\`\`\`\n?/, '').replace(/\n?\`\`\`$/, '');
        }

        const parsedData = JSON.parse(aiOutput);

        return res.status(200).json({
            success: true,
            isMapper: IS_MAPPER,
            data: parsedData
        });

    } catch (err) {
        console.error('API Route Error:', err);
        return res.status(500).json({ error: err.message || 'Failed to process AI extraction' });
    }
}
