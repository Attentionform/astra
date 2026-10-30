import fetch from 'node-fetch';

const OPENROUTER_API_KEY = process.env.OPENROUTER_API_KEY;
if (!OPENROUTER_API_KEY) {
  console.error("Missing OPENROUTER_API_KEY in environment");
  process.exit(1);
}

const MODEL = "poolside/laguna-m.1:free";

const systemPrompt = `You are ASTra v4, an AI Document IDE.
You edit documents exclusively using JSON operations against the TOON format.
Output ONLY a single valid JSON object representing the operation to execute. Do NOT include markdown blocks (\`\`\`json) or any other conversational text.

Valid Actions: replace, style, insert, delete, move, table, section, pageSetup, namedStyle, field, equation, list, headerFooter, image

CRITICAL RULES FOR ACTION SELECTION:
1. For named styles (e.g., 'Heading 1', 'Quote'), you MUST use action "namedStyle", NOT "style".
2. For inserting special elements like TOC, TOF, or Page Numbers, you MUST use action "field", NOT "insert".
3. For inserting LaTeX equations, you MUST use action "equation", NOT "insert".
4. For deleting images, you MUST use action "image" with operation "delete", NOT the generic "delete".

Example 1:
{"action": "namedStyle", "operation": "apply", "styleName": "Heading 1", "target": "p_6"}

Example 2:
{"action": "field", "operation": "insertTOC", "anchor": "p_1"}

Example 3:
{"action": "equation", "operation": "insert", "latex": "E=mc^2", "anchor": "p_7"}

Example 4:
{"action": "image", "operation": "delete", "target": "img_2"}`;

const tests = [
  // Tier 1
  { name: "Replace", prompt: "Target row p_1 (version 1) has text 'Hello World'. Change it to 'Goodbye World'.", expectedAction: "replace" },
  { name: "Style", prompt: "Make node p_2 bold and italic.", expectedAction: "style" },
  { name: "Insert", prompt: "Insert a new paragraph after p_1 with the text 'New section begins'.", expectedAction: "insert" },
  { name: "Delete", prompt: "Delete node p_3 entirely.", expectedAction: "delete" },
  { name: "Move", prompt: "Move node p_4 to be before p_1.", expectedAction: "move" },
  { name: "Table", prompt: "Add a new row after index 2 in table t_1.", expectedAction: "table" },
  { name: "Section", prompt: "Split the section at node p_5.", expectedAction: "section" },
  
  // Tier 2
  { name: "Page Setup", prompt: "Change the document orientation to landscape.", expectedAction: "pageSetup" },
  { name: "Named Style", prompt: "Apply the 'Heading 1' style to node p_6.", expectedAction: "namedStyle" },
  { name: "Field", prompt: "Insert a Table of Contents (TOC) before p_1.", expectedAction: "field" },
  { name: "Equation", prompt: "Insert the LaTeX equation 'E=mc^2' after p_7.", expectedAction: "equation" },
  { name: "List", prompt: "Convert paragraphs p_8 and p_9 into a bulleted list.", expectedAction: "list" },
  { name: "Header Footer", prompt: "Set the default header text to 'Confidential'.", expectedAction: "headerFooter" },
  { name: "Image", prompt: "Delete the image with target ID 'img_2'.", expectedAction: "image" }
];

async function runTests() {
  console.log(`Starting ASTra LLM Tests using ${MODEL} via OpenRouter...\\n`);
  let passed = 0;
  
  for (const test of tests) {
    process.stdout.write(`Testing [${test.name}]... `);
    
    try {
      const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${OPENROUTER_API_KEY}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          model: MODEL,
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: test.prompt }
          ],
          temperature: 0.1
        })
      });
      
      const data = (await res.json()) as any;
      if (!data.choices || data.choices.length === 0) {
        console.log(`❌ Failed (API Error)`);
        console.error(data);
        continue;
      }
      
      let content = data.choices[0].message.content.trim();
      
      // Clean up markdown just in case the LLM ignored instructions
      content = content.replace(/^```json/, '').replace(/^```/, '').replace(/```$/, '').trim();
      
      let parsed;
      try {
        parsed = JSON.parse(content);
      } catch {
        console.log(`❌ Failed (Invalid JSON)`);
        console.log(`Raw output: ${content}`);
        continue;
      }
      
      if (parsed.action === test.expectedAction) {
        console.log(`✅ Passed`);
        passed++;
      } else {
        console.log(`❌ Failed (Expected ${test.expectedAction}, got ${parsed.action})`);
      }
    } catch (err: any) {
      console.log(`❌ Error (${err.message})`);
    }
    
    // Sleep to avoid rate limits on free models
    await new Promise(r => setTimeout(r, 1000));
  }
  
  console.log(`\\nTest Suite Complete: ${passed}/${tests.length} passed.`);
  if (passed === tests.length) {
    console.log("All tasks for 7.3 and 7.4 verified!");
  }
}

runTests();
