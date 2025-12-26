from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, ValidationError
from typing import Union
import json
import re
import google.generativeai as genai
from config import GOOGLE_GEMINI_API_KEY

# --------------------------------------------------
# FASTAPI INIT
# --------------------------------------------------
app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --------------------------------------------------
# STARTUP LOG
# --------------------------------------------------
@app.on_event("startup")
async def startup_event():
    print("\n" + "=" * 55)
    print("🎤 VOICE → EXCEL COMMAND SERVER STARTED")
    print("✅ POST http://localhost:8000/command")
    print("=" * 55 + "\n")

# --------------------------------------------------
# INPUT MODEL
# --------------------------------------------------
class TextCommand(BaseModel):
    text: str

# --------------------------------------------------
# STRICT AI OUTPUT SCHEMA
# --------------------------------------------------
class ExcelCommand(BaseModel):
    action: str

    # write / delete
    cell: str | None = None
    value: Union[str, int, float, bool, None] = None

    # insert operations
    column: str | None = None
    row: int | None = None

    # range-based operations
    range: str | None = None
    target: str | None = None

    # sort / filter
    order: str | None = None
    condition: str | None = None

    # chart
    x_column: str | None = None
    y_column: str | None = None

    confidence: float | None = None

# --------------------------------------------------
# NORMALIZE USER TEXT
# --------------------------------------------------


# --------------------------------------------------
# CLEAN GEMINI OUTPUT
# --------------------------------------------------
def clean_json(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = text.replace("```json", "").replace("```", "").strip()
    return text

# --------------------------------------------------
# GOOGLE GEMINI API CALL (SINGLE CALL ONLY)
# --------------------------------------------------
def ask_gemini(prompt: str) -> str | None:
    genai.configure(api_key=GOOGLE_GEMINI_API_KEY)

    system_prompt = """
You are an Excel command parser.

Convert user instructions into structured Excel actions.

Return ONLY valid JSON.
NO explanation. NO markdown. NO extra text.

Do not assume fixed cell positions unless explicitly stated.
Prefer semantic targets such as:
- "same column bottom"
- "entire column"
- "next available row"

Allowed actions and formats:

1. Write value
{ "action": "write", "cell": "A1", "value": 5, "confidence": 0.95 }

2. Delete cell
{ "action": "delete_cell", "cell": "A1", "confidence": 0.95 }

3. Insert row
{ "action": "insert_row", "row": 3, "confidence": 0.95 }

4. Insert column
{ "action": "insert_column", "column": "C", "confidence": 0.95 }

5. Sum column
{ "action": "sum", "range": "A:A", "target": "same column bottom", "confidence": 0.95 }

6. Average column
{ "action": "average", "range": "A:A", "target": "same column bottom", "confidence": 0.95 }

7. Standard deviation
{ "action": "stddev", "range": "A:A", "target": "same column bottom", "confidence": 0.95 }

8. Bold column
{ "action": "bold", "range": "B:B", "confidence": 0.95 }

9. Sort column
{ "action": "sort", "range": "A:A", "order": "asc", "confidence": 0.95 }

10. Filter values
{ "action": "filter", "range": "A:A", "condition": ">10", "confidence": 0.95 }

11. Create chart
{ "action": "create_chart", "x_column": "A", "y_column": "B", "confidence": 0.95 }
"""

    try:
        model = genai.GenerativeModel(
            model_name="gemini-2.5-flash-lite",
            generation_config=genai.types.GenerationConfig(
                temperature=0,
            ),
            system_instruction=system_prompt,
        )

        response = model.generate_content(prompt)
        raw = response.text

        print("\n--- RAW GEMINI RESPONSE ---")
        print(raw)
        print("--------------------------")

        return clean_json(raw)

    except Exception as e:
        print("❌ GOOGLE GEMINI API ERROR:", e)
        return None

# --------------------------------------------------
# HEALTH CHECK
# --------------------------------------------------
@app.get("/health")
def health():
    return {"status": "running"}

# --------------------------------------------------
# MAIN COMMAND ENDPOINT
# --------------------------------------------------
@app.post("/command")
def command(input: TextCommand):
    
    print("\n🎤 USER SAID:", input.text)



    ai_output = ask_gemini(input.text)

    if not ai_output:
        return {
            "action": "error",
            "message": "Gemini API unreachable"
        }

    print("\n🤖 AI CLEAN OUTPUT:")
    print(ai_output)

    try:
        parsed = json.loads(ai_output)
        validated = ExcelCommand(**parsed)

        print("\n✅ PARSED & VALIDATED:")
        print(validated.model_dump())

        return validated.model_dump()

    except (json.JSONDecodeError, ValidationError) as e:
        print("\n❌ INVALID AI RESPONSE")
        print(e)

        return {
            "action": "error",
            "message": "Invalid AI response",
            "raw_response": ai_output
        }

# --------------------------------------------------
# RUN SERVER
# --------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
