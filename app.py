import uuid
import json
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from refactor import RefactorExcel


class TaskStatus(BaseModel):
    task_id: str
    status: str
    error: str | None = None
app = FastAPI()
BASE_DIR = Path(__file__).parent
STORAGE = BASE_DIR / 'storage'
INPUT_DIR = STORAGE / 'input'
OUTPUT_DIR = STORAGE / 'output'
STATUS_FILE = STORAGE / 'tasks.json'

for d in [INPUT_DIR, OUTPUT_DIR]:
    d.mkdir(parents=True, exist_ok=True)

def load_status():
    if STATUS_FILE.exists():
        return json.loads(STATUS_FILE.read_text())
    return {}

def save_status(data):
    STATUS_FILE.write_text(json.dumps(data, indent=2))

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    task_id = str(uuid.uuid4())

    # сохраняем расширение исходного файла
    suffix = Path(file.filename).suffix
    input_path = INPUT_DIR / f"{task_id}{suffix}"
    output_path = OUTPUT_DIR / f"{task_id}.xlsx"

    status = load_status()
    status[task_id] = {"status": "pending"}
    save_status(status)

    try:
        with input_path.open("wb") as f:
            f.write(await file.read())

        processor = RefactorExcel(input_path)
        processor.run(output_path)

        status[task_id] = {"status": "success"}
    except Exception as e:
        print("ERROR in /upload:", e)
        status[task_id] = {"status": "failed", "error": str(e)}
    finally:
        save_status(status)

    return {"task_id": task_id}

@app.get("/status/{task_id}", response_model=TaskStatus)
def get_status(task_id: str):
    status = load_status()
    if task_id not in status:
        raise HTTPException(status_code=404, detail="Task not found")
    return {"task_id": task_id, **status[task_id]}

@app.get("/result/{task_id}")
def get_result(task_id: str):
    output_path = OUTPUT_DIR / f"{task_id}.xlsx"
    status = load_status()

    if status.get(task_id, {}).get("status") != "success":
        raise HTTPException(status_code=404, detail="Result not ready")

    if not output_path.exists():
        raise HTTPException(status_code=404, detail="File not found")

    return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=f"{task_id}.xlsx")
