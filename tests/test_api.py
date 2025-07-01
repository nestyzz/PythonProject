from fastapi.testclient import TestClient
from pathlib import Path
from app import app

client = TestClient(app)

class TestAPI:

    def test_full_flow_xlsb(self):
        file_path = Path(__file__).parent.parent / "storage" / "input" / "тестовое_пример_входа.xlsb"
        assert file_path.exists()

        with open(file_path, "rb") as f:
            response = client.post("/upload", files={
                "file": (file_path.name, f, "application/vnd.ms-excel.sheet.binary.macroEnabled.12")
            })

        assert response.status_code == 200
        task_id = response.json()["task_id"]

        status_resp = client.get(f"/status/{task_id}")
        assert status_resp.status_code == 200
        assert status_resp.json()["status"] == "success"

        result = client.get(f"/result/{task_id}")
        assert result.status_code == 200
        assert result.headers["content-type"].startswith("application/vnd.openxmlformats")
