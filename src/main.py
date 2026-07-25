"""로컬 웹 화면 실행 진입점.

리포지토리 루트에서 `python src/main.py`로 실행하면 http://127.0.0.1:8000 에서
통제 등록/실행 화면을 사용할 수 있다.
"""
from __future__ import annotations

import uvicorn

if __name__ == "__main__":
    uvicorn.run("web.backend:app", host="127.0.0.1", port=8000, reload=False)
