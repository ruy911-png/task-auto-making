"""pytest가 src/ 아래 패키지(excel_io, sap_automation, batch, lookup, validation, web)를
`import excel_io` 형태로 바로 임포트할 수 있도록 경로를 추가한다."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))
