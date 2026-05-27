from __future__ import annotations

from io import BytesIO
import zipfile

from fastapi.testclient import TestClient

from web_app import app


def test_web_index_renders_upload_and_actions() -> None:
    client = TestClient(app)

    response = client.get("/")

    assert response.status_code == 200
    assert "上传文件" in response.text
    assert "自然语言规则" in response.text
    assert "分析文件" in response.text
    assert "生成填充计划" in response.text
    assert "执行填充" in response.text
    assert "下载已填充 Excel" in response.text


def test_web_upload_saves_files(tmp_path, monkeypatch) -> None:
    import web_app

    monkeypatch.setattr(web_app, "MAIN_DIR", tmp_path / "input" / "main_excel")
    monkeypatch.setattr(web_app, "TEMPLATE_DIR", tmp_path / "input" / "templates")
    monkeypatch.setattr(web_app, "CONFIG_DIR", tmp_path / "config")
    client = TestClient(app)

    response = client.post(
        "/upload",
        data={
            "rules_text": "主表 Sheet「发票数据」填入 template.xlsx 模板。\n客户名称填入 B5。",
        },
        files=[
            ("main_excel", ("main.xlsx", b"main", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
            ("templates", ("template.xlsx", b"template", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
        ],
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert (tmp_path / "input" / "main_excel" / "main.xlsx").exists()
    assert (tmp_path / "input" / "templates" / "template.xlsx").exists()
    assert (tmp_path / "config" / "mapping_config.json").exists()
    assert (tmp_path / "config" / "rules_prompt.txt").exists()


def test_web_download_only_includes_filled_excel_files(tmp_path, monkeypatch) -> None:
    import web_app

    output_dir = tmp_path / "output"
    filled_dir = output_dir / "filled_files"
    logs_dir = output_dir / "logs"
    filled_dir.mkdir(parents=True)
    logs_dir.mkdir(parents=True)
    (filled_dir / "filled.xlsx").write_bytes(b"filled")
    (logs_dir / "process_log.xlsx").write_bytes(b"log")
    (output_dir / "summary.txt").write_text("summary", encoding="utf-8")

    monkeypatch.setattr(web_app, "OUTPUT_DIR", output_dir)
    monkeypatch.setattr(web_app, "WEB_DIR", output_dir / "web")
    client = TestClient(app)

    response = client.get("/download")

    assert response.status_code == 200
    assert response.headers["content-type"] == "application/zip"
    with zipfile.ZipFile(BytesIO(response.content)) as archive:
        assert archive.namelist() == ["filled.xlsx"]
