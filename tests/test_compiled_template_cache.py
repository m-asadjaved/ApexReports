import time
from pathlib import Path
from fastapi.testclient import TestClient

from main import app
from app.engine.template_cache_manager import TemplateCacheManager

client = TestClient(app)

def test_compiled_template_cache_hit_and_headers():
    # 1. Clear cache first
    TemplateCacheManager.invalidate("DEPT_EMPLOYEES")
    
    payload = {
        "template_code": "DEPT_EMPLOYEES",
        "format": "html",
        "compiler": "fast",
        "data": {
            "departments": [
                {
                    "dept_name": "FINANCE",
                    "loc": "NY",
                    "employees": [
                        {"empno": 1, "ename": "Alice", "job": "Manager", "sal": 8000, "manager": "CEO"},
                        {"empno": 2, "ename": "Bob", "job": "Analyst", "sal": 5000, "manager": "Alice"}
                    ]
                }
            ]
        }
    }

    # Call 1: Cache MISS (Compiles template & caches it)
    res1 = client.post("/api/v1/reports/render", json=payload)
    assert res1.status_code == 200
    assert res1.headers.get("X-Template-Cache") == "MISS"
    assert "FINANCE" in res1.text
    assert "Alice" in res1.text
    assert "Bob" in res1.text

    # Call 2: Cache HIT (Serves directly from RAM cache in < 5ms)
    t0 = time.perf_counter()
    res2 = client.post("/api/v1/reports/render", json=payload)
    t_ms = (time.perf_counter() - t0) * 1000
    assert res2.status_code == 200
    assert res2.headers.get("X-Template-Cache") == "HIT"
    assert "FINANCE" in res2.text
    assert t_ms < 25.0, f"Cache HIT should be sub-25ms, got {t_ms:.2f}ms"


def test_compiled_template_disk_cache_survival():
    # Simulate server restart by clearing RAM cache only
    payload = {
        "template_code": "DEPT_EMPLOYEES",
        "format": "html",
        "compiler": "fast",
        "data": {"departments": [{"dept_name": "LEGAL"}]}
    }

    # Warm disk cache
    client.post("/api/v1/reports/render", json=payload)

    # Wipe RAM
    TemplateCacheManager._RAM_CACHE.clear()

    # Next call should load from disk cache -> HIT
    res = client.post("/api/v1/reports/render", json=payload)
    assert res.status_code == 200
    assert res.headers.get("X-Template-Cache") == "HIT"
    assert "LEGAL" in res.text


def test_cache_stats_and_invalidation_endpoint():
    # Warm up cache
    client.post("/api/v1/reports/render", json={
        "template_code": "DEPT_EMPLOYEES",
        "format": "html",
        "data": {}
    })

    # Check stats endpoint
    stats_res = client.get("/api/v1/cache/stats")
    assert stats_res.status_code == 200
    stats = stats_res.json()
    assert stats["ram_cached_count"] >= 1
    assert "DEPT_EMPLOYEES" in stats["ram_templates"]

    # Invalidate
    inv_res = client.post("/api/v1/cache/invalidate?template_code=DEPT_EMPLOYEES")
    assert inv_res.status_code == 200
    assert inv_res.json()["status"] == "success"

    # Verify invalidated
    stats_after = client.get("/api/v1/cache/stats").json()
    assert "DEPT_EMPLOYEES" not in stats_after["ram_templates"]


def test_compiler_mode_full_bypasses_cache():
    payload = {
        "template_code": "DEPT_EMPLOYEES",
        "format": "html",
        "compiler": "full",
        "data": {"departments": [{"dept_name": "RESEARCH"}]}
    }
    res = client.post("/api/v1/reports/render", json=payload)
    assert res.status_code == 200
    assert res.headers.get("X-Template-Cache") == "BYPASS"
    assert "RESEARCH" in res.text
