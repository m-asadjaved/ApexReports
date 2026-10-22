import atexit
import asyncio
import base64
import json
import logging
import os
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Optional

import httpx
import websockets

from app.renderer.browser_pdf_renderer import BrowserPdfRenderer
from app.core.exceptions import RenderingException

logger = logging.getLogger(__name__)


class ChromeCdpPool:
    """
    Manages a persistent headless Chrome/Chromium daemon with Chrome DevTools Protocol (CDP).
    Enables blazing-fast (~120ms) vector PDF rendering without process startup overhead.
    """

    _process: Optional[subprocess.Popen] = None
    _port: int = 9222
    _profile_dir: Optional[str] = None
    _started: bool = False
    _lock: asyncio.Lock = asyncio.Lock()

    @classmethod
    def get_port(cls) -> int:
        return cls._port

    @classmethod
    def is_available(cls) -> bool:
        """Returns True if a compatible browser executable is installed."""
        return BrowserPdfRenderer.is_available()

    @classmethod
    def is_running(cls) -> bool:
        """Checks if the Chrome CDP daemon is alive and responding on its debugging port."""
        try:
            with httpx.Client(timeout=0.8) as client:
                res = client.get(f"http://127.0.0.1:{cls._port}/json/version")
                return res.status_code == 200
        except Exception:
            return False

    @classmethod
    def ensure_started(cls, timeout: float = 4.0) -> bool:
        """
        Ensures the persistent headless Chrome daemon is running.
        If not, spawns it in the background and waits for CDP initialization.
        """
        if cls.is_running():
            cls._started = True
            return True

        browser_bin = BrowserPdfRenderer.get_browser_executable()
        if not browser_bin:
            logger.warning("No Chrome/Edge executable found; CDP pool unavailable.")
            return False

        if not cls._profile_dir:
            cls._profile_dir = tempfile.mkdtemp(prefix="apex_cdp_profile_")

        logger.info(f"Launching persistent headless Chrome CDP daemon on port {cls._port}...")

        cmd = [
            browser_bin,
            "--headless=new",
            f"--remote-debugging-port={cls._port}",
            "--disable-gpu",
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-background-networking",
            "--disable-sync",
            "--disable-translate",
            "--metrics-recording-only",
            f"--user-data-dir={cls._profile_dir}",
            "about:blank"
        ]

        try:
            cls._process = subprocess.Popen(
                cmd,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
            )

            # Wait for CDP endpoint to respond
            deadline = time.time() + timeout
            while time.time() < deadline:
                if cls.is_running():
                    cls._started = True
                    logger.info(f"Chrome CDP daemon successfully initialized on port {cls._port}.")
                    # Register shutdown handler
                    atexit.register(cls.stop)
                    return True
                time.sleep(0.15)

            logger.warning("Chrome CDP daemon did not start within deadline.")
            return False

        except Exception as e:
            logger.error(f"Failed to launch Chrome CDP daemon: {e}")
            return False

    @classmethod
    async def render_html_to_pdf_async(cls, html_content: str, timeout: float = 12.0) -> bytes:
        """
        Renders an HTML string into PDF bytes using a warm tab over CDP WebSocket.
        Typical execution latency: 100ms - 180ms.
        """
        if not cls.ensure_started():
            raise RenderingException("Chrome CDP daemon is not running.")

        async with httpx.AsyncClient(timeout=5.0) as http_client:
            # 1. Create a new blank tab target
            new_tab_res = await http_client.put(f"http://127.0.0.1:{cls._port}/json/new")
            if new_tab_res.status_code != 200:
                raise RenderingException(f"Failed to create CDP tab: {new_tab_res.text}")

            tab_info = new_tab_res.json()
            tab_id = tab_info["id"]
            ws_url = tab_info.get("webSocketDebuggerUrl")

            if not ws_url:
                raise RenderingException("CDP did not return a WebSocket debugger URL.")

            try:
                # 2. Connect via WebSocket and execute commands
                async with websockets.connect(ws_url, max_size=50 * 1024 * 1024) as ws:
                    cmd_id = 0

                    async def send_command(method: str, params: Optional[dict] = None) -> dict:
                        nonlocal cmd_id
                        cmd_id += 1
                        msg_id = cmd_id
                        payload = {"id": msg_id, "method": method}
                        if params:
                            payload["params"] = params
                        await ws.send(json.dumps(payload))
                        while True:
                            raw = await ws.recv()
                            msg = json.loads(raw)
                            if msg.get("id") == msg_id:
                                if "error" in msg:
                                    raise RenderingException(f"CDP command '{method}' failed: {msg['error']}")
                                return msg

                    # Enable Page domain
                    await send_command("Page.enable")

                    # Navigate to base64 data URL
                    b64_html = base64.b64encode(html_content.encode("utf-8")).decode("ascii")
                    data_url = f"data:text/html;charset=utf-8;base64,{b64_html}"
                    await send_command("Page.navigate", {"url": data_url})

                    # Wait for page load event
                    load_deadline = time.time() + timeout
                    while time.time() < load_deadline:
                        raw = await asyncio.wait_for(ws.recv(), timeout=timeout)
                        msg = json.loads(raw)
                        if msg.get("method") == "Page.loadEventFired":
                            break

                    # Print to PDF via CDP
                    print_options = {
                        "printBackground": True,
                        "preferCSSPageSize": True,
                        "marginTop": 0,
                        "marginBottom": 0,
                        "marginLeft": 0,
                        "marginRight": 0
                    }
                    pdf_res = await send_command("Page.printToPDF", print_options)
                    pdf_b64 = pdf_res["result"]["data"]
                    pdf_bytes = base64.b64decode(pdf_b64)
                    return pdf_bytes

            finally:
                # 3. Always close the tab to avoid memory leaks
                try:
                    await http_client.put(f"http://127.0.0.1:{cls._port}/json/close/{tab_id}")
                except Exception:
                    pass

    @classmethod
    def render_html_to_pdf(cls, html_content: str, timeout: float = 12.0) -> bytes:
        """
        Synchronous wrapper for render_html_to_pdf_async.
        Safe for use inside sync FastAPI route handlers or background tasks.
        """
        try:
            try:
                loop = asyncio.get_running_loop()
            except RuntimeError:
                loop = None

            if loop and loop.is_running():
                # If already inside an active event loop, run in a separate thread pool
                import concurrent.futures
                with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
                    future = pool.submit(lambda: asyncio.run(cls.render_html_to_pdf_async(html_content, timeout)))
                    return future.result()
            else:
                return asyncio.run(cls.render_html_to_pdf_async(html_content, timeout))
        except Exception as e:
            logger.warning(f"CDP rendering failed: {e}. Falling back to CLI browser renderer.")
            return BrowserPdfRenderer.convert_html_to_pdf(html_content)

    @classmethod
    def stop(cls):
        """Terminates the Chrome CDP process cleanly."""
        if cls._process:
            try:
                cls._process.terminate()
                cls._process.wait(timeout=2.0)
            except Exception:
                try:
                    cls._process.kill()
                except Exception:
                    pass
            cls._process = None
            cls._started = False
            logger.info("Chrome CDP daemon stopped.")
