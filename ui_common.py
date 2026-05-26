"""UI helpers dùng chung cho cả 2 tab: timer box, log adder, stat box, cost."""
import html as _html
from datetime import datetime

from config import PRICE_INPUT, PRICE_OUTPUT, USD_TO_VND


def calc_cost(tok_in: int, tok_out: int) -> tuple[float, float]:
    """Quy đổi token → (USD, VND).

    Source-of-truth giá Gemini ở `config.py` (PRICE_INPUT/PRICE_OUTPUT).
    Cập nhật pricing → chỉ sửa config.py, không cần đụng hàm này.
    """
    usd = (tok_in / 1e6) * PRICE_INPUT + (tok_out / 1e6) * PRICE_OUTPUT
    vnd = usd * USD_TO_VND
    return usd, vnd


def timer_box_html(elapsed: float, status: str,
                   border="#6c63ff", val="#a78bfa", status_color="#818cf8",
                   prefix="⏱") -> str:
    """HTML timer box hiển thị thời gian chạy + dòng status (status được escape)."""
    safe_status = _html.escape(str(status))
    safe_prefix = _html.escape(str(prefix))
    return f"""<div class='timer-box' style='border-color:{border}'>
        <div class='timer-val' style='color:{val}'>{safe_prefix} {elapsed:.0f}s</div>
        <div class='timer-lbl'>Tổng thời gian đã chạy</div>
        <div class='timer-status' style='color:{status_color}'>{safe_status}</div>
    </div>"""


def timer_done_html(elapsed: float, status: str) -> str:
    safe_status = _html.escape(str(status))
    return f"""<div class='timer-box' style='border-color:#059669'>
        <div class='timer-val' style='color:#10b981'>✅ {elapsed:.1f}s</div>
        <div class='timer-lbl'>Tổng thời gian</div>
        <div class='timer-status' style='color:#10b981'>{safe_status}</div>
    </div>"""


def timer_error_html(message: str) -> str:
    safe_msg = _html.escape(str(message)[:200])
    return f"""<div class='timer-box' style='border-color:#ef4444'>
        <div class='timer-val' style='color:#ef4444'>❌ Lỗi</div>
        <div class='timer-lbl'>Đã dừng</div>
        <div class='timer-status' style='color:#ef4444'>{safe_msg}</div>
    </div>"""


def stat_box_html(value: str, label: str) -> str:
    safe_value = _html.escape(str(value))
    safe_label = _html.escape(str(label))
    return f"<div class='stat-box'><div class='stat-val'>{safe_value}</div><div class='stat-lbl'>{safe_label}</div></div>"


def make_log_adder(log_lines: list, log_ph, max_show: int = 40):
    """
    Tạo closure `add_log(msg)` — append timestamp + msg vào list và re-render.
    `log_ph` là `st.empty()` để paint vào.
    Mỗi `msg` được HTML-escape vì có thể chứa filename/exception/text từ AI.
    """
    def add_log(msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        log_lines.append(f"[{ts}] {_html.escape(str(msg))}")
        log_ph.markdown(
            f"<div class='log-box'>{'<br>'.join(log_lines[-max_show:])}</div>",
            unsafe_allow_html=True,
        )
    return add_log
