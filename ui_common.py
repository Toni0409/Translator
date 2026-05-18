"""UI helpers dùng chung cho cả 2 tab: timer box, log adder, stat box, cost."""
from datetime import datetime

from config import PRICE_INPUT, PRICE_OUTPUT, USD_TO_VND


def calc_cost(tok_in: int, tok_out: int) -> tuple[float, float]:
    """Quy đổi token → (USD, VND). PDF + Word dùng chung công thức."""
    usd = ((tok_in / 1e6) * PRICE_INPUT + (tok_out / 1e6) * PRICE_OUTPUT) * 10
    vnd = usd * USD_TO_VND
    return usd, vnd


def timer_box_html(elapsed: float, status: str,
                   border="#6c63ff", val="#a78bfa", status_color="#818cf8",
                   prefix="⏱") -> str:
    """HTML timer box hiển thị thời gian chạy + dòng status."""
    return f"""<div class='timer-box' style='border-color:{border}'>
        <div class='timer-val' style='color:{val}'>{prefix} {elapsed:.0f}s</div>
        <div class='timer-lbl'>Tổng thời gian đã chạy</div>
        <div class='timer-status' style='color:{status_color}'>{status}</div>
    </div>"""


def timer_done_html(elapsed: float, status: str) -> str:
    return f"""<div class='timer-box' style='border-color:#059669'>
        <div class='timer-val' style='color:#10b981'>✅ {elapsed:.1f}s</div>
        <div class='timer-lbl'>Tổng thời gian</div>
        <div class='timer-status' style='color:#10b981'>{status}</div>
    </div>"""


def timer_error_html(message: str) -> str:
    return f"""<div class='timer-box' style='border-color:#ef4444'>
        <div class='timer-val' style='color:#ef4444'>❌ Lỗi</div>
        <div class='timer-lbl'>Đã dừng</div>
        <div class='timer-status' style='color:#ef4444'>{message[:80]}</div>
    </div>"""


def stat_box_html(value: str, label: str) -> str:
    return f"<div class='stat-box'><div class='stat-val'>{value}</div><div class='stat-lbl'>{label}</div></div>"


def make_log_adder(log_lines: list, log_ph, max_show: int = 40):
    """
    Tạo closure `add_log(msg)` — append timestamp + msg vào list và re-render.
    `log_ph` là `st.empty()` để paint vào.
    """
    def add_log(msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        log_lines.append(f"[{ts}] {msg}")
        log_ph.markdown(
            f"<div class='log-box'>{'<br>'.join(log_lines[-max_show:])}</div>",
            unsafe_allow_html=True,
        )
    return add_log
