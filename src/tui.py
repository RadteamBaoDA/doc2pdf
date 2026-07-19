from collections import deque
from datetime import datetime
from threading import Lock

from rich import box
from rich.console import RenderableType
from rich.layout import Layout
from rich.panel import Panel
from rich.table import Table
from rich.text import Text


class LogBuffer:
    """Captures logs for display in the TUI."""
    def __init__(self, maxlen=1000): # Increased buffer
        self.queue = deque(maxlen=maxlen)
        self.scroll_offset = 0
        self.view_height = 20 # Approximate view height, adjustable
        self._lock = Lock()
        self._changed = False
    
    def write(self, message: str):
        if message.strip():
             with self._lock:
                 self.queue.append(message.strip())
                 # Auto-scroll if at bottom (offset 0)
                 if self.scroll_offset > 0:
                     self.scroll_offset += 1
                 self._changed = True

    def consume_changed(self) -> bool:
        """Return whether new logs arrived since the last TUI refresh."""
        with self._lock:
            changed = self._changed
            self._changed = False
            return changed

    def scroll_up(self):
        """Scroll up (view older logs)."""
        with self._lock:
            if self.scroll_offset < len(self.queue) - self.view_height:
                self.scroll_offset += 1
                self._changed = True

    def scroll_down(self):
        """Scroll down (view newer logs)."""
        with self._lock:
            if self.scroll_offset > 0:
                self.scroll_offset -= 1
                self._changed = True

    def get_renderable(self) -> RenderableType:
        with self._lock:
            # Calculate slice from a consistent snapshot while worker threads log.
            total = len(self.queue)
            if total == 0:
                text = ""
            else:
                if self.scroll_offset == 0:
                    # Show latest
                    visible = list(self.queue)[-self.view_height:]
                else:
                    # Show history
                    end = total - self.scroll_offset
                    start = max(0, end - self.view_height)
                    visible = list(self.queue)[start:end]
                text = "\n".join(visible)
            scrolled = self.scroll_offset > 0
            
        return Panel(
            Text.from_markup(text),
            title=f"Application Logs {'(SCROLLED)' if scrolled else ''}",
            border_style="blue",
            box=box.ROUNDED
        )

def make_layout() -> Layout:
    """Create the main TUI layout."""
    layout = Layout()
    layout.split(
        Layout(name="header", size=4),
        Layout(name="main")
    )
    layout["main"].split_column(
        Layout(name="progress", size=6),
        Layout(name="logs") # Logs take remaining space
    )
    return layout

def get_header() -> RenderableType:
    """Create the header panel."""
    # Using a simple text banner since we can't easily display images
    grid = Table.grid(expand=True)
    grid.add_column(justify="left", ratio=1)
    grid.add_column(justify="right")
    
    title = Text(" DOC2PDF CONVERTER ", style="bold white on blue")
    meta = Text(f" {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ", style="dim")
    
    grid.add_row(title, meta)
    
    return Panel(
        grid,
        style="white on blue",
        box=box.HEAVY
    )

class TUIContext:
    """Context manager helper for TUI state."""
    def __init__(self, log_buffer: LogBuffer):
        self.log_buffer = log_buffer
        self.layout = make_layout()
        self.layout["header"].update(get_header())
        self.layout["logs"].update(log_buffer.get_renderable())

    def update_progress(self, renderable: RenderableType):
        self.layout["progress"].update(Panel(renderable, title="Progress", border_style="green", box=box.ROUNDED))
        
    def update_logs(self):
        self.layout["logs"].update(self.log_buffer.get_renderable())
        self.layout["header"].update(get_header()) # Update time
