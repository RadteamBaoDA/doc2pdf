from collections import deque
from datetime import datetime
from typing import Optional

from rich.console import RenderableType, Group
from rich.layout import Layout
from rich.panel import Panel
from rich.text import Text
from rich.align import Align
from rich.table import Table
from rich import box

class LogBuffer:
    """Captures logs for display in the TUI."""
    def __init__(self, maxlen=100):
        self.queue = deque(maxlen=maxlen)
    
    def write(self, message: str):
        if message.strip():
             self.queue.append(message.strip())

    def get_renderable(self) -> RenderableType:
        return Panel(
            Text.from_markup("\n".join(self.queue)),
            title="Application Logs",
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
