"""Tooltip utility for tkinter widgets"""
import tkinter as tk


class ToolTip:
    """Create a tooltip for a given widget"""
    
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0
        self.widget.bind('<Enter>', self.enter)
        self.widget.bind('<Leave>', self.leave)
        self.widget.bind('<ButtonPress>', self.leave)

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(500, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        # Try to get cursor position for text widgets, fallback to widget position
        try:
            bbox = self.widget.bbox("insert")
            if bbox is not None:
                x, y, cx, cy = bbox
                x += self.widget.winfo_rootx() + 25
                y += self.widget.winfo_rooty() + 20
            else:
                # Fallback for widgets without text cursor (buttons, labels, etc.)
                x = self.widget.winfo_rootx() + 25
                y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        except (tk.TclError, TypeError):
            # Fallback for any other errors
            x = self.widget.winfo_rootx() + 25
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        
        # Create tooltip window
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))
        
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                        background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                        font=("tahoma", "8", "normal"), wraplength=300)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

