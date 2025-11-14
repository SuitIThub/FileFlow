import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
import json
import time
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from abc import ABC, abstractmethod
from typing import List, Dict, Any
import threading
from PIL import Image, ImageTk
import io


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


class Rule(ABC):
    """Abstract base class for renaming rules"""
    
    def __init__(self, tag_name: str):
        self.tag_name = tag_name
    
    @abstractmethod
    def get_value(self, file_index: int, batch_count: int) -> str:
        """Get the replacement value for this rule"""
        pass
    
    @abstractmethod
    def reset(self):
        """Reset the rule state for a new batch"""
        pass
    
    @abstractmethod
    def to_dict(self) -> Dict[str, Any]:
        """Serialize rule to dictionary"""
        pass
    
    @classmethod
    @abstractmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Rule':
        """Deserialize rule from dictionary"""
        pass


class CounterRule(Rule):
    """Rule that counts up with each file in the batch"""
    
    def __init__(self, tag_name: str, start_value: int = 0, increment: int = 1, step: int = 1, max_value: int = None):
        super().__init__(tag_name)
        self.start_value = start_value
        self.increment = increment
        self.step = step  # How many operations before incrementing
        self.max_value = max_value  # Maximum value before wrapping to start_value
        self.current_value = start_value
        self.operation_count = 0
    
    def get_value(self, file_index: int, batch_count: int) -> str:
        value = self.current_value
        self.operation_count += 1
        
        # Only increment when we've reached the step threshold
        if self.operation_count % self.step == 0:
            self.current_value += self.increment
            
            # Handle max value wrapping
            if self.max_value is not None and self.current_value > self.max_value:
                self.current_value = self.start_value
        
        return str(value)
    
    def reset(self):
        self.current_value = self.start_value
        self.operation_count = 0
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'type': 'counter',
            'tag_name': self.tag_name,
            'start_value': self.start_value,
            'increment': self.increment,
            'step': self.step,
            'max_value': self.max_value
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'CounterRule':
        return cls(data['tag_name'], data['start_value'], data['increment'], data.get('step', 1), data.get('max_value'))


class ListRule(Rule):
    """Rule that iterates through a list of values"""
    
    def __init__(self, tag_name: str, values: List[str], step: int = 1):
        super().__init__(tag_name)
        self.values = values
        self.step = step  # How many operations before advancing to next value
        self.current_index = 0
        self.operation_count = 0
    
    def get_value(self, file_index: int, batch_count: int) -> str:
        if not self.values:
            return ""
        
        value = self.values[self.current_index % len(self.values)]
        self.operation_count += 1
        
        # Only advance to next value when we've reached the step threshold
        if self.operation_count % self.step == 0:
            self.current_index += 1
        
        return value
    
    def reset(self):
        self.current_index = 0
        self.operation_count = 0
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'type': 'list',
            'tag_name': self.tag_name,
            'values': self.values,
            'step': self.step
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ListRule':
        return cls(data['tag_name'], data['values'], data.get('step', 1))


class BatchRule(Rule):
    """Rule that counts up with each batch"""
    
    def __init__(self, tag_name: str, start_value: int = 0, increment: int = 1, step: int = 1, max_value: int = None):
        super().__init__(tag_name)
        self.start_value = start_value
        self.increment = increment
        self.step = step  # How many batches before incrementing
        self.max_value = max_value  # Maximum value before wrapping to start_value
        self.current_value = start_value
        self.batch_count = 0
    
    def get_value(self, file_index: int, batch_count: int) -> str:
        return str(self.current_value)
    
    def reset(self):
        pass  # Batch counter doesn't reset per batch
    
    def increment_batch(self):
        self.batch_count += 1
        
        # Only increment when we've reached the step threshold
        if self.batch_count % self.step == 0:
            self.current_value += self.increment
            
            # Handle max value wrapping
            if self.max_value is not None and self.current_value > self.max_value:
                self.current_value = self.start_value
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'type': 'batch',
            'tag_name': self.tag_name,
            'start_value': self.start_value,
            'increment': self.increment,
            'step': self.step,
            'max_value': self.max_value,
            'current_value': self.current_value,
            'batch_count': self.batch_count
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'BatchRule':
        rule = cls(data['tag_name'], data['start_value'], data['increment'], data.get('step', 1), data.get('max_value'))
        rule.current_value = data.get('current_value', data['start_value'])
        rule.batch_count = data.get('batch_count', 0)
        return rule


class FileWatcher(FileSystemEventHandler):
    """Handles file system events"""
    
    def __init__(self, app):
        self.app = app
        self.last_check_time = time.time()
    
    def on_created(self, event):
        if not event.is_directory:
            # Add a small delay to ensure file is fully written
            threading.Timer(0.5, self._process_new_file, args=[event.src_path]).start()
    
    def _process_new_file(self, file_path):
        # Wait a bit longer to ensure file is fully written
        time.sleep(0.5)
        
        # Check if file is still being written by trying to open it
        max_retries = 3
        for attempt in range(max_retries):
            try:
                # Try to open the file to ensure it's not being written
                with open(file_path, 'rb') as f:
                    f.read(1)  # Try to read a byte
                break
            except (PermissionError, IOError):
                if attempt < max_retries - 1:
                    time.sleep(1.0)  # Wait longer between retries
                else:
                    return  # Skip this file if we can't access it
        
        self.app.add_tracked_file(file_path)


class FileManagerApp:
    """Main application class"""
    
    def __init__(self):
        # Force working directory to script location for tkinterdnd2 compatibility
        script_dir = os.path.dirname(os.path.abspath(__file__))
        original_cwd = os.getcwd()
        
        # Always change to script directory for tkinterdnd2
        os.chdir(script_dir)
        
        # Use tkinterdnd2 for drag and drop support
        try:
            import tkinterdnd2 as tkdnd
            self.root = tkdnd.Tk()
            self.drag_drop_available = True
            print(f"Drag and drop initialized successfully from: {script_dir}")
            
        except Exception as e:
            print(f"Drag and drop initialization failed: {e}")
            # Fall back to regular Tk if tkinterdnd2 fails to initialize
            self.root = tk.Tk()
            self.drag_drop_available = False
        
        # Store for potential restoration (but don't restore yet - tkinterdnd2 needs the directory)
        self.original_cwd = original_cwd
        self.script_dir = script_dir
        
        self.root.title("File Manager with Rule-based Renaming")
        self.root.geometry("900x700")
        
        # Application state
        self.source_folder = tk.StringVar()
        self.dest_folder = tk.StringVar()
        self.file_formats = tk.StringVar(value="*")
        self.naming_pattern = tk.StringVar(value="file_{counter}")
        self.tracked_files = []
        self.rules = []
        self.observer = None
        self.last_file_time = 0
        self.status_text = tk.StringVar(value="Ready")
        
        # View state
        self.view_mode = tk.StringVar(value="list")  # "list" or "grid"
        self.thumbnail_cache = {}  # Cache for generated thumbnails
        
        # Widget tracking for incremental updates
        self.file_widgets = {}  # Index -> widget_data mapping for files
        self.rule_widgets = {}  # Index -> widget_data mapping for rules
        self.last_files_state = []  # Last known state of files for change detection
        self.last_rules_state = []  # Last known state of rules for change detection
        
        # Button references for state management
        self.start_tracking_btn = None
        self.stop_tracking_btn = None
        self.copy_rename_btn = None
        self.clear_tracked_btn = None
        
        # UI label references for file count and latest rename
        self.file_count_label = None
        self.latest_rename_label = None
        self.latest_rename_info = None  # Tuple of (original_name, new_name)
        
        # Word separator symbols for Ctrl+Arrow navigation
        # Characters in this string will be treated as word separators
        # Whitespace is always treated as a separator
        self.word_separators = "{}[]().,;:!@#$%^&*-+=|\\/<>?~`\"'_"
        
        # Settings file - ensure it's in the same directory as the script
        self.settings_file = os.path.join(self.script_dir, "file_manager_settings.json")
        
        self.create_ui()
        self.load_settings()
        self.update_path_labels()  # Set initial path label states
        self.update_button_states()  # Set initial button states
        self.update_naming_pattern_label()  # Set initial label state
        self.update_file_count_label()  # Set initial file count
        self.update_latest_rename_label()  # Set initial latest rename label
        
    def create_ui(self):
        """Create the user interface"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Source folder selection
        self.source_label = ttk.Label(main_frame, text="Source Folder:")
        self.source_label.grid(row=row, column=0, sticky=tk.W, pady=5)
        ToolTip(self.source_label, "The folder to monitor for new files. When tracking is started, new files appearing in this folder will be automatically added to the tracked files list.\n\n✅ = Path exists and is valid\n❌ = Path does not exist\nNo symbol = Path is empty")
        
        self.source_entry = ttk.Entry(main_frame, textvariable=self.source_folder, width=50)
        self.source_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5)
        self._setup_custom_word_navigation(self.source_entry)
        ToolTip(self.source_entry, "Path to the folder containing files to track. You can type a path or use the Browse button to select a folder.\n\nNote: This field is disabled while tracking is active to prevent conflicts.")
        
        self.source_create = ttk.Button(main_frame, text="Create", command=self.create_source_folder)
        self.source_create.grid(row=row, column=2, padx=(5, 2))
        self.source_create.grid_remove()  # Hide initially
        ToolTip(self.source_create, "Create the source folder and all necessary parent directories.")
        
        self.source_browse = ttk.Button(main_frame, text="Browse", command=self.browse_source)
        self.source_browse.grid(row=row, column=3, padx=5)
        ToolTip(self.source_browse, "Click to open a folder selection dialog to choose the source folder.\n\nNote: This button is disabled while tracking is active to prevent conflicts.")
        row += 1
        
        # Destination folder selection
        self.dest_label = ttk.Label(main_frame, text="Destination Folder:")
        self.dest_label.grid(row=row, column=0, sticky=tk.W, pady=5)
        ToolTip(self.dest_label, "The folder where renamed copies of tracked files will be saved. Files in this folder may be overwritten if they have the same name as generated filenames.\n\n✅ = Path exists and is valid\n❌ = Path does not exist\nNo symbol = Path is empty")
        
        self.dest_entry = ttk.Entry(main_frame, textvariable=self.dest_folder, width=50)
        self.dest_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5)
        self._setup_custom_word_navigation(self.dest_entry)
        ToolTip(self.dest_entry, "Path to the destination folder. You can type a path or use the Browse button to select a folder.\n\nNote: This field is disabled while tracking is active to prevent conflicts.")
        
        self.dest_create = ttk.Button(main_frame, text="Create", command=self.create_dest_folder)
        self.dest_create.grid(row=row, column=2, padx=(5, 2))
        self.dest_create.grid_remove()  # Hide initially
        ToolTip(self.dest_create, "Create the destination folder and all necessary parent directories.")
        
        self.dest_browse = ttk.Button(main_frame, text="Browse", command=self.browse_dest)
        self.dest_browse.grid(row=row, column=3, padx=5)
        ToolTip(self.dest_browse, "Click to open a folder selection dialog to choose the destination folder.\n\nNote: This button is disabled while tracking is active to prevent conflicts.")
        row += 1
        
        # File formats
        formats_label = ttk.Label(main_frame, text="File Formats (semicolon separated):")
        formats_label.grid(row=row, column=0, sticky=tk.W, pady=5)
        ToolTip(formats_label, "Filter which file types to track. Use '*' for all files, or specify extensions like '.jpg;.png;.gif' or patterns like '*.txt;*.doc'.")
        
        self.formats_entry = ttk.Entry(main_frame, textvariable=self.file_formats, width=50)
        self.formats_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5)
        self._setup_custom_word_navigation(self.formats_entry)
        ToolTip(self.formats_entry, "Semicolon-separated list of file formats to track. Examples:\n• '*' = all files\n• '.jpg;.png' = only jpg and png files\n• '*.txt;*.doc' = text and document files\n\nNote: This field is disabled while tracking is active to prevent conflicts.")
        row += 1
        
        # Naming pattern
        self.naming_pattern_label = ttk.Label(main_frame, text="Naming Pattern:")
        self.naming_pattern_label.grid(row=row, column=0, sticky=tk.W, pady=5)
        ToolTip(self.naming_pattern_label, "Template for generating new filenames. Use {tag_name} to insert values from rules. ⚠️ appears when tags are used without corresponding rules. Example: 'photo_{counter}_{batch}' creates 'photo_1_0.jpg', 'photo_2_0.jpg', etc.")
        
        pattern_entry = ttk.Entry(main_frame, textvariable=self.naming_pattern, width=50)
        pattern_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5)
        self._setup_custom_word_navigation(pattern_entry)
        ToolTip(pattern_entry, "Enter the filename pattern using {tag_name} placeholders. Tags will be replaced with values from rules you create below. Example patterns:\n• 'file_{counter}' → file_1, file_2, etc.\n• '{list}_{counter}' → value1_1, value2_2, etc.\n• 'batch{batch}_item{counter}' → batch0_item1, batch0_item2, etc.")
        row += 1
        
        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, pady=10)
        
        self.start_tracking_btn = ttk.Button(button_frame, text="Start Tracking", command=self.start_tracking)
        self.start_tracking_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(self.start_tracking_btn, "Begin monitoring the source folder for new files. Any new files that appear will be automatically added to the tracked files list. Requires a valid source folder.")
        
        self.stop_tracking_btn = ttk.Button(button_frame, text="Stop Tracking", command=self.stop_tracking)
        self.stop_tracking_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(self.stop_tracking_btn, "Stop monitoring the source folder. No new files will be added to the tracked list, but existing tracked files remain.")
        
        self.copy_rename_btn = ttk.Button(button_frame, text="Copy & Rename Files", command=self.copy_and_rename)
        self.copy_rename_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(self.copy_rename_btn, "Copy all tracked files to the destination folder with new names based on the naming pattern and rules. Disabled when there are duplicate names among tracked files. May show confirmation dialogs for existing files or missing rules.")
        
        self.clear_tracked_btn = ttk.Button(button_frame, text="Clear Tracked Files", command=self.clear_tracked)
        self.clear_tracked_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(self.clear_tracked_btn, "Remove all files from the tracked files list. This does not affect the actual files, only clears the list.")
        
        self.add_files_btn = ttk.Button(button_frame, text="Add Files", command=self.add_files_manually)
        self.add_files_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(self.add_files_btn, "Manually select one or more files to add to the tracked files list. Files must match the current file format filter.")
        row += 1
        
        # Tracked files section header and controls
        files_header_frame = ttk.Frame(main_frame)
        files_header_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))
        files_header_frame.columnconfigure(1, weight=1)
        
        # Tracked files label and count in a frame
        tracked_label_frame = ttk.Frame(files_header_frame)
        tracked_label_frame.grid(row=0, column=0, sticky=tk.W)
        
        tracked_label = ttk.Label(tracked_label_frame, text="Tracked Files:")
        tracked_label.pack(side=tk.LEFT)
        ToolTip(tracked_label, "List of files to be copied and renamed. Files show as 'original_name → preview_name'. Colors indicate conflicts:\n• Red background = duplicate preview names\n• Blue background = preview name exists in destination\n• Normal = no conflicts")
        
        # File count label
        self.file_count_label = ttk.Label(tracked_label_frame, text="(0)")
        self.file_count_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # View toggle buttons
        view_frame = ttk.Frame(files_header_frame)
        view_frame.grid(row=0, column=1, padx=10)
        
        list_view_btn = ttk.Radiobutton(view_frame, text="List View", variable=self.view_mode, 
                                       value="list", command=self.update_files_display)
        list_view_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(list_view_btn, "Show files in a compact list format with file details and controls on each row.")
        
        grid_view_btn = ttk.Radiobutton(view_frame, text="Grid View", variable=self.view_mode, 
                                       value="grid", command=self.update_files_display)
        grid_view_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(grid_view_btn, "Show files in a grid layout with thumbnails, file names, and preview names.")
        
        # Latest rename info label (frame with multiple colored labels)
        self.latest_rename_frame = ttk.Frame(files_header_frame)
        self.latest_rename_frame.grid(row=0, column=2, sticky=tk.E, padx=(10, 0))
        
        self.latest_rename_prefix = tk.Label(self.latest_rename_frame, text="Latest: ", font=("Arial", 9))
        self.latest_rename_prefix.pack(side=tk.LEFT)
        
        self.latest_rename_original = tk.Label(self.latest_rename_frame, text="", fg="red", font=("Arial", 9))
        self.latest_rename_original.pack(side=tk.LEFT)
        
        self.latest_rename_arrow = tk.Label(self.latest_rename_frame, text=" → ", font=("Arial", 9))
        self.latest_rename_arrow.pack(side=tk.LEFT)
        
        self.latest_rename_preview = tk.Label(self.latest_rename_frame, text="", fg="green", font=("Arial", 9))
        self.latest_rename_preview.pack(side=tk.LEFT)
        
        # Keep reference for compatibility
        self.latest_rename_label = self.latest_rename_frame
        

        
        row += 1
        
        # Create frame for files list with scrollbar
        files_container = ttk.Frame(main_frame)
        files_container.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        files_container.columnconfigure(0, weight=1)
        files_container.rowconfigure(0, weight=1)
        
        # Create a canvas and scrollbar for files
        self.files_canvas = tk.Canvas(files_container, height=200)
        files_scrollbar = ttk.Scrollbar(files_container, orient="vertical", command=self.files_canvas.yview)
        self.files_scrollable_frame = ttk.Frame(self.files_canvas)
        
        # Configure the scrollable frame to expand
        self.files_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.files_canvas.configure(scrollregion=self.files_canvas.bbox("all"))
        )
        
        # Configure canvas to resize its content when canvas size changes
        self.files_canvas.bind(
            "<Configure>",
            lambda e: self.files_canvas.itemconfig(self.files_canvas_frame_id, width=e.width)
        )
        
        self.files_canvas_frame_id = self.files_canvas.create_window((0, 0), window=self.files_scrollable_frame, anchor="nw")
        self.files_canvas.configure(yscrollcommand=files_scrollbar.set)
        
        self.files_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Set up mouse wheel scrolling for files
        self.setup_mouse_wheel_scrolling(self.files_canvas, self.files_scrollable_frame)
        
        # Set up drag and drop functionality
        self.setup_drag_and_drop()
        
        # Add tooltip to the files canvas
        tooltip_text = "Scrollable list of tracked files with preview names. Each file has:\n• ↑↓ buttons to reorder files\n• Preview of original → renamed filename\n• ✕ button to remove individual files"
        if self.drag_drop_available:
            tooltip_text += "\n• Drag and drop files here to add them to the list"
        else:
            tooltip_text += "\n• Use 'Add Files' button to manually add files (drag and drop not available)"
        tooltip_text += "\n\nBackground colors indicate naming conflicts."
        ToolTip(self.files_canvas, tooltip_text)
        
        main_frame.rowconfigure(row, weight=1)
        row += 1
        
        # Rules section
        rules_frame = ttk.LabelFrame(main_frame, text="Renaming Rules", padding="10")
        rules_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        rules_frame.columnconfigure(0, weight=1)
        
        # Rules container with scrollbar
        rules_container = ttk.Frame(rules_frame)
        rules_container.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        rules_container.columnconfigure(0, weight=1)
        
        # Create a canvas and scrollbar for rules
        self.rules_canvas = tk.Canvas(rules_container, height=200)
        rules_scrollbar = ttk.Scrollbar(rules_container, orient="vertical", command=self.rules_canvas.yview)
        self.rules_scrollable_frame = ttk.Frame(self.rules_canvas)
        
        # Configure the scrollable frame to expand
        self.rules_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox("all"))
        )
        
        # Configure canvas to resize its content when canvas size changes
        self.rules_canvas.bind(
            "<Configure>",
            lambda e: self.rules_canvas.itemconfig(self.canvas_frame_id, width=e.width)
        )
        
        self.canvas_frame_id = self.rules_canvas.create_window((0, 0), window=self.rules_scrollable_frame, anchor="nw")
        self.rules_canvas.configure(yscrollcommand=rules_scrollbar.set)
        
        self.rules_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        rules_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Set up mouse wheel scrolling for rules
        self.setup_mouse_wheel_scrolling(self.rules_canvas, self.rules_scrollable_frame)
        
        rules_container.rowconfigure(0, weight=1)
        
        # Rule buttons
        rule_buttons = ttk.Frame(rules_frame)
        rule_buttons.grid(row=1, column=0, pady=5)
        add_rule_btn = ttk.Button(rule_buttons, text="Add Rule", command=self.add_rule)
        add_rule_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(add_rule_btn, "Add a new renaming rule. Rules define how tags in the naming pattern are replaced with values. Three types available:\n• CounterRule: increments with each file\n• ListRule: cycles through a list of values\n• BatchRule: increments with each batch operation")
        row += 1
        
        # Status area
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="5")
        status_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        status_frame.columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(status_frame, textvariable=self.status_text, 
                                     relief=tk.SUNKEN, anchor=tk.W, padding=(5, 3))
        self.status_label.grid(row=0, column=0, sticky=(tk.W, tk.E))
        row += 1
        
        # Settings buttons
        settings_frame = ttk.Frame(main_frame)
        settings_frame.grid(row=row, column=0, columnspan=3, pady=10)
        
        export_btn = ttk.Button(settings_frame, text="Export Settings", command=self.export_settings)
        export_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(export_btn, "Save current settings (folders, naming pattern, rules) to a JSON file that can be shared or backed up.")
        
        import_btn = ttk.Button(settings_frame, text="Import Settings", command=self.import_settings)
        import_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(import_btn, "Load settings from a previously exported JSON file, replacing current configuration.")
        
        save_btn = ttk.Button(settings_frame, text="Save Settings", command=self.save_settings)
        save_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(save_btn, "Save current settings to the default settings file. Settings are automatically loaded when the application starts.")
        
        # Set up trace callbacks to update button states when fields change
        self.source_folder.trace('w', lambda *args: self.update_path_labels())
        self.source_folder.trace('w', lambda *args: self.update_button_states())
        self.dest_folder.trace('w', lambda *args: self.update_path_labels())
        self.dest_folder.trace('w', lambda *args: self.update_button_states())
        self.dest_folder.trace('w', lambda *args: self.update_files_display())  # Update display when dest changes to check existing files
        self.naming_pattern.trace('w', lambda *args: self.update_files_display())
        self.naming_pattern.trace('w', lambda *args: self.update_rules_display())
        self.naming_pattern.trace('w', lambda *args: self.update_naming_pattern_label())
        self.naming_pattern.trace('w', lambda *args: self.update_latest_rename_label())  # Update preview when naming pattern changes
    
    def show_status(self, message, message_type="info"):
        """Show a status message in the status area"""
        if message_type == "error":
            prefix = "❌ Error: "
        elif message_type == "warning":
            prefix = "⚠️ Warning: "
        elif message_type == "success":
            prefix = "✅ Success: "
        else:
            prefix = "ℹ️ "
        
        self.status_text.set(prefix + message)
        
        # Auto-clear status after 10 seconds for non-error messages
        if message_type != "error":
            self.root.after(10000, lambda: self.status_text.set("Ready"))
    
    def is_tag_used_in_pattern(self, tag_name):
        """Check if a tag is used in the current naming pattern"""
        pattern = self.naming_pattern.get()
        tag = f"{{{tag_name}}}"
        return tag in pattern
    
    def update_button_states(self):
        """Update button enabled/disabled states based on current application state"""
        # Start Tracking button: enabled only if source folder is set and valid
        has_source = bool(self.source_folder.get() and os.path.exists(self.source_folder.get()))
        is_tracking = self.observer is not None
        
        if self.start_tracking_btn:
            self.start_tracking_btn.config(state='normal' if has_source and not is_tracking else 'disabled')
        
        # Stop Tracking button: enabled only if currently tracking
        if self.stop_tracking_btn:
            self.stop_tracking_btn.config(state='normal' if is_tracking else 'disabled')
        
        # Path fields and browse buttons: disabled while tracking
        if hasattr(self, 'source_entry'):
            self.source_entry.config(state='disabled' if is_tracking else 'normal')
        if hasattr(self, 'source_browse'):
            self.source_browse.config(state='disabled' if is_tracking else 'normal')
        if hasattr(self, 'source_create'):
            self.source_create.config(state='disabled' if is_tracking else 'normal')
        if hasattr(self, 'dest_entry'):
            self.dest_entry.config(state='disabled' if is_tracking else 'normal')
        if hasattr(self, 'dest_browse'):
            self.dest_browse.config(state='disabled' if is_tracking else 'normal')
        if hasattr(self, 'dest_create'):
            self.dest_create.config(state='disabled' if is_tracking else 'normal')
        
        # File formats field: disabled while tracking
        if hasattr(self, 'formats_entry'):
            self.formats_entry.config(state='disabled' if is_tracking else 'normal')
        
        # Copy & Rename Files button: enabled only if has tracked files, destination folder, and no conflicts
        has_dest = bool(self.dest_folder.get() and os.path.exists(self.dest_folder.get()))
        has_tracked_files = len(self.tracked_files) > 0
        has_conflicts = self.has_any_conflicts()
        
        if self.copy_rename_btn:
            self.copy_rename_btn.config(state='normal' if has_tracked_files and has_dest and not has_conflicts else 'disabled')
        
        # Clear Tracked Files button: enabled only if has tracked files
        if self.clear_tracked_btn:
            self.clear_tracked_btn.config(state='normal' if has_tracked_files else 'disabled')
    
    def update_file_count_label(self):
        """Update the file count label to show the number of tracked files"""
        if self.file_count_label:
            count = len(self.tracked_files)
            self.file_count_label.config(text=f"({count})")
    
    def update_latest_rename_label(self):
        """Update the latest rename label to show the preview of the most recently added file"""
        if hasattr(self, 'latest_rename_original') and hasattr(self, 'latest_rename_preview'):
            # Show preview of the last file in the tracked files list (most recently added)
            if self.tracked_files:
                last_index = len(self.tracked_files) - 1
                file_path = self.tracked_files[last_index]
                original_filename = os.path.basename(file_path)
                
                # Generate preview name
                preview_name = self.generate_filename_preview(last_index, len(self.tracked_files))
                file_ext = os.path.splitext(file_path)[1]
                preview_full_name = preview_name + file_ext
                
                # Update colored labels
                self.latest_rename_original.config(text=original_filename)
                self.latest_rename_preview.config(text=preview_full_name)
            elif self.latest_rename_info:
                # Fall back to showing the last actual rename operation if no files are tracked
                original_name, new_name = self.latest_rename_info
                self.latest_rename_original.config(text=original_name)
                self.latest_rename_preview.config(text=new_name)
            else:
                # Clear the labels
                self.latest_rename_original.config(text="")
                self.latest_rename_preview.config(text="")
    
    def browse_source(self):
        # Get initial directory - use current field value or script directory
        initial_dir = self.source_folder.get()
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = os.path.dirname(os.path.abspath(__file__))
        
        folder = filedialog.askdirectory(title="Select Source Folder", initialdir=initial_dir)
        if folder:
            self.source_folder.set(folder)
    
    def browse_dest(self):
        # Get initial directory - use current field value or script directory
        initial_dir = self.dest_folder.get()
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = os.path.dirname(os.path.abspath(__file__))
        
        folder = filedialog.askdirectory(title="Select Destination Folder", initialdir=initial_dir)
        if folder:
            self.dest_folder.set(folder)
    
    def create_source_folder(self):
        """Create the source folder and all necessary parent directories"""
        source_path = self.source_folder.get().strip()
        if not source_path:
            self.show_status("Source folder path is empty", "error")
            return
        
        try:
            os.makedirs(source_path, exist_ok=True)
            self.show_status(f"Created source folder: {source_path}", "success")
            self.update_path_labels()  # Update the display
        except Exception as e:
            self.show_status(f"Failed to create source folder: {str(e)}", "error")
    
    def create_dest_folder(self):
        """Create the destination folder and all necessary parent directories"""
        dest_path = self.dest_folder.get().strip()
        if not dest_path:
            self.show_status("Destination folder path is empty", "error")
            return
        
        try:
            os.makedirs(dest_path, exist_ok=True)
            self.show_status(f"Created destination folder: {dest_path}", "success")
            self.update_path_labels()  # Update the display
        except Exception as e:
            self.show_status(f"Failed to create destination folder: {str(e)}", "error")
    
    def start_tracking(self):
        """Start tracking files in the source folder"""
        if not self.source_folder.get():
            self.show_status("Please select a source folder", "error")
            return
        
        if not os.path.exists(self.source_folder.get()):
            self.show_status("Source folder does not exist", "error")
            return
        
        # Stop existing tracking
        self.stop_tracking()
        
        # Get the most recent file time for baseline
        self.last_file_time = self.get_most_recent_file_time()
        
        # Start file watcher
        self.observer = Observer()
        event_handler = FileWatcher(self)
        self.observer.schedule(event_handler, self.source_folder.get(), recursive=False)
        self.observer.start()
        
        self.update_button_states()
        self.show_status("File tracking started", "success")
    
    def stop_tracking(self):
        """Stop tracking files"""
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.observer = None
            self.update_button_states()
            self.show_status("File tracking stopped", "info")
    
    def get_most_recent_file_time(self):
        """Get the modification time of the most recent file in source folder"""
        try:
            files = [f for f in os.listdir(self.source_folder.get()) 
                    if os.path.isfile(os.path.join(self.source_folder.get(), f))]
            if not files:
                return time.time()
            
            most_recent = max(files, key=lambda f: os.path.getmtime(os.path.join(self.source_folder.get(), f)))
            return os.path.getmtime(os.path.join(self.source_folder.get(), most_recent))
        except:
            return time.time()
    
    def add_tracked_file(self, file_path):
        """Add a file to the tracked files list"""
        if not self.should_track_file(file_path):
            return
        
        # Check if file was created after we started tracking
        file_time = os.path.getmtime(file_path)
        if file_time <= self.last_file_time:
            return
        
        # Add to tracked files if not already present
        if file_path not in self.tracked_files:
            self.tracked_files.append(file_path)
            self.update_files_display()
            self.update_file_count_label()
            self.update_latest_rename_label()
    
    def generate_thumbnail(self, file_path, size=(150, 150), retry_count=0):
        """Generate a thumbnail for an image file"""
        cache_key = f"{file_path}_{size[0]}x{size[1]}"
        if cache_key in self.thumbnail_cache:
            return self.thumbnail_cache[cache_key]
        
        try:
            # Check if file exists and is accessible
            if not os.path.exists(file_path):
                return self.create_placeholder_thumbnail(size, "File not found")
            
            # Check if it's an image file
            file_ext = os.path.splitext(file_path)[1].lower()
            image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'}
            
            if file_ext not in image_extensions:
                # Not an image, return placeholder
                return self.create_placeholder_thumbnail(size, "Non-image")
            
            # Try to open and resize the image with retry logic
            max_retries = 2
            last_exception = None
            
            for attempt in range(max_retries + 1):
                try:
                    # Try to open the file
                    with Image.open(file_path) as img:
                        # Verify the image is valid by loading it
                        img.verify()
                    
                    # Reopen for processing (verify() invalidates the image)
                    with Image.open(file_path) as img:
                        # Convert to RGB if necessary (for PNG with transparency, etc.)
                        if img.mode not in ('RGB', 'RGBA'):
                            img = img.convert('RGB')
                        
                        # Calculate size to maintain aspect ratio
                        img.thumbnail(size, Image.Resampling.LANCZOS)
                        
                        # Convert to PhotoImage for tkinter
                        photo = ImageTk.PhotoImage(img)
                        
                        # Cache the thumbnail
                        self.thumbnail_cache[cache_key] = photo
                        return photo
                        
                except (IOError, OSError, PermissionError) as e:
                    last_exception = e
                    if attempt < max_retries:
                        # Wait a bit before retrying (file might still be being written)
                        time.sleep(0.5 * (attempt + 1))
                        continue
                    else:
                        break
                
        except Exception as e:
            last_exception = e
        
        # If we get here, thumbnail generation failed
        if retry_count == 0 and last_exception:
            # Schedule a retry after a delay for the first failure
            self.root.after(2000, lambda: self._retry_thumbnail_generation(file_path, size))
            return self.create_placeholder_thumbnail(size, "Loading...")
        
        # Failed to create thumbnail even after retry, return error placeholder
        error_msg = "Load failed"
        if isinstance(last_exception, PermissionError):
            error_msg = "Access denied"
        elif isinstance(last_exception, (IOError, OSError)):
            error_msg = "File error"
        
        return self.create_placeholder_thumbnail(size, error_msg)
    
    def _retry_thumbnail_generation(self, file_path, size):
        """Retry thumbnail generation after a delay"""
        try:
            # Remove any existing cache entry
            cache_key = f"{file_path}_{size[0]}x{size[1]}"
            if cache_key in self.thumbnail_cache:
                del self.thumbnail_cache[cache_key]
            
            # Try to generate thumbnail again (with retry_count=1 to prevent infinite recursion)
            self.generate_thumbnail(file_path, size, retry_count=1)
            
            # Update the display to show the new thumbnail
            self.update_files_display()
        except Exception:
            # Silently fail on retry
            pass
    
    def create_placeholder_thumbnail(self, size=(150, 150), message="No preview"):
        """Create a placeholder thumbnail for non-image files"""
        # Create a simple placeholder image
        img = Image.new('RGB', size, color='lightgray')
        
        # Add text to the placeholder
        try:
            from PIL import ImageDraw, ImageFont
            draw = ImageDraw.Draw(img)
            
            # Try to use a default font, fall back to basic if not available
            try:
                font_size = max(10, min(size[0] // 10, size[1] // 10))
                font = ImageFont.truetype("arial.ttf", font_size)
            except:
                try:
                    font = ImageFont.load_default()
                except:
                    font = None
            
            if font:
                # Calculate text position to center it
                bbox = draw.textbbox((0, 0), message, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                x = (size[0] - text_width) // 2
                y = (size[1] - text_height) // 2
                
                draw.text((x, y), message, fill='darkgray', font=font)
            else:
                # Fallback without font
                draw.text((10, size[1]//2), message, fill='darkgray')
                
        except ImportError:
            # PIL might not have ImageDraw/ImageFont, skip text
            pass
        
        photo = ImageTk.PhotoImage(img)
        return photo
    
    def should_track_file(self, file_path):
        """Check if a file should be tracked based on format filters"""
        file_name = os.path.basename(file_path)
        formats = [f.strip() for f in self.file_formats.get().split(';') if f.strip()]
        
        if not formats or '*' in formats:
            return True
        
        file_ext = os.path.splitext(file_name)[1].lower()
        for fmt in formats:
            if fmt.startswith('.'):
                if file_ext == fmt.lower():
                    return True
            elif fmt.startswith('*.'):
                if file_ext == fmt[1:].lower():
                    return True
            elif fmt == os.path.splitext(file_name)[0]:
                return True
        
        return False
    
    def update_files_display(self):
        """Update the files display with preview names and controls"""
        # Detect what changed compared to last state
        changes = self._detect_file_changes()
        
        # Check if view mode changed to preserve scroll position
        view_mode_changed = changes.get('view_mode_changed', False)
        
        # Preserve scroll position if view mode is changing
        saved_scroll_position = None
        if view_mode_changed:
            saved_scroll_position = self._get_relative_scroll_position(self.files_canvas)
        
        # Apply incremental updates
        if changes['full_rebuild_needed']:
            self._full_rebuild_files()
        else:
            self._incremental_update_files(changes)
        
        # Restore relative scroll position if view mode changed
        if view_mode_changed and saved_scroll_position is not None:
            self._restore_relative_scroll_position(self.files_canvas, saved_scroll_position)
        
        # Update last known state
        self.last_files_state = self._get_current_files_state()
        
        self.update_button_states()
        self.update_latest_rename_label()  # Update preview when display changes (rules/pattern changes)
    
    def show_file_conflict_dialog(self, existing_files):
        """Show dialog for handling file conflicts with multiple options"""
        dialog = tk.Toplevel(self.root)
        dialog.title("File Conflicts")
        dialog.geometry("600x550")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        result = {"action": "cancel"}
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="File Conflicts Detected", font=("Arial", 12, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Message
        message_text = f"The following {len(existing_files)} file(s) already exist in the destination folder:"
        message_label = ttk.Label(main_frame, text=message_text, wraplength=550)
        message_label.pack(pady=(0, 10))
        
        # Files list
        files_frame = ttk.Frame(main_frame)
        files_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Create scrollable text widget for file list
        files_text = tk.Text(files_frame, height=8, wrap=tk.WORD, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, command=files_text.yview)
        files_text.configure(yscrollcommand=scrollbar.set)
        
        files_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Populate file list
        files_text.config(state=tk.NORMAL)
        for filename in existing_files:
            files_text.insert(tk.END, f"• {filename}\n")
        files_text.config(state=tk.DISABLED)
        
        # Options explanation
        explanation_label = ttk.Label(main_frame, 
                                    text="Choose how to handle these conflicts:",
                                    font=("Arial", 10, "bold"))
        explanation_label.pack(pady=(10, 5))
        
        # Options frame
        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill=tk.X, pady=(0, 15))
        
        def set_action_and_close(action):
            result["action"] = action
            dialog.destroy()
        
        # Buttons with explanations
        btn_frame1 = ttk.Frame(options_frame)
        btn_frame1.pack(fill=tk.X, pady=2)
        overwrite_btn = ttk.Button(btn_frame1, text="Overwrite", width=12,
                                  command=lambda: set_action_and_close("overwrite"))
        overwrite_btn.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(btn_frame1, text="Replace existing files with new ones").pack(side=tk.LEFT)
        
        btn_frame2 = ttk.Frame(options_frame)
        btn_frame2.pack(fill=tk.X, pady=2)
        rename_btn = ttk.Button(btn_frame2, text="Rename", width=12,
                               command=lambda: set_action_and_close("rename"))
        rename_btn.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(btn_frame2, text="Add numbers to new files (e.g., file_1.jpg, file_2.jpg)").pack(side=tk.LEFT)
        
        btn_frame3 = ttk.Frame(options_frame)
        btn_frame3.pack(fill=tk.X, pady=2)
        ignore_btn = ttk.Button(btn_frame3, text="Ignore", width=12,
                               command=lambda: set_action_and_close("ignore"))
        ignore_btn.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(btn_frame3, text="Skip existing files, only copy new ones").pack(side=tk.LEFT)
        
        btn_frame4 = ttk.Frame(options_frame)
        btn_frame4.pack(fill=tk.X, pady=2)
        cancel_btn = ttk.Button(btn_frame4, text="Cancel", width=12,
                               command=lambda: set_action_and_close("cancel"))
        cancel_btn.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(btn_frame4, text="Cancel the entire copy operation").pack(side=tk.LEFT)
        
        # Make Overwrite the default (focused) button
        overwrite_btn.focus()
        
        # Handle window close button
        dialog.protocol("WM_DELETE_WINDOW", lambda: set_action_and_close("cancel"))
        
        # Wait for dialog to close
        dialog.wait_window()
        
        return result["action"]
    
    def _get_current_files_state(self):
        """Get current state of files for change detection"""
        state = []
        for i, file_path in enumerate(self.tracked_files):
            preview_name = self.generate_filename_preview(i, len(self.tracked_files))
            file_ext = os.path.splitext(file_path)[1]
            preview_full_name = preview_name + file_ext
            
            state.append({
                'file_path': file_path,
                'preview_full_name': preview_full_name,
                'view_mode': self.view_mode.get()
            })
        return state
    
    def _detect_file_changes(self):
        """Detect what changed in the files list"""
        current_state = self._get_current_files_state()
        last_state = self.last_files_state
        
        # Check if view mode changed (requires full rebuild)
        view_mode_changed = False
        if last_state and current_state and len(last_state) > 0 and len(current_state) > 0:
            view_mode_changed = last_state[0].get('view_mode') != current_state[0].get('view_mode')
        
        # Check if widget count is completely wrong (requires full rebuild)
        widget_count_mismatch = len(self.file_widgets) > len(self.tracked_files)
        
        if view_mode_changed or widget_count_mismatch:
            return {'full_rebuild_needed': True, 'view_mode_changed': view_mode_changed}
        
        # Handle file additions (incremental)
        files_added = len(current_state) > len(last_state)
        files_removed = len(current_state) < len(last_state)
        
        if files_removed:
            # Files were removed - need full rebuild to handle index changes
            return {'full_rebuild_needed': True, 'view_mode_changed': False}
        
        # Find specific changes
        updated_indices = []
        new_indices = []
        
        # Check existing files for changes
        for i in range(min(len(current_state), len(last_state))):
            if current_state[i] != last_state[i]:
                updated_indices.append(i)
        
        # Handle new files
        if files_added:
            for i in range(len(last_state), len(current_state)):
                new_indices.append(i)
            
            # Also update the previously last item's button state (if it exists)
            if len(last_state) > 0:
                prev_last_index = len(last_state) - 1
                if prev_last_index not in updated_indices:
                    updated_indices.append(prev_last_index)
        
        return {
            'full_rebuild_needed': False,
            'view_mode_changed': False,
            'updated_indices': updated_indices,
            'new_indices': new_indices
        }
    
    def _full_rebuild_files(self):
        """Perform full rebuild of files display"""
        # Clear existing widgets
        for widget_data in self.file_widgets.values():
            if 'frame' in widget_data:
                widget_data['frame'].destroy()
        self.file_widgets.clear()
        
        # Clear the container
        for widget in self.files_scrollable_frame.winfo_children():
            widget.destroy()
        
        # Reset grid configuration to ensure clean slate
        for i in range(10):  # Clear potential old column configurations
            try:
                self.files_scrollable_frame.columnconfigure(i, weight=0, uniform="")
            except:
                break
        
        # Render based on current view mode
        if self.view_mode.get() == "grid":
            self.create_grid_view()
        else:
            self.create_list_view()
    
    def _incremental_update_files(self, changes):
        """Perform incremental update of files display"""
        # Update existing changed items
        for index in changes['updated_indices']:
            if index in self.file_widgets:
                if self.view_mode.get() == "grid":
                    self._update_grid_item(index)
                else:
                    self._update_list_item(index)
        
        # Add new items (process in reverse order for better performance)
        new_indices = changes['new_indices']
        for index in reversed(new_indices):
            file_path = self.tracked_files[index]
            if self.view_mode.get() == "grid":
                # Calculate grid position for new item
                cols = max(1, min(4, len(self.tracked_files)))
                row = index // cols
                col = index % cols
                max_width = 200  # Use a reasonable default width
                self.add_file_to_grid_display(file_path, index, row, col, max_width)
            else:
                self.add_file_to_list_display(file_path, index)
        
        # Auto-scroll to show the newest file if any were added
        if new_indices:
            self.scroll_to_newest_file(max(new_indices))
    
    def scroll_to_newest_file(self, file_index):
        """Scroll the files view to show the newly added file"""
        # Schedule the scroll after the UI has updated
        self.root.after(100, lambda: self._perform_scroll_to_file(file_index))
    
    def _perform_scroll_to_file(self, file_index):
        """Actually perform the scroll to show the specified file"""
        try:
            # Update the scroll region first
            self.files_canvas.configure(scrollregion=self.files_canvas.bbox("all"))
            
            # Check if we have a widget for this file index
            if file_index in self.file_widgets:
                widget_data = self.file_widgets[file_index]
                
                # Get the main widget for this file
                main_widget = None
                if 'frame' in widget_data:
                    main_widget = widget_data['frame']
                elif 'cell_frame' in widget_data:
                    main_widget = widget_data['cell_frame']
                
                if main_widget and main_widget.winfo_exists():
                    # Get the widget's position relative to the scrollable frame
                    widget_y = main_widget.winfo_y()
                    widget_height = main_widget.winfo_height()
                    
                    # Get the scrollable frame's total height
                    frame_height = self.files_scrollable_frame.winfo_reqheight()
                    
                    # Get the canvas viewport height
                    canvas_height = self.files_canvas.winfo_height()
                    
                    if frame_height > canvas_height:
                        # Calculate the position to scroll to (show the widget at the bottom of the view)
                        # This ensures the new file is visible
                        target_y = widget_y + widget_height - canvas_height + 20  # 20px padding
                        target_y = max(0, target_y)  # Don't scroll past the top
                        
                        # Convert to fraction of total scrollable area
                        scroll_fraction = target_y / (frame_height - canvas_height)
                        scroll_fraction = min(1.0, max(0.0, scroll_fraction))  # Clamp to [0, 1]
                        
                        # Scroll to show the new file
                        self.files_canvas.yview_moveto(scroll_fraction)
        except Exception as e:
            # Silently handle any scrolling errors
            pass
    
    def _update_list_item(self, index):
        """Update a specific list item without rebuilding"""
        if index not in self.file_widgets:
            return
        
        widget_data = self.file_widgets[index]
        file_path = self.tracked_files[index]
        
        # Update preview name
        preview_name = self.generate_filename_preview(index, len(self.tracked_files))
        file_ext = os.path.splitext(file_path)[1]
        preview_full_name = preview_name + file_ext
        original_name = os.path.basename(file_path)
        display_text = f"{original_name} → {preview_full_name}"
        
        if 'file_label' in widget_data:
            widget_data['file_label'].config(text=display_text)
        
        # Update background color based on conflicts
        has_duplicate = self.has_duplicate_preview_name(index, preview_full_name)
        exists_in_dest = self.preview_exists_in_destination(preview_full_name)
        
        if has_duplicate:
            bg_color = "#ffcccc"
            style_name = f"FileDuplicate{index}.TFrame"
        elif exists_in_dest:
            bg_color = "#ccccff"
            style_name = f"FileExists{index}.TFrame"
        else:
            bg_color = "#f0f0f0" if index % 2 == 0 else "#ffffff"
            style_name = f"File{index % 2}.TFrame"
        
        if 'frame' in widget_data:
            widget_data['frame'].configure(style=style_name)
            style = ttk.Style()
            style.configure(style_name, background=bg_color)
        
        if 'file_label' in widget_data:
            widget_data['file_label'].config(background=bg_color)
        
        # Update button states
        if 'up_btn' in widget_data:
            widget_data['up_btn'].config(state='normal' if index > 0 else 'disabled')
        if 'down_btn' in widget_data:
            widget_data['down_btn'].config(state='normal' if index < len(self.tracked_files) - 1 else 'disabled')
    
    def _update_grid_item(self, index):
        """Update a specific grid item without rebuilding"""
        if index not in self.file_widgets:
            return
        
        widget_data = self.file_widgets[index]
        file_path = self.tracked_files[index]
        
        # Update preview name
        preview_name = self.generate_filename_preview(index, len(self.tracked_files))
        file_ext = os.path.splitext(file_path)[1]
        preview_full_name = preview_name + file_ext
        original_name = os.path.basename(file_path)
        
        if 'orig_label' in widget_data:
            widget_data['orig_label'].config(text=original_name)
        if 'prev_label' in widget_data:
            widget_data['prev_label'].config(text=preview_full_name)
        
        # Update conflict states
        has_duplicate = self.has_duplicate_preview_name(index, preview_full_name)
        exists_in_dest = self.preview_exists_in_destination(preview_full_name)
        
        # Determine background color
        if has_duplicate:
            bg_color = "#ffcccc"
        elif exists_in_dest:
            bg_color = "#ccccff"
        else:
            bg_color = "#f0f0f0"
        
        if 'cell_frame' in widget_data:
            if has_duplicate:
                widget_data['cell_frame'].configure(relief='solid', borderwidth=2, style="Duplicate.TFrame")
            elif exists_in_dest:
                widget_data['cell_frame'].configure(relief='solid', borderwidth=2, style="Exists.TFrame")
            else:
                widget_data['cell_frame'].configure(relief='ridge', borderwidth=2)
        
        # Update spacer label background color to match
        if 'spacer_label' in widget_data:
            widget_data['spacer_label'].config(background=bg_color)
        
        # Update button states
        if 'left_btn' in widget_data:
            widget_data['left_btn'].config(state='normal' if index > 0 else 'disabled')
        if 'right_btn' in widget_data:
            widget_data['right_btn'].config(state='normal' if index < len(self.tracked_files) - 1 else 'disabled')
    
    def create_list_view(self):
        """Create the list view of tracked files"""
        for i, file_path in enumerate(self.tracked_files):
            self.add_file_to_list_display(file_path, i)
    
    def create_grid_view(self):
        """Create the grid view of tracked files"""
        if not self.tracked_files:
            return
        
        # Calculate grid dimensions and cell sizes
        num_files = len(self.tracked_files)
        cols = max(1, min(4, num_files))  # Maximum 4 columns
        
        # Calculate maximum cell width based on file names
        max_width = 200  # Minimum width
        for i, file_path in enumerate(self.tracked_files):
            original_name = os.path.basename(file_path)
            preview_name = self.generate_filename_preview(i, len(self.tracked_files))
            file_ext = os.path.splitext(file_path)[1]
            preview_full_name = preview_name + file_ext
            
            # Calculate text width (approximate)
            name_width = max(len(original_name), len(preview_full_name)) * 8 + 20
            max_width = max(max_width, name_width)
        
        max_width = min(max_width, 300)  # Maximum width limit
        
        # Configure all grid columns at once for consistent sizing
        for col in range(cols):
            self.files_scrollable_frame.columnconfigure(col, weight=1, uniform="grid_column")
        
        # Create grid cells
        for i, file_path in enumerate(self.tracked_files):
            row = i // cols
            col = i % cols
            self.add_file_to_grid_display(file_path, i, row, col, max_width)
    
    def add_file_to_list_display(self, file_path, index):
        """Add a single file entry to the files display"""
        # Create a frame for the file entry with alternating background colors
        file_frame = ttk.Frame(self.files_scrollable_frame)
        file_frame.pack(fill=tk.X, padx=5, pady=1)
        
        # Configure the scrollable frame to expand its children
        self.files_scrollable_frame.columnconfigure(0, weight=1)
        
        # Check for conflicts and determine background color
        preview_name = self.generate_filename_preview(index, len(self.tracked_files))
        file_ext = os.path.splitext(file_path)[1]
        preview_full_name = preview_name + file_ext
        
        has_duplicate = self.has_duplicate_preview_name(index, preview_full_name)
        exists_in_dest = self.preview_exists_in_destination(preview_full_name)
        
        # Determine background color based on conflicts or alternating pattern
        if has_duplicate:
            bg_color = "#ffcccc"  # Light red for duplicates
            style_name = f"FileDuplicate{index}.TFrame"
        elif exists_in_dest:
            bg_color = "#ccccff"  # Light blue for existing files
            style_name = f"FileExists{index}.TFrame"
        else:
            bg_color = "#f0f0f0" if index % 2 == 0 else "#ffffff"
            style_name = f"File{index % 2}.TFrame"
        
        file_frame.configure(style=style_name)
        
        # Configure styles using the bg_color variable
        style = ttk.Style()
        style.configure(style_name, background=bg_color)
        
        col = 0
        
        # Up button
        up_btn = ttk.Button(file_frame, text="↑", width=3,
                           command=lambda: self.move_file_up(index))
        up_btn.grid(row=0, column=col, padx=(0, 2), sticky=tk.W)
        up_btn.config(state='normal' if index > 0 else 'disabled')
        ToolTip(up_btn, "Move this file up in the list. Files are processed in the order shown, so this affects the numbering in counter rules.")
        col += 1
        
        # Down button
        down_btn = ttk.Button(file_frame, text="↓", width=3,
                             command=lambda: self.move_file_down(index))
        down_btn.grid(row=0, column=col, padx=(0, 5), sticky=tk.W)
        down_btn.config(state='normal' if index < len(self.tracked_files) - 1 else 'disabled')
        ToolTip(down_btn, "Move this file down in the list. Files are processed in the order shown, so this affects the numbering in counter rules.")
        col += 1
        
        # File preview text
        original_name = os.path.basename(file_path)
        
        # Display format: "original_name → preview_name"
        display_text = f"{original_name} → {preview_full_name}"
        
        file_label = ttk.Label(file_frame, text=display_text, background=bg_color)
        file_label.grid(row=0, column=col, padx=(5, 5), sticky=(tk.W, tk.E))
        file_frame.columnconfigure(col, weight=1)  # Make this column expandable
        col += 1
        
        # Remove button
        remove_btn = ttk.Button(file_frame, text="✕", width=3,
                               command=lambda: self.remove_file_at_index(index))
        remove_btn.grid(row=0, column=col, padx=(5, 0), sticky=tk.E)
        ToolTip(remove_btn, "Remove this file from the tracked files list. The original file is not affected, only removed from processing.")
        
        # Store widget references for incremental updates
        self.file_widgets[index] = {
            'frame': file_frame,
            'up_btn': up_btn,
            'down_btn': down_btn,
            'file_label': file_label,
            'remove_btn': remove_btn
        }
        
        # Bind scroll events to the new widgets immediately
        self._bind_scroll_to_new_widget(file_frame)
    
    def add_file_to_grid_display(self, file_path, index, row, col, cell_width):
        """Add a file to the grid display"""
        # Create a frame for the grid cell
        cell_frame = ttk.Frame(self.files_scrollable_frame, borderwidth=2, relief='ridge')
        cell_frame.grid(row=row, column=col, padx=5, pady=5, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        # Configure cell frame with consistent dimensions
        cell_frame.configure(width=cell_width, height=250)
        cell_frame.grid_propagate(False)  # Prevent frame from shrinking
        
        # Get file info
        original_name = os.path.basename(file_path)
        preview_name = self.generate_filename_preview(index, len(self.tracked_files))
        file_ext = os.path.splitext(file_path)[1]
        preview_full_name = preview_name + file_ext
        
        # Check conflicts
        has_duplicate = self.has_duplicate_preview_name(index, preview_full_name)
        exists_in_dest = self.preview_exists_in_destination(preview_full_name)
        
        # Determine background color and border style based on conflicts
        if has_duplicate:
            bg_color = "#ffcccc"
            cell_frame.configure(relief='solid', borderwidth=2)
            cell_frame.configure(style=f"Duplicate.TFrame")
        elif exists_in_dest:
            bg_color = "#ccccff"
            cell_frame.configure(relief='solid', borderwidth=2)
            cell_frame.configure(style=f"Exists.TFrame")
        else:
            bg_color = "#f0f0f0"  # Default background color
        
        # Configure styles
        style = ttk.Style()
        style.configure("Duplicate.TFrame", background=bg_color, borderwidth=2)
        style.configure("Exists.TFrame", background=bg_color, borderwidth=2)
        
        grid_row = 0
        
        # Thumbnail
        thumbnail_size = (cell_width - 20, 120)  # Leave some margin
        thumbnail = self.generate_thumbnail(file_path, thumbnail_size)
        
        # Create thumbnail label
        thumb_label = tk.Label(cell_frame, image=thumbnail)
        thumb_label.grid(row=grid_row, column=0, padx=5, pady=5)
        thumb_label.image = thumbnail  # Keep a reference to prevent garbage collection
        
        ToolTip(thumb_label, "File thumbnail preview")
        grid_row += 1
        
        # Original name (red text)
        orig_label = tk.Label(cell_frame, text=original_name, fg="red", 
                             wraplength=cell_width-10, justify=tk.CENTER, font=("Arial", 8))
        orig_label.grid(row=grid_row, column=0, padx=5, pady=2, sticky=(tk.W, tk.E))
        grid_row += 1
        
        # Preview name (green text)
        prev_label = tk.Label(cell_frame, text=preview_full_name, fg="green",
                             wraplength=cell_width-10, justify=tk.CENTER, font=("Arial", 8))
        prev_label.grid(row=grid_row, column=0, padx=5, pady=2, sticky=(tk.W, tk.E))
        grid_row += 1
        
        # Add a spacer row to push buttons to bottom
        spacer_label = tk.Label(cell_frame, text="", background=bg_color)
        spacer_label.grid(row=grid_row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        cell_frame.rowconfigure(grid_row, weight=1)  # Make spacer row expandable
        grid_row += 1
        
        # Button row at the bottom with left/right arrows and delete button
        button_frame = ttk.Frame(cell_frame)
        button_frame.grid(row=grid_row, column=0, sticky=(tk.W, tk.E, tk.S), padx=5, pady=5)
        button_frame.columnconfigure(1, weight=1)  # Spacer in middle
        
        # Left button (move up in list)
        left_btn = ttk.Button(button_frame, text="←", width=3,
                             command=lambda: self.move_file_up(index))
        left_btn.grid(row=0, column=0, sticky=tk.W)
        left_btn.config(state='normal' if index > 0 else 'disabled')
        ToolTip(left_btn, "Move this file up in the list (left in grid = up in list order)")
        
        # Remove button in the center
        remove_btn = ttk.Button(button_frame, text="✕", width=3,
                               command=lambda: self.remove_file_at_index(index))
        remove_btn.grid(row=0, column=1)
        ToolTip(remove_btn, "Remove this file from the tracked files list")
        
        # Right button (move down in list)
        right_btn = ttk.Button(button_frame, text="→", width=3,
                              command=lambda: self.move_file_down(index))
        right_btn.grid(row=0, column=2, sticky=tk.E)
        right_btn.config(state='normal' if index < len(self.tracked_files) - 1 else 'disabled')
        ToolTip(right_btn, "Move this file down in the list (right in grid = down in list order)")
        
        # Configure cell frame columns to expand
        cell_frame.columnconfigure(0, weight=1)
        
        # Store widget references for incremental updates
        self.file_widgets[index] = {
            'cell_frame': cell_frame,
            'left_btn': left_btn,
            'right_btn': right_btn,
            'thumb_label': thumb_label,
            'orig_label': orig_label,
            'prev_label': prev_label,
            'spacer_label': spacer_label,
            'remove_btn': remove_btn
        }
        
        # Bind scroll events to the new widgets immediately
        self._bind_scroll_to_new_widget(cell_frame)
    
    def move_file_up(self, index):
        """Move a file up in the list"""
        if index > 0:
            # Swap files
            self.tracked_files[index], self.tracked_files[index - 1] = \
                self.tracked_files[index - 1], self.tracked_files[index]
            
            # Force a complete rebuild by calling the rebuild method directly
            self._full_rebuild_files()
            
            # Update the state tracking after rebuild
            self.last_files_state = self._get_current_files_state()
    
    def move_file_down(self, index):
        """Move a file down in the list"""
        if index < len(self.tracked_files) - 1:
            # Swap files
            self.tracked_files[index], self.tracked_files[index + 1] = \
                self.tracked_files[index + 1], self.tracked_files[index]
            
            # Force a complete rebuild by calling the rebuild method directly
            self._full_rebuild_files()
            
            # Update the state tracking after rebuild
            self.last_files_state = self._get_current_files_state()
    
    def remove_file_at_index(self, index):
        """Remove a file at the specified index"""
        if 0 <= index < len(self.tracked_files):
            del self.tracked_files[index]
            # Clear widget tracking to force rebuild on removal
            self.file_widgets.clear()
            self.update_files_display()
            self.update_file_count_label()
            self.update_latest_rename_label()
    
    def has_duplicate_preview_name(self, current_index, preview_full_name):
        """Check if the preview name is duplicated among other tracked files"""
        for i, file_path in enumerate(self.tracked_files):
            if i != current_index:
                other_preview_name = self.generate_filename_preview(i, len(self.tracked_files))
                other_file_ext = os.path.splitext(file_path)[1]
                other_preview_full_name = other_preview_name + other_file_ext
                if preview_full_name == other_preview_full_name:
                    return True
        return False
    
    def preview_exists_in_destination(self, preview_full_name):
        """Check if the preview filename already exists in the destination folder"""
        if not self.dest_folder.get() or not os.path.exists(self.dest_folder.get()):
            return False
        
        dest_path = os.path.join(self.dest_folder.get(), preview_full_name)
        return os.path.exists(dest_path)
    
    def has_any_conflicts(self):
        """Check if there are any blocking naming conflicts (only duplicates now)"""
        if not self.tracked_files:
            return False
        
        preview_names = []
        for i, file_path in enumerate(self.tracked_files):
            preview_name = self.generate_filename_preview(i, len(self.tracked_files))
            file_ext = os.path.splitext(file_path)[1]
            preview_full_name = preview_name + file_ext
            
            # Check for duplicates (blocking conflict)
            if preview_full_name in preview_names:
                return True
            preview_names.append(preview_full_name)
        
        return False
    
    def has_existing_files_in_destination(self):
        """Check if any files would overwrite existing files in destination"""
        if not self.tracked_files:
            return False
        
        for i, file_path in enumerate(self.tracked_files):
            preview_name = self.generate_filename_preview(i, len(self.tracked_files))
            file_ext = os.path.splitext(file_path)[1]
            preview_full_name = preview_name + file_ext
            
            if self.preview_exists_in_destination(preview_full_name):
                return True
        
        return False
    
    def get_missing_rule_tags(self):
        """Get list of tags used in naming pattern that don't have corresponding rules"""
        import re
        pattern = self.naming_pattern.get()
        # Find all tags in the pattern using regex
        tags_in_pattern = re.findall(r'\{([^}]+)\}', pattern)
        existing_rule_tags = {rule.tag_name for rule in self.rules}
        missing_tags = [tag for tag in tags_in_pattern if tag not in existing_rule_tags]
        return missing_tags
    
    def update_path_labels(self):
        """Update path labels to show error symbols for invalid paths"""
        # Check source folder
        source_path = self.source_folder.get()
        is_tracking = self.observer is not None
        
        if source_path and not os.path.exists(source_path):
            self.source_label.config(text="Source Folder: ❌")
            self.source_label.config(foreground="red")
            # Show create button if not tracking and path is not empty
            if not is_tracking and hasattr(self, 'source_create'):
                self.source_create.grid()
        elif source_path and os.path.exists(source_path):
            self.source_label.config(text="Source Folder: ✅")
            self.source_label.config(foreground="green")
            # Hide create button when path exists
            if hasattr(self, 'source_create'):
                self.source_create.grid_remove()
        else:
            self.source_label.config(text="Source Folder:")
            self.source_label.config(foreground="black")
            # Hide create button when path is empty
            if hasattr(self, 'source_create'):
                self.source_create.grid_remove()
        
        # Check destination folder
        dest_path = self.dest_folder.get()
        if dest_path and not os.path.exists(dest_path):
            self.dest_label.config(text="Destination Folder: ❌")
            self.dest_label.config(foreground="red")
            # Show create button if not tracking and path is not empty
            if not is_tracking and hasattr(self, 'dest_create'):
                self.dest_create.grid()
        elif dest_path and os.path.exists(dest_path):
            self.dest_label.config(text="Destination Folder: ✅")
            self.dest_label.config(foreground="green")
            # Hide create button when path exists
            if hasattr(self, 'dest_create'):
                self.dest_create.grid_remove()
        else:
            self.dest_label.config(text="Destination Folder:")
            self.dest_label.config(foreground="black")
            # Hide create button when path is empty
            if hasattr(self, 'dest_create'):
                self.dest_create.grid_remove()
    
    def update_naming_pattern_label(self):
        """Update the naming pattern label to show warning if rules are missing"""
        missing_tags = self.get_missing_rule_tags()
        if missing_tags:
            self.naming_pattern_label.config(text="Naming Pattern: ⚠️")
        else:
            self.naming_pattern_label.config(text="Naming Pattern:")
    
    def clear_tracked(self):
        """Clear the tracked files list"""
        self.tracked_files.clear()
        # Clear widget tracking
        self.file_widgets.clear()
        self.update_files_display()
        self.update_file_count_label()
        self.update_latest_rename_label()
    
    def add_files_manually(self):
        """Open file dialog to manually select files to add"""
        # Get initial directory - use source folder if set, otherwise script directory
        initial_dir = self.source_folder.get()
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Open file selection dialog allowing multiple files
        file_paths = filedialog.askopenfilenames(
            title="Select Files to Add",
            initialdir=initial_dir,
            filetypes=[
                ("All files", "*.*"),
                ("Image files", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.tiff;*.webp"),
                ("Text files", "*.txt;*.doc;*.docx"),
                ("Video files", "*.mp4;*.avi;*.mov;*.wmv;*.flv;*.mkv")
            ]
        )
        
        if file_paths:
            added_count = 0
            skipped_count = 0
            
            for file_path in file_paths:
                if self.should_track_file(file_path):
                    if file_path not in self.tracked_files:
                        self.tracked_files.append(file_path)
                        added_count += 1
                    else:
                        skipped_count += 1
                else:
                    skipped_count += 1
            
            if added_count > 0:
                self.update_files_display()
                self.update_file_count_label()
                self.update_latest_rename_label()
                if skipped_count > 0:
                    self.show_status(f"Added {added_count} files, skipped {skipped_count} (already tracked or format filter)", "success")
                else:
                    self.show_status(f"Added {added_count} files", "success")
            else:
                if skipped_count > 0:
                    self.show_status("No files added - all were already tracked or don't match format filter", "warning")
                else:
                    self.show_status("No files selected", "info")
    
    def setup_mouse_wheel_scrolling(self, canvas, scrollable_frame=None):
        """Set up mouse wheel scrolling for a canvas widget"""
        def on_mousewheel(event):
            # Check if there's content to scroll
            if canvas.winfo_exists():
                # Get current scroll region
                scroll_region = canvas.cget("scrollregion")
                if scroll_region and scroll_region != "0 0 0 0":
                    # Calculate scroll amount (negative for natural scrolling)
                    delta = -1 * (event.delta / 120) if hasattr(event, 'delta') else -1 * event.num
                    canvas.yview_scroll(int(delta), "units")
        
        def on_shift_mousewheel(event):
            # Horizontal scrolling with Shift+wheel (if needed in future)
            if canvas.winfo_exists():
                scroll_region = canvas.cget("scrollregion")
                if scroll_region and scroll_region != "0 0 0 0":
                    delta = -1 * (event.delta / 120) if hasattr(event, 'delta') else -1 * event.num
                    canvas.xview_scroll(int(delta), "units")
        
        # Bind mouse wheel events for different platforms to the canvas
        # Windows and MacOS
        canvas.bind("<MouseWheel>", on_mousewheel)
        canvas.bind("<Shift-MouseWheel>", on_shift_mousewheel)
        
        # Linux
        canvas.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))
        canvas.bind("<Shift-Button-4>", lambda e: canvas.xview_scroll(-1, "units"))
        canvas.bind("<Shift-Button-5>", lambda e: canvas.xview_scroll(1, "units"))
        
        # Store the canvas reference and scroll functions for use in child widget binding
        scroll_data = {
            'canvas': canvas,
            'on_mousewheel': on_mousewheel,
            'on_shift_mousewheel': on_shift_mousewheel
        }
        
        # Also bind to the provided scrollable frame for better UX
        if scrollable_frame:
            # Store scroll data as an attribute of the scrollable frame for easy access
            scrollable_frame._scroll_data = scroll_data
            
            # Bind to the scrollable frame directly
            self._bind_scroll_events(scrollable_frame, scroll_data)
            
            # Set up a mechanism to automatically bind scroll events to new child widgets
            self._setup_auto_scroll_binding(scrollable_frame, scroll_data)
    
    def _bind_scroll_events(self, widget, scroll_data):
        """Bind scroll events to a specific widget"""
        canvas = scroll_data['canvas']
        on_mousewheel = scroll_data['on_mousewheel']
        on_shift_mousewheel = scroll_data['on_shift_mousewheel']
        
        try:
            # Skip binding to Entry and Text widgets to preserve their normal behavior
            if isinstance(widget, (tk.Entry, ttk.Entry, tk.Text)):
                return
            
            # Check if widget exists and can be bound to
            if not widget.winfo_exists():
                return
            
            # Bind mouse wheel events
            widget.bind("<MouseWheel>", on_mousewheel)
            widget.bind("<Shift-MouseWheel>", on_shift_mousewheel)
            
            # Linux scroll events
            widget.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
            widget.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))
            widget.bind("<Shift-Button-4>", lambda e: canvas.xview_scroll(-1, "units"))
            widget.bind("<Shift-Button-5>", lambda e: canvas.xview_scroll(1, "units"))
            
        except Exception:
            # Silently handle any binding errors
            pass
    
    def _setup_auto_scroll_binding(self, scrollable_frame, scroll_data):
        """Set up automatic binding of scroll events to new child widgets"""
        def bind_to_all_children():
            """Recursively bind scroll events to all current children"""
            try:
                self._recursive_bind_children(scrollable_frame, scroll_data)
            except Exception:
                pass
        
        # Initial binding after a short delay
        self.root.after(100, bind_to_all_children)
        
        # Set up periodic re-binding to catch new widgets
        def periodic_rebind():
            bind_to_all_children()
            # Schedule next rebind
            self.root.after(1000, periodic_rebind)  # Check every second
        
        # Start periodic rebinding
        self.root.after(1000, periodic_rebind)
    
    def _recursive_bind_children(self, widget, scroll_data):
        """Recursively bind scroll events to all child widgets"""
        try:
            # First bind to the widget itself
            self._bind_scroll_events(widget, scroll_data)
            
            # Get all children of this widget
            children = widget.winfo_children()
            
            for child in children:
                # Bind scroll events to this child
                self._bind_scroll_events(child, scroll_data)
                
                # Recursively process grandchildren
                self._recursive_bind_children(child, scroll_data)
                
        except Exception:
            # Silently handle any errors during recursive binding
            pass
    
    def _bind_scroll_to_new_widget(self, widget):
        """Bind scroll events to a newly created widget and its children"""
        # Determine which scroll data to use based on the widget's ancestry
        scroll_data = None
        
        # Check if this widget is a descendant of the files scrollable frame
        if hasattr(self, 'files_scrollable_frame') and self._is_descendant_of(widget, self.files_scrollable_frame):
            scroll_data = getattr(self.files_scrollable_frame, '_scroll_data', None)
        
        # Check if this widget is a descendant of the rules scrollable frame
        elif hasattr(self, 'rules_scrollable_frame') and self._is_descendant_of(widget, self.rules_scrollable_frame):
            scroll_data = getattr(self.rules_scrollable_frame, '_scroll_data', None)
        
        # If we found scroll data, bind the events
        if scroll_data:
            self._recursive_bind_children(widget, scroll_data)
    
    def _is_descendant_of(self, widget, ancestor):
        """Check if widget is a descendant of ancestor widget"""
        try:
            current = widget
            while current:
                if current == ancestor:
                    return True
                current = current.master
            return False
        except Exception:
            return False
    
    def _get_relative_scroll_position(self, canvas):
        """Get the current relative scroll position (0.0 to 1.0) of a canvas"""
        try:
            # Get the current view position
            view_top, view_bottom = canvas.yview()
            
            # Calculate the relative position (0.0 = top, 1.0 = bottom)
            if view_bottom - view_top >= 1.0:
                # Content fits entirely in view, return 0.0 (top)
                return 0.0
            else:
                # Calculate relative position based on the top of the view
                return view_top
        except Exception:
            return 0.0
    
    def _restore_relative_scroll_position(self, canvas, relative_position):
        """Restore a relative scroll position (0.0 to 1.0) to a canvas"""
        try:
            # Schedule the scroll restoration after UI has updated
            self.root.after(100, lambda: self._perform_scroll_restoration(canvas, relative_position))
        except Exception:
            # Silently handle any errors
            pass
    
    def _perform_scroll_restoration(self, canvas, relative_position):
        """Actually perform the scroll restoration"""
        try:
            # Update the scroll region first to ensure accurate positioning
            canvas.configure(scrollregion=canvas.bbox("all"))
            
            # Get the current view bounds to check if content is scrollable
            view_top, view_bottom = canvas.yview()
            
            # Only scroll if there's content that extends beyond the view
            if view_bottom - view_top < 1.0:
                # Clamp the relative position to valid range
                relative_position = max(0.0, min(1.0, relative_position))
                
                # Apply the relative scroll position
                canvas.yview_moveto(relative_position)
        except Exception:
            # Silently handle any errors
            pass
    
    def setup_drag_and_drop(self):
        """Set up drag and drop functionality for the files canvas"""
        if not self.drag_drop_available:
            print("Drag and drop not available, skipping setup")
            return
        
        try:
            # Import tkinterdnd2 for drag and drop functionality
            import tkinterdnd2 as tkdnd
            
            # Enable drag and drop primarily on the canvas
            # The canvas is the main visible area users will drop onto
            self.files_canvas.drop_target_register(tkdnd.DND_FILES)
            self.files_canvas.dnd_bind('<<DropEnter>>', self.on_drag_enter)
            self.files_canvas.dnd_bind('<<DropPosition>>', self.on_drag_position)
            self.files_canvas.dnd_bind('<<DropLeave>>', self.on_drag_leave)
            self.files_canvas.dnd_bind('<<Drop>>', self.on_drop)
            
            # Also enable on the scrollable frame for when files are visible
            self.files_scrollable_frame.drop_target_register(tkdnd.DND_FILES)
            self.files_scrollable_frame.dnd_bind('<<DropEnter>>', self.on_drag_enter)
            self.files_scrollable_frame.dnd_bind('<<DropPosition>>', self.on_drag_position)
            self.files_scrollable_frame.dnd_bind('<<DropLeave>>', self.on_drag_leave)
            self.files_scrollable_frame.dnd_bind('<<Drop>>', self.on_drop)
            
            print("Drag and drop setup completed successfully")
            
        except Exception as e:
            print(f"Failed to set up drag and drop: {e}")
            self.drag_drop_available = False
            # Update tooltip to reflect the change
            self.show_status("Drag and drop setup failed, use 'Add Files' button instead", "warning")
    
    def on_drag_enter(self, event):
        """Handle drag enter event"""
        # Change cursor or visual feedback when files are dragged over
        try:
            event.widget.configure(cursor="plus")
        except:
            pass
        
        # Return the action for tkinterdnd2
        return "copy"
    
    def on_drag_position(self, event):
        """Handle drag position event"""
        # Return the action for tkinterdnd2
        return "copy"
    
    def on_drag_leave(self, event):
        """Handle drag leave event"""
        # Reset cursor when drag leaves
        try:
            event.widget.configure(cursor="")
        except:
            pass
        return None
    
    def on_drop(self, event):
        """Handle file drop event"""
        # Reset cursor
        try:
            event.widget.configure(cursor="")
        except:
            pass
        
        # Get dropped files - event.data contains the file paths
        files = self.parse_drop_data(event.data)
        
        if files:
            added_count = 0
            skipped_count = 0
            
            for file_path in files:
                # Check if it's a file (not directory)
                if os.path.isfile(file_path):
                    if self.should_track_file(file_path):
                        if file_path not in self.tracked_files:
                            self.tracked_files.append(file_path)
                            added_count += 1
                        else:
                            skipped_count += 1
                    else:
                        skipped_count += 1
            
            if added_count > 0:
                self.update_files_display()
                self.update_file_count_label()
                if skipped_count > 0:
                    self.show_status(f"Added {added_count} files, skipped {skipped_count} (already tracked or format filter)", "success")
                else:
                    self.show_status(f"Added {added_count} files via drag and drop", "success")
            else:
                if skipped_count > 0:
                    self.show_status("No files added - all were already tracked or don't match format filter", "warning")
        else:
            self.show_status("No valid files found in drop data", "warning")
        
        # Return the action for tkinterdnd2
        return "copy"
    
    def parse_drop_data(self, data):
        """Parse dropped file data into file paths"""
        files = []
        
        # Handle different data formats that might be received from tkinterdnd2
        if isinstance(data, (list, tuple)):
            # Data is already a list of file paths
            for item in data:
                if isinstance(item, str) and os.path.exists(item):
                    files.append(item)
        elif isinstance(data, str):
            # Data is a string with file paths - handle different formats
            # tkinterdnd2 often provides paths in braces format like: {path1} {path2}
            if data.startswith('{') and '}' in data:
                # Handle brace-enclosed paths
                import re
                paths = re.findall(r'\{([^}]+)\}', data)
                for path in paths:
                    path = path.strip()
                    if path and os.path.exists(path):
                        files.append(path)
            else:
                # Handle space-separated or newline-separated paths
                # Try both space and newline separators
                separators = ['\n', ' ']
                for separator in separators:
                    if separator in data:
                        for line in data.split(separator):
                            path = line.strip()
                            # Remove file:// protocol if present
                            if path.startswith('file://'):
                                path = path[7:]
                            # Handle URL-encoded paths (spaces as %20, etc.)
                            try:
                                import urllib.parse
                                path = urllib.parse.unquote(path)
                            except:
                                pass
                            
                            if path and os.path.exists(path):
                                files.append(path)
                        break  # If we found files with this separator, don't try others
                else:
                    # Single file path
                    path = data.strip()
                    if path.startswith('file://'):
                        path = path[7:]
                    try:
                        import urllib.parse
                        path = urllib.parse.unquote(path)
                    except:
                        pass
                    if path and os.path.exists(path):
                        files.append(path)
        
        return files
    

    
    def copy_and_rename(self):
        """Copy and rename tracked files"""
        if not self.tracked_files:
            self.show_status("No files to copy", "warning")
            return
        
        if not self.dest_folder.get():
            self.show_status("Please select a destination folder", "error")
            return
        
        if not os.path.exists(self.dest_folder.get()):
            self.show_status("Destination folder does not exist", "error")
            return
        
        # Check for missing rules and ask for confirmation
        missing_tags = self.get_missing_rule_tags()
        if missing_tags:
            message = f"The naming pattern contains tags without corresponding rules: {', '.join(missing_tags)}\n\n"
            message += "These tags will be left as-is in the filenames.\n\n"
            message += "Do you want to continue?"
            
            if not messagebox.askyesno("Missing Rules Warning", message, icon='warning'):
                return
        
        # Check for existing files and ask for action
        existing_files_action = None
        if self.has_existing_files_in_destination():
            existing_files = []
            for i, file_path in enumerate(self.tracked_files):
                preview_name = self.generate_filename_preview(i, len(self.tracked_files))
                file_ext = os.path.splitext(file_path)[1]
                preview_full_name = preview_name + file_ext
                if self.preview_exists_in_destination(preview_full_name):
                    existing_files.append(preview_full_name)
            
            existing_files_action = self.show_file_conflict_dialog(existing_files)
            if existing_files_action == "cancel":
                return
        
        try:
            # Reset all rules for new batch
            for rule in self.rules:
                rule.reset()
            
            copied_count = 0
            skipped_count = 0
            
            for i, file_path in enumerate(self.tracked_files):
                if os.path.exists(file_path):
                    new_name = self.generate_filename(i, len(self.tracked_files))
                    file_ext = os.path.splitext(file_path)[1]
                    dest_path = os.path.join(self.dest_folder.get(), new_name + file_ext)
                    
                    # Handle existing files based on user's choice
                    if os.path.exists(dest_path):
                        if existing_files_action == "ignore":
                            skipped_count += 1
                            continue
                        elif existing_files_action == "rename":
                            # Find a unique name by adding numbers
                            counter = 1
                            base_dest_path = dest_path
                            while os.path.exists(dest_path):
                                name_part = os.path.splitext(base_dest_path)[0]
                                dest_path = f"{name_part}_{counter}{file_ext}"
                                counter += 1
                        # For "overwrite", we just proceed with the original dest_path
                    
                    shutil.copy2(file_path, dest_path)
                    copied_count += 1
                    
                    # Track the latest rename (store original filename and new filename)
                    original_filename = os.path.basename(file_path)
                    new_filename = os.path.basename(dest_path)
                    self.latest_rename_info = (original_filename, new_filename)
            
            # Increment batch counters
            for rule in self.rules:
                if isinstance(rule, BatchRule):
                    rule.increment_batch()
            
            # Clear tracked files and update display
            self.tracked_files.clear()
            self.update_files_display()
            self.update_file_count_label()
            self.update_rules_display()
            self.update_latest_rename_label()
            
            # Show completion message
            if skipped_count > 0:
                self.show_status(f"Copied {copied_count} files, skipped {skipped_count} existing files", "success")
            else:
                self.show_status(f"Copied and renamed {copied_count} files", "success")
            
        except Exception as e:
            self.show_status(f"Failed to copy files: {str(e)}", "error")
    
    def generate_filename(self, file_index, total_files):
        """Generate filename using the naming pattern and rules"""
        pattern = self.naming_pattern.get()
        
        # Replace rule tags
        for rule in self.rules:
            tag = f"{{{rule.tag_name}}}"
            if tag in pattern:
                value = rule.get_value(file_index, total_files)
                pattern = pattern.replace(tag, value)
        
        return pattern
    
    def generate_filename_preview(self, file_index, total_files):
        """Generate filename preview without affecting rule states"""
        pattern = self.naming_pattern.get()
        
        # Create temporary rule copies to simulate without affecting actual state
        for rule in self.rules:
            tag = f"{{{rule.tag_name}}}"
            if tag in pattern:
                # Create a temporary copy of the rule for preview
                temp_rule = self._create_temp_rule_copy(rule)
                
                # Simulate the operations up to this file index
                for j in range(file_index + 1):
                    value = temp_rule.get_value(j, total_files)
                    if j == file_index:  # Only use the value for the current file
                        pattern = pattern.replace(tag, value)
        
        return pattern
    
    def _create_temp_rule_copy(self, rule):
        """Create a temporary copy of a rule for preview purposes"""
        if isinstance(rule, CounterRule):
            temp_rule = CounterRule(rule.tag_name, rule.start_value, rule.increment, rule.step, rule.max_value)
        elif isinstance(rule, ListRule):
            temp_rule = ListRule(rule.tag_name, rule.values.copy(), rule.step)
        elif isinstance(rule, BatchRule):
            temp_rule = BatchRule(rule.tag_name, rule.start_value, rule.increment, rule.step, rule.max_value)
            temp_rule.current_value = rule.current_value
            temp_rule.batch_count = rule.batch_count
        else:
            return rule
        
        temp_rule.reset()
        return temp_rule
    
    def add_rule(self):
        """Add a new rule"""
        # Generate a unique tag name
        base_name = "counter"
        counter = 1
        tag_name = base_name
        
        # Find a unique tag name
        while any(rule.tag_name == tag_name for rule in self.rules):
            tag_name = f"{base_name}{counter}"
            counter += 1
        
        # Create a default counter rule
        rule = CounterRule(tag_name, 0, 1, 1, None)
        self.rules.append(rule)
        self.update_rules_display()
        self.update_files_display()  # Update preview
    
    def create_rule_from_dialog(self, dialog_data):
        """Create a rule from dialog data"""
        rule_type = dialog_data['rule_type']
        tag_name = dialog_data['tag_name']
        step = dialog_data.get('step', 1)
        
        if rule_type == 'counter':
            return CounterRule(tag_name, dialog_data['start_value'], dialog_data['increment'], step)
        elif rule_type == 'list':
            values = [v.strip() for v in dialog_data['values'].split(';') if v.strip()]
            return ListRule(tag_name, values, step)
        elif rule_type == 'batch':
            return BatchRule(tag_name, dialog_data['start_value'], dialog_data['increment'], step)
        else:
            return CounterRule(tag_name, 0, 1, 1)
    
    def update_rules_display(self):
        """Update the rules display"""
        # Detect what changed compared to last state
        changes = self._detect_rule_changes()
        
        # Apply incremental updates
        if changes['full_rebuild_needed']:
            self._full_rebuild_rules()
        else:
            self._incremental_update_rules(changes)
        
        # Update last known state
        self.last_rules_state = self._get_current_rules_state()
        
        # Update naming pattern label warning
        self.update_naming_pattern_label()
    
    def _get_current_rules_state(self):
        """Get current state of rules for change detection"""
        state = []
        for i, rule in enumerate(self.rules):
            tag_used = self.is_tag_used_in_pattern(rule.tag_name)
            rule_data = rule.to_dict()
            rule_data['tag_used'] = tag_used
            state.append(rule_data)
        return state
    
    def _detect_rule_changes(self):
        """Detect what changed in the rules list"""
        current_state = self._get_current_rules_state()
        last_state = self.last_rules_state
        
        # Check if widget count is wrong (requires full rebuild)
        # This handles cases where widgets were cleared or count doesn't match
        widget_count_mismatch = len(self.rule_widgets) != len(self.rules)
        
        if widget_count_mismatch:
            return {'full_rebuild_needed': True}
        
        # Handle rule additions/removals
        rules_added = len(current_state) > len(last_state)
        rules_removed = len(current_state) < len(last_state)
        
        if rules_removed:
            # Rules were removed - need full rebuild to handle index changes
            return {'full_rebuild_needed': True}
        
        # Check if any rule type changed (requires full rebuild because widgets are different)
        for i in range(min(len(current_state), len(last_state))):
            if current_state[i].get('type') != last_state[i].get('type'):
                return {'full_rebuild_needed': True}
        
        # Find specific changes
        updated_indices = []
        new_indices = []
        
        # Check existing rules for changes
        for i in range(min(len(current_state), len(last_state))):
            if current_state[i] != last_state[i]:
                updated_indices.append(i)
        
        # Handle new rules
        if rules_added:
            for i in range(len(last_state), len(current_state)):
                new_indices.append(i)
            
            # Also update the previously last item's state (if it exists)
            if len(last_state) > 0:
                prev_last_index = len(last_state) - 1
                if prev_last_index not in updated_indices:
                    updated_indices.append(prev_last_index)
        
        return {
            'full_rebuild_needed': False,
            'updated_indices': updated_indices,
            'new_indices': new_indices
        }
    
    def _full_rebuild_rules(self):
        """Perform full rebuild of rules display"""
        # Clear existing widgets
        for widget_data in self.rule_widgets.values():
            if 'frame' in widget_data:
                widget_data['frame'].destroy()
        self.rule_widgets.clear()
        
        # Clear the container
        for widget in self.rules_scrollable_frame.winfo_children():
            widget.destroy()
        
        # Add rules
        for i, rule in enumerate(self.rules):
            self.add_rule_to_display(rule, i)
    
    def _incremental_update_rules(self, changes):
        """Perform incremental update of rules display"""
        # Update existing changed items
        for index in changes['updated_indices']:
            if index in self.rule_widgets:
                self._update_rule_item(index)
        
        # Add new items (process in reverse order for better performance)
        new_indices = changes['new_indices']
        for index in reversed(new_indices):
            rule = self.rules[index]
            self.add_rule_to_display(rule, index)
        
        # Auto-scroll to show the newest rule if any were added
        if new_indices:
            self.scroll_to_newest_rule(max(new_indices))
    
    def scroll_to_newest_rule(self, rule_index):
        """Scroll the rules view to show the newly added rule"""
        # Schedule the scroll after the UI has updated
        self.root.after(100, lambda: self._perform_scroll_to_rule(rule_index))
    
    def _perform_scroll_to_rule(self, rule_index):
        """Actually perform the scroll to show the specified rule"""
        try:
            # Update the scroll region first
            self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox("all"))
            
            # Check if we have a widget for this rule index
            if rule_index in self.rule_widgets:
                widget_data = self.rule_widgets[rule_index]
                
                # Get the main widget for this rule
                main_widget = widget_data.get('frame')
                
                if main_widget and main_widget.winfo_exists():
                    # Get the widget's position relative to the scrollable frame
                    widget_y = main_widget.winfo_y()
                    widget_height = main_widget.winfo_height()
                    
                    # Get the scrollable frame's total height
                    frame_height = self.rules_scrollable_frame.winfo_reqheight()
                    
                    # Get the canvas viewport height
                    canvas_height = self.rules_canvas.winfo_height()
                    
                    if frame_height > canvas_height:
                        # Calculate the position to scroll to (show the widget at the bottom of the view)
                        # This ensures the new rule is visible
                        target_y = widget_y + widget_height - canvas_height + 20  # 20px padding
                        target_y = max(0, target_y)  # Don't scroll past the top
                        
                        # Convert to fraction of total scrollable area
                        scroll_fraction = target_y / (frame_height - canvas_height)
                        scroll_fraction = min(1.0, max(0.0, scroll_fraction))  # Clamp to [0, 1]
                        
                        # Scroll to show the new rule
                        self.rules_canvas.yview_moveto(scroll_fraction)
        except Exception as e:
            # Silently handle any scrolling errors
            pass
    
    def _update_rule_item(self, index):
        """Update a specific rule item without rebuilding"""
        if index not in self.rule_widgets or index >= len(self.rules):
            return
        
        widget_data = self.rule_widgets[index]
        rule = self.rules[index]
        
        # Update tag usage background color
        tag_used = self.is_tag_used_in_pattern(rule.tag_name)
        
        if not tag_used:
            bg_color = "#fffacd"  # Light yellow background for unused tags
            style_name = f"RuleUnused{index}.TFrame"
        else:
            bg_color = "#f0f0f0" if index % 2 == 0 else "#ffffff"
            style_name = f"Rule{index % 2}.TFrame"
        
        if 'frame' in widget_data:
            widget_data['frame'].configure(style=style_name)
            style = ttk.Style()
            style.configure(style_name, background=bg_color)
        
        # Update rule type if it changed
        if 'rule_type_var' in widget_data:
            widget_data['rule_type_var'].set(rule.__class__.__name__)
        
        # Update tag name if it changed
        if 'tag_var' in widget_data:
            widget_data['tag_var'].set(rule.tag_name)
        
        # Update rule-specific fields
        if isinstance(rule, CounterRule):
            if 'start_var' in widget_data:
                widget_data['start_var'].set(rule.start_value)
            if 'inc_var' in widget_data:
                widget_data['inc_var'].set(rule.increment)
            if 'max_var' in widget_data:
                widget_data['max_var'].set(str(rule.max_value) if rule.max_value is not None else "")
            if 'step_var' in widget_data:
                widget_data['step_var'].set(rule.step)
        elif isinstance(rule, ListRule):
            if 'values_var' in widget_data:
                widget_data['values_var'].set('; '.join(rule.values))
            if 'step_var' in widget_data:
                widget_data['step_var'].set(rule.step)
        elif isinstance(rule, BatchRule):
            if 'current_var' in widget_data:
                widget_data['current_var'].set(rule.current_value)
            if 'inc_var' in widget_data:
                widget_data['inc_var'].set(rule.increment)
            if 'max_var' in widget_data:
                widget_data['max_var'].set(str(rule.max_value) if rule.max_value is not None else "")
            if 'step_var' in widget_data:
                widget_data['step_var'].set(rule.step)
    
    def add_rule_to_display(self, rule, index):
        """Add a rule to the rules display with inline editing"""
        # Create a frame for the rule with alternating background colors
        rule_frame = ttk.Frame(self.rules_scrollable_frame)
        rule_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # Configure the scrollable frame to expand its children
        self.rules_scrollable_frame.columnconfigure(0, weight=1)
        
        # Configure column weights to make the frame responsive
        # For list rules, make the values field expandable (will be set later in add_list_fields)
        # For other rules, the spacer column weight will be set dynamically when spacer is added
        
        # Determine background color based on tag usage and alternating pattern
        tag_used = self.is_tag_used_in_pattern(rule.tag_name)
        
        if not tag_used:
            # Light yellow background for unused tags
            bg_color = "#fffacd"
            style_name = f"RuleUnused{index}.TFrame"
        else:
            # Alternating colors for used tags
            bg_color = "#f0f0f0" if index % 2 == 0 else "#ffffff"
            style_name = f"Rule{index % 2}.TFrame"
        
        rule_frame.configure(style=style_name)
        
        # Configure styles using the bg_color variable
        style = ttk.Style()
        style.configure(style_name, background=bg_color)
        
        col = 0
        
        # Rule type dropdown
        rule_type_var = tk.StringVar()
        rule_type_dropdown = ttk.Combobox(rule_frame, textvariable=rule_type_var, 
                                         values=['CounterRule', 'ListRule', 'BatchRule'], 
                                         state='readonly', width=12)
        rule_type_dropdown.grid(row=0, column=col, padx=(0, 5), sticky=tk.W)
        ToolTip(rule_type_dropdown, "Type of rule:\n• CounterRule: increments with each file (0,1,2...)\n• ListRule: cycles through custom values (A,B,C,A,B,C...)\n• BatchRule: increments with each copy operation (stays same within batch)")
        col += 1
        
        # Set current rule type
        rule_type_var.set(rule.__class__.__name__)
        rule_type_dropdown.bind('<<ComboboxSelected>>', lambda e, idx=index, var=rule_type_var: self.change_rule_type(idx, var.get()))
        
        # Tag name field
        tag_label = ttk.Label(rule_frame, text="Tag:")
        tag_label.grid(row=0, column=col, padx=(5, 2), sticky=tk.W)
        ToolTip(tag_label, "Tag name used in the naming pattern. Use {tag_name} in the pattern to insert this rule's value.")
        col += 1
        tag_var = tk.StringVar(value=rule.tag_name)
        tag_entry = ttk.Entry(rule_frame, textvariable=tag_var, width=12)
        tag_entry.grid(row=0, column=col, padx=(0, 5), sticky=tk.W)
        self._setup_custom_word_navigation(tag_entry)
        ToolTip(tag_entry, "Name of this rule's tag. Use {tag_name} in the naming pattern to insert values from this rule. Must be unique among all rules.")
        col += 1
        tag_entry.bind('<FocusOut>', lambda e, idx=index, var=tag_var: self.update_rule_tag(idx, var.get()))
        tag_entry.bind('<Return>', lambda e, idx=index, var=tag_var, entry=tag_entry: self._handle_enter_update(entry, lambda: self.update_rule_tag(idx, var.get())))
        
        # Rule-specific fields and variables
        rule_specific_vars = {}
        if isinstance(rule, CounterRule):
            col, counter_vars = self.add_counter_fields(rule_frame, rule, index, col)
            rule_specific_vars.update(counter_vars)
        elif isinstance(rule, ListRule):
            col, list_vars = self.add_list_fields(rule_frame, rule, index, col)
            rule_specific_vars.update(list_vars)
        elif isinstance(rule, BatchRule):
            col, batch_vars = self.add_batch_fields(rule_frame, rule, index, col)
            rule_specific_vars.update(batch_vars)
        
        # Add spacer for non-list rules to push step and delete button to the right
        if not isinstance(rule, ListRule):
            spacer = ttk.Label(rule_frame, text="")
            spacer.grid(row=0, column=col, sticky=(tk.W, tk.E))
            # Update the column weight to match the spacer column
            rule_frame.columnconfigure(col, weight=1)
            col += 1
        
        # Step field (always before delete button)
        step_label = ttk.Label(rule_frame, text="Step:")
        step_label.grid(row=0, column=col, padx=(5, 2), sticky=tk.W)
        ToolTip(step_label, "How many files to process before advancing this rule. Step=1 advances every file, Step=2 advances every other file, etc.")
        col += 1
        step_var = tk.IntVar(value=rule.step)
        step_entry = ttk.Entry(rule_frame, textvariable=step_var, width=6)
        step_entry.grid(row=0, column=col, padx=(0, 5), sticky=tk.W)
        self._setup_custom_word_navigation(step_entry)
        ToolTip(step_entry, "How many files to process before advancing this rule. Useful for applying the same value to groups of files.")
        col += 1
        
        # Bind step field events based on rule type
        if isinstance(rule, CounterRule):
            step_entry.bind('<FocusOut>', lambda e, idx=index, var=step_var: self.update_counter_step(idx, var.get()))
            step_entry.bind('<Return>', lambda e, idx=index, var=step_var, entry=step_entry: self._handle_enter_update(entry, lambda: self.update_counter_step(idx, var.get())))
        elif isinstance(rule, ListRule):
            step_entry.bind('<FocusOut>', lambda e, idx=index, var=step_var: self.update_list_step(idx, var.get()))
            step_entry.bind('<Return>', lambda e, idx=index, var=step_var, entry=step_entry: self._handle_enter_update(entry, lambda: self.update_list_step(idx, var.get())))
        elif isinstance(rule, BatchRule):
            step_entry.bind('<FocusOut>', lambda e, idx=index, var=step_var: self.update_batch_step(idx, var.get()))
            step_entry.bind('<Return>', lambda e, idx=index, var=step_var, entry=step_entry: self._handle_enter_update(entry, lambda: self.update_batch_step(idx, var.get())))
        
        # Delete button (always at the end)
        delete_button = ttk.Button(rule_frame, text="✕", width=3,
                                  command=lambda: self.delete_rule_by_index(index))
        delete_button.grid(row=0, column=col, padx=(5, 0), sticky=tk.E)
        ToolTip(delete_button, "Remove this rule. Any tags using this rule name in the naming pattern will be left as-is in filenames.")
        
        # Store widget references for incremental updates
        widget_storage = {
            'frame': rule_frame,
            'rule_type_var': rule_type_var,
            'tag_var': tag_var,
            'step_var': step_var,
            'delete_button': delete_button
        }
        
        # Add rule-specific widget references
        widget_storage.update(rule_specific_vars)
        
        self.rule_widgets[index] = widget_storage
        
        # Bind scroll events to the new widgets immediately
        self._bind_scroll_to_new_widget(rule_frame)
    
    def add_counter_fields(self, parent, rule, index, start_col):
        """Add counter rule specific fields"""
        col = start_col
        
        start_label = ttk.Label(parent, text="Start:")
        start_label.grid(row=0, column=col, padx=(5, 2), sticky=tk.W)
        ToolTip(start_label, "Starting value for the counter. This is the first number that will be used.")
        col += 1
        start_var = tk.IntVar(value=rule.start_value)
        start_entry = ttk.Entry(parent, textvariable=start_var, width=6)
        start_entry.grid(row=0, column=col, padx=(0, 5), sticky=tk.W)
        self._setup_custom_word_navigation(start_entry)
        ToolTip(start_entry, "The counter begins with this value for the first file.")
        col += 1
        start_entry.bind('<FocusOut>', lambda e, idx=index, var=start_var: self.update_counter_start(idx, var.get()))
        start_entry.bind('<Return>', lambda e, idx=index, var=start_var, entry=start_entry: self._handle_enter_update(entry, lambda: self.update_counter_start(idx, var.get())))
        
        inc_label = ttk.Label(parent, text="Inc:")
        inc_label.grid(row=0, column=col, padx=(5, 2), sticky=tk.W)
        ToolTip(inc_label, "How much to add to the counter for each step. Can be negative to count down.")
        col += 1
        inc_var = tk.IntVar(value=rule.increment)
        inc_entry = ttk.Entry(parent, textvariable=inc_var, width=6)
        inc_entry.grid(row=0, column=col, padx=(0, 5), sticky=tk.W)
        self._setup_custom_word_navigation(inc_entry)
        ToolTip(inc_entry, "Counter increases by this amount each step. Use negative values to count down.")
        col += 1
        inc_entry.bind('<FocusOut>', lambda e, idx=index, var=inc_var: self.update_counter_increment(idx, var.get()))
        inc_entry.bind('<Return>', lambda e, idx=index, var=inc_var, entry=inc_entry: self._handle_enter_update(entry, lambda: self.update_counter_increment(idx, var.get())))
        
        max_label = ttk.Label(parent, text="Max:")
        max_label.grid(row=0, column=col, padx=(5, 2), sticky=tk.W)
        ToolTip(max_label, "Maximum value before wrapping back to start value. Leave empty for no limit.")
        col += 1
        max_var = tk.StringVar(value=str(rule.max_value) if rule.max_value is not None else "")
        max_entry = ttk.Entry(parent, textvariable=max_var, width=6)
        max_entry.grid(row=0, column=col, padx=(0, 5), sticky=tk.W)
        self._setup_custom_word_navigation(max_entry)
        ToolTip(max_entry, "When reached, counter wraps back to start value. Leave empty for unlimited counting.")
        col += 1
        max_entry.bind('<FocusOut>', lambda e, idx=index, var=max_var: self.update_counter_max(idx, var.get()))
        max_entry.bind('<Return>', lambda e, idx=index, var=max_var, entry=max_entry: self._handle_enter_update(entry, lambda: self.update_counter_max(idx, var.get())))
        
        # Return column position and variables
        variables = {
            'start_var': start_var,
            'inc_var': inc_var,
            'max_var': max_var
        }
        return col, variables
    
    def add_list_fields(self, parent, rule, index, start_col):
        """Add list rule specific fields"""
        col = start_col
        
        values_label = ttk.Label(parent, text="Values:")
        values_label.grid(row=0, column=col, padx=(5, 2), sticky=tk.W)
        ToolTip(values_label, "Semicolon-separated list of values to cycle through. The rule will use these values in order and repeat.")
        col += 1
        values_var = tk.StringVar(value='; '.join(rule.values))
        values_entry = ttk.Entry(parent, textvariable=values_var)
        values_entry.grid(row=0, column=col, padx=(0, 5), sticky=(tk.W, tk.E))
        self._setup_custom_word_navigation(values_entry)
        ToolTip(values_entry, "Separate values with semicolons (;). Examples: 'A;B;C' or 'red;green;blue' or '1;10;100'")
        # Configure this column to expand for list rules
        parent.columnconfigure(col, weight=1)
        col += 1
        values_entry.bind('<FocusOut>', lambda e, idx=index, var=values_var: self.update_list_values(idx, var.get()))
        values_entry.bind('<Return>', lambda e, idx=index, var=values_var, entry=values_entry: self._handle_enter_update(entry, lambda: self.update_list_values(idx, var.get())))
        
        # Return column position and variables
        variables = {
            'values_var': values_var
        }
        return col, variables
    
    def add_batch_fields(self, parent, rule, index, start_col):
        """Add batch rule specific fields"""
        col = start_col
        
        current_label = ttk.Label(parent, text="Current:")
        current_label.grid(row=0, column=col, padx=(5, 2), sticky=tk.W)
        ToolTip(current_label, "Current value of the batch counter. This stays the same for all files in a batch and increments when you copy files.")
        col += 1
        current_var = tk.IntVar(value=rule.current_value)
        current_entry = ttk.Entry(parent, textvariable=current_var, width=6)
        current_entry.grid(row=0, column=col, padx=(0, 5), sticky=tk.W)
        self._setup_custom_word_navigation(current_entry)
        ToolTip(current_entry, "All files in this batch will use this value. Increments after each copy operation.")
        col += 1
        current_entry.bind('<FocusOut>', lambda e, idx=index, var=current_var: self.update_batch_current(idx, var.get()))
        current_entry.bind('<Return>', lambda e, idx=index, var=current_var, entry=current_entry: self._handle_enter_update(entry, lambda: self.update_batch_current(idx, var.get())))
        
        inc_label = ttk.Label(parent, text="Inc:")
        inc_label.grid(row=0, column=col, padx=(5, 2), sticky=tk.W)
        ToolTip(inc_label, "How much to increase the batch counter after each copy operation. Can be negative.")
        col += 1
        inc_var = tk.IntVar(value=rule.increment)
        inc_entry = ttk.Entry(parent, textvariable=inc_var, width=6)
        inc_entry.grid(row=0, column=col, padx=(0, 5), sticky=tk.W)
        self._setup_custom_word_navigation(inc_entry)
        ToolTip(inc_entry, "Batch counter increases by this amount after each copy operation.")
        col += 1
        inc_entry.bind('<FocusOut>', lambda e, idx=index, var=inc_var: self.update_batch_increment(idx, var.get()))
        inc_entry.bind('<Return>', lambda e, idx=index, var=inc_var, entry=inc_entry: self._handle_enter_update(entry, lambda: self.update_batch_increment(idx, var.get())))
        
        max_label = ttk.Label(parent, text="Max:")
        max_label.grid(row=0, column=col, padx=(5, 2), sticky=tk.W)
        ToolTip(max_label, "Maximum batch value before wrapping back to start. Leave empty for no limit.")
        col += 1
        max_var = tk.StringVar(value=str(rule.max_value) if rule.max_value is not None else "")
        max_entry = ttk.Entry(parent, textvariable=max_var, width=6)
        max_entry.grid(row=0, column=col, padx=(0, 5), sticky=tk.W)
        self._setup_custom_word_navigation(max_entry)
        ToolTip(max_entry, "When reached, batch counter wraps back to start. Leave empty for unlimited.")
        col += 1
        max_entry.bind('<FocusOut>', lambda e, idx=index, var=max_var: self.update_batch_max(idx, var.get()))
        max_entry.bind('<Return>', lambda e, idx=index, var=max_var, entry=max_entry: self._handle_enter_update(entry, lambda: self.update_batch_max(idx, var.get())))
        
        # Return column position and variables
        variables = {
            'current_var': current_var,
            'inc_var': inc_var,
            'max_var': max_var
        }
        return col, variables
    
    def change_rule_type(self, index, new_type):
        """Change the type of a rule"""
        old_rule = self.rules[index]
        tag_name = old_rule.tag_name
        
        if new_type == 'CounterRule':
            new_rule = CounterRule(tag_name, 0, 1, 1, None)
        elif new_type == 'ListRule':
            new_rule = ListRule(tag_name, ['value1'], 1)
        elif new_type == 'BatchRule':
            new_rule = BatchRule(tag_name, 0, 1, 1, None)
        else:
            return
        
        self.rules[index] = new_rule
        # Clear widget tracking to force rebuild when rule type changes
        self.rule_widgets.clear()
        self.update_rules_display()
        self.update_files_display()  # Update preview
    
    def _handle_enter_update(self, entry_widget, update_func):
        """Helper method to handle Enter key press: update and remove focus"""
        update_func()
        # Remove focus from the entry widget after update
        entry_widget.master.focus_set()
    
    def _is_word_char(self, char):
        """Check if a character is part of a word (alphanumeric or underscore, and not in separator string)"""
        # Whitespace is always a separator
        if char.isspace():
            return False
        # Check if character is in the separator string
        if char in self.word_separators:
            return False
        # Word characters are alphanumeric or underscore
        return char.isalnum() or char == '_'
    
    def _find_word_start(self, text, pos):
        """Find the start of the word at position pos, treating symbols as separators"""
        if pos <= 0:
            return 0
        
        # Move back to find the start of the current word
        # A word starts when we transition from non-word to word character
        i = pos - 1
        
        # Skip any non-word characters at current position
        while i >= 0 and not self._is_word_char(text[i]):
            i -= 1
        
        # Now find the start of the word
        while i > 0 and self._is_word_char(text[i - 1]):
            i -= 1
        
        return i
    
    def _find_word_end(self, text, pos):
        """Find the end of the word at position pos, treating symbols as separators"""
        if pos >= len(text):
            return len(text)
        
        # Move forward to find the end of the current word
        # A word ends when we transition from word to non-word character
        i = pos
        
        # Skip any non-word characters at current position
        while i < len(text) and not self._is_word_char(text[i]):
            i += 1
        
        # Now find the end of the word
        while i < len(text) and self._is_word_char(text[i]):
            i += 1
        
        return i
    
    def _on_ctrl_left(self, event):
        """Handle Ctrl+Left: move to start of previous word"""
        widget = event.widget
        if not isinstance(widget, (tk.Entry, ttk.Entry)):
            return
        
        try:
            text = widget.get()
            cursor_pos = widget.index(tk.INSERT)
            
            # Find start of current word
            word_start = self._find_word_start(text, cursor_pos)
            
            # If we're already at the start of a word, move to start of previous word
            if word_start == cursor_pos and cursor_pos > 0:
                word_start = self._find_word_start(text, cursor_pos - 1)
            
            widget.icursor(word_start)
            return "break"
        except Exception:
            return None
    
    def _on_ctrl_right(self, event):
        """Handle Ctrl+Right: move to end of next word"""
        widget = event.widget
        if not isinstance(widget, (tk.Entry, ttk.Entry)):
            return
        
        try:
            text = widget.get()
            cursor_pos = widget.index(tk.INSERT)
            
            # Find end of current word
            word_end = self._find_word_end(text, cursor_pos)
            
            # If we're already at the end of a word, move to end of next word
            if word_end == cursor_pos and cursor_pos < len(text):
                word_end = self._find_word_end(text, cursor_pos + 1)
            
            widget.icursor(word_end)
            return "break"
        except Exception:
            return None
    
    def _setup_custom_word_navigation(self, entry_widget):
        """Set up custom word navigation for an Entry widget"""
        entry_widget.bind('<Control-Left>', self._on_ctrl_left)
        entry_widget.bind('<Control-Right>', self._on_ctrl_right)
    
    def update_rule_tag(self, index, new_tag):
        """Update a rule's tag name"""
        if not new_tag.strip():
            return
        
        # Check for duplicate tag names
        if any(rule.tag_name == new_tag for i, rule in enumerate(self.rules) if i != index):
            self.show_status("Tag name already exists", "error")
            self.update_rules_display()  # Reset display
            return
        
        self.rules[index].tag_name = new_tag
        self.update_files_display()  # Update preview
    
    def update_counter_start(self, index, value):
        """Update counter rule start value"""
        if isinstance(self.rules[index], CounterRule):
            self.rules[index].start_value = value
            self.rules[index].current_value = value
            self.update_files_display()  # Update preview
    
    def update_counter_increment(self, index, value):
        """Update counter rule increment"""
        if isinstance(self.rules[index], CounterRule):
            self.rules[index].increment = value
            self.update_files_display()  # Update preview
    
    def update_counter_step(self, index, value):
        """Update counter rule step"""
        if isinstance(self.rules[index], CounterRule):
            self.rules[index].step = max(1, value)  # Ensure step is at least 1
            self.update_files_display()  # Update preview
    
    def update_counter_max(self, index, value):
        """Update counter rule max value"""
        if isinstance(self.rules[index], CounterRule):
            if value.strip() == "":
                self.rules[index].max_value = None
            else:
                try:
                    max_val = int(value)
                    self.rules[index].max_value = max_val if max_val >= self.rules[index].start_value else None
                except ValueError:
                    self.rules[index].max_value = None
            self.update_files_display()  # Update preview
    
    def update_list_values(self, index, values_text):
        """Update list rule values"""
        if isinstance(self.rules[index], ListRule):
            values = [v.strip() for v in values_text.split(';') if v.strip()]
            self.rules[index].values = values if values else ['value1']
            self.update_files_display()  # Update preview
    
    def update_list_step(self, index, value):
        """Update list rule step"""
        if isinstance(self.rules[index], ListRule):
            self.rules[index].step = max(1, value)  # Ensure step is at least 1
            self.update_files_display()  # Update preview
    
    def update_batch_current(self, index, value):
        """Update batch rule current value"""
        if isinstance(self.rules[index], BatchRule):
            self.rules[index].current_value = value
            self.update_files_display()  # Update preview
    
    def update_batch_increment(self, index, value):
        """Update batch rule increment"""
        if isinstance(self.rules[index], BatchRule):
            self.rules[index].increment = value
            self.update_files_display()  # Update preview
    
    def update_batch_step(self, index, value):
        """Update batch rule step"""
        if isinstance(self.rules[index], BatchRule):
            self.rules[index].step = max(1, value)  # Ensure step is at least 1
            self.update_files_display()  # Update preview
    
    def update_batch_max(self, index, value):
        """Update batch rule max value"""
        if isinstance(self.rules[index], BatchRule):
            if value.strip() == "":
                self.rules[index].max_value = None
            else:
                try:
                    max_val = int(value)
                    self.rules[index].max_value = max_val if max_val >= self.rules[index].start_value else None
                except ValueError:
                    self.rules[index].max_value = None
            self.update_files_display()  # Update preview
    
    def delete_rule_by_index(self, index):
        """Delete a rule by index"""
        if 0 <= index < len(self.rules):
            del self.rules[index]
            # Clear widget tracking to force rebuild on removal
            self.rule_widgets.clear()
            self.update_rules_display()
            self.update_files_display()  # Update preview
    
    def save_settings(self):
        """Save settings to file"""
        settings = {
            'source_folder': self.source_folder.get(),
            'dest_folder': self.dest_folder.get(),
            'file_formats': self.file_formats.get(),
            'naming_pattern': self.naming_pattern.get(),
            'view_mode': self.view_mode.get(),
            'rules': [rule.to_dict() for rule in self.rules]
        }
        
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
            self.show_status("Settings saved", "success")
        except Exception as e:
            self.show_status(f"Failed to save settings: {str(e)}", "error")
    
    def load_settings(self):
        """Load settings from file"""
        if not os.path.exists(self.settings_file):
            return
        
        try:
            with open(self.settings_file, 'r') as f:
                settings = json.load(f)
            
            self.source_folder.set(settings.get('source_folder', ''))
            self.dest_folder.set(settings.get('dest_folder', ''))
            self.file_formats.set(settings.get('file_formats', '*'))
            self.naming_pattern.set(settings.get('naming_pattern', 'file_{counter}'))
            self.view_mode.set(settings.get('view_mode', 'list'))
            
            # Load rules
            self.rules.clear()
            for rule_data in settings.get('rules', []):
                if rule_data['type'] == 'counter':
                    rule = CounterRule.from_dict(rule_data)
                elif rule_data['type'] == 'list':
                    rule = ListRule.from_dict(rule_data)
                elif rule_data['type'] == 'batch':
                    rule = BatchRule.from_dict(rule_data)
                else:
                    continue
                self.rules.append(rule)
            
            # Clear widget tracking and force full rebuild for initial load
            self.rule_widgets.clear()
            self.file_widgets.clear()
            self.last_rules_state = []
            self.last_files_state = []
            
            self.update_rules_display()
            self.update_button_states()
            
        except Exception as e:
            self.show_status(f"Failed to load settings: {str(e)}", "error")
    
    def export_settings(self):
        """Export settings to a file"""
        file_path = filedialog.asksaveasfilename(
            title="Export Settings",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            settings = {
                'source_folder': self.source_folder.get(),
                'dest_folder': self.dest_folder.get(),
                'file_formats': self.file_formats.get(),
                'naming_pattern': self.naming_pattern.get(),
                'view_mode': self.view_mode.get(),
                'rules': [rule.to_dict() for rule in self.rules]
            }
            
            try:
                with open(file_path, 'w') as f:
                    json.dump(settings, f, indent=2)
                self.show_status("Settings exported successfully", "success")
            except Exception as e:
                self.show_status(f"Failed to export settings: {str(e)}", "error")
    
    def import_settings(self):
        """Import settings from a file"""
        file_path = filedialog.askopenfilename(
            title="Import Settings",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'r') as f:
                    settings = json.load(f)
                
                self.source_folder.set(settings.get('source_folder', ''))
                self.dest_folder.set(settings.get('dest_folder', ''))
                self.file_formats.set(settings.get('file_formats', '*'))
                self.naming_pattern.set(settings.get('naming_pattern', 'file_{counter}'))
                self.view_mode.set(settings.get('view_mode', 'list'))
                
                # Load rules
                self.rules.clear()
                for rule_data in settings.get('rules', []):
                    if rule_data['type'] == 'counter':
                        rule = CounterRule.from_dict(rule_data)
                    elif rule_data['type'] == 'list':
                        rule = ListRule.from_dict(rule_data)
                    elif rule_data['type'] == 'batch':
                        rule = BatchRule.from_dict(rule_data)
                    else:
                        continue
                    self.rules.append(rule)
                
                self.update_rules_display()
                self.update_button_states()
                self.show_status("Settings imported successfully", "success")
                
            except Exception as e:
                self.show_status(f"Failed to import settings: {str(e)}", "error")
    
    def run(self):
        """Run the application"""
        try:
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            self.root.mainloop()
        except KeyboardInterrupt:
            self.on_closing()
    
    def on_closing(self):
        """Handle application closing"""
        self.stop_tracking()
        self.save_settings()
        self.root.destroy()





if __name__ == "__main__":
    app = FileManagerApp()
    app.run()
