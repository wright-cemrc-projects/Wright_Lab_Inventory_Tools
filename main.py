import tkinter as tk
import sys
from console_helper import ConsoleFooter, ConsoleRedirector
import app_context
from inventory_helper import GridDewarManager, Freezer20Manager, CellDewarManager, Freezer80Manager
from atexit_helper import TempFileManager
from data_helper import IDManager, ExcelHelper
import atexit

# -------------------------------------------------------------------
# INITIALIZATION
# -------------------------------------------------------------------

# Redirect terminal output to shared ConsoleRedirector so messages also show up in GUI console
app_context.console_redirector = ConsoleRedirector()
sys.stdout = app_context.console_redirector
sys.stderr = app_context.console_redirector

# Initialize managers for global app state
id_manager = IDManager()  # Handles Google Drive file IDs
temp_file_manager = TempFileManager()  # Manages temporary files
app_context.temp_file_manager = temp_file_manager  # Make temp_file_manager globally accesible
excel_manager = ExcelHelper() # Manages file actions

# Register cleanup at exit for temporary files
atexit.register(temp_file_manager.cleanup_temp_files)

# -------------------------------------------------------------------
# ROOT WINDOW CONFIGURATION
# -------------------------------------------------------------------

# Create main application window
root = tk.Tk()
root.title("Wright Lab Data Entry")
root.geometry("700x700+100+100")

# Configure root grid layout
# Row 0 will contain main content
# Row 1 will contian the console (fixed height at bottom)
root.grid_rowconfigure(0, weight=1)  # Content area expands
root.grid_rowconfigure(1, weight=0)  # Console stays fixed
root.grid_columnconfigure(0, weight=1)

# -------------------------------------------------------------------
# FRAMES
# -------------------------------------------------------------------

# Main frame (content area above the console)
main_frame = tk.Frame(root)
main_frame.grid(row=0, column=0, sticky="nsew")

# Console footer at the bottom of the root window
console = ConsoleFooter(root)
console.grid(row=1, column=0, sticky="ew")

# -------------------------------------------------------------------
# WINDOW CLOSING BEHAVIOR
# -------------------------------------------------------------------

def on_closing():
    """Cleanup temp files and fully exit when root window is closed"""

    temp_file_manager.cleanup_temp_files()
    root.destroy()
    sys.exit()

# Bind close button (X) to custom handler on_closing
root.protocol("WM_DELETE_WINDOW", on_closing)

# -------------------------------------------------------------------
# HELP BUTTON
# -------------------------------------------------------------------
def click_Help():
    """Open the user manual word file"""

    excel_manager.open_excel_file("UserManual.docx")

help_button = tk.Button(
    main_frame,
    text="?",
    font=("Arial", 12, "bold"),
    width=3,
    relief="raised",
    command=click_Help)
help_button.grid(row=0, column=0, sticky="ne", padx=10, pady=(10, 0))

# -------------------------------------------------------------------
# HEADER
# -------------------------------------------------------------------

label = tk.Label(main_frame, text="What Category of Data Would You Like to Enter?", font=("Arial", 16))
# Place the label in the first row, spanning the full width
label.grid(row=1, column=0, pady=(0, 10), padx=20, sticky="n")

# -------------------------------------------------------------------
# BUTTON CALLBACKS
# -------------------------------------------------------------------

# Each button launches the "main menu" for its respective manager class.
# These managers handle UI/workflow for each inventory type.

def click_Grid():
    """Open the main menu for the Grid Dewar inventory."""

    grid_dewar_manager = GridDewarManager(root, id_manager=id_manager)
    grid_dewar_manager.open_main_menu()

def click_80():
    """Open the main menu for the -80°C Freezer inventory."""

    freezer80_manager = Freezer80Manager(root, id_manager=id_manager)
    freezer80_manager.open_main_menu()

def click_20():
    """Open the main menu for the -20°C Freezer inventory."""

    freezer20_manager = Freezer20Manager(root, id_manager=id_manager)
    freezer20_manager.open_main_menu()

def click_Cell():
    """Open the main menu for the Cell Dewar inventory."""
    
    cell_manager = CellDewarManager(root, id_manager=id_manager)
    cell_manager.open_main_menu()

# Define buttons (label + callback) in a list for clean iteration
buttons = [
    ("Grid Dewar", click_Grid),
    ("-80 Freezer", click_80),
    ("-20 Freezer", click_20),
    ("Cell Culture Dewar", click_Cell)
]

# Create buttons inside main_frame
for i, (text, cmd) in enumerate(buttons, start=2): # Starts below the label
    btn = tk.Button(
        main_frame,
        text=text,
        command=cmd,
        font=("Arial", 14),
        relief="groove",
        activebackground="gray"
    )
    btn.grid(row=i, column=0, sticky="nsew", padx=100, pady=10)

# Make the column expand properly
main_frame.grid_columnconfigure(0, weight=1)

# Give each button rows some "stretch" weight
for i in range(2, len(buttons) + 2):
    main_frame.grid_rowconfigure(i, weight=1)

# -------------------------------------------------------------------
# LAUNCH GUI
# -------------------------------------------------------------------

# Launch GUI
root.mainloop()