import tkinter as tk
import sys
from console_helper import ConsoleFooter
import app_context

class ToplevelWindowHelper:
    """
    A helper class for creating Tkinter Toplevel windows with navigation history and an attached console footer.

    This class provides a structured way to open new top-level windows that:
    - Include a main content frame for custom widgets
    - Attach a console footer for live program output
    - Optionally provide a "Back" button for navigation
    - Manage a history stack for returning to previous windows
    - Handle clean program exit when the top-level window is closed

    Parameters
    ----------
    parent_window (tk.Widget):  
        The parent Tkinter window or frame.  
    title (str, optional):  Default = "New Window"
        Title of the new window (default = "New Window").  
    size (str, optional):  Default = "700x700"
        Window size in "widthxheight" format (default = "700x700").  
    show_back_button (bool, optional):  Default = True
        Whether to include a "Back" button in the window (default = True).  

    Attributes
    ----------
    parent (tk.Widget):  
        Reference to the parent window.  
    window (tk.Toplevel):  
        The created top-level window instance.  
    container (tk.Frame):  
        Container frame managing layout for main content and console.  
    main_frame (tk.Frame):  
        Frame where user content/widgets should be added.  
    console (ConsoleFooter):  
        Console footer widget at the bottom of the window.  
    history_stack (list[tk.Widget]):  
        Class-level stack tracking previously opened parent windows.  

    Methods
    -------
    _on_close():  
        Closes all Tk windows and exits the program.  
    _go_back():  
        Closes the current window and reopens the previous window in history.  
    get_window() -> tk.Toplevel:  
        Returns the Toplevel window object for advanced customization.  
    get_main_frame() -> tk.Frame:  
        Returns the main content frame for adding user widgets.  
    """

    # List to track the stack of Tkinter window names to allow for "Back" functionality
    history_stack = []

    def __init__(self, parent_window, title="New Window", size="700x700", show_back_button=True):
        """Initializes ToplevelWindowHelper. See class docstring for prameter/attribute details."""

        self.parent = parent_window
        self.window = tk.Toplevel(parent_window)
        self.window.title(title)

        # Get parent geometry
        parent_x = parent_window.winfo_x()
        parent_y = parent_window.winfo_y()
        # Set window inventory
        self.window.geometry(f"{size}+{parent_x}+{parent_y}")

        # Hide parent window if it has that method
        if hasattr(self.parent, "withdraw"):
            self.parent.withdraw()

        # Push parent to the history stack
        ToplevelWindowHelper.history_stack.append(self.parent)

        # Create a container for layout
        self.container = tk.Frame(self.window)
        self.container.grid(row=0, column=0, sticky="nsew")
        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_columnconfigure(0, weight=1)

        # Main content frame (you can access this for your UI)
        self.main_frame = tk.Frame(self.container)
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        # Console at bottom
        self.console = ConsoleFooter(self.container)
        self.console.grid(row=1, column=0, sticky="ew")
        # Configure row weights to stretch main_frame, not console
        self.container.grid_rowconfigure(0, weight=1)  # main_frame expands
        self.container.grid_rowconfigure(1, weight=0)  # console stays fixed
        self.container.grid_columnconfigure(0, weight=1)

        if show_back_button:
            back_button = tk.Button(self.main_frame, text="Back", command=self._go_back)
            back_button.grid(row=0, column=0, sticky="nw", padx=10, pady=10)

        self.window.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        """Close all Tk windows and eit program when the top-level is closed."""

        for widget in tk._default_root.winfo_children():
            widget.destroy()
        sys.exit()

    def _go_back(self):
        """Return to the previous window in history, if available."""

        # TODO: disabled for now.
        # Give a warning message with "Retry" and "Cancel" options.

        #Save the position of the current window
        x, y = self.window.winfo_x(), self.window.winfo_y()

        self.window.destroy()

        # Open previous window
        if ToplevelWindowHelper.history_stack:
            previous = ToplevelWindowHelper.history_stack.pop()
            if hasattr(previous, 'deiconify'):
                # Keep the current window location
                previous.geometry(f"+{x}+{y}")
                # Open previous window
                previous.deiconify()
                


    def get_window(self):
        """Return the Tkinter Toplevel window object (for advanced customization)."""

        return self.window

    def get_main_frame(self):
        """
        Return the main content frame (where user widgets should be placed).
        
        - Allows for the addition of widgets into their own fram without messing with the console below.
        """

        return self.main_frame