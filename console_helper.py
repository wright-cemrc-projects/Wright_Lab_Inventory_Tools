import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import sys
import app_context

class ConsoleRedirector:
    """
    Redirect printed terminal output (stdout/stderr) to the bottom Tkinter widget while still appearing in the terminal.

    - Allows multiple Tkinter Text widgets to receive live console output.  
    - Maintains a history buffer of printed messages in `app_context.console_history`.  
    - Ensures output is also written to the original system terminal.  

    Attributes
    ----------
    targets (list[tk.Text]):  
        Tkinter text widgets that will receive redirected output.  
    original_stdout (io.TextIOBase):  
        Reference to the actual terminal stdout object.  

    Methods
    -------
    add_target(widget (tk.Text)) -> None:  
        Adds a Tkinter widget as a target for terminal output.  
    write(message (str)) -> None:  
        Writes text to all registered widgets and to the original terminal.  
    flush() -> None:  
        Dummy flush method required for sys.stdout redirection compatibility.  
    """

    def __init__(self):
        """Initialize ConsoleRedirector. See class docstring for attributes."""

        # Target Tkinter widgets for terminal output
        self.targets = []
        # Keep reference to the actual terminal stdout
        self.original_stdout = sys.__stdout__

    def add_target(self, widget):
        """Add a Tkinter Text widget as a target for redirected output"""

        self.targets.append(widget)

    def write(self, message):
        """Called whenever text is written to stdout/stderr. Distributes text to all widgets and the terminal."""

        # Keep a history buffer of the terminal output
        app_context.console_history.append(message)

        # Send message to each widget in targets
        for widget in self.targets[:]:
            try:
                widget.configure(state="normal")  # Enable editing temporarily
                widget.insert(tk.END, message)  # Insert the new message
                widget.configure(state="disabled")  # Lock again
                widget.see(tk.END)  # Auto-scroll to bottom
                widget.update_idletasks()   # Force refresh immediately
            except tk.TclError:
                # If the widget is destroyed, remove it from targets
                self.targets.remove(widget)

        # Always write to the original terminal as well
        try:
            sys.__stdout__.write(message)  # Always write to terminal too
        except Exception:
            pass

    def flush(self):
        """Dummy flush method required for compatibility with sys.stdout redirection"""

        pass

# Initialize global redirector here (not in app_context.py to avoid loop)
app_context.console_redirector = ConsoleRedirector()

class ConsoleWindow(tk.Toplevel):
    """
    A floating window that displays the redirected console output from a ConsoleRedirector.

    - Attaches to a parent Tkinter window and docks directly below it.  
    - Automatically follows parent window movement and resizing.  
    - Displays console text in a read-only, scrollable Text widget.  

    Parameters
    ----------
    parent_window (tk.Widget):  
        The parent Tkinter window this console is attached to.  
    output_redirector (ConsoleRedirector):  
        The redirector instance responsible for forwarding console output into this window.  

    Attributes
    ----------
    parent_window (tk.Widget):  
        The parent window that this console is linked to.  
    output_redirector (ConsoleRedirector):  
        Redirector that manages console output forwarding.  
    text (tk.Text):  
        The read-only text widget that displays redirected output.  
    bind_id (str):  
        The binding ID for tracking parent window movement/resize events.  

    Methods
    -------
    _position_console() -> None:  
        Positions the console window directly below its parent window.  
    _on_parent_configure(event (tk.Event)) -> None:  
        Repositions the console whenever the parent window moves or resizes.  
    on_close() -> None:  
        Cleans up event bindings and removes this console from the redirector’s targets before destroying the window.  
    """

    def __init__(self, parent_window, output_redirector):
        """Initializes ConsoleWindow. See class docstring for parameter/attribute details."""

        super().__init__(parent_window)
        self.parent_window = parent_window
        self.title("Console Output")
        self.geometry("700x150")
        self.output_redirector = output_redirector
        
        # Create a text widget for the output
        self.text = tk.Text(self, bg="#f5f5f5", fg="#333333", font=("Courier New", 9), wrap="word", state="disabled")
        self.text.pack(fill="both", expand=True)

        output_redirector.add_target(self.text)

        # Bind parent window movements/resizes and save binding ID
        self.bind_id = self.parent_window.bind("<Configure>", self._on_parent_configure)
        # Position console initially
        self._position_console()

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _position_console(self):
        """ Position the window directly below its parent window."""

        x = self.parent_window.winfo_x()
        y = self.parent_window.winfo_y() + self.parent_window.winfo_height()
        self.geometry(f"+{x}+{y}")

    def _on_parent_configure(self, event):
        """Called whenever the parent window moves/resizes to reposition the console."""

        self._position_console()

    def on_close(self):
        """Clean up event that binds and removes this text widget from the redirector targets."""

        if hasattr(self, 'bind_id'):
            self.parent_window.unbind("<Configure>", self.bind_id)
        if self.text in self.output_redirector.targets:
            self.output_redirector.targets.remove(self.text)
        self.destroy()

class ConsoleFooter(tk.Frame):
    """
    A footer-style console widget with scrollback, attached to the bottom of an application window.

    - Redirects console output (stdout and stderr) into a scrollable Tkinter Text widget.  
    - Preserves history using `app_context.console_history`.  
    - Integrates with `app_context.console_redirector` for automatic output redirection.  

    Parameters
    ----------
    parent (tk.Widget):  
        The parent Tkinter widget that will contain this console footer.  

    Attributes
    ----------
    text (tk.Text):  
        The read-only text widget used to display redirected console output.  
    scrollbar (tk.Scrollbar):  
        The vertical scrollbar linked to the text widget for navigation.  

    Methods
    -------
    _load_history() -> None:  
        Loads previous console output from `app_context.console_history` into this widget during initialization.
    """

    def __init__(self, parent):
        """Initializes ConsoleFooter. See class docstring for parameter/attribute details."""

        # Initialize this widget as a Tkinter Frame with the given parent and background color
        super().__init__(parent, bg="#f5f5f5")

        # Console text area
        self.text = tk.Text(self, height=6, bg="#f5f5f5", fg="#333", font=("Courier New", 9), wrap="word", state="disabled")
        self.text.pack(side="left", fill="both", expand=True)

        # Add scrollbar
        scrollbar = tk.Scrollbar(self, command=self.text.yview)
        scrollbar.pack(side="right", fill="y")
        self.text.config(yscrollcommand=scrollbar.set)

        # Add this widget as a redirector target
        app_context.console_redirector.add_target(self.text)

        # Load any saved console history
        self._load_history()

    def _load_history(self):
        """Load previous console output from global history into this widget."""

        self.text.config(state="normal")
        for line in app_context.console_history:
            self.text.insert(tk.END, line)
        self.text.config(state="disabled")
        self.text.see(tk.END)