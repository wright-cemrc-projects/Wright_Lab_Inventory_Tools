import tkinter as tk

class AutocompleteEntry(tk.Entry):
    """
    A tkinter Entry widget with dropdown menu and autocomplete functionality.

    When the user types in an AutocompleteEntry box, a dropdown appears with 
    suggestions generated from the provided list. The user can navigate the 
    dropdown with arrow keys and select a suggestion using Enter or mouse click. 
    The options dynamically filter to match what the user types.

    Parameters
    ----------
    master (tk.Widget):  
        Parent tkinter widget that this Entry belongs to.  
    suggestions (list[str]):  
        List of suggestion strings to display in the dropdown.  
    on_select (callable) Default = None:  
        Callback function that is invoked when a suggestion is selected.  
    *args (tuple):  
        Additional positional arguments passed to `tk.Entry`.  
    **kwargs (dict):  
        Additional keyword arguments passed to `tk.Entry`.  

    Attributes
    ----------
    var (tk.StringVar):  
        Tracks the current text value of the entry.  
    suggestions (list[str]):  
        Sorted list of suggestion strings for autocomplete.  
    on_select (callable | None):  
        Function to call when a suggestion is chosen.  
    listbox_visible (bool):  
        Whether the dropdown listbox is currently shown.  
    listbox (tk.Toplevel | None):  
        The floating window containing the suggestion listbox.  
    lb (tk.Listbox):  
        The Listbox widget inside `listbox` that holds suggestions.  

    Methods
    -------
    def _sort_key(val: str) -> tuple[int, any]:
        Custom key for sorting mixed numbers and strings.
    _on_change(*args) -> None:  
        Triggered whenever the entry text changes; filters suggestions.  
    _on_focus_in(event) -> None:  
        Shows the full suggestion list when the entry gains focus.  
    _show_listbox(matches (list[str])) -> None:  
        Displays the dropdown with the provided matches.  
    _hide_listbox(event=None) -> None:  
        Hides the dropdown if it is visible.  
    _on_click(event) -> None:  
        Handles mouse clicks on dropdown items.  
    _select_item() -> None:  
        Inserts the selected suggestion into the entry and triggers callback.  
    _on_down(event) -> str | None:  
        Moves selection down in the suggestion list.  
    _on_up(event) -> str | None:  
        Moves selection up in the suggestion list.  
    _on_return(event) -> str:  
        Finalizes selection when Enter is pressed and moves focus forward.  
    _on_focus_out(event) -> None:  
        Delays hiding the listbox when focus is lost (to allow clicks).  
    update_suggestions(new_suggestions (list[str])) -> None:  
        Updates the suggestion list with new values.  
    """

    def __init__(self, master, suggestions, on_select=None, *args, **kwargs):
        """
        Initialize the AutocompleteEntry Object. See class docstring for parameter details.
        """

        # Create a tkinter StringVar to track the Entry's text
        self.var = tk.StringVar()
        # Link the StringVar to this Entry widget
        kwargs['textvariable'] = self.var
        # Initialize with given arguments
        super().__init__(master, *args, **kwargs)

        # Store the suggestions alphabetically and case insensitive for perdictable display
        self.suggestions = sorted(suggestions, key=self._sort_key)

        # Add a trace on the StringVar to call _on_change whenever the text is changed
        self.var.trace_add("write", self._on_change)

        # Store the on_select callback 
        self.on_select = on_select 

        # Keeps track if the dropdown list is visible
        self.listbox_visible = False
        # Holds the dropdown Toplevel when shown
        self.listbox = None

        # Key bindings for navigation and interaction
        self.bind("<Down>", self._on_down)  # Arrow down
        self.bind("<Up>", self._on_up)  # Arrow up
        self.bind("<Return>", self._on_return)  # Enter
        self.bind("<Escape>", self._hide_listbox)  # Escape
        self.bind("<FocusOut>", self._on_focus_out)  # Loose focus
        self.bind("<FocusIn>", self._on_focus_in)  # Gain focus

    @staticmethod
    def _sort_key(val: str):
        """
        Custom key for sorting mixed numbers and strings.
        
        - Blank strings ("", " ") come first.
        - Numbers next, in numeric order.
        - Strings last, in alphabetic order.
        """
            
        val = str(val).strip()

        # Handle blanks first
        if val == "":
            return (0, "")
    
        try:
            # Sort numbers by numeric order
            return (1, int(val))
        except ValueError:
            # Sort strings by alphabetic order
            return (2, val.lower())

    def _on_change(self, *args):
        """
        Called when the Entry's text changes. Filters suggestions and shows the listbox.
        """

        # Get the current text form the entry widget
        typed = self.var.get()

        # Show all suggestions if nothing is typed
        if typed == "":
            matches = [" "] + [s for s in self.suggestions if s != ""]
        # Find suggestions that contain current text
        else:
            # Case-insensitive substring match
            matches = [s for s in self.suggestions if typed.lower() in s.lower() and s != ""]

        # If we found matches, display them in the dropdown listbox
        if matches:
            self._show_listbox(matches)
        # If no matches, hide the listbox entirely
        else:
            self._hide_listbox()

    def _on_focus_in(self, event):
        """
        Show listbox on focus.
        """

        self._show_listbox(self.suggestions)

    def _show_listbox(self, matches):
        """
        Display a dropdown list of suggestions.
        
        :param matches: List of suggestion strings that match the current entry text.
        """

        # Create a listbox if not already visible
        if not self.listbox_visible:
            # Create a floating window attatched to the main window with not decorations (title, boarder)
            self.listbox = tk.Toplevel(self.winfo_toplevel())
            self.listbox.wm_overrideredirect(True)

            # Keep the listbox above the main window
            self.listbox.wm_attributes("-topmost", True)

            # Ensure geometery information is up to date
            self.update_idletasks()

            # Calculate position that is just below the entry widget
            x = self.winfo_rootx()
            y = self.winfo_rooty() + self.winfo_height()
            width = self.winfo_width()

            # Set the popup size (width = entry width, height = 100px)
            self.listbox.wm_geometry(f"{width}x100+{x}+{y}")

            # Raise the popup above siblings in the stacking order
            self.listbox.lift()

            # Create the listbox inside the popup window
            self.lb = tk.Listbox(self.listbox)
            self.lb.pack(fill="both", expand=True)

            # Bind mouse click and Enter key to selection handlers
            self.lb.bind("<ButtonRelease-1>", self._on_click)
            self.lb.bind("<Return>", self._on_return)

            # mark that the listbox is now visible
            self.listbox_visible = True

        # Clear previous suggestions
        self.lb.delete(0, tk.END)
        # Update listbox with matching suggestions
        for item in matches:
            self.lb.insert(tk.END, item)

        # Select the first item by default
        self.lb.select_set(0)
        self.lb.activate(0)

    def _hide_listbox(self, event=None):
        """
        Hides the listbox dropdown if visible.
        """

        if self.listbox_visible:
            self.listbox.destroy()
            self.listbox_visible = False

    def _on_click(self, event):
        """
        Handles user clicking an item in the listbox.
        """

        if self.lb is not None:
            index = self.lb.nearest(event.y)
            if index >= 0:
                self.lb.select_clear(0, tk.END)
                self.lb.select_set(index)
                self.lb.activate(index)
                self._select_item()
                self.icursor(tk.END)
                self.focus_set()

    def _select_item(self):
        """
        Inserts the selected suggestion into the Entry and hides the listbox.
        """

        # If no item is currently selected in the listbox, exit early
        if not self.lb.curselection():
            return
        
        # Get the index of the currently selected item and retrieve the value (text) from that index
        index = self.lb.curselection()[0]
        value = self.lb.get(index)

        # Set the Entry widget's text to the chosen suggestion and hide the listbox
        self.var.set(value)
        self._hide_listbox()

        # If an on_select callback was provided, trigger it
        if self.on_select:
            self.on_select(value)

    def _on_down(self, event):
        """
        Moves selection down in the suggestion list.
        """

        if self.listbox_visible:
            curr = self.lb.curselection()
            index = 0 if not curr else (curr[0] + 1) % self.lb.size()
            self.lb.select_clear(0, tk.END)
            self.lb.select_set(index)
            self.lb.activate(index)
            self.lb.see(index)
            return "break"

    def _on_up(self, event):
        """
        Moves selection up in the suggestion list.
        """

        if self.listbox_visible:
            curr = self.lb.curselection()
            index = self.lb.size() - 1 if not curr else (curr[0] - 1) % self.lb.size()
            self.lb.select_clear(0, tk.END)
            self.lb.select_set(index)
            self.lb.activate(index)
            self.lb.see(index)
            return "break"

    def _on_return(self, event):
        """
        Finalizes selection when Enter is pressed.
        """

        if self.listbox_visible:
            self._select_item()

        # Move focus to next widget
        self.tk_focusNext().focus()

        return "break"

    def _on_focus_out(self, event):
        """
        Delays hiding the listbox slightly to allow click events.
        """

        self.after(100, self._hide_listbox)

    def update_suggestions(self, new_suggestions):
        """
        Updates the internal suggestions dropdown list.

        :param new_suggestions: List of new suggestion strings.
        """

        self.suggestions = sorted(new_suggestions, key=self._sort_key)