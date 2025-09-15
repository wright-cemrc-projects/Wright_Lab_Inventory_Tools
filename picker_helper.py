import tkinter as tk
from tkinter import messagebox
import pandas as pd

class PickerHelper:
    """
    A helper class that provides interactive pickers for selecting inventory locations in a tkinter-based application.

    Two picker styles are supported:
      1. Rectangular picker (grid of buttons) for box-based storage.
      2. Circular picker (canvas of circle "buttons") for grid-style dewars.

    Behavior:
    - In "adding" mode, prevents selection of already-occupied locations.
    - In "removing" mode, prevents selection of empty locations.
    - Highlights selected locations in light blue.
    - Updates the linked tkinter Entry field with the chosen values when done.

    Parameters
    ----------
    parent (tk.Widget):
        The parent Tkinter window from which the picker is launched.
    entries (dict):
        Dictionary mapping entry field names to their corresponding Tkinter Entry widgets.
    rows_df (pd.DataFrame):
        DataFrame containing the inventory data.
    top_label (tk.Widget, optional): Default = None
        Label widget for displaying messages or warnings.
    adding (bool, optional): Default = True
        True if selecting locations to add inventory, False if removing.
    grid_coords (tuple[int, int], optional): Default = (9, 9)
        Number of rows and columns for the rectangular picker grid.
    for_filled (list, optional): Default = []
        List of entry field names used to determine filled locations.
    location_column (str, optional): Default = ""
        Column name in `rows_df` representing the location ID.
    label_column (str, optional): Default = ""
        Column name in `rows_df` representing the label for each location.
    rectangle_picker (bool, optional): Default = True
        True to use a rectangular grid picker, False to use a circular grid picker.

    Attributes
    ----------
    required_to_validate (list):
        Entry fields that must be filled before opening the picker.
    selected (set[str]):
        Set of currently selected location IDs.
    circle_items (dict):
        Maps location IDs to canvas item IDs for circular picker buttons.
    filled_locations (set[str]):
        Set of location IDs already occupied in the inventory.
    picker (tk.Toplevel):
        The popup window displaying the picker.
    message_label (tk.Label):
        Label displaying instructions or warnings within the picker.

    Methods
    -------
    _hide_all_dropdowns() -> None:
        Hides autocomplete dropdowns for all entry fields.
    _validate_and_extract() -> bool:
        Validates required entries, extracts values, and identifies filled locations.
    _toggle(cell (str), btn (tk.Widget, optional) = None) -> None:
        Toggles the selection state of a given location and updates its appearance.
    _finish_selection(int_type_location (bool, optional) = False) -> None:
        Finalizes the selection and writes selected locations back to the linked entry.
    _disabled_click_warning(coord (str), mode (str)) -> None:
        Shows a warning when trying to select an unavailable location.
    get_box_dimensions() -> tuple[int, int]:
        Retrieves box dimensions (rows, columns) from the DataFrame for variable-sized boxes.
    _select_all(self) -> None:
        Select all valid locations based on mode: empty for adding, filled for removing.
    _update_button_color(self, coord, color):
        Updates the visual color of a button or canvas circle for a given location.
    _create_picker() -> None:
        Builds a rectangular picker popup with button grid for location selection.
    _create_circle_button(canvas (tk.Canvas), x (int), y (int), r (int), text (str), command (callable), filled (bool, optional) = False) -> int:
        Draws a circular clickable button on a Canvas and returns its ID.
    _create_grid_picker() -> None:
        Builds a circular grid picker popup with canvas-based buttons.
    """
    
    def __init__(self, parent, entries, rows_df, top_label=None, adding=True, grid_coords=(9,9), for_filled=[], location_column="", label_column="", rectangle_picker=True):
        """Initializes the PickerHelper. See class docstring for parameter/attribute details."""

        # Store references
        self.parent = parent
        self.entries = entries
        self.df = rows_df
        self.top_label = top_label
        self.adding = adding
        self.grid_coords = grid_coords
        self.required_to_validate = for_filled
        self.location_column = location_column
        self.label_column = label_column

        # Track selected locations
        self.selected = set()
        # Create a dict to store coordinates for grid dewar picker buttons
        self.circle_items = {}

        # Hide AutocompleteEntry dropdowns when the picker opens
        self._hide_all_dropdowns()

        # Ensure required fields exist and extract constraints from entries
        if not self._validate_and_extract():
            if self.top_label:
                self.top_label.config(text=f"Error: Required fields (" + ", ".join(for_filled) + f") are missing or invalid.", fg="red")
            return

        # Launch appropriate picker style
        if rectangle_picker:
            self._create_picker()
        else:
            self._create_grid_picker()

    def _hide_all_dropdowns(self):
        """Hides dropdown suggestion boxes for all input fields if they exist."""

        for entry in self.entries.values():
            if hasattr(entry, "_hide_listbox"):
                entry._hide_listbox()

    def _validate_and_extract(self):
        """
        Validates the required input fields and extracts their values.
        Also determines which locations are already filled.

        :return (bool): True if validation succeeds, False otherwise.
        """
        self.required_values = {}

        # --- STEP 1: Validate user input from required fields ---
        # Ensure that all required values are filled and store them in a dictionary after cleaning
        try:
            for field in self.required_to_validate:
                # Get the raw user input from the Tkinter Entry widget 
                value = self.entries[field].get().strip()
                # If it is empty, fail validation
                if not value:
                    raise ValueError(f"{field} is required.")
                
                # Convert to int if possible, otherwise keep as string. Falls back to string if not convertible
                try:
                    value = int(value)
                except ValueError:
                    pass  # Leave as string if not a number
                
                # Store the cleaned value
                self.required_values[field] = value
        # Show error messages in the GUI label if one occurs
        except ValueError as e:
            if self.top_label:
                self.top_label.config(
                    text=str(e),
                    fg="red",
                )
            return False

        # --- STEP 2: Normalize location column reference ---
        # location_column may be specified either by index (int) or name (str).
        if isinstance(self.location_column, int):
            # Convert the index to the column string name
            try:
                self.normalized_location_column = self.df.columns[self.location_column]
            # Show an error for an invalid index
            except IndexError:
                if self.top_label:
                    self.top_label.config(text=f"Error: location_column index {self.location_column} is out of range.", fg="red")
                return False
            
        elif isinstance(self.location_column, str):
            # Verify that the provided name is actually in the DataFrame header. Return False if not.
            if self.location_column not in self.df.columns:
                if self.top_label:
                    self.top_label.config(text=f"Error: column '{self.location_column}' not found in DataFrame.", fg="red")
                return False
            self.normalized_location_column = self.location_column

        else:
            if self.top_label:
                self.top_label.config(text=f"Error: location_column must be str or int.", fg="red")
            return False

        # --- STEP 3: Apply row filtering based on required field values ---
        # Start with a mask that selects all rows
        mask = pd.Series(True, index=self.df.index)

        for field, value in self.required_values.items():
            if field not in self.df.columns:
                # If the required field is not in the DataFrame, provide a warning, but continue
                print(f"Warning: {field} not found in DataFrame columns.")
                continue
            
            col_val = self.df[field]

            try:
                # Try numeric comparison if both are numeric
                col_numeric = pd.to_numeric(col_val, errors='coerce')
                val_numeric = float(value)
                # Compare as integers
                mask &= col_numeric.fillna(-1).astype(int) == int(val_numeric)
            except (ValueError, TypeError):
                # Fallback to string comparison
                mask &= col_val.astype(str).str.strip() == str(value).strip()

        # --- STEP 4: Extract filled locations from filtered rows ---
        self.filled_locations = set(
            self.df.loc[mask, self.normalized_location_column].dropna().astype(str).str.upper().str.strip()
        )

        # Return True if everything succeded
        return True

    def _toggle(self, cell, btn=None):
        """
        Toggle the selection state of a cell (e.g., 'A1') and update its color based on: selected, unselected, or already filled.

        :param cell (str): The cell location (e.g., 'A1').
        :param btn (tk.Widget): tkinter Button widget to update its background.

        Legend
        -------------
        - Selected cell .......... lightblue
        - Unselected + filled .... dark gray (#A9A9A9)
        - Unselected + empty ..... default system button color ("SystemButtonFace")
        """

        # If cell is already selected, deslect it and store its fill color
        # Gray (#A9A9A9) for filled or system button color for empty
        if cell in self.selected:
            self.selected.remove(cell)
            fill = "#A9A9A9" if cell in self.filled_locations else "SystemButtonFace"
        # If cell is not already selected, select it and set its fill color to light blue
        else:
            self.selected.add(cell)
            fill = "lightblue"

        # Update the fill color for tkinter button (rectangle picker)
        if btn:
            btn.config(bg=fill)
        # Update the fill color for canvas button (grid dewar picker)
        else:
            circle_id = self.circle_items.get(cell)
            if circle_id:
                self.picker.children['!canvas'].itemconfig(circle_id, fill=fill)

    def _finish_selection(self, int_type_location=False):
        """
        Finalize the selection process.

        Collects the user's chosen locations, formats them into a string,
        inserts that string into the designated entry field, and closes the picker window.

        :param int_type_location (bool): If True, formats the selected values as integers and sorts them in numeric order.
        """

        # Sort and join the selected location IDs into a single comma seperated string
        if int_type_location:
            value = ", ".join(sorted(self.selected, key=int))
        else:
            value = ", ".join(sorted(self.selected))

        # Delete what is currently in the Entry box
        self.entries[self.location_column].delete(0, tk.END)
        # Insert the new formatted string
        self.entries[self.location_column].insert(0, value)
        # Close the picker pop-up
        self.picker.destroy()
        # Move the focus to the next entry box
        self.entries[self.label_column].focus_set()

    def _disabled_click_warning(self, coord, mode):
        """
        Displays a warning when a user tries to select an unavailable location (occupied/empty).

        :param coord (str): The location (Ex: A1, C3).
        :param mode (str): 'add' or 'remove' mode.
        """

        # Provides a message if you attempt to add to a location that is already occupied
        if mode == "add":
            msg = f"Location {coord} is already occupied. You can't add there."
        # Provides a message if you attempt to remove from a location that is empty
        elif mode == "remove":
            msg = f"Location {coord} is empty. You can't remove from there."
        else:
            msg = f"Location {coord} is not available."

        # Format the message in red
        if self.top_label:
            self.message_label.config(text=msg, fg="red")
        else:
            messagebox.showwarning("Invalid Action", msg)

    def get_box_dimensions(self):
        """Retrieve box dimensions (rows x cols) from the DataFrame for inventories that have changing box dimensions (-80C)."""

        # Start mask as all True
        mask = pd.Series(True, index=self.df.index)

        # Apply filters based on user-selected entries
        for col in self.required_to_validate:
            val = self.entries[col].get()
            mask &= self.df[col].astype(str).str.strip() == str(val).strip()

        # Check if any row matches
        if not mask.any():
            print("No matching rows found")
            return (0, 0)  # No matching box found

        # Get first matching row's dimensions
        dimensions_str = self.df.loc[mask, "Box Dimensions"].iloc[0]
        
        # Catch invalid dimensions format (must be rrxcc)
        if not isinstance(dimensions_str, str) or "x" not in dimensions_str:
            print("Dimensions string is invalid format")
            return (0, 0)  # Invalid format
        
        # Parse "rows x cols" formated string to return a tuple (rows, cols)
        try:
            rows, cols = map(int, dimensions_str.lower().split("x"))
            return (rows, cols)
        except ValueError:
            print("Error parsing dimensions string to int")
            return (0, 0)
        
    def _select_all(self):
        """
        Select all valid locations based on mode: empty for adding, filled for removing.
        
        If all valid locations are already selected, then deselect all.
        """

        # Determine if it is the circular picker
        if getattr(self, "circle_items", None):
            coords = self.circle_items.keys()
        # Determine if it is the rectangular picker
        elif getattr(self, "rloc_buttons", None):
            coords = self.rloc_buttons.keys()
        else:
            return  # no buttons found

        # Determine which coords are valid to select
        valid_coords = set()
        for coord in coords:
            # Determine if the coordinate is a filled location
            is_filled = coord in self.filled_locations
            # Detirmine if valid (adding and empty) or (removing and filled)
            if (self.adding and not is_filled) or (not self.adding and is_filled):
                valid_coords.add(coord)

        # If all valid locations are already selected, deselect all
        if valid_coords.issubset(self.selected):
            for coord in valid_coords:
                self.selected.remove(coord)
                self._update_button_color(coord, "#A9A9A9" if coord in self.filled_locations else "SystemButtonFace")
        else:
            # Otherwise, select all valid locations
            for coord in valid_coords:
                self.selected.add(coord)
                self._update_button_color(coord, "lightblue")
    
    def _update_button_color(self, coord, color):
        """Updates the visual color of a button or canvas circle for a given location."""

        # Handles rectangle picker button color
        if coord in getattr(self, "rloc_buttons", {}):
            self.rloc_buttons[coord].config(bg=color)
        # Handles circular picker button color
        elif coord in getattr(self, "circle_items", {}):
            circle_id = self.circle_items[coord]
            self.picker.children['!canvas'].itemconfig(circle_id, fill=color)
                    
    def _create_picker(self):
        """Build rectangular picker grid picker popup with buttons for inventory locations."""

        # If the grid dimensions (rows x cols) are not yet know (such as -80C) fetch them dynamically from the DataFrame
        if self.grid_coords == "Check":
            num_rows, num_cols = self.get_box_dimensions()
            # Catch if there are missing or invalid dimensions.
            # get_box_dimensions() returns zeros if it does not find coordinates 
            if num_rows ==0 or num_rows ==0:
                messagebox.showerror("error", "Invalid box dimensions")
                return
            # Save the determined grid dimensions to the grid_coords variable
            self.grid_coords = (num_rows, num_cols)

        # Parse any existing entries into a set of selected locations
        existing = self.entries[self.location_column].get()
        if existing:
            self.selected = set(loc.strip().upper() for loc in existing.split(",") if loc.strip())
        else:
            self.selected = set()

        # Create a pop-up window
        self.picker = tk.Toplevel(self.parent)

        # Build a descriptive title of your current inventory context
        title_parts = ", ".join(f"{k}: {v}" for k, v in self.required_values.items())
        self.picker.title(f"Selection ({title_parts})")
        
        # Show an instruction message for the user
        instruction_text = f"Pick {self.location_column}(s) you would like to add" if self.adding else f"Pick {self.location_column}(s) you would like to remove"
        self.message_label = tk.Label(self.picker, text=instruction_text, fg="black", font=("Arial", 12, "bold"))
        self.message_label.grid(row=0, column=0, columnspan=9, pady=(10, 5))

        # Allow each grid column to expand evenly within the window
        for col in range(self.grid_coords[1]):
            self.picker.grid_columnconfigure(col, weight=1)

        # Generate grid labels
        # Rows = letters (A, B, C...), Cols = numbers (1, 2, 3...)
        rows = [chr(ord('A') + i) for i in range(self.grid_coords[0])]
        cols = list(range(1, self.grid_coords[1] + 1))

        # Store rectangular location buttons here
        self.rloc_buttons = {}

        # Build a button for every grid cell (e.g. A1, B2...)
        for r, row_letter in enumerate(rows, start=1):
            for c, col_num in enumerate(cols):
                coord = f"{row_letter}{col_num}"
                is_filled = coord in self.filled_locations
                btn = tk.Button(self.picker, text=coord, width=4)

                # Button behavior depends on if inventory is being added or removed
                if self.adding:  # Adding mode
                    if is_filled:
                        if coord in self.selected:
                            # If an already filled location is in selected, show a warning message and remove it
                            self._disabled_click_warning(coord, mode="add")
                            self.selected.remove(coord)
                        # If cell is occupied create a disabled gray button. Provide a warning if the user clicks on it.
                        btn.config(bg="#A9A9A9", command=lambda c=coord: self._disabled_click_warning(c, mode="add")) 
                    else:
                        # If cell is unoccupied but selected (alrady in entry box) fill it lightblue
                        if coord in self.selected:
                            btn.config(bg="lightblue")
                        # An empty location can be toggled
                        btn.config(command=lambda c=coord, b=btn: self._toggle(c, b))

                elif self.adding == False: # Removing mode
                    if is_filled:
                        if coord in self.selected:
                            # If a cell is occupied and selected fill it lightblue
                            btn.config(bg="lightblue")
                        else:
                            # If only occupied fill it gray
                            btn.config(bg="#A9A9A9")
                        # Occupied location can be toggled
                        btn.config(command=lambda c=coord, b=btn: self._toggle(c, b))
                    else:
                        if coord in self.selected:
                            # If a cell is unoccupied but selected, show a warning message and remove it
                            self._disabled_click_warning(coord, mode="remove")
                            self.selected.remove(coord)
                        # If user clicks on an empty cell provide a warning
                        btn.config(command=lambda c=coord: self._disabled_click_warning(c, mode="remove"))

                else:
                    print("'adding' argument must either be True or False!")

                # Place the button into the grid layout.
                btn.grid(row=r, column=c, padx=2, pady=2, sticky="ew")
                # Store button as a location button
                self.rloc_buttons[coord] = btn

        # Create a frame at the bottom for the action buttons (reset, add/remove)
        button_frame = tk.Frame(self.picker)
        button_frame.grid(row=(self.grid_coords[0] + 1), column=0, columnspan=self.grid_coords[1], pady=10, padx=5, sticky="ew")
        self.picker.grid_columnconfigure(0, weight=1)

        # Allow equal resizing of both columns in button row
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        # Add "Select/Deselect All" button to select all non disabled location buttons
        select_all_btn = tk.Button(button_frame, text="Select/Deselect All", command=self._select_all)
        select_all_btn.grid(row=0, column=0, padx=5, sticky="nsew")
        # Add "Done" button at the bottom to finish selection
        done_btn = tk.Button(button_frame, text="Done", command=lambda: self._finish_selection(int_type_location=False))
        done_btn.grid(row=0, column=1, padx=5, sticky="nsew")

    def _create_circle_button(self, canvas, x, y, r, text, command, filled=False):
        """
        Draw a circular "button" on a tkinter Canvas.

        :param canvas: The tkinter Canvas widget where the circle is drawn.
        :param x: X coordinate of the circle center.
        :param y: Y coordinate of the circle center.
        :param r: Radius of the circle.
        :param text: Label to display inside the circle.
        :param command: Command to executewhen the circle or label is clicked.
        :param filled: If True, fill the circle gray and treat as occupied. If False treat as empty.
        :return: The canvas item ID of the circle drawn.
        """

        # Determine fill color: gray if occupied, otherwise system default
        fill_color = "#A9A9A9" if filled else "SystemButtonFace"

        # Draw the cicular shape and label on the canvas
        circle_id = canvas.create_oval(x - r, y - r, x + r, y + r, fill=fill_color, outline="black")
        label_id = canvas.create_text(x, y, text=text, font=("Arial", 10, "bold"))

        # Define a click handler that calls the provided command
        # and passes the circle's label (text) as argument.
        def on_click(event, label=text):
            command(label)

        # Bind both circle and label to the same click handler so clicking anywhere on the button triggers the action
        canvas.tag_bind(circle_id, "<Button-1>", on_click)
        canvas.tag_bind(label_id, "<Button-1>", on_click)

        # Return the ID of the circle so its fill color can be updated later.
        return circle_id

    def _create_grid_picker(self):
        """
        Creates a custom grid picker with circular "buttons" on a Canvas.

        This method builds a pop-up (Toplevel) window where the user can select
        locations represented by circles. Circles can be:
        - gray if already occupied,
        - light blue if currently selected by the user,
        - default if empty.
        """

        # Parse any existing entries into a set of selected locations
        existing = self.entries[self.location_column].get()

        if existing:
            self.selected = set(loc.strip().upper() for loc in existing.split(",") if loc.strip())
        else:
            self.selected = set()

        # Create the pop-up selection window
        self.picker = tk.Toplevel(self.parent)

        # Build a descriptive title of your current inventory context
        title_parts = ", ".join(f"{k}: {v}" for k, v in self.required_values.items())
        self.picker.title(f"Selection ({title_parts})")

        # Show an instruction message for the user
        instruction_text = f"Pick {self.location_column}(s) you would like to add" if self.adding else f"Pick {self.location_column}(s) you would like to remove"
        self.message_label = tk.Label(self.picker, text=instruction_text, fg="black", font=("Arial", 12, "bold"))
        self.message_label.grid(row=0, column=0, columnspan=9, pady=(10, 5))

        # Logical description of layout (row, starting column, count),
        # not directly used for drawing, but helps track rows and columns.
        layout = [
            (1, 1, 2),
            (2, 0, 3),
            (3, 0, 4),
            (4, 1, 3),
        ]

        # Canvas that holds all circular buttons
        canvas = tk.Canvas(self.picker, width=300, height=300)
        canvas.grid(row=1, column=0, columnspan=9)

        # Fixed coordinates of circle centers for visual arrangement
        positions = [
            (120, 40), (180, 40),
            (100, 90), (150, 90), (200, 90),
            (75, 140), (125, 140), (175, 140), (225, 140),
            (100, 190), (150, 190), (200, 190)
        ]

        # Create each circular button in the grid
        for i, (x, y) in enumerate(positions, start=1):
            coord = str(i)

            # Determine if the location is filled
            is_filled = coord in self.filled_locations

            # Button behavior depends on if inventory is being added or removed
            if self.adding:  # Adding mode
                if is_filled:
                    if coord in self.selected:
                        # If an already filled location is in selected, show a warning message and remove it
                        self._disabled_click_warning(coord, mode="add")
                        self.selected.remove(coord)
                    # Draw disabled gray circle and provide a warnign if clicked
                    circle_id = self._create_circle_button(
                        canvas, x, y, 20, coord,
                        lambda c=coord: self._disabled_click_warning(c, mode="add"),
                        filled=True
                    )
                    canvas.itemconfig(circle_id, fill="#A9A9A9")  # gray
                else:
                    # An empty location can be toggled
                    circle_id = self._create_circle_button(
                        canvas, x, y, 20, coord,
                        lambda c=coord: self._toggle(c, None),
                        filled=False
                    )
                    if coord in self.selected:
                        # If cell is unoccupied but selected (alrady in entry box) fill it lightblue
                        canvas.itemconfig(circle_id, fill="lightblue")
                
            elif self.adding == False:  # Removing mode
                if is_filled:
                    # Occupied location can be toggled
                    circle_id = self._create_circle_button(
                        canvas, x, y, 20, coord,
                        lambda c=coord: self._toggle(c, None),
                        filled=True
                    )
                    if coord in self.selected:
                        # If a cell is occupied and selected fill it lightblue
                        canvas.itemconfig(circle_id, fill="lightblue")
                    else:
                        # If only occupied fill it gray
                        canvas.itemconfig(circle_id, fill="#A9A9A9")
                else:
                    if coord in self.selected:
                        # If a cell is unoccupied but selected, show a warning message and remove it
                        self._disabled_click_warning(coord, mode="remove")
                        self.selected.remove(coord)
                    # If user clicks on an empty cell provide a warning
                    circle_id = self._create_circle_button(
                        canvas, x, y, 20, coord,
                        lambda c=coord: self._disabled_click_warning(c, mode="remove"),
                        filled=False
                    )
            
            else:
                print("'adding' argument must either be True or False!")
            
            # Store circle reference to apply changes later
            self.circle_items[coord] = circle_id

        # Place the "Done" button below the grid
        final_row = max(r for r, _, _ in layout) + 1
        total_columns = max(c + count for r, c, count in layout)

        
        # Create a frame at the bottom for the action buttons (select all/done)
        button_frame = tk.Frame(self.picker)
        button_frame.grid(row=2, column=0, columnspan=9, pady=10, padx=5, sticky="ew")

        # Allow equal resizing of both columns in button row
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        # Add "Select/Deselect All" button to select all non disabled location buttons
        select_all_btn = tk.Button(button_frame, text="Select/Deselect All", command=self._select_all)
        select_all_btn.grid(row=0, column=0, padx=5, sticky="ew")
        # Add "Done" button at the bottom to finish selection
        done_btn = tk.Button(button_frame, text="Done", command=lambda: self._finish_selection(int_type_location=False))
        done_btn.grid(row=0, column=1, padx=5, sticky="ew")