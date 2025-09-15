import tkinter as tk
from tkinter import messagebox
import pandas as pd
from autocomplete import AutocompleteEntry
from dropdown_helper import DropdownHelper
from window_helper import ToplevelWindowHelper
from picker_helper import PickerHelper
import re
from window_helper import ToplevelWindowHelper

class dataAddWindows:
    """
    Manage the setup of inventory add and remove windows in a tkinter application.

    This class dynamically builds top-level windows with entry fields and dropdowns
    for inventory management. It validates user input, handles data cleaning, checks 
    for conflicts (e.g., occupied locations), and updates the underlying pandas DataFrame.
    Optional callbacks can be provided to synchronize changes with external logic.

    Parameters
    ----------
    parent (tk.Widget):  
        Parent Tkinter widget to attach the window to.  
    rows_df (pandas.DataFrame):  
        DataFrame containing the current inventory data.  
    columns_to_sort_by (list of str, optional): Default = []  
        Columns to sort the DataFrame by after updates.  
    for_filled (list of str, optional): Default = []  
        Columns used to determine if a location is already filled.  
    location_column (str, optional): Default = ""  
        Column containing the location identifier for inventory units.  
    int_fields (list of str, optional): Default = []  
        Columns that must contain integer values.  
    remove_fields (list of str, optional): Default = []  
        Columns used to match rows when removing items.  
    label_column (str, optional): Default = ""  
        Column containing a descriptive label for inventory units.  
    letterNums (list of str, optional): Default = []  
        Columns that must follow a letter+number format (e.g., A1).  
    date_column (str, optional): Default = ""  
        Column containing dates that should be validated in MM/DD/YYYY format.  
    grid_coords (tuple of int, optional): Default = (9, 9)  
        Grid size for rectangular picker buttons (rows, columns).  
    rectangle_picker (bool, optional): Default = True  
        Whether to use a rectangular picker (True) or circular picker (False).  

    Attributes
    ----------
    parent (tk.Widget):  
        Parent window.  
    rows_df (pandas.DataFrame):  
        Current inventory data, updated after add/remove actions.  
    entries (dict):  
        Maps column names to their Tkinter entry widgets.  
    add_callback (callable, optional): Default = None  
        Callback function invoked after adding rows; receives the updated DataFrame.  
    remove_callback (callable, optional): Default = None  
        Callback function invoked after removing rows; receives the updated DataFrame.  

    Methods
    -------
    Configure_AddRemove_Window(required_fields (list of str), top_label (str), adding (bool), unused_columns (list of str), window_name (str)):  
        Builds and configures the Tkinter data entry window with dynamic entry fields, labels, dropdowns, and buttons.  
    add_data(label (tk.Widget), add_window (tk.Widget)) -> (None):  
        Validates, cleans, and adds user input to the DataFrame, checking for conflicts and invoking add_callback.  
    remove_data(label (tk.Widget), remove_window (tk.Widget)) -> (None):  
        Validates, cleans, and removes matching rows from the DataFrame, invoking remove_callback.  
    validate_required_fields(message_label (tk.Widget)) -> (bool):  
        Ensures that all required entry fields are filled and updates message_label on error.  
    reset_fields(labels (dict)) -> (None):  
        Clears all entries and resets dropdown suggestions.  
    validate_entries(message_label (tk.Widget)) -> (dict | None):  
        Cleans and validates user input, enforcing integer, letter+number, and date formats; returns cleaned data or None if invalid.  
    """

    def __init__(self, parent, rows_df, columns_to_sort_by=[], for_filled=[], location_column="", int_fields=[], remove_fields=[], 
                 label_column="", letterNums=[], date_column="", grid_coords=(9, 9), rectangle_picker=True):
        """Initialize dataAddWindows. See class docstring for prameter details."""

        self.parent = parent
        self.rows_df = rows_df
        self.entries = {}
        self.columns_to_sort_by = columns_to_sort_by
        self.for_filled = for_filled
        self.location_column = location_column
        self.int_fields = int_fields
        self.remove_fields = remove_fields
        self.label_column = label_column
        self.letterNums = letterNums
        self.date_column = date_column
        self.grid_coords= grid_coords
        self.rectangle_picker = rectangle_picker
        
        # Callbacks that can be assigned externally to handle add/remove logic
        self.add_callback = None
        self.remove_callback = None

    def Configure_AddRemove_Window(self, required_fields=[], top_label="Enter Information Below", adding=True, unused_columns=[], window_name="Data Entry"):
        """
        Configures the tkinter window, buttons, and entry boxes for the data addition and removal windows.

        Parameters:
            required_fields (list[str]): a list of the entry fields (columns) that the user must enter in order to add/remove.
            top_label (str): The string for the label that will appear at the top of the window.
            adding (bool): True if adding to inventory and False if removing from inventory.
            unused_columns (list[str]): a list of the columns you would like not to be added as an entry field for the user.
            window_name (str): The name of the window to be made.
        """
        # Store arguments for later use
        self.required_fields = required_fields
        self.adding = adding
        self.top_label = top_label
        self.unused_columns = unused_columns
        self.window_name = window_name

        # Create a new top-level window using helper clas
        window_helper = ToplevelWindowHelper(self.parent, window_name)
        window_main_frame = window_helper.get_main_frame()
        window = window_helper.window

        # Create a frame at the top of the window for the title/label
        top_frame = tk.Frame(window_main_frame)
        top_frame.grid(row=0, column=0, columnspan=4, sticky="ew", padx=50, pady=(10, 0))

        # Initialize dropdown helper to handle entry filtering
        dropdown_helper = DropdownHelper(self.rows_df, self.entries)

        # Configure grid column weights for layout stretching
        window_main_frame.grid_columnconfigure(0, weight=0)
        window_main_frame.grid_columnconfigure(1, weight=0)
        window_main_frame.grid_columnconfigure(2, weight=1)
        window_main_frame.grid_columnconfigure(3, weight=0)

        # Add the top instructional label
        message_label = tk.Label(top_frame, text=self.top_label)
        message_label.pack(pady=(5, 0))

        # Build dictionary of all usable columns (minus excluded ones).
        labels_and_options = {col: None for col in self.rows_df.columns if col not in self.unused_columns}
        # Add the options for each field
        labels_and_options = dropdown_helper.add_dropdown_options(labels_and_options)
        
        # Create entry widgets dynamically for each column
        for i, (label_text, options) in enumerate(labels_and_options.items()):
            window_main_frame.grid_rowconfigure((i+1), weight=1)

            # The entry requirments to be added to the label
            if label_text in self.int_fields:
                requirement = " (integer)"
            elif label_text in self.letterNums:
                if label_text == self.location_column:
                    requirement = " (format: A8, B1\nComma seperated for multiple)"
                else:
                    requirement = " (format: A8, P23)"
            elif label_text == self.date_column:
                requirement = " (MM/DD/YYYY)"
            else:
                requirement = ""

            # Add label for the entry field with any requirments
            label = tk.Label(window_main_frame, text=label_text + requirement)
            label.grid(row=(i+1), column=0, sticky="w", padx=10, pady=(10, 0))

            # Add red asterisk if this field is required
            if label_text in self.required_fields:
                star_label = tk.Label(window_main_frame, text="*", fg="red", font=("Arial", 12, "bold"))
                star_label.grid(row=(i+1), column=1, sticky="e",padx=(0,2), pady=(10,0))

            # Special case: location field gets an extra "Pick" button
            if label_text == self.location_column:
                # Make the entry box an AutocompleteEntry box with dropdown menu
                entry = AutocompleteEntry(window_main_frame, options, on_select=lambda val, k=label_text: dropdown_helper.filter_dropdowns(val))
                # Re-filter dropdowns when user leaves the field
                entry.bind("<FocusOut>", lambda e, k=label_text, ent=entry: dropdown_helper.filter_dropdowns(ent.get()), add="+")
                entry.grid(row=(i+1), column=2, sticky="ew", padx=(10, 0), pady=(10, 0))
                # Add the entry as the value for the entries dictionary
                self.entries[label_text] = entry

                # Add the picker button which opens the picker popup
                pick_button = tk.Button(window_main_frame, text="Pick", command=lambda e=entry: PickerHelper(window_main_frame, self.entries, self.rows_df, message_label, 
                                                                                                      adding=adding, grid_coords=self.grid_coords, for_filled=self.for_filled, 
                                                                                                      location_column=self.location_column, label_column=self.label_column, rectangle_picker=self.rectangle_picker), takefocus=False)
                pick_button.grid(row=(i+1), column=3, padx=(5, 10), pady=(10, 0), sticky="w")
            else:
                # Make the entry box an AutocompleteEntry box with dropdown menu
                entry = AutocompleteEntry(window_main_frame, options, on_select=lambda val, k=label_text: dropdown_helper.filter_dropdowns(val))
                # Re-filter dropdowns when user leaves the field
                entry.bind("<FocusOut>", lambda e, k=label_text, ent=entry: dropdown_helper.filter_dropdowns(ent.get()), add="+")
                entry.grid(row=(i+1), column=2, sticky="ew", padx=(10, 0), pady=(10, 0))
                # Add the entry as the value for the entries dictionary
                self.entries[label_text] = entry

        # Create a frame at the bottom for the action buttons (reset, add/remove)
        button_frame = tk.Frame(window_main_frame)
        button_frame.grid(row=(len(labels_and_options) + 1), column=0, columnspan=4, pady=20, padx=10, sticky="ew")

        # Allow equal resizing of both columns in button row
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        # Create Reset button to clear all fields
        reset_button = tk.Button(button_frame, text="Reset Fields", command=lambda: self.reset_fields(labels_and_options))
        reset_button.grid(row=0, column=0, padx=5, sticky="ew")
        # If adding data create an "Add" button to add to inventory
        if adding == True:
            add_button = tk.Button(button_frame, text="Add", command=lambda: self.add_data(message_label, window))
            add_button.grid(row=0, column=1, padx=5, sticky="ew")
        # If removing data create a "Remove" button to remove from inventory
        elif adding == False:
            remove_button = tk.Button(button_frame, text="Remove", command=lambda: self.remove_data(message_label, window))
            remove_button.grid(row=0, column=1, padx=5, sticky="ew")
        else:
            print("Do not recognize field. Must do 'True' or 'False'")

    def add_data(self, label, add_window):
        """
        Takes the user entered data and adds it to the existing inventory in rows_df.

        :param label (tk.Widget): The label widget that is at the top of the add_window (data entry window) to show error messages.
        :param add_window (tk.Widget): The data entry window to allow its eventual withdrawl.
        :return (pd.DataFrame): returns the rows_df to the outside callback function.
        """

        # Ensure the user has entered the required fields.
        valid = self.validate_required_fields(label)
        if not valid:
            return
        
        # Clean and validate user entries (strip whitespace, check formats, cast types).
        self.cleaned_data = self.validate_entries(label)
        if self.cleaned_data is None:
            # If validation fails, an error message is shown within validate_entries()
            return
        
        # Confirm action with the user before committing changes.
        confirm = messagebox.askokcancel("Confirm", "Are you sure you want to add this to inventory?")
        if not confirm:
            return
        
        # Build new rows list[dict] for each location provided by the user.
        new_rows = []
        for loc in self.cleaned_data[self.location_column]:
            row = {}
            for col in self.rows_df.columns:
                col = col.strip()

                # Assign the values in the dictionary of an individual unit of inventory
                if col in self.cleaned_data:
                    val = loc if col == self.location_column else self.cleaned_data[col]

                    # Cast to int if this column is expected to be numeric
                    if col in self.int_fields:
                        try:
                            val = int(val)
                        except ValueError:
                            val = None
                    row[col] = val
                else:
                    row[col] = None
            new_rows.append(row)
        # Convert the stored new rows into a dataframe
        new_rows = pd.DataFrame(new_rows)
        
        # Create a DataFrame to track user entered locations that confilct with already filled locations
        conflict_locations = pd.DataFrame()

        # Check conflicts for each new row seperatly
        for i, row in new_rows.iterrows():
            # Start with all True mask
            mask = pd.Series(True, index=self.rows_df.index)

            # Narrow mask by checking each column in for_filled
            for col in self.for_filled:
                mask &= self.rows_df[col].astype(str).str.strip() == str(row[col]).strip()

            # Also filter using the location column
            mask &= self.rows_df[self.location_column].astype(str).str.strip() == str(row[self.location_column]).strip()

            # If there are any conflicts for this row add it to the conflict tracking DataFrame
            if mask.any():
                conflict_locations = pd.concat([conflict_locations, self.rows_df[mask]])

        # Create an error message if any conflicts are found and return without adding
        if not conflict_locations.empty:
            label.config(text=f"Error: {self.location_column}(s) {', '.join(conflict_locations[self.location_column])} already occupied.", fg="red")
            return

        # Append new rows to the DataFrame and re-sort for consistency.
        self.rows_df = pd.concat([self.rows_df, pd.DataFrame(new_rows)], ignore_index=True)
        self.rows_df.sort_values(by=self.columns_to_sort_by, inplace=True) 

        # Send the updated DataFrame back via callback (so parent GUI stays in sync).
        if self.add_callback:
            self.add_callback(self.rows_df)

        # Hide the add_window after successful submission
        add_window.withdraw()
    
    def remove_data(self, label, remove_window):
        """
        Takes the user entered data and removes it from the existing inventory in rows_df.

        :param label (tk.Widget): The label widget that is at the top of the remove_window (data entry window) to display error messages.
        :param remove_window (tk.Widget): The data removal window.
        :return (pd.DataFrame): Updated rows_df is passed back thorugh self.remove_ callback.
        """

        # Ensure the user has entered the required fields.
        valid = self.validate_required_fields(label)
        if not valid:
            return
        
        # Clean and validate user entries (strip whitespace, check formats, cast types).
        self.cleaned_data = self.validate_entries(label)
        if self.cleaned_data is None:
            # If validation fails, an error message is shown within validate_entries()
            return

        # Initialize collections for tracking:
        matched_indices = []  # Indices of rows that match and will be removed
        labels_at_locations = []  # Human-readable summaries of matched rows
        conflict_locations = []  # Locations that could not be matched

        # Iterate over each Location the user wants to remove. 
        for loc in self.cleaned_data[self.location_column]:
            # Create a copy of the users data containing only a single location at a time.
            local_data = self.cleaned_data.copy()
            local_data[self.location_column] = loc

            # Build mask that matches all fields for this specific location (all keys must match to be removed)
            mask = pd.Series(True, index=self.rows_df.index)
            for field in self.remove_fields:
                val = local_data[field]
                mask &= self.rows_df[field].astype(str).str.strip() == str(val).strip()

            # If no matching row is found, flag this location as invalid/conflict.
            match = self.rows_df[mask]
            if match.empty:
                conflict_locations.append(loc)
                continue

            # Otherwise, collect indices of matching rows for deletion.
            matched_indices.extend(match.index.tolist())

            # Build summary for confirmation dialog using the first matched row.
            row = match.iloc[0]
            parts = []
            for field in self.remove_fields + [self.label_column]:
                if field in row:
                    parts.append(f"{field}: {row.get(field, 'N/A')}")
            labels_at_locations.append("\n".join(parts))

        # If any vial locations could not be matched, show error and stop.
        if conflict_locations:
            label.config(
                text=f"Error: {self.location_column}(s) {', '.join(conflict_locations)} do not exist in the inventory.",
                fg="red"
            )
            return

        # === Ask the user to confirm removal with a popup ===
        # maximum number of entires to show is 8
        max_entries = 8
        lines = labels_at_locations[:max_entries]  # show only first 8
        # Create the summary message
        summary = "\n\n".join(lines)

        if len(labels_at_locations) > max_entries:
            # Add an additional message if entries where cut off
            summary += f"\n\n...and {len(labels_at_locations) - max_entries} more items."

        # Create the popup message
        confirm = messagebox.askokcancel(
            "Confirm Removal",
            f"Are you sure you want to remove the following from inventory?\n\n{summary}"
        )
        if not confirm:
            return

        # Drop all matched rows at once and re-sort the DataFrame.
        self.rows_df.drop(index=matched_indices, inplace=True)
        self.rows_df.sort_values(by=self.columns_to_sort_by, inplace=True)

        # Send the updated DataFrame back via callback (so parent GUI stays in sync).
        if self.remove_callback:
            self.remove_callback(self.rows_df)

        # Hide the remove_window after successful removal.
        remove_window.withdraw()

    def validate_required_fields(self, message_label):
        """
        Determines if all required fields are filled in by the user.

        :param message_label (str): The label at the top of the add_data or remove_data window to return messages.
        :return (bool): True if all required fields are filled in and False if not.
        """

        missing = [field for field in self.required_fields if not self.entries[field].get().strip()]
        if missing:
            error_text = "Please fill in: " + ", ".join(missing)
            message_label.config(text=error_text, fg="red", justify="left")
            return False
        return True

    def reset_fields(self, labels):
        """
        Clears user entries and resets all dropdown options

        :param labels: The user entry boxes
        """

        for key, entry in self.entries.items():
            entry.var.set("")  # clear the AutocompleteEntry
            entry._hide_listbox()  # Hide dropdown if it's visible
            entry.update_suggestions(labels[key])  # reset dropdown suggestions

    def validate_entries(self, message_label):
        """
        Validates and cleans user entries from the entry fields.

        - Strips whitespace
        - Type casts integers
        - Validates letter+number formats (e.g., A1)
        - Validates date format (MM/DD/YYYY)
        - Returns cleaned data if valid, otherwise shows error in the message_label.

        :param message_label (tk.Widget): The label widget at the top of the add_data or remove_data window to display messages.
        :return cleaned_data (list): A dictionary of the cleaned and verified users data. Key(the column header):Value(cleaned user entry).
            or None if validation fails.
        """
        
        cleaned_data = {}

        # Iterate over each field/entry pair
        for field, entry in self.entries.items():
            raw_value = entry.get().strip() # Remove leading/trailing spaces

            # Skip fields with no user input
            if not raw_value:
                continue

            # ---------- Case 1: Integer-only fields ----------
            if field in self.int_fields:
                # split the location field by commas, validate each one individually and add them to a list, and place the list into the cleaned_data dict.
                if field == self.location_column:
                    # Split and clean each location.
                    user_locations = [loc.strip() for loc in raw_value.split(",") if loc.strip()]
                    cleaned_locations = []
                    for loc in user_locations:
                        try:
                            cleaned_locations.append(int(loc))
                        except ValueError:
                            message_label.config(text=f"Invalid value for {field}: '{loc}' must be an integer", fg="red")
                            return None
                    cleaned_data[field] = cleaned_locations  # Store the list of ints here.
                # Single integer fields
                else:
                    try:
                        cleaned_data[field] = int(raw_value)
                    except ValueError:
                        message_label.config(text=f"Invalid value for {field}: must be an integer", fg="red")
                        return None
                    
            # ---------- Case 2: Letter+number format fields (e.g., A1, P23) ----------
            # Validate fields that should contain a single letter followed by only numbers (Ex: A1 or P23). Also Capitalizes the letter.
            elif field in self.letterNums:
                # split the location field by commas, validate each one individually and add them to a list, and place the list into the cleaned_data dict.
                if field == self.location_column:
                    # Split multiple comma-separated values (Make uppercase)
                    user_locations = [loc.strip().upper() for loc in raw_value.split(",") if loc.strip()]
                    pattern = r"^[A-Za-z]{1}\d+$"
                    # Validate each location
                    for loc in user_locations:
                        if not re.match(pattern, loc):
                            message_label.config(text=f"Invalid format for {field}: must be one letter followed by a number (e.g., A1)", fg="red")
                            return None 
                    # Convert the letter to uppercase
                    cleaned_data[field] = user_locations
                # Single entry letter number fields.
                else:
                    pattern = r"^[A-Za-z]{1}\d+$"
                    if not re.match(pattern, raw_value):
                        message_label.config(text=f"Invalid format for {field}: must be one letter followed by a number (e.g., A1)", fg="red")
                        return None
                    # Convert the letter to uppercase
                    cleaned_data[field] = raw_value.upper()

            # ---------- Case 3: Date validation (MM/DD/YYYY) ----------
            # Validate that dates follow the required format of MM/DD/YYYY.
            elif field == self.date_column:
                try:
                    # Try parsing the provided date into MM/DD/YYYY format
                    parsed_date = pd.to_datetime(str(raw_value).strip(), errors="raise", dayfirst=False)
                    # Reformat to MM/DD/YYYY if possible and save in the cleaned data
                    cleaned_data[field] = parsed_date.strftime("%m/%d/%y")
                except ValueError:
                    message_label.config(
                        text=f"Invalid format for {field}. Please use MM/DD/YYYY (e.g., 06/30/2025).",
                        fg="red"
                    )
                    return None
            
            # ---------- Case 4: Plain string fields ----------
            # Keep as string if no special category.
            else:
                cleaned_data[field] = raw_value

        # Return dictionary of validated and cleaned entries
        return cleaned_data