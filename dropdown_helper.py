import pandas as pd

class DropdownHelper:
    """
    Create and manage dropdown options for AutocompleteEntry objects.

    - Normalizes values for display consistency.  
    - Populates dropdowns with unique values from a DataFrame.  
    - Dynamically filters options across interdependent fields.  

    Parameters
    ----------
    full_df (pd.DataFrame):  
        The full DataFrame containing all possible values.  
    entries (dict):  
        A dictionary mapping entry field (column) names to their associated AutocompleteEntry object.  

    Methods
    -------------
    normalize_display(val (str)) -> str:  
        Convert a raw value into a display-friendly string.  
    add_dropdown_options(labels_dict (dict)) -> dict:  
        Populate dropdown menus with unique options from the DataFrame.  
    filter_dropdowns(selected_value (str)) -> None:  
        Filter dropdown options dynamically based on current user selections.  
    """

    def __init__(self, full_df: pd.DataFrame, entries: dict):
        """
        Initializes the DropdownHelper. See class docstring for parameter details.
        """

        self.full_df = full_df
        self.entries = entries

    def normalize_display(self, val: str):
        """
        Converts values to clean, display-friendly strings.

        Converts floats with no decimal into integers (e.g., 1.0 → "1").
        Strips leading/trailing spaces.

        :param val (str): The original value.
        :return (str): A normalized string value.
        """

        # Change float values to display integers if no decimal
        if isinstance(val, float) and val.is_integer():
            return str(int(val))
        # Return the value with no leading/trailing space
        return str(val).strip()

    def add_dropdown_options(self, labels_dict: dict):
        """
        Populates dropdown options for each entry field based on unique values from the DataFrame.

        :param labels_dict (dict): A dictionarywith entry fields as the keysand empty lists as the values.
        :return labels_dict (dict): The updated dictionary with sorted lists of dropdown options for each entry field stored as their values.
        """
        # Clean column header leading/trailing space
        self.full_df.columns = [col.strip() for col in self.full_df.columns]

        # Create a unique set of values for each column and store it in the dict value for its associated entry field.
        for label in labels_dict:
            if label in self.full_df.columns:
                # Extract non-empty unique values and normalize them
                values = sorted(
                    set([""]) | set(
                        self.normalize_display(val)
                        for val in self.full_df[label].dropna()
                        if str(val).strip() not in ("", " ")
                    ),
                    key=str.lower
                )  
                labels_dict[label] = values
            else:
                print(f"⚠️ Column '{label}' not found in DataFrame.")
                labels_dict[label] = []

        return labels_dict

    def filter_dropdowns(self, selected_value: str):
        """
        Dynamically filters all dropdown options based on the current selection in each entry.

        - Uses all non-empty entry fields to filter the DataFrame.
        - If filtering results in an empty DataFrame, resets all dropdowns to full options.
        - Otherwise, updates each dropdown to only include values from the filtered results.

        :param selected_value: The newly selected value in that field.
        """

        # Remove leading/trailing spaces and cast floats as ints
        selected_value = self.normalize_display(selected_value)

        # Create a copy of the full_df to be filtered
        filtered_df = self.full_df.copy()

        # Apply filtering based on all current non-empty entry values
        for key, entry in self.entries.items():
            if key not in self.full_df.columns:
                continue

            value = self.normalize_display(entry.get())
            # If the value exists filter the df to exclude rows that do not contain that value
            if value:
                filtered_df = filtered_df[filtered_df[key].apply(self.normalize_display) == value]

        # If no df rows match the current filters, reset all dropdowns to the full options
        if filtered_df.empty:
            for key, entry in self.entries.items():
                if key not in self.full_df.columns:
                    continue
                # Get set of all unique values
                all_values = set(
                    self.normalize_display(val)
                    for val in self.full_df[key].dropna()
                    if str(val).strip()
                )
                # Adds an empty string to the set and then calls update_suggestions() on the AutocompleteEntry object to update the dropdown
                entry.update_suggestions(sorted(set([""]) | all_values, key=str.lower))
            return  # Exit early to avoid filtering an empty DataFrame

        # Otherwise, update each dropdown based on the filtered results
        for key, entry in self.entries.items():
            if key not in self.full_df.columns:
                continue
            # Get set of unique filtered values
            filtered_values = set(
                self.normalize_display(val)
                for val in filtered_df[key].dropna()
                if str(val).strip()
            )
            # Adds an empty string to the set and then calls update_suggestions() on the AutocompleteEntry object to update the dropdown
            entry.update_suggestions(sorted(set([""]) | filtered_values, key=str.lower))