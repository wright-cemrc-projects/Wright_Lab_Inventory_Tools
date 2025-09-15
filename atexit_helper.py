import os
import tkinter as tk
from tkinter import messagebox

class TempFileManager:
    """
    Manage deletion of temporary files.

    - Ensures that files are closed before continuing.  
    - Deletes any marked files when the program closes.  

    Attributes
    ----------
    files_to_delete (list):  
        Stores file paths marked as temporary.  

    Methods
    -------------
    mark_for_deletion(path (str)) -> None:  
        Mark a file path to be deleted when the script exits.  
    is_file_in_use(path (str)) -> bool:  
        Check if a file is currently open/in use.  
    get_open_files() -> list:  
        Return a list of files marked for deletion that are still open.  
    notify_if_open_files() -> None:  
        Show a popup warning the user of open files that must be closed.  
    cleanup_temp_files() -> None:  
        Attempt to delete all marked files, retrying until success or user cancels.  
    """
    
    def __init__(self):
        """Initializes TempFileManager. See class docstring for attribute details."""

        # Keeps track of file paths that should be deleted when the program exits
        self.files_to_delete = []

    def mark_for_deletion(self, path):
        """
        Marks a file path to be deleted when the script exits.

        :param path: Path to the file to be deleted.
        """

        # Add the file path to be deleted if it is not already added
        if path not in self.files_to_delete:
            self.files_to_delete.append(path)

    def is_file_in_use(self, path):
        """
        Checks if a file is currently in use (open) by attempting to rename the file to istself.

        - Assumes that if it can not rename the file it is open.
        - Raises an exception if file is open.

        :param path (str): Path to the file.
        :return (bool): True if the file is open, False otherwise.
        """

        # Return False if path does not exist
        if not os.path.exists(path):
            return False
        # Return False if path is not open
        try:
            os.rename(path, path)
            return False
        # Return True if file is open
        except Exception as e:
            return True

    def get_open_files(self):
        """Returns a list of files marked for deletion that are still open/in use."""

        open_files = []
        for path in self.files_to_delete:
            if self.is_file_in_use(path):
                open_files.append(path)
        return open_files

    def notify_if_open_files(self):
        """
        Check for open files and notify the user via a popup of which files to close before continueing.
        
        Prevents the user form continueing until these files are closed.
        """

        open_files = self.get_open_files()

        # Create a window for the popup message
        root = tk.Tk()
        root.withdraw()  # Hide root window

        # Show a popup of the files that are currently open and need to be closed
        while open_files:
            retry = messagebox.askretrycancel(
                "Files Open",
                "Some files are currently open or locked:\n\n" +
                "\n".join(f"• {file}" for file in open_files) +
                "\n\nPlease close these files to avoid issues.\n\n"
                "Click Retry after closing them, or Cancel to stop."
            )

            if not retry:
                # Exit the check and stay at the present window if user chose cancel
                break

            # Refresh the list to see if all the files are closed
            open_files = self.get_open_files()
            
        root.destroy()

    def cleanup_temp_files(self):
        """
        Delete all files marked for deletion at program exit.

        - Skips files that are currently open/in use.
        - If deletion fails, prtomps the user with a retry/cancel popup warning the user to close open files.
        - Keps retrying until files are deleted or user clicks cancel.

        On success: files are removed and list is cleared.
        On failure: remaining undeleted files are reported to the user.
        """

        # Get a list of open files
        open_files = self.get_open_files()

        # Try deleting all files not currently open
        for path in self.files_to_delete:
            if path not in open_files:
                try:
                    os.remove(path)
                    print(f"Deleted: {path}")
                except Exception as e:
                    print(f"Failed to delete {path}: {e}")
                    open_files.append(path)

        # If all files were deleted, return silently
        if not open_files:
            return

        # Retry loop for undeleted files
        try:
            root = tk.Tk()
            root.withdraw()  # Hide the main Tkinter window

            while open_files:
                # Ask user to retry or cancel
                retry = messagebox.askretrycancel(
                    "Cleanup Failed",
                    "Some temporary files could not be deleted (maybe still open):\n\n" +
                    "\n".join(open_files) + "\n\nTo close program or continue, close any open files and click Retry."
                )

                # If user choses cancel, close without deleting remaining files
                if not retry:
                    break

                # Attempt to delete the remaining files again
                still_undeleted = []
                for path in open_files:
                    try:
                        if os.path.exists(path):
                            os.remove(path)
                            print(f"Deleted on retry: {path}")
                    except Exception:
                        still_undeleted.append(path)

                # Update the list for the next retry attempt
                open_files = still_undeleted

            # Close the hidden Tkinter root window
            root.destroy()
            # Clear the stored file paths
            self.files_to_delete = []

        except Exception as e:
            print("Could not show retry popup:", e) 