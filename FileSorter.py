import os
import shutil
import customtkinter
from tkinter import filedialog
import sys
import ctypes
import json

# --- Application Name and Configuration Path Setup ---
APP_NAME = "ProFileOrganizer"  # Used for creating the settings folder


def get_config_file_path():
    """Determines the platform-specific path for the configuration file."""
    if sys.platform == "win32":
        # Typically C:\Users\<Username>\AppData\Roaming\AppName
        base_path = os.getenv('APPDATA', os.path.expanduser('~'))
    elif sys.platform == "darwin":  # macOS
        # Typically ~/Library/Application Support/AppName
        base_path = os.path.join(os.path.expanduser(
            '~/Library'), 'Application Support')
    else:  # Linux and other Unix-like
        # Typically ~/.config/AppName or ~/.AppName
        base_path = os.getenv('XDG_CONFIG_HOME', os.path.join(
            os.path.expanduser('~'), '.config'))

    app_config_dir = os.path.join(base_path, APP_NAME)

    try:
        # Ensure the directory exists
        os.makedirs(app_config_dir, exist_ok=True)
    except OSError as e:
        # Fallback to current directory if creating app_config_dir fails (e.g. permissions)
        print(
            f"Warning: Could not create app config directory at {app_config_dir}: {e}. Using current directory for config.")
        app_config_dir = "."  # Current directory as fallback

    return os.path.join(app_config_dir, "organizer_custom_categories.json")


# Get the dynamic path for the config file
CONFIG_FILE_PATH = get_config_file_path()

# --- Custom Category Mappings ---
custom_category_mappings = {}

# --- Load and Save Custom Category Settings ---


def load_custom_category_settings():
    global custom_category_mappings
    print(f"Attempting to load custom categories from: {CONFIG_FILE_PATH}")
    try:
        if os.path.exists(CONFIG_FILE_PATH):
            with open(CONFIG_FILE_PATH, 'r') as f:
                loaded_mappings_from_json = json.load(f)
                custom_category_mappings = {
                    category: set(extensions)
                    for category, extensions in loaded_mappings_from_json.items()
                }
                print(
                    f"Custom categories loaded successfully from {CONFIG_FILE_PATH}.")
        else:
            custom_category_mappings = {  # Default categories if no config file
                "IMAGES": {"jpg", "jpeg", "png", "gif", "bmp", "webp"},
                "DOCUMENTS": {"pdf", "doc", "docx", "txt", "rtf", "odt"},
                "AUDIO": {"mp3", "wav", "aac"}, "VIDEO": {"mp4", "mov", "avi", "mkv"},
                "ARCHIVES": {"zip", "rar", "7z"}, "NO_EXTENSION_FILES": {"no_extension"}
            }
            print(
                f"Config file '{CONFIG_FILE_PATH}' not found. Using default categories. It will be created when settings are saved.")
    except (json.JSONDecodeError, IOError, TypeError, ValueError) as e:
        print(
            f"Error loading/parsing {CONFIG_FILE_PATH}: {e}. Using/Reverting to default categories.")
        custom_category_mappings = {
            "IMAGES": {"jpg", "jpeg", "png"}, "DOCUMENTS": {"pdf", "txt"},
            "NO_EXTENSION_FILES": {"no_extension"}
        }


def save_custom_category_settings():
    global custom_category_mappings
    try:
        mappings_to_save_to_json = {
            category: sorted(list(extensions))
            for category, extensions in custom_category_mappings.items()
        }
        with open(CONFIG_FILE_PATH, 'w') as f:  # Use the full path
            json.dump(mappings_to_save_to_json, f, indent=4)

        feedback_message = f"INFO: Custom category settings saved to {CONFIG_FILE_PATH}.\n"
        if 'status_textbox' in globals() and status_textbox and status_textbox.winfo_exists():
            status_textbox.insert("end", feedback_message)
            status_textbox.see("end")
        else:
            print(feedback_message.strip())
    except IOError as e:
        error_message = f"ERROR: Could not save custom category settings to {CONFIG_FILE_PATH}: {e}\n"
        if 'status_textbox' in globals() and status_textbox and status_textbox.winfo_exists():
            status_textbox.insert("end", error_message)
            status_textbox.see("end")
        else:
            print(error_message.strip())


load_custom_category_settings()  # Load settings at startup

# --- Constants ---
NO_EXTENSION_KEYWORD = "no_extension"
# NO_EXTENSION_CATEGORY_NAME is now just a key in custom_category_mappings, e.g., "NO_EXTENSION_FILES"

# --- Basic CustomTkinter Set Up ---
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")
app = customtkinter.CTk()
app.title(f"{APP_NAME} (Settings in AppData)")
app.geometry("800x750")

selected_folder_path_var = customtkinter.StringVar(value="No Folder Selected")

# --- Helper Functions (get_forbidden_paths, is_system_or_hidden) ---


def get_forbidden_paths():
    paths = []
    system_drive_guess = os.path.splitdrive(
        sys.executable)[0] if sys.executable else "C:"
    system_drive = os.path.normpath(system_drive_guess + os.sep)
    if sys.platform == "win32":
        paths.extend([
            os.path.join(system_drive, "Windows"), os.path.join(
                system_drive, "Program Files"),
            os.path.join(system_drive, "Program Files (x86)"),
            os.getenv("SystemRoot", ""), os.getenv(
                "ProgramFiles", ""), os.getenv("ProgramFiles(x86)", ""),
        ])
        user_profile = os.getenv("UserProfile")
        if user_profile:
            paths.append(os.path.join(user_profile, "AppData"))
    elif sys.platform == "darwin":
        paths.extend(["/System", "/Library", "/usr", "/private", os.path.expanduser(
            "~/Library/Application Support")])  # Note: App Support for consistency
    elif sys.platform.startswith("linux"):
        paths.extend(["/boot", "/etc", "/lib", "/proc", "/sys", "/usr", "/var",
                     # Note: .config for consistency
                      os.path.expanduser("~/.config")])
    normalized_paths = [os.path.normpath(p) for p in paths if p]
    return [p for p in normalized_paths if os.path.exists(p)]


def is_system_or_hidden(filepath):
    if sys.platform == "win32":
        try:
            FILE_ATTRIBUTE_HIDDEN = 0x02
            FILE_ATTRIBUTE_SYSTEM = 0x04
            attrs = ctypes.windll.kernel32.GetFileAttributesW(str(filepath))
            return attrs != -1 and (bool(attrs & FILE_ATTRIBUTE_HIDDEN) or bool(attrs & FILE_ATTRIBUTE_SYSTEM))
        except Exception:
            return False
    elif sys.platform in ["darwin", "linux", "linux2"]:
        return os.path.basename(filepath).startswith('.')
    return False

# --- GUI-Adapted Backend Functions (get_files_in_folder, get_file_extension, create_folder_if_not_exists_gui, move_file_to_folder_gui, delete_empty_folders_recursively) ---
# These functions remain largely the same. For brevity, I'm assuming they are correctly defined as in previous versions.
# Ensure they are present in your actual combined script.


def get_files_in_folder(folder_path, textbox_widget):
    all_file_paths = []
    skipped_for_safety_count = 0
    textbox_widget.insert(
        "end", f"Phase 1: Scanning folder and subfolders: {folder_path}...\n")
    app.update_idletasks()
    try:
        for dirpath, dirnames, filenames in os.walk(folder_path):
            original_dirnames = list(dirnames)
            dirnames[:] = []
            for d_name in original_dirnames:
                d_full_path = os.path.join(dirpath, d_name)
                if is_system_or_hidden(d_full_path):
                    skipped_for_safety_count += 1
                    continue
                dirnames.append(d_name)
            for filename in filenames:
                full_path = os.path.join(dirpath, filename)
                if is_system_or_hidden(full_path):
                    skipped_for_safety_count += 1
                    continue
                all_file_paths.append(full_path)
        if skipped_for_safety_count > 0:
            textbox_widget.insert(
                "end", f"Scan note: Skipped {skipped_for_safety_count} hidden/system files/folders.\n")
        if not all_file_paths:
            textbox_widget.insert(
                "end", "Scan complete. No processable files found.\n")
        else:
            textbox_widget.insert(
                "end", f"Scan complete. Found {len(all_file_paths)} potentially processable files.\n")
    except Exception as e:
        textbox_widget.insert(
            "end", f"ERROR during file scan in {folder_path}: {e}\n")
    finally:
        textbox_widget.see("end")
        return all_file_paths


def get_file_extension(file_name):
    name_part, extension_part = os.path.splitext(file_name)
    return extension_part.lower()


def create_folder_if_not_exists_gui(folder_path, sub_folder_name, textbox_widget):
    new_folder_path = os.path.join(folder_path, sub_folder_name)
    try:
        os.makedirs(new_folder_path, exist_ok=True)
        return new_folder_path
    except OSError as e:
        textbox_widget.insert(
            "end", f"  ERROR Creating Directory {new_folder_path}: {e}\n")
        textbox_widget.see("end")
        return None


def move_file_to_folder_gui(full_source_file_path, original_target_dir, destination_category_name, textbox_widget):
    file_name_only = os.path.basename(full_source_file_path)
    final_destination_folder = os.path.join(
        original_target_dir, destination_category_name)
    final_destination_path = os.path.join(
        final_destination_folder, file_name_only)
    moved_successfully = False
    try:
        if os.path.normpath(os.path.dirname(full_source_file_path)) == os.path.normpath(final_destination_folder) and \
           os.path.normpath(full_source_file_path) == os.path.normpath(final_destination_path):
            return True
        if os.path.exists(final_destination_path):  # Collision handling
            name_part, ext_part = os.path.splitext(file_name_only)
            count = 1
            new_file_name_base = f"{name_part}_{count}"
            new_final_destination_path = os.path.join(
                final_destination_folder, f"{new_file_name_base}{ext_part}")
            while os.path.exists(new_final_destination_path):
                count += 1
                new_file_name_base = f"{name_part}_{count}"
                new_final_destination_path = os.path.join(
                    final_destination_folder, f"{new_file_name_base}{ext_part}")
            final_destination_path = new_final_destination_path
            textbox_widget.insert(
                "end", f"  INFO: File '{file_name_only}' exists. Renaming to '{os.path.basename(final_destination_path)}'.\n")
        textbox_widget.insert(
            "end", f"  Moving '{full_source_file_path}'\n    to '{final_destination_path}'...\n")
        textbox_widget.see("end")
        app.update_idletasks()
        shutil.move(full_source_file_path, final_destination_path)
        textbox_widget.insert("end", f"  Successfully Moved.\n")
        moved_successfully = True
    except Exception as e:
        textbox_widget.insert(
            "end", f"  ERROR Moving File '{file_name_only}': {e}\n")
    finally:
        textbox_widget.see("end")
        return moved_successfully


def delete_empty_folders_recursively(path_to_check, main_organize_root, actual_created_category_folder_names, textbox_widget):
    textbox_widget.insert(
        "end", f"\nPhase 3: Attempting to delete empty subfolders in {path_to_check}...\n")
    app.update_idletasks()
    deleted_count = 0
    protected_category_paths = {os.path.normpath(os.path.join(
        main_organize_root, name)) for name in actual_created_category_folder_names}
    normalized_main_root = os.path.normpath(main_organize_root)
    for dirpath, dirnames, filenames in os.walk(path_to_check, topdown=False):
        norm_dirpath = os.path.normpath(dirpath)
        if norm_dirpath == normalized_main_root or norm_dirpath in protected_category_paths:
            continue
        try:
            if not os.listdir(dirpath):
                try:
                    os.rmdir(dirpath)
                    textbox_widget.insert(
                        "end", f"  Deleted empty folder: {dirpath}\n")
                    deleted_count += 1
                except OSError as e:
                    textbox_widget.insert(
                        "end", f"  ERROR deleting folder {dirpath}: {e}\n")
        except OSError as e:
            textbox_widget.insert(
                "end", f"  ERROR accessing folder {dirpath} for emptiness check: {e}\n")
        finally:
            textbox_widget.see("end")
            app.update_idletasks()
    if deleted_count > 0:
        textbox_widget.insert(
            "end", f"Deleted {deleted_count} empty subfolder(s).\n")
    else:
        textbox_widget.insert(
            "end", "No empty subfolders were found or deleted.\n")
    textbox_widget.see("end")


# --- Settings Window Functionality (references global custom_category_mappings) ---
settings_window_instance = None
_categories_scrollable_frame_settings = None
_cat_name_entry_settings = None
_cat_exts_entry_settings = None
_add_update_button_settings = None
_feedback_label_settings = None


def open_settings_window():
    global settings_window_instance, custom_category_mappings, status_textbox
    global _categories_scrollable_frame_settings, _cat_name_entry_settings, _cat_exts_entry_settings
    global _add_update_button_settings, _feedback_label_settings

    if settings_window_instance is not None and settings_window_instance.winfo_exists():
        settings_window_instance.focus()
        return

    settings_win = customtkinter.CTkToplevel(app)
    settings_win.title("Custom Category Settings")
    settings_win.geometry("650x700")
    settings_win.transient(app)
    settings_win.grab_set()
    settings_window_instance = settings_win

    manage_frame = customtkinter.CTkFrame(settings_win)
    manage_frame.pack(pady=10, padx=10, fill="x")
    customtkinter.CTkLabel(manage_frame, text="Category Name:").grid(
        row=0, column=0, padx=5, pady=5, sticky="w")
    _cat_name_entry_settings = customtkinter.CTkEntry(
        manage_frame, placeholder_text="e.g., My Pictures, Work Reports")
    _cat_name_entry_settings.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    customtkinter.CTkLabel(manage_frame, text="Extensions (comma-separated):").grid(
        row=1, column=0, padx=5, pady=5, sticky="w")
    _cat_exts_entry_settings = customtkinter.CTkEntry(
        manage_frame, placeholder_text=f"e.g., jpg, png (or '{NO_EXTENSION_KEYWORD}')")
    _cat_exts_entry_settings.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
    manage_frame.columnconfigure(1, weight=1)
    _feedback_label_settings = customtkinter.CTkLabel(
        manage_frame, text="", text_color=("red", "pink"))
    _feedback_label_settings.grid(
        row=2, column=0, columnspan=2, pady=2, sticky="ew")
    _add_update_button_settings = customtkinter.CTkButton(
        manage_frame, text="Add New Category", command=lambda: add_or_update_category_command_settings())
    _add_update_button_settings.grid(row=3, column=0, columnspan=2, pady=10)

    list_label = customtkinter.CTkLabel(
        settings_win, text="Current Custom Categories & Assigned Extensions:")
    list_label.pack(pady=(10, 0))
    _categories_scrollable_frame_settings = customtkinter.CTkScrollableFrame(
        settings_win, height=300)
    _categories_scrollable_frame_settings.pack(
        pady=5, padx=10, fill="both", expand=True)

    refresh_category_list_display_settings()

    def on_settings_close():
        global settings_window_instance
        save_custom_category_settings()
        settings_window_instance = None
        settings_win.destroy()

    close_button = customtkinter.CTkButton(
        settings_win, text="Save and Close Settings", command=on_settings_close)
    close_button.pack(pady=20)
    settings_win.protocol("WM_DELETE_WINDOW", on_settings_close)


# For settings UI validation
def get_all_assigned_extensions_settings(exclude_category=None):
    global custom_category_mappings
    assigned = set()
    for cat, exts_set in custom_category_mappings.items():
        if exclude_category and cat == exclude_category:
            continue
        assigned.update(exts_set)
    return assigned


# Renamed for clarity
def add_or_update_category_command_settings(existing_category_name_to_update=None):
    global custom_category_mappings, status_textbox, _feedback_label_settings, _cat_name_entry_settings, _cat_exts_entry_settings, _add_update_button_settings

    _feedback_label_settings.configure(text="")
    new_cat_name = _cat_name_entry_settings.get().strip()
    exts_str = _cat_exts_entry_settings.get().strip().lower()

    if not new_cat_name:
        _feedback_label_settings.configure(
            text="Category name cannot be empty.")
        return
    if not exts_str:
        _feedback_label_settings.configure(text="Extensions cannot be empty.")
        return

    is_no_ext_cat_name = any(NO_EXTENSION_KEYWORD in s for s in custom_category_mappings.get(
        # Heuristic
        new_cat_name, {})) or new_cat_name.upper() == "NO_EXTENSION_FILES"

    if new_cat_name.upper() == "NO_EXTENSION_FILES" and exts_str != NO_EXTENSION_KEYWORD:  # Enforce rule for special category
        _feedback_label_settings.configure(
            text=f"'{new_cat_name}' category must ONLY contain the special keyword '{NO_EXTENSION_KEYWORD}'.")
        return
    if new_cat_name.upper() != "NO_EXTENSION_FILES" and NO_EXTENSION_KEYWORD in {e.strip() for e in exts_str.split(',')}:
        _feedback_label_settings.configure(
            text=f"The special keyword '{NO_EXTENSION_KEYWORD}' can only be used with a category named like 'NO_EXTENSION_FILES'.")
        return

    current_exts_set = {e.strip() for e in exts_str.split(
        ',') if e.strip() and e.strip() != "."}
    if not current_exts_set:
        _feedback_label_settings.configure(
            text="No valid extensions provided after parsing.")
        return

    # Name conflict check
    if not existing_category_name_to_update and new_cat_name in custom_category_mappings:  # Adding new, name exists
        _feedback_label_settings.configure(
            text=f"Category '{new_cat_name}' already exists. Choose a different name or edit existing.")
        return
    if existing_category_name_to_update and new_cat_name != existing_category_name_to_update and new_cat_name in custom_category_mappings:  # Renaming, new name conflicts
        _feedback_label_settings.configure(
            text=f"New category name '{new_cat_name}' already exists. Choose another.")
        return

    all_other_assigned_exts = get_all_assigned_extensions_settings(
        exclude_category=existing_category_name_to_update or new_cat_name)
    conflicting_exts = current_exts_set.intersection(all_other_assigned_exts)
    if conflicting_exts:
        _feedback_label_settings.configure(
            text=f"Extensions {conflicting_exts} are in other categories. Unique assignment enforced.")
        return

    # If renaming, delete the old entry first
    if existing_category_name_to_update and existing_category_name_to_update != new_cat_name and existing_category_name_to_update in custom_category_mappings:
        del custom_category_mappings[existing_category_name_to_update]

    custom_category_mappings[new_cat_name] = current_exts_set  # Add or update
    status_textbox.insert(
        "end", f"INFO: Category '{new_cat_name}' {'updated' if existing_category_name_to_update else 'added'}.\n")
    status_textbox.see("end")

    _cat_name_entry_settings.delete(0, "end")
    _cat_exts_entry_settings.delete(0, "end")
    _add_update_button_settings.configure(
        text="Add New Category", command=lambda: add_or_update_category_command_settings())  # Reset button
    refresh_category_list_display_settings()


def populate_edit_fields_settings(category_name_to_edit):
    global custom_category_mappings, _cat_name_entry_settings, _cat_exts_entry_settings, _add_update_button_settings, _feedback_label_settings
    if category_name_to_edit in custom_category_mappings:
        _cat_name_entry_settings.delete(0, "end")
        _cat_name_entry_settings.insert(0, category_name_to_edit)
        exts_to_edit = ", ".join(
            sorted(list(custom_category_mappings[category_name_to_edit])))
        _cat_exts_entry_settings.delete(0, "end")
        _cat_exts_entry_settings.insert(0, exts_to_edit)
        _add_update_button_settings.configure(
            text=f"Update '{category_name_to_edit}'", command=lambda cn=category_name_to_edit: add_or_update_category_command_settings(existing_category_name_to_update=cn))
        _feedback_label_settings.configure(text="")


def refresh_category_list_display_settings():
    global custom_category_mappings, _categories_scrollable_frame_settings
    for widget in _categories_scrollable_frame_settings.winfo_children():
        widget.destroy()
    if not custom_category_mappings:
        customtkinter.CTkLabel(_categories_scrollable_frame_settings,
                               text="No custom categories defined.").pack(pady=10)
        return
    sorted_categories = sorted(custom_category_mappings.keys())
    for cat_name in sorted_categories:
        exts_set = custom_category_mappings[cat_name]
        display_exts_list = []
        has_no_ext_keyword = False
        for ext in sorted(list(exts_set)):
            if ext == NO_EXTENSION_KEYWORD:
                has_no_ext_keyword = True
            else:
                display_exts_list.append(f".{ext}")
        display_exts_str = ", ".join(display_exts_list)
        if has_no_ext_keyword:
            display_exts_str = f"(Files with No Ext){', ' if display_exts_list else ''}{display_exts_str}"

        item_frame = customtkinter.CTkFrame(
            _categories_scrollable_frame_settings)
        item_frame.pack(fill="x", pady=2, padx=2)
        info_label = customtkinter.CTkLabel(
            item_frame, text=f"{cat_name}:  {display_exts_str}", wraplength=350, justify="left")
        info_label.pack(side="left", padx=5, pady=2, expand=True, fill="x")
        edit_btn = customtkinter.CTkButton(
            item_frame, text="Edit", width=50, height=24, command=lambda cn=cat_name: populate_edit_fields_settings(cn))
        edit_btn.pack(side="left", padx=(0, 5), pady=2)
        remove_btn = customtkinter.CTkButton(item_frame, text="Remove", width=60, height=24, command=lambda cn=cat_name: remove_category_command_settings(
            cn), fg_color=("gray70", "gray25"), hover_color=("gray60", "gray35"))
        remove_btn.pack(side="left", padx=(0, 5), pady=2)
    if hasattr(_categories_scrollable_frame_settings, '_parent_canvas'):
        _categories_scrollable_frame_settings._parent_canvas.update_idletasks()


def remove_category_command_settings(cat_to_remove):
    global custom_category_mappings, status_textbox, _cat_name_entry_settings, _cat_exts_entry_settings, _add_update_button_settings, _feedback_label_settings
    if cat_to_remove in custom_category_mappings:
        del custom_category_mappings[cat_to_remove]
        refresh_category_list_display_settings()
        if _cat_name_entry_settings.get() == cat_to_remove:
            _cat_name_entry_settings.delete(0, "end")
            _cat_exts_entry_settings.delete(0, "end")
            _add_update_button_settings.configure(
                text="Add New Category", command=lambda: add_or_update_category_command_settings())
            _feedback_label_settings.configure(text="")
        status_textbox.insert(
            "end", f"INFO: Custom category '{cat_to_remove}' removed.\n")
        status_textbox.see("end")

# --- Main GUI Command Functions ---


def browse_folder_command():
    progress_bar.set(0)
    folder = filedialog.askdirectory()
    if folder:
        selected_folder_path_var.set(folder)
        if status_textbox:
            status_textbox.delete("1.0", "end")
            status_textbox.insert("end", f"Selected folder: {folder}\n")
            status_textbox.see("end")
    elif status_textbox:
        status_textbox.insert("end", "Folder selection cancelled.\n")
        status_textbox.see("end")


def start_organization_command():
    global custom_category_mappings
    target_dir = selected_folder_path_var.get()
    progress_bar.set(0)
    if not target_dir or target_dir == "No Folder Selected" or not os.path.isdir(target_dir):
        status_textbox.insert("end", "ERROR: Select a valid target folder.\n")
        status_textbox.see("end")
        return
    normalized_target_dir = os.path.normpath(target_dir)
    for forbidden_path in get_forbidden_paths():
        if normalized_target_dir == forbidden_path or (normalized_target_dir + os.sep).startswith(forbidden_path + os.sep):
            status_textbox.insert(
                "end", f"CRITICAL ERROR: Folder '{target_dir}' is in a protected system directory ('{forbidden_path}').\nOrganization ABORTED.\n")
            status_textbox.see("end")
            return

    status_textbox.insert(
        "end", f"\n--- Starting Full Organization for: {target_dir} ---\n")
    app.update_idletasks()
    files_to_process = []
    total_files = 0
    processed_files_count = 0
    actual_created_category_folder_names = set()
    try:
        files_to_process = get_files_in_folder(target_dir, status_textbox)
        total_files = len(files_to_process)
        if not files_to_process:
            status_textbox.insert(
                "end", "Organization finished: No files to process after scan.\n")
            progress_bar.set(1.0)
        else:
            status_textbox.insert(
                "end", f"\nPhase 2: Processing {total_files} files based on custom categories...\n")
            app.update_idletasks()
            for full_file_path in files_to_process:
                current_filename_only = os.path.basename(full_file_path)
                current_file_extension_with_dot = get_file_extension(
                    current_filename_only)
                target_custom_category_name = None

                processed_ext_key = ""
                if not current_file_extension_with_dot or (current_file_extension_with_dot == "." and not current_filename_only.endswith("..")):
                    processed_ext_key = NO_EXTENSION_KEYWORD
                else:
                    processed_ext_key = current_file_extension_with_dot[1:].lower(
                    )

                # truly empty extension from a filename like "myfile." when NO_EXTENSION_KEYWORD isn't ""
                if not processed_ext_key and processed_ext_key != NO_EXTENSION_KEYWORD:
                    processed_ext_key = NO_EXTENSION_KEYWORD  # map it to the keyword for lookup

                for cat_name, exts_set in custom_category_mappings.items():
                    if processed_ext_key in exts_set:
                        target_custom_category_name = cat_name
                        break

                if not target_custom_category_name:
                    display_key = processed_ext_key if processed_ext_key != NO_EXTENSION_KEYWORD else "No Extension"
                    status_textbox.insert(
                        "end", f"  -Skipping '{current_filename_only}' (type '{display_key}' not in any custom category).\n")
                    processed_files_count += 1
                    if total_files > 0:
                        progress_bar.set(processed_files_count / total_files)
                    app.update_idletasks()
                    continue

                if os.path.normpath(os.path.dirname(full_file_path)) == os.path.normpath(os.path.join(target_dir, target_custom_category_name)):
                    processed_files_count += 1
                    if total_files > 0:
                        progress_bar.set(processed_files_count / total_files)
                    app.update_idletasks()
                    continue

                status_textbox.insert(
                    "end", f"\nProcessing: {full_file_path}\n")
                status_textbox.insert(
                    "end", f"  -Target Custom Category: {target_custom_category_name}\n")
                app.update_idletasks()
                actual_created_category_folder_names.add(
                    target_custom_category_name)
                destination_category_root_path = create_folder_if_not_exists_gui(
                    target_dir, target_custom_category_name, status_textbox)
                if destination_category_root_path:
                    move_file_to_folder_gui(
                        full_file_path, target_dir, target_custom_category_name, status_textbox)
                processed_files_count += 1
                if total_files > 0:
                    progress_bar.set(processed_files_count / total_files)
                app.update_idletasks()

            status_textbox.insert(
                "end", "\n--- File processing complete! ---\n")
            if total_files > 0:
                progress_bar.set(1.0)
                app.update_idletasks()
            delete_empty_folders_recursively(
                target_dir, target_dir, actual_created_category_folder_names, status_textbox)
    except Exception as e:
        status_textbox.insert(
            "end", f"\n--- AN UNEXPECTED ERROR OCCURRED: {e} ---\n")
        if total_files > 0:
            progress_bar.set(processed_files_count / total_files)
        else:
            progress_bar.set(0)
    finally:
        if total_files == processed_files_count and total_files >= 0:
            progress_bar.set(1.0)
        status_textbox.insert(
            "end", "\n--- Organization attempt finished. ---\n")
        status_textbox.see("end")


# --- GUI Layout ---
top_frame = customtkinter.CTkFrame(master=app)
top_frame.pack(pady=(10, 5), padx=20, fill="x")
browse_button = customtkinter.CTkButton(
    master=top_frame, text="Select Main Folder", command=browse_folder_command)
browse_button.pack(side="left", padx=(10, 10), pady=10)
folder_display_label = customtkinter.CTkLabel(
    master=top_frame, textvariable=selected_folder_path_var, wraplength=350)
folder_display_label.pack(side="left", padx=0, pady=10, expand=True, fill="x")
settings_button = customtkinter.CTkButton(
    master=top_frame, text="Category Settings", command=open_settings_window, width=150)
settings_button.pack(side="left", padx=(10, 10), pady=10)

organize_button = customtkinter.CTkButton(
    master=app, text="Start Organizing Files", command=start_organization_command, height=40)
organize_button.pack(pady=5, padx=20)
progress_bar = customtkinter.CTkProgressBar(master=app, height=20)
progress_bar.pack(pady=(5, 10), padx=20, fill="x")
progress_bar.set(0)
status_textbox = customtkinter.CTkTextbox(master=app, height=350, wrap="word")
status_textbox.pack(pady=(0, 10), padx=20, fill="both", expand=True)
status_textbox.insert("0.0", f"Welcome To {APP_NAME}!\n\n- Select main folder.\n- Configure custom categories and their extensions in Settings.\n  (Settings are saved to {CONFIG_FILE_PATH} when you close settings window).\n- Only files matching an extension in a custom category will be moved.\n- Hidden/System files & critical paths are protected.\n- Empty subfolders will be deleted.\n\n")

app.mainloop()
