"""commands/smart_post_dialog/entry.py"""

# pylint: disable=too-many-lines,too-many-locals,too-many-statements,too-many-branches,too-many-arguments,too-many-positional-arguments,global-statement,broad-exception-caught

import json
import logging
import os
import shutil
import subprocess
import tempfile
import time

import adsk.cam
import adsk.core

try:
    from ... import config
    from ...lib import fusionAddInUtils as futil
except ImportError:  # Fallback for Fusion's varying sys.path/package loading.
    import config  # type: ignore
    from lib import fusionAddInUtils as futil  # type: ignore

try:
    from . import fusion_helpers, personal_pipeline
except ImportError:  # Fallback when Fusion loads the command as a top-level module.
    import fusion_helpers  # type: ignore
    import personal_pipeline  # type: ignore


# =============================================================================
# GLOBAL VARIABLES
# =============================================================================
# region

# Get the active document
app = adsk.core.Application.get()
ui = app.userInterface

# List to store local event handlers
local_handlers = []

# Command configuration constants
CMD_ID = f"{config.COMPANY_NAME}_{config.ADDIN_NAME}_cmdDialog"
CMD_NAME = "Smart Post"
CMD_DESCRIPTION = "Fast G-code and tool change unlock"
IS_PROMOTED = False

# Workspace and panel configuration
WORKSPACE_ID = "CAMEnvironment"
PANEL_ID = "SmartPostPanel"
PANEL_NAME = "SmartPost"
POSITION_ID = "CAMManagePanel"
COMMAND_BESIDE_ID = ""

# Path to the folder containing command icons
ICON_FOLDER = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "./resources/icon", ""
)
BUTTON_ICON = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "./resources/open_button", ""
)

# Path to the postprocessor
POST_PATH = ""

# Personal-mode intermediate postprocessors.  Keep the former SmartPost post as
# an explicit fallback while using the current Autodesk-derived post by default.
XML_POST_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "xml_last.cps")
OLD_XML_POST_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "xml.cps")

# List of available options for high feedrate mapping behavior
HIGH_FEED_MAPPING_ITEMS = [
    "Preserve rapid movement",
    "Preserve single axis rapid movement",
    "Preserve axial and radial rapid movement",
    "Always use high feed",
]

# List of available units for the postprocessor
UNIT_ITEMS = ["Inches", "Millimeters", "Document Unit"]


MESSAGE_BOX_TITLE = CMD_NAME
MESSAGE_BOX_BUTTONS = adsk.core.MessageBoxButtonTypes.OKButtonType
MESSAGE_BOX_ICON = adsk.core.MessageBoxIconTypes.NoIconIconType


def show_message(
    text,
    title=MESSAGE_BOX_TITLE,
    buttons=MESSAGE_BOX_BUTTONS,
    icon=MESSAGE_BOX_ICON,
):
    """Wrapper around Fusion `ui.messageBox` for pylint compatibility."""
    ui.messageBox(str(text), title, buttons, icon)


def do_events():
    """Call Fusion `adsk.doEvents` across API versions."""
    do_events_fn = getattr(adsk, "doEvents", None)
    if do_events_fn is None:
        return

    try:
        do_events_fn()
    except TypeError:
        args = []
        args.append(0)
        do_events_fn(*args)


def as_bool(value):
    """Convert persisted JSON/default values without treating "false" as true."""
    if isinstance(value, str):
        return value.strip().lower() in ("1", "true", "yes", "on")
    return bool(value)


def get_personal_xml_post(use_old_xml):
    """Select the intermediate CPS used only by the Personal workflow."""
    return OLD_XML_POST_FILE if as_bool(use_old_xml) else XML_POST_FILE



def progress_show(dialog, title, message, minimum, maximum):
    """Call Fusion progress dialog `show` across API versions."""
    show_fn = getattr(dialog, "show", None)
    if show_fn is None:
        return

    args = (title, message, minimum, maximum)
    try:
        show_fn(*args)
    except TypeError:
        args = (title, message, minimum, maximum, 0)
        show_fn(*args)


# Error codes for postprocessing (post.exe)
ERROR_CODES = {
    0: "Successful processing.",
    1: "Unspecified failure.",
    100: "Failed to load post configuration.",
    101: "Empty configuration.",
    102: "Failed to initialize.",
    103: "Failed to evaluate configuration.",
    104: "Invalid machine configuration.",
    200: "Failed to load intermediate NC data.",
    201: "Unknown format.",
    300: "Failed to open output file.",
    400: "Failed to open log file.",
    500: "Post processing failed.",
    501: "Post processing was aborted.",
    502: "Post processing timed out.",
}

# Path to the configuration file
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
# Global variable to store cached configuration data
CONFIG_DATA = None

# endregion

# =============================================================================
# CONFIGURATION
# =============================================================================
# region


def load_config():
    """Load configuration from file or return default settings."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "PERSONAL_LICENSE": config.DEFAULT_PERSONAL_LICENSE,
        "OLD_XML": config.DEFAULT_OLD_XML,
        "PROGRAM_NAME": config.DEFAULT_PROGRAM_NAME,
        "PROGRAM_NUMBER": config.DEFAULT_PROGRAM_NUMBER,
        "COMMENT": config.DEFAULT_COMMENT,
        "POST_NAME": config.DEFAULT_POST_NAME,
        "POST_FOLDER": config.DEFAULT_POST_FOLDER,
        "OUTPUT_FOLDER": config.DEFAULT_OUTPUT_FOLDER,
        "UNIT": config.DEFAULT_UNIT,
        "IS_OPEN_IN_EDITOR": config.DEFAULT_IS_OPEN_IN_EDITOR,
        "ALLOW_HELICAL_MOVES": config.DEFAULT_ALLOW_HELICAL_MOVES,
        "HIGH_FEEDRATE_MAPPING_VALUE": config.DEFAULT_HIGH_FEEDRATE_MAPPING_VALUE,
        "MINIMUM_CHORD_LENGTH": config.DEFAULT_MINIMUM_CHORD_LENGTH,
        "HIGH_FEEDRATE": config.DEFAULT_HIGH_FEEDRATE,
        "MAXIMUM_CIRCULAR_RADIUS": config.DEFAULT_MAXIMUM_CIRCULAR_RADIUS,
        "MINIMUM_CIRCULAR_RADIUS": config.DEFAULT_MINIMUM_CIRCULAR_RADIUS,
        "TOLERANCE": config.DEFAULT_TOLERANCE,
    }


def save_config(new_config):
    """Save configuration to file with error handling and atomic write."""
    try:
        # Ensure config directory exists
        config_dir = os.path.dirname(CONFIG_FILE)
        if config_dir:
            os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)

        # Atomic write using temp file
        temp_file = f"{CONFIG_FILE}.tmp"
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(new_config, f, indent=4, ensure_ascii=False)

        # Replace existing config atomically
        if os.path.exists(CONFIG_FILE):
            os.remove(CONFIG_FILE)
            futil.log(f"Replacing existing config at: {normalize_path(CONFIG_FILE)}")
        os.rename(temp_file, CONFIG_FILE)
        futil.log("Config successfully saved")
        return True
    except Exception as e:
        show_message(f"Failed to save settings:\n{str(e)}")

        # Clean up temp file if it exists
        if os.path.exists(temp_file):
            try:
                os.remove(temp_file)
                futil.log(f"Cleaned up temp file: {temp_file}")
            except Exception as cleanup_e:
                futil.log(f"Failed to clean up temp file: {str(cleanup_e)}")
        return False


def config_value(key, value=None):
    """Safe get/set for config values with validation."""
    global CONFIG_DATA
    if CONFIG_DATA is None:
        CONFIG_DATA = load_config()
    if value is None:
        return CONFIG_DATA.get(key)
    try:
        # Normalize path-style values
        if key.endswith("_PATH") or key.endswith("_FOLDER"):
            value = normalize_path(str(value))
        # Update the config dictionary with the new value
        CONFIG_DATA[key] = value
        futil.log(f"Updated config key '{key}' with value: {value}")
        if not save_config(CONFIG_DATA):
            show_message(f"Failed to save config for key: {key}")
        return value
    except Exception as e:
        show_message(f"Config update failed for {key}: {str(e)}")
        return None


def save_command_configuration(inputs):
    """Save all configuration values from UI inputs."""
    try:
        global CONFIG_DATA

        # Map UI input IDs to configuration keys
        config_mapping = {
            "personal_input": "PERSONAL_LICENSE",
            "old_xml_input": "OLD_XML",
            "program_name_input": "PROGRAM_NAME",
            "program_number_input": "PROGRAM_NUMBER",
            "comment_input": "COMMENT",
            "post_name_input": "POST_NAME",
            "output_folder_input": "OUTPUT_FOLDER",
            "unit_input": "UNIT",
            "open_in_editor_input": "IS_OPEN_IN_EDITOR",
            "allow_helical_moves_input": "ALLOW_HELICAL_MOVES",
            "high_feedrate_mapping_input": "HIGH_FEEDRATE_MAPPING_VALUE",
            "minimum_chord_length_input": "MINIMUM_CHORD_LENGTH",
            "high_feedrate_input": "HIGH_FEEDRATE",
            "maximum_circular_radius_input": "MAXIMUM_CIRCULAR_RADIUS",
            "minimum_circular_radius_input": "MINIMUM_CIRCULAR_RADIUS",
            "tolerance_input": "TOLERANCE",
        }

        # Extract values from UI inputs based on the mapping
        new_config = {
            config_key: get_input_value(inputs, input_id, config_key)
            for input_id, config_key in config_mapping.items()
        }
        futil.log("Parsed configuration values from UI inputs")

        # Determine post folder path
        post_folder = config_value("POST_FOLDER") or config.DEFAULT_POST_FOLDER
        futil.log(f"Resolved post folder path: {normalize_path(post_folder)}")

        updated_config = {}

        # Assign all config values, and include POST_FOLDER when POST_NAME is used
        for key, value in new_config.items():
            updated_config[key] = value
            if key == "POST_NAME":
                updated_config["POST_FOLDER"] = post_folder

        # Save updated configuration
        CONFIG_DATA = updated_config

        if save_config(CONFIG_DATA):
            futil.log("Configuration saved successfully")
        else:
            futil.log("Failed to persist configuration")

        return True
    except Exception as e:
        show_message(f"Failed to save configuration: {str(e)}")
        return False


# endregion

# =============================================================================
# EVENT HANDLERS
# =============================================================================
# region


def start():
    """Initialize the add-in."""
    global CONFIG_DATA

    # Load configuration from file
    CONFIG_DATA = load_config()
    futil.log("Configuration loaded")

    # Save default config if config file does not exist
    if not os.path.exists(CONFIG_FILE):
        save_config(CONFIG_DATA)
        futil.log(f"Created default config file at {normalize_path(CONFIG_FILE)}")

    # Attach event handler to command creation
    cmd_def = ui.commandDefinitions.addButtonDefinition(
        CMD_ID, CMD_NAME, CMD_DESCRIPTION, ICON_FOLDER
    )
    futil.add_handler(cmd_def.commandCreated, command_created)

    # Retrieve the target workspace
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    if workspace is None:
        show_message(f"Workspace '{WORKSPACE_ID}' not found.")
        return
    futil.log(f"Workspace '{WORKSPACE_ID}' found")

    # Retrieve or create the target panel
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    if panel is None:
        panel = workspace.toolbarPanels.add(PANEL_ID, PANEL_NAME, POSITION_ID, False)
        futil.log(f"Created new panel '{PANEL_NAME}' in workspace")

    # Add the command control (button) to the panel
    control = panel.controls.addCommand(cmd_def, COMMAND_BESIDE_ID, False)
    control.isPromoted = IS_PROMOTED
    futil.log(f"Added command control to panel with promotion = {IS_PROMOTED}")


def stop():
    """Remove the command and UI elements from Fusion 360"""
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    if not workspace:
        futil.log(f"Workspace '{WORKSPACE_ID}' not found")
        return

    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    if not panel:
        futil.log(f"Panel '{PANEL_ID}' not found in workspace '{WORKSPACE_ID}'")
        return

    # Remove the command control (button) from the panel
    command_control = panel.controls.itemById(CMD_ID)
    if command_control:
        command_control.deleteMe()
        futil.log(f"Removed command control '{CMD_ID}' from panel '{PANEL_ID}'")
    else:
        futil.log(f"Command control '{CMD_ID}' not found in panel '{PANEL_ID}'")

    # Remove the command definition
    command_definition = ui.commandDefinitions.itemById(CMD_ID)
    if command_definition:
        command_definition.deleteMe()
        futil.log(f"Deleted command definition '{CMD_ID}'")
    else:
        futil.log(f"Command definition '{CMD_ID}' not found")


def command_created(args: adsk.core.CommandCreatedEventArgs):
    """Handles the event when the command is created."""
    inputs = args.command.commandInputs

    # Get the active document and CAM product
    doc = app.activeDocument
    product = doc.products.itemByProductType("CAMProductType")
    cam = adsk.cam.CAM.cast(product)

    # Check if CAM product exists and has setups
    if not cam or cam.setups.count == 0:
        show_message("No setups found in the current document.")
        return

    # Add personal license checkbox input
    personal_input = inputs.addBoolValueInput(
        "personal_input",
        "License Personal (Testing)",
        True,
        "",
        as_bool(config_value("PERSONAL_LICENSE")),
    )
    old_xml_input = inputs.addBoolValueInput(
        "old_xml_input",
        "Old xml.cps",
        True,
        "",
        as_bool(config_value("OLD_XML")),
    )
    old_xml_input.isEnabled = personal_input.value

    # Create dropdown for setup selection
    setups_combo = inputs.addDropDownCommandInput(
        "setup_selector_input",
        "Select Setup",
        adsk.core.DropDownStyles.TextListDropDownStyle,
    )

    # Populate setups dropdown with available setups
    first_item = True
    for idx in range(cam.setups.count):
        setup = cam.setups.item(idx)
        setups_combo.listItems.add(setup.name, first_item)
        first_item = False

    # Add "Selected Operations" option if any operations are selected
    has_select_op = False
    for idx in range(cam.allOperations.count):
        op = cam.allOperations.item(idx)
        if op and getattr(op, "isSelected", False):
            has_select_op = True
            break
    setups_combo.listItems.add("Selected Operations", has_select_op)

    # Add program information inputs
    inputs.addStringValueInput(
        "program_name_input", "Program Name", config_value("PROGRAM_NAME")
    )
    inputs.addStringValueInput(
        "program_number_input", "Program Number", config_value("PROGRAM_NUMBER")
    )
    inputs.addStringValueInput("comment_input", "Comment", config_value("COMMENT"))

    # Add postprocessor selection inputs
    post_name_input = inputs.addStringValueInput(
        "post_name_input", "Postprocessor", config_value("POST_NAME")
    )
    post_name_input.isReadOnly = True
    inputs.addBoolValueInput(
        "select_post_button", "Select Postprocessor", False, BUTTON_ICON, True
    )

    # Add output folder selection inputs
    def_out_folder = normalize_path(config_value("OUTPUT_FOLDER"))
    output_folder = inputs.addStringValueInput(
        "output_folder_input", "Output Folder", def_out_folder
    )
    output_folder.isReadOnly = True
    inputs.addBoolValueInput(
        "select_output_folder_button", "Select Output Folder", False, BUTTON_ICON, True
    )

    # Add units selection dropdown
    unit_input = inputs.addDropDownCommandInput(
        "unit_input", "Unit", adsk.core.DropDownStyles.TextListDropDownStyle
    )
    default_unit = config_value("UNIT")
    for unit_item in UNIT_ITEMS:
        is_selected = unit_item == default_unit
        unit_input.listItems.add(unit_item, is_selected)

    # Add option to open NC file in editor after generation
    inputs.addBoolValueInput(
        "open_in_editor_input",
        "Open NC file in Editor",
        True,
        "",
        as_bool(config_value("IS_OPEN_IN_EDITOR")),
    )

    # Create a collapsible group for built-in post parameters
    group_built_in = inputs.addGroupCommandInput(
        "group_built_in", "Built-in Post Parameters"
    )
    group_built_in.isExpanded = False
    built_in_items = group_built_in.children

    built_in_items.addBoolValueInput(
        "allow_helical_moves_input",
        "Allow Helical Moves",
        True,
        "",
        as_bool(config_value("ALLOW_HELICAL_MOVES")),
    )

    # Add feedrate mapping dropdown
    feedrate_mapping_input = built_in_items.addDropDownCommandInput(
        "high_feedrate_mapping_input",
        "High Feedrate Mapping",
        adsk.core.DropDownStyles.TextListDropDownStyle,
    )

    default_feedrate_item = config_value("HIGH_FEEDRATE_MAPPING_VALUE")
    if default_feedrate_item not in HIGH_FEED_MAPPING_ITEMS:
        default_feedrate_item = HIGH_FEED_MAPPING_ITEMS[0]
    # Populate feedrate mapping dropdown
    for feedrate_mapping_item in HIGH_FEED_MAPPING_ITEMS:
        is_selected = feedrate_mapping_item == default_feedrate_item
        feedrate_mapping_input.listItems.add(feedrate_mapping_item, is_selected)

    # Add 'Minimum Chord Length' input
    built_in_items.addStringValueInput(
        "minimum_chord_length_input",
        "Minimum Chord Length (mm)",
        config_value("MINIMUM_CHORD_LENGTH"),
    )
    # Add 'High Feedrate' input
    built_in_items.addStringValueInput(
        "high_feedrate_input", "High Feedrate (mm/min)", config_value("HIGH_FEEDRATE")
    )
    # Add 'Maximum circular radius' input
    built_in_items.addStringValueInput(
        "maximum_circular_radius_input",
        "Maximum Circular Radius (mm)",
        config_value("MAXIMUM_CIRCULAR_RADIUS"),
    )
    # Add 'Minimum circular radius' input
    built_in_items.addStringValueInput(
        "minimum_circular_radius_input",
        "Minimum Circular Radius (mm)",
        config_value("MINIMUM_CIRCULAR_RADIUS"),
    )
    # Add 'Tolerance' input
    built_in_items.addStringValueInput(
        "tolerance_input", "Tolerance (mm)", config_value("TOLERANCE")
    )

    futil.add_handler(
        args.command.execute, command_execute, local_handlers=local_handlers
    )
    futil.add_handler(
        args.command.inputChanged, command_input_changed, local_handlers=local_handlers
    )
    futil.add_handler(
        args.command.validateInputs,
        command_validate_input,
        local_handlers=local_handlers,
    )
    futil.add_handler(
        args.command.destroy, command_destroy, local_handlers=local_handlers
    )
    futil.log("Command created event completed successfully")


def command_validate_input(args: adsk.core.ValidateInputsEventArgs):
    """Validates inputs and blocks OK button if invalid."""
    inputs = args.inputs
    args.areInputsValid = True

    # Validate program name (must not be empty)
    if not inputs.itemById("program_name_input").value.strip():
        args.areInputsValid = False

    program_number = inputs.itemById("program_number_input").value
    if not (program_number.isdigit() and int(program_number) >= 0):
        args.areInputsValid = False

    # Validate program number (must be positive integer)
    post_name = inputs.itemById("post_name_input").value

    global POST_PATH
    POST_PATH = os.path.join(config_value("POST_FOLDER"), post_name)

    # Validate post processor exists
    if not post_name.strip() or not os.path.exists(POST_PATH):
        inputs.itemById("post_name_input").value = ""
        POST_PATH = ""
        config_value("POST_NAME", "")
        config_value("POST_FOLDER", config.DEFAULT_POST_FOLDER)
        futil.log("Post processor not found and resetting to default: {POST_PATH}")
        args.areInputsValid = False

    # List of float fields
    float_fields = [
        ("minimum_chord_length_input", is_positive_float),
        ("high_feedrate_input", is_non_negative_float),
        ("maximum_circular_radius_input", is_positive_float),
        ("minimum_circular_radius_input", is_positive_float),
        ("tolerance_input", is_positive_float),
    ]

    # Validate all numeric fields
    for field_id, validator in float_fields:
        input_item = inputs.itemById(field_id)
        if input_item:
            try:
                value = float(input_item.value)
                if not validator(value):
                    args.areInputsValid = False
                    break
            except (ValueError, TypeError):
                args.areInputsValid = False
                break


def command_input_changed(args: adsk.core.InputChangedEventArgs):
    """Handles changes in input fields within the command dialog."""
    changed_input = args.input
    inputs = args.inputs

    if changed_input.id == "personal_input":
        inputs.itemById("old_xml_input").isEnabled = bool(changed_input.value)

    # Handle postprocessor selection button click
    if changed_input.id == "select_post_button":

        # Set file dialog properties
        file_dlg = ui.createFileDialog()
        file_dlg.title = "Select Postprocessor"
        file_dlg.initialDirectory = os.path.abspath(config_value("POST_FOLDER"))

        # Show dialog and process result
        if file_dlg.showOpen() == adsk.core.DialogResults.DialogOK:

            file_path = normalize_path(file_dlg.filename)
            file_name = normalize_path(os.path.basename(file_path))
            file_folder = normalize_path(os.path.dirname(file_path))

            global POST_PATH

            # Validate selected file exists
            if not os.path.exists(file_path):
                show_message("Selected postprocessor file does not exist!")
                inputs.itemById("post_name_input").value = ""
                POST_PATH = ""
                config_value("POST_NAME", "")
                config_value("POST_FOLDER", config.DEFAULT_POST_FOLDER)
                futil.log(
                    "Post processor not found and resetting to default: {POST_PATH}"
                )
                return

            # Update UI and configuration

            inputs.itemById("post_name_input").value = file_name
            config_value("POST_NAME", file_name)
            config_value("POST_FOLDER", file_folder)
            futil.log(f"Postprocessor selected: {file_name}")

    # Handle output folder selection button click
    elif changed_input.id == "select_output_folder_button":

        # Set folder dialog properties
        folder_dlg = ui.createFolderDialog()
        folder_dlg.title = "Select Output Folder"
        initial_directory = inputs.itemById("output_folder_input").value
        folder_dlg.initialDirectory = os.path.abspath(initial_directory)

        # Show dialog and process result
        if folder_dlg.showDialog() == adsk.core.DialogResults.DialogOK:

            # Update UI and configuration
            output_folder = normalize_path(folder_dlg.folder)
            inputs.itemById("output_folder_input").value = output_folder
            config_value("OUTPUT_FOLDER", output_folder)
            futil.log(f"Output folder selected: {output_folder}")

    futil.log("Input changed event processed successfully")


def command_destroy(_args: adsk.core.CommandEventArgs):
    """Cleans up event handlers when the command is destroyed."""
    global local_handlers
    local_handlers = []
    futil.log("Command destroy event completed - handlers cleaned up")


def command_execute(args: adsk.core.CommandEventArgs):
    """Handles the command execution with improved error handling and workflow separation."""
    try:
        inputs = args.command.commandInputs

        # Save configuration
        if not save_command_configuration(inputs):
            return

        # Get active document and CAM product
        doc = app.activeDocument
        cam = adsk.cam.CAM.cast(doc.products.itemByProductType("CAMProductType"))
        if not cam or cam.setups.count == 0:
            show_message("No CAM setups found")
            return

        # Determine selected operations
        setup_selector = inputs.itemById("setup_selector_input").selectedItem.name
        if setup_selector == "Selected Operations":
            # Get operations from selected operations
            operations = []
            for idx in range(cam.allOperations.count):
                op = cam.allOperations.item(idx)
                if op and getattr(op, "isSelected", False):
                    operations.append(op)
            if not operations:
                show_message("No operations selected")
                return
            operations = [op for op in operations if op.parent == operations[0].parent]
            futil.log(f"Found {len(operations)} selected operations")
        else:
            # Get operations from specific setup
            setup_number = get_setup_number(setup_selector, cam)
            if setup_number is None:
                show_message(f"Setup '{setup_selector}' not found")
                return

            operations = []
            setup = cam.setups.item(setup_number)
            setup_ops = getattr(setup, 'allOperations', None)
            setup_ops_count = getattr(setup_ops, 'count', 0) if setup_ops else 0
            for idx in range(setup_ops_count):
                op = setup_ops.item(idx)
                if op and getattr(op, 'hasToolpath', False):
                    operations.append(op)
            futil.log(f"Found {len(operations)} operations in setup {setup_selector}")

        if not operations:
            show_message("No valid operations with toolpaths found")
            return

        # Collect and validate processing parameters
        params = collect_processing_parameters(inputs)
        if not params:
            show_message("Failed to collect processing parameters")
            return
        # Execute appropriate workflow based on license type
        # if params['personal_license'] and is_hobbyist_license():
        if params["personal_license"]:
            execute_personal_workflow(cam, operations, params)
        else:
            execute_standard_workflow(cam, operations, params)
    except Exception as e:
        show_message(f"Error: {str(e)}")


# endregion

# =============================================================================
# PARAMETERS
# =============================================================================
# region


def setup_logging(log_file):
    """Configure logging system"""

    mode = logging.DEBUG if config.DEBUG else logging.INFO
    futil.log(
        f"{'Debug mode enabled' if config.DEBUG else 'Debug mode disabled'}. Logging to {log_file}"
    )

    logging.basicConfig(
        filename=log_file,
        level=mode,
        format="%(asctime)s - %(levelname)s - %(message)s",
        filemode="w",
        force=True,
    )


def collect_processing_parameters(inputs):
    """Collect and validate all processing parameters from UI."""
    try:
        # Get the selected unit
        selected_unit_item = get_input_value(inputs, "unit_input", "Unit")
        if selected_unit_item not in UNIT_ITEMS:
            raise ValueError(
                f"Invalid selection for 'UNIT_ITEMS': {selected_unit_item}"
            )

        unit_num = UNIT_ITEMS.index(selected_unit_item)
        unit_text = (
            UNIT_ITEMS[unit_num].replace(" ", "")
            if unit_num == 2
            else UNIT_ITEMS[unit_num]
        )
        futil.log(f"Unit selected: {unit_text} (index {unit_num})")

        # Get the selected "High Feedrate Mapping" value
        selected_hfm_name = get_input_value(
            inputs, "high_feedrate_mapping_input", "High Feedrate Mapping"
        )
        if selected_hfm_name not in HIGH_FEED_MAPPING_ITEMS:
            raise ValueError(
                f"Invalid selection for 'HIGH_FEED_MAPPING_ITEMS': {selected_hfm_name}"
            )

        feed_map_idx = HIGH_FEED_MAPPING_ITEMS.index(selected_hfm_name)
        high_feedrate_mapping = str(
            feed_map_idx + 2 if feed_map_idx == 3 else feed_map_idx
        )  # Adjust for specific index
        futil.log(f"High feedrate mapping selected: {high_feedrate_mapping}")

        # Create a dictionary params
        params = {
            "personal_license": get_input_value(inputs, "personal_input", "License"),
            "old_xml": get_input_value(inputs, "old_xml_input", "Old xml.cps"),
            "program_name": get_input_value(
                inputs, "program_name_input", "Program Name"
            ),
            "program_number": get_input_value(
                inputs, "program_number_input", "Program Number"
            ),
            "comment": get_input_value(inputs, "comment_input", "Comment"),
            "output_folder": get_input_value(
                inputs, "output_folder_input", "Output Folder"
            ),
            "unit_text": unit_text,
            "unit_num": unit_num,
            "post_path": POST_PATH,
            "open_in_editor": get_input_value(
                inputs, "open_in_editor_input", "Open in Editor"
            ),
            "allow_helical_moves": get_input_value(
                inputs, "allow_helical_moves_input", "Allow Helical Moves"
            ),
            "high_feedrate_mapping": high_feedrate_mapping,
            "min_chord_length": get_input_value(
                inputs, "minimum_chord_length_input", "Minimum Chord Length"
            ),
            "high_feedrate": get_input_value(
                inputs, "high_feedrate_input", "High Feedrate"
            ),
            "max_circ_radius": get_input_value(
                inputs, "maximum_circular_radius_input", "Maximum Circular Radius"
            ),
            "min_circ_radius": get_input_value(
                inputs, "minimum_circular_radius_input", "Minimum Circular Radius"
            ),
            "tolerance_value": get_input_value(inputs, "tolerance_input", "Tolerance"),
        }

        if not os.path.exists(params["post_path"]):
            show_message("Invalid post processor path")
            return None

        return params
    except Exception as e:
        show_message(f"Parameter error: {str(e)}")
        return None


def execute_personal_workflow(cam, operations, params):
    """Execute workflow for Personal/Hobbyist license."""

    # Validate parameters and convert them to floats
    unit = params["unit_num"]
    high_feedrate = float(params["high_feedrate"])  # Default value in mm
    min_chord_length = float(params["min_chord_length"])
    max_circ_radius = float(params["max_circ_radius"])
    min_circ_radius = float(params["min_circ_radius"])
    tolerance = float(params["tolerance_value"])

    # Get document unit if necessary
    if unit == 2:  # Document Unit
        doc_unit = get_document_units()
        if doc_unit is None:
            show_message("Document unit not found")
            return False
        unit = doc_unit

    # Convert parameters from mm to inches if unit is 0 (Inches)
    if unit == 0:  # Inches
        conversion_factor = 25.4  # Conversion from mm to inches
        high_feedrate /= conversion_factor
        min_chord_length /= conversion_factor
        max_circ_radius /= conversion_factor
        min_circ_radius /= conversion_factor
        tolerance /= conversion_factor

    # Prepare post-processing parameters
    post_params = {
        "program_number": params["program_number"],
        "program_name": params["program_name"],
        "comment": params["comment"],
        "post_path": params["post_path"],
        "output_folder": params["output_folder"],
        "unit": unit,
        "open_in_editor": params["open_in_editor"],
        "allowHelicalMoves": params["allow_helical_moves"],
        "highFeedMapping": params["high_feedrate_mapping"],
        "minimumChordLength": min_chord_length,
        "highFeedrate": high_feedrate,
        "maximumCircularRadius": max_circ_radius,
        "minimumCircularRadius": min_circ_radius,
        "tolerance": tolerance,
        "old_xml": params["old_xml"],
    }

    futil.log("Post-processing parameters prepared")

    # Execute batch post-processing with the prepared parameters
    if not batch_post(cam, operations, **post_params):
        show_message("Failed to process operations in Personal mode")
        return False

    return True


# endregion


# =============================================================================
# POSTPROCESSING
# =============================================================================
# region
def execute_standard_workflow(cam, operations, params):
    """Execute standard NC Program workflow."""
    try:
        futil.log("===================================", force_console=True)
        futil.log("=== Standard G-code generation ===", force_console=True)
        futil.log("===================================", force_console=True)

        # Create NC Program input
        start_time = time.time()
        nc_input = cam.ncPrograms.createInput()
        nc_input.displayName = get_unique_nc_program_name(cam)
        nc_input.operations = operations

        # Set NC Program parameters
        nc_params = nc_input.parameters
        nc_params.itemByName("nc_program_name").value.value = params["program_number"]
        nc_params.itemByName("nc_program_filename").value.value = params["program_name"]
        nc_params.itemByName("nc_program_comment").value.value = params["comment"]
        nc_params.itemByName("nc_program_output_folder").value.value = params[
            "output_folder"
        ]
        nc_params.itemByName("nc_program_unit").value.value = params["unit_text"]
        nc_params.itemByName("nc_program_openInEditor").value.value = bool(
            params["open_in_editor"]
        )
        nc_params.itemByName("nc_program_info_nc_extension").value.value = "nc"

        futil.log(
            f"Setting NC Program parameters:\n"
            f"  nc_program_name: {params['program_number']}\n"
            f"  nc_program_filename: {params['program_name']}\n"
            f"  nc_program_comment: {params['comment']}\n"
            f"  nc_program_output_folder: {params['output_folder']}\n"
            f"  nc_program_unit: {params['unit_text']}\n"
            f"  nc_program_openInEditor: {bool(params['open_in_editor'])}\n"
            f"  nc_program_info_nc_extension: 'nc'"
        )

        # Add and validate the NC Program
        new_program = cam.ncPrograms.add(nc_input)
        if not new_program:
            raise RuntimeError("Failed to create NC Program")
        new_program.name = get_unique_nc_program_name(cam)

        # Configure post processor
        post_config = get_post(params["post_path"])
        if not post_config:
            raise RuntimeError("Post processor not found")
        new_program.postConfiguration = post_config

        # Add post parameters
        post_params = new_program.postParameters
        post_params.itemByName("builtin_allowHelicalMoves").value.value = bool(
            params["allow_helical_moves"]
        )
        post_params.itemByName("builtin_highFeedMapping").value.value = params[
            "high_feedrate_mapping"
        ]
        post_params.itemByName("builtin_minimumChordLength").value.value = fix_units(
            params["min_chord_length"]
        )
        post_params.itemByName("builtin_highFeedrate").value.value = float(
            params["high_feedrate"]
        )
        post_params.itemByName("builtin_maximumCircularRadius").value.value = fix_units(
            params["max_circ_radius"]
        )
        post_params.itemByName("builtin_minimumCircularRadius").value.value = fix_units(
            params["min_circ_radius"]
        )
        post_params.itemByName("builtin_tolerance").value.value = fix_units(
            params["tolerance_value"]
        )

        futil.log(
            f"Setting Post parameters:\n"
            f"  builtin_allowHelicalMoves: {bool(params['allow_helical_moves'])}\n"
            f"  builtin_highFeedMapping: {params['high_feedrate_mapping']}\n"
            f"  builtin_minimumChordLength: {fix_units(params['min_chord_length'])}\n"
            f"  builtin_highFeedrate: {float(params['high_feedrate'])}\n"
            f"  builtin_maximumCircularRadius: {fix_units(params['max_circ_radius'])}\n"
            f"  builtin_minimumCircularRadius: {fix_units(params['min_circ_radius'])}\n"
            f"  builtin_tolerance: {fix_units(params['tolerance_value'])}"
        )

        # Update post parameters
        new_program.updatePostParameters(post_params)
        post_options_type = getattr(adsk.cam, "NCProgramPostProcessOptions", None)
        if post_options_type is None:
            raise AttributeError("NCProgramPostProcessOptions not available")
        post_options = post_options_type.create()

        # Postprocess NC Program
        new_program.postProcess(post_options)

        # Verify output
        if not new_program.hasError:
            nc_file = normalize_path(
                os.path.join(params["output_folder"], f"{params['program_name']}.nc")
            )

            if os.path.exists(nc_file):
                file_size = os.path.getsize(nc_file)
                futil.log(
                    f"Successfully generated NC file: {params['program_name']} ({file_size} bytes)",
                    force_console=True,
                )
                futil.log(f"File path: {nc_file}", force_console=True)
            else:
                futil.log("Error: Output NC file was not created", force_console=True)
        else:
            futil.log("Error: Post processing failed", force_console=True)

        exec_time = time.time() - start_time
        futil.log(
            f"G-code generation completed in {exec_time:.2f} seconds",
            force_console=True,
        )
    except Exception as e:
        show_message(f"Standard workflow error: {str(e)}")
        futil.log(f"Standard workflow error: {str(e)}", force_console=True)


def batch_post(cam, operations, **post_params):
    """Batch postprocessing with XML merging for Fusion 360 Personal license."""
    start_time = time.time()

    # Validate critical paths
    missing_files = []
    xml_post_file = get_personal_xml_post(post_params.get("old_xml", False))
    xml_post_mode = "legacy" if as_bool(post_params.get("old_xml", False)) else "current"
    futil.log(
        f"Personal intermediate post ({xml_post_mode}): "
        f"{normalize_path(xml_post_file)}",
        force_console=True,
    )
    if not os.path.exists(xml_post_file):
        missing_files.append(normalize_path(xml_post_file))

    post_exe_path = find_fusion_post_exe()
    if not post_exe_path:
        missing_files.append(normalize_path("post.exe"))

    if missing_files:
        error_msg = "Missing required files:\n" + "\n".join(
            f"• {f}" for f in missing_files
        )
        show_message(error_msg)
        return False

    # Get parameters from **post_params
    output_folder = normalize_path(post_params["output_folder"])
    program_name = post_params["program_name"]
    pgm_num = post_params["program_number"]
    comment = post_params["comment"]
    post_processor = normalize_path(post_params["post_path"])
    unit = post_params["unit"]

    futil.log("=== Batch Post Parameters ===")
    futil.log(f"Output folder: {output_folder}")
    futil.log(f"Program name: {program_name}")
    futil.log(f"Program number: {pgm_num}")
    futil.log(f"Comment: {comment}")
    futil.log(f"Post processor path: {post_processor}")
    futil.log(f"Unit: {unit}")

    # Isolate every intermediate artifact so cleanup can never remove or
    # overwrite a user's existing XML/log/backup files.
    os.makedirs(output_folder, exist_ok=True)
    work_folder = tempfile.mkdtemp(
        prefix=f".{program_name}_smartpost_", dir=output_folder
    )
    log_fd, log_path = tempfile.mkstemp(
        prefix="post_", suffix=".log", dir=work_folder
    )
    os.close(log_fd)
    progress_fd, progress_path = tempfile.mkstemp(
        prefix="progress_", suffix=".tmp", dir=work_folder
    )
    os.close(progress_fd)
    log_path = normalize_path(log_path)
    progress_path = normalize_path(progress_path)
    setup_logging(progress_path)

    try:
        # Create progress dialog
        progress_dialog = ui.createProgressDialog()
        progress_dialog.isCancelButtonShown = False
        progress_show(progress_dialog, "Batch Post Processing", "Initializing...", 0, 3)
        do_events()
        time.sleep(0.05)

        # Process each operation to generate XML files
        processed_ops = process_operations(
            cam,
            operations,
            program_name,
            xml_post_file,
            work_folder,
            unit,
            post_params,
        )

        if not processed_ops:
            raise RuntimeError("No XML files generated for merging")

        progress_dialog.message = "Processing operation completed"
        progress_dialog.progressValue = 1
        do_events()
        time.sleep(0.05)

        merged_xml = normalize_path(
            os.path.join(work_folder, f"{program_name}_merged.xml")
        )

        try:
            if len(processed_ops) == 1:
                # For single file
                os.replace(processed_ops[0], merged_xml)
            else:
                # For multiple files
                if not merge_xml_files(processed_ops, merged_xml):
                    raise RuntimeError("XML merging failed")

        except OSError as e:
            raise RuntimeError(f"Failed to create merged XML file: {str(e)}") from e

        progress_dialog.message = "Merging XML files completed"
        progress_dialog.progressValue = 2
        do_events()
        time.sleep(0.05)

        # G-code generation
        nc_file = normalize_path(os.path.join(output_folder, f"{program_name}.nc"))

        if not generate_gcode(
            post_exe_path,
            post_processor,
            merged_xml,
            nc_file,
            pgm_num,
            unit,
            post_params,
            log_path,
        ):
            raise RuntimeError("G-code generation failed")

        exec_time = time.time() - start_time
        futil.log(
            f"G-code generation completed in {exec_time:.2f} seconds",
            force_console=True,
        )
        progress_dialog.message = "G-code generation completed"
        progress_dialog.progressValue = 3
        do_events()
        time.sleep(0.05)
        progress_dialog.hide()
        return True
    except Exception as e:
        if "progress_dialog" in locals():
            progress_dialog.hide()
        futil.log(f"Batch Post error:\n{str(e)}", force_console=True)
        show_message(f"Batch Post error:\n{str(e)}")
        return False
    finally:
        # Only remove files whose unique names were created by this invocation.
        logging.shutdown()
        for temporary_path in (log_path, progress_path):
            try:
                if os.path.exists(temporary_path):
                    os.remove(temporary_path)
                    futil.log(f"Removed SmartPost temporary file: {temporary_path}")
            except OSError as cleanup_error:
                futil.log(
                    f"Warning: could not remove {temporary_path}: {cleanup_error}",
                    force_console=True,
                )
        resolved_output = os.path.realpath(output_folder)
        resolved_work = os.path.realpath(work_folder)
        is_own_work_folder = (
            os.path.commonpath((resolved_output, resolved_work)) == resolved_output
            and os.path.dirname(resolved_work) == resolved_output
            and os.path.basename(resolved_work).startswith(f".{program_name}_smartpost_")
        )
        if is_own_work_folder:
            try:
                shutil.rmtree(resolved_work)
            except OSError as cleanup_error:
                futil.log(
                    f"Warning: could not remove SmartPost work folder "
                    f"{resolved_work}: {cleanup_error}",
                    force_console=True,
                )
        else:
            futil.log(
                f"Refusing to remove unexpected work folder: {resolved_work}",
                force_console=True,
            )


def merge_xml_files(file_paths, output_file):
    """Merge operation-level intermediate XML files."""
    return personal_pipeline.merge_xml_files(file_paths, output_file, show_message)


def process_operations(
    cam, operations, program_name, post_processor, output_folder, unit, post_params
):
    """Post individual operations to numbered intermediate XML files."""
    return personal_pipeline.process_operations(
        cam,
        operations,
        program_name,
        post_processor,
        output_folder,
        unit,
        post_params,
        show_message,
    )

def generate_gcode(
    post_exe_path,
    post_processor,
    merged_xml,
    nc_file,
    pgm_num,
    unit,
    post_params,
    log_path,
):
    """Execute post.exe to generate final G-code"""
    futil.log("==================================", force_console=True)
    futil.log("=== Starting G-code generation ===", force_console=True)
    futil.log("==================================", force_console=True)

    # Build command parameters
    params = [
        f'"{normalize_path(post_exe_path)}"',
        "--log",
        f'"{normalize_path(log_path)}"',
        "--nobackup",
        "--allowui",
        "--sandbox",
        "--lang",
        "en",
    ]

    # Open NC File in Editor
    if not post_params.get("open_in_editor", False):
        params.append("--noeditor")

    # Set unit and rate suffixes
    unit_suffix, unit_rate = ("in", "in/min") if unit == 0 else ("mm", "mm/min")

    # Extended parameters for post.exe
    params.extend(
        [
            "--property",
            "allowHelicalMoves",
            str(post_params["allowHelicalMoves"]).lower(),
            "--property",
            "highFeedMapping",
            str(post_params["highFeedMapping"]),
            "--property",
            "minimumChordLength",
            f"{post_params.get('minimumChordLength', 0)}{unit_suffix}",
            "--property",
            "highFeedrate",
            f"{post_params.get('highFeedrate', 0)}{unit_rate}",
            "--property",
            "maximumCircularRadius",
            f"{post_params.get('maximumCircularRadius', 0)}{unit_suffix}",
            "--property",
            "minimumCircularRadius",
            f"{post_params.get('minimumCircularRadius', 0)}{unit_suffix}",
            "--property",
            "tolerance",
            f"{post_params.get('tolerance', 0)}{unit_suffix}",
            "--property",
            "programComment",
            f"'{post_params['comment']}'",
            "--property",
            "programName",
            str(pgm_num),
            "--property",
            "unit",
            str(unit),
            f'"{normalize_path(post_processor)}"',
            f'"{normalize_path(merged_xml)}"',
            f'"{normalize_path(nc_file)}"',
        ]
    )

    # Debug mode
    if logging.getLogger().level == logging.DEBUG:
        params.insert(3, "--debug")

    futil.log("Final post.exe command:")
    futil.log(" ".join(params))

    try:
        # Execute post processor
        futil.log("Starting post.exe process...")
        result = subprocess.run(
            " ".join(params),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=60,  # 1 min
            shell=True,
            check=False,
        )

        # Process results
        if result.returncode != 0:
            error_message = ERROR_CODES.get(result.returncode, "Unknown error code")
            logging.error(
                "post.exe failed with code %s: %s",
                result.returncode,
                error_message,
            )
            futil.log(
                f"post.exe failed with return code {result.returncode}: {error_message}",
                force_console=True,
            )
            if os.path.exists(log_path):
                with open(log_path, "r", encoding="utf-8", errors="replace") as f:
                    logging.error("%s", f.read())
            return False

        # Verify output
        if not os.path.exists(nc_file):
            futil.log("Error: Output NC file was not created")
            return False

        file_size = os.path.getsize(nc_file)
        futil.log(
            f"Successfully generated NC file ({file_size} bytes)", force_console=True
        )
        futil.log(f"File path: {nc_file}", force_console=True)

        # Clean up: delete temporary merged XML file
        try:
            if os.path.exists(merged_xml):
                os.remove(merged_xml)
                futil.log(f"Deleted temporary file: {merged_xml}")
            if os.path.exists(log_path):
                os.remove(log_path)
                futil.log(f"Deleted temporary log file: {log_path}")
        except OSError as e:
            futil.log(
                f"Warning: Could not delete temporary file {merged_xml}: {str(e)}",
                force_console=True,
            )

        return True

    except subprocess.TimeoutExpired:
        logging.error("Error: Post processing timed out after 1 minute")
        futil.log("Error: Post processing timed out after 1 minute", force_console=True)
        return False
    except Exception as e:
        logging.error("Post execution error: %s", str(e))
        futil.log(f"Post execution error: {str(e)}", force_console=True)
        return False


# endregion

# =============================================================================
# OTHER FUNCTIONS
# =============================================================================
# region


find_fusion_post_exe = fusion_helpers.find_fusion_post_exe
get_setup_number = fusion_helpers.get_setup_number
get_unique_nc_program_name = fusion_helpers.get_unique_nc_program_name
get_input_value = fusion_helpers.get_input_value
normalize_path = fusion_helpers.normalize_path
is_positive_float = fusion_helpers.is_positive_float
is_non_negative_float = fusion_helpers.is_non_negative_float


def get_post(post_path):
    return fusion_helpers.get_post(post_path, show_message)


def is_hobbyist_license():
    return fusion_helpers.is_hobbyist_license(app)


def get_document_units():
    return fusion_helpers.get_document_units(app)


def fix_units(value: float):
    return fusion_helpers.fix_units(value, show_message)


# endregion
