"""Small Fusion and path helpers shared by the Smart Post command."""

import glob
import json
import os
import shutil

import adsk.cam
import adsk.core
import adsk.fusion

try:
    from ...lib import fusionAddInUtils as futil
except ImportError:  # Fallback for Fusion's varying sys.path/package loading.
    from lib import fusionAddInUtils as futil  # type: ignore


def normalize_path(path):
    """Normalize a file path and use separators accepted by Fusion."""
    normalized = os.path.normpath(os.path.expandvars(path))
    return normalized.replace("\\", "/")


def find_fusion_post_exe():
    """Auto-detect the newest installed Fusion post.exe."""
    local_app_data = os.getenv("LOCALAPPDATA")
    if not local_app_data:
        return None
    fusion_appdata = os.path.join(local_app_data, r"Autodesk\webdeploy")
    if not os.path.exists(fusion_appdata):
        return None
    possible_paths = glob.glob(
        os.path.join(fusion_appdata, "*", "*", "Applications", "CAM360", "post.exe")
    )
    return max(possible_paths, key=os.path.getmtime) if possible_paths else None


def get_setup_number(setup_name, cam):
    """Find the index of a setup by name in the CAM environment."""
    setup_name = setup_name.strip().lower()
    for index, setup in enumerate(cam.setups):
        if setup.name.strip().lower() == setup_name:
            return index
    return None


def get_post(post_path, show_message):
    """Retrieve a post configuration from the Fusion user library."""
    post_library_path = os.path.join(
        os.path.expanduser("~"),
        "AppData",
        "Roaming",
        "Autodesk",
        "Fusion 360 CAM",
        "Posts",
    )
    target_post_name = os.path.basename(post_path)
    target_path = os.path.join(post_library_path, target_post_name)

    if not os.path.exists(target_path):
        shutil.copy(post_path, post_library_path)
        show_message(
            f"Postprocessor '{target_post_name}' not found in "
            f"{normalize_path(post_library_path)}. Copying from library..."
        )

    cam_manager = adsk.cam.CAMManager.get()
    post_library = getattr(cam_manager.libraryManager, "postLibrary", None)
    if post_library is None:
        raise AttributeError("Post library manager is unavailable")

    locations_type = getattr(
        adsk.cam, "LibraryLocations", getattr(adsk.cam, "LibraryLocation", None)
    )
    if locations_type is None:
        raise AttributeError("LibraryLocations is unavailable")
    local_location = getattr(
        locations_type,
        "LocalLibraryLocation",
        getattr(locations_type, "LocalLibrary", None),
    )
    user_folder = post_library.urlByLocation(local_location)

    for user_post in post_library.childAssetURLs(user_folder):
        post_name = user_post.toString()
        if target_post_name in post_name:
            url_type = getattr(adsk.core, "URL", None)
            post_url = url_type.create(post_name) if url_type else post_name
            return post_library.postConfigurationAtURL(post_url)

    show_message(f"Could not find Postprocessor '{target_post_name}' in user library")
    return None


def get_unique_nc_program_name(cam, base_name="NCProgram"):
    """Generate a unique NC Program name."""
    existing_names = {program.name for program in cam.ncPrograms}
    counter = 1
    while True:
        new_name = f"{base_name}{counter}"
        if new_name not in existing_names:
            return new_name
        counter += 1


def get_input_value(inputs, input_id, param_name):
    """Safely retrieve a value from a UI input, including dropdowns."""
    input_item = inputs.itemById(input_id)
    if not input_item:
        raise KeyError(f"Input for '{param_name}' (ID: {input_id}) not found")
    if hasattr(input_item, "selectedItem"):
        return input_item.selectedItem.name
    return input_item.value


def is_hobbyist_license(app):
    """Check whether the active Fusion license is Personal/Hobbyist."""
    try:
        license_info = app.executeTextCommand("Application.LicenseInformation")
        license_data = json.loads(license_info)
        result = any(
            service.get(".isHobbyistLicense", "false") == "true"
            for service in license_data.values()
        )
        futil.log(f"Hobbyist license check result: {result}")
        return result
    except (ValueError, TypeError, RuntimeError, AttributeError):
        return False


def get_document_units(app):
    """Return document length units as 0 for inches or 1 for millimeters."""
    try:
        document = app.activeDocument
        design = adsk.fusion.Design.cast(
            document.products.itemByProductType("DesignProductType")
        )
        if not design:
            return 1
        return 0 if design.unitsManager.defaultLengthUnits == "inch" else 1
    except (AttributeError, RuntimeError, TypeError, ValueError):
        return 1


def is_positive_float(value):
    """Return whether a value represents a positive float."""
    try:
        return float(value) > 0
    except (ValueError, TypeError):
        return False


def is_non_negative_float(value):
    """Return whether a value represents a non-negative float."""
    try:
        return float(value) >= 0
    except (ValueError, TypeError):
        return False


def fix_units(value, show_message):
    """Correct parameters that Fusion converts to centimeters instead of mm."""
    try:
        return float(value) / 10
    except (ValueError, TypeError):
        show_message(f"Cannot convert value {value} to a number.")
        return value
