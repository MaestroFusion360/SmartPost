"""Intermediate XML generation and merging for Fusion Personal mode."""

# The Fusion API requires a broad parameter set for each post operation.
# pylint: disable=too-many-arguments,too-many-positional-arguments,too-many-locals,too-many-branches

import os

import adsk.cam
import adsk.core

try:
    from ...lib import fusionAddInUtils as futil
except ImportError:  # Fallback for Fusion's varying sys.path/package loading.
    from lib import fusionAddInUtils as futil  # type: ignore

try:
    from .fusion_helpers import normalize_path
except ImportError:  # Fallback when Fusion loads the command as a top-level module.
    from fusion_helpers import normalize_path  # type: ignore


def merge_xml_files(file_paths, output_file, show_message):
    """Merge operation-level intermediate XML files into one NC document."""
    futil.log("==============================", force_console=True)
    futil.log("======= Merging files ========", force_console=True)
    futil.log("==============================", force_console=True)

    for file_path in file_paths:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"XML file not found: {file_path}")

    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        with open(output_file, "w", encoding="utf-8") as out_file:
            with open(file_paths[0], "r", encoding="utf-8") as first_file:
                first_content = first_file.read()
                nc_end = first_content.rfind("</nc>")
                if nc_end == -1:
                    raise ValueError(
                        "First file is not valid NC XML (missing </nc> tag)"
                    )
                out_file.write(first_content[:nc_end])

            for index, file_path in enumerate(file_paths[1:], 1):
                with open(file_path, "r", encoding="utf-8") as current_file:
                    content = current_file.read()
                    nc_end = content.rfind("</nc>")
                    if nc_end == -1:
                        futil.log(
                            f"Warning: Invalid NC XML in {file_path}, skipping",
                            force_console=True,
                        )
                        continue

                    spindle_param = max(
                        content.find("<parameter name='areBothSpindlesGrabbed'"),
                        content.find('<parameter name="areBothSpindlesGrabbed"'),
                    )
                    section_start = max(content.find("<tool"), content.find("<section"))

                    if spindle_param != -1 and section_start != -1:
                        out_file.write(
                            "\n" + content[spindle_param:section_start].strip()
                        )
                    if section_start != -1:
                        out_file.write("\n" + content[section_start:nc_end].strip())
                    elif spindle_param == -1:
                        out_file.write("\n" + content[:nc_end].strip())

                    futil.log(f"Merged file {index}: {file_path}", force_console=True)

            out_file.write("\n</nc>")
            if out_file.tell() == 0:
                raise ValueError("Merged file is empty")
            futil.log(
                f"Successfully merged XML files into: {output_file}",
                force_console=True,
            )

        for file_path in file_paths:
            try:
                os.remove(file_path)
                futil.log(f"Removed temporary file: {file_path}")
            except OSError as error:
                futil.log(f"Warning: Could not remove {file_path} - {error}")
        return True
    except Exception as error:  # pylint: disable=broad-exception-caught
        show_message(f"XML merge error: {error}")
        if os.path.exists(output_file):
            os.remove(output_file)
        return False


def process_operations(
    cam,
    operations,
    program_name,
    post_processor,
    output_folder,
    unit,
    post_params,
    show_message,
):
    """Post individual operations to numbered intermediate XML files."""
    output_units = {
        0: adsk.cam.PostOutputUnitOptions.InchesOutput,
        1: adsk.cam.PostOutputUnitOptions.MillimetersOutput,
    }.get(unit)
    if output_units is None:
        raise ValueError(f"Unsupported resolved output unit: {unit}")

    futil.log("===============================", force_console=True)
    futil.log("=== Starting XML generation ===", force_console=True)
    futil.log("===============================", force_console=True)
    futil.log(f"Program name: {program_name}")
    futil.log(f"Output folder: {output_folder}")

    output_folder = normalize_path(output_folder)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    generated_files = []
    operations = list(operations)

    def create_value_input(value, value_type):
        value = value_type(value)
        if value_type is bool:
            return adsk.core.ValueInput.createByBoolean(value)
        return adsk.core.ValueInput.createByReal(float(value))

    param_mapping = {
        "allowHelicalMoves": bool,
        "highFeedMapping": int,
        "minimumChordLength": float,
        "highFeedrate": float,
        "maximumCircularRadius": float,
        "minimumCircularRadius": float,
        "tolerance": float,
    }

    for index, operation in enumerate(operations, 1):
        operation_name = getattr(operation, "name", f"Op_{index}")
        numbered_name = f"{program_name}_{index}"
        xml_path = normalize_path(os.path.join(output_folder, f"{numbered_name}.xml"))
        try:
            post_input = adsk.cam.PostProcessInput.create(
                numbered_name, post_processor, output_folder, output_units
            )
            post_input.isOpenInEditor = False
            post_properties = adsk.core.NamedValues.create()
            for param, param_type in param_mapping.items():
                if param in post_params:
                    value_input = create_value_input(post_params[param], param_type)
                    post_properties.add(param, value_input)
            post_input.postProperties = post_properties

            if not cam.postProcess(operation, post_input):
                raise RuntimeError("CAM post processing returned False")
            if not os.path.exists(xml_path):
                raise FileNotFoundError(f"Output file was not created: {xml_path}")

            generated_files.append(xml_path)
            futil.log(
                f"Successfully processed: {operation_name} -> {xml_path}",
                force_console=True,
            )
        except Exception as error:  # pylint: disable=broad-exception-caught
            error_message = f"Failed to process {operation_name}: {error}"
            futil.log(error_message, force_console=True)
            show_message(error_message, "Processing XML Error")
            return None

    return generated_files
