import sys
import os
import json
import subprocess
import traceback
import re

def slice_with_prusa(step_file_path):
    """
    Slices a STEP file using PrusaSlicer CLI twice (with and without supports)
    to accurately calculate object and support weights.

    Args:
        step_file_path (str): Path to the STEP file

    Returns:
        dict: Dictionary containing the slicing metrics or None if slicing failed
    """
    try:
        # Validate file exists
        if not os.path.exists(step_file_path):
            print(f"ERROR: File not found: {step_file_path}")
            return None

        # Get file info
        file_name = os.path.basename(step_file_path)
        base_name = os.path.splitext(file_name)[0]
        output_dir = os.path.dirname(step_file_path)

        # Create output directory for slicer stats
        stats_dir = os.path.join(output_dir, "Slicer_Stats")
        os.makedirs(stats_dir, exist_ok=True)

        stats_file_path = os.path.join(stats_dir, f"{base_name}_stats.json")
        gcode_path_with_supports = os.path.join(stats_dir, f"{base_name}_with_supports.gcode")
        gcode_path_without_supports = os.path.join(stats_dir, f"{base_name}_without_supports.gcode")

        # Read config settings
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
        if not os.path.exists(config_path):
            print(f"ERROR: Config file not found: {config_path}")
            return None

        with open(config_path, "r") as f:
            config_content = f.read()

        # Extract settings from config
        # Consider making filament_cost_per_kg an argument or env variable
        filament_cost_per_kg = extract_config_value(config_content, "filament_cost", 2500.0)
        supports_enabled = extract_config_value(config_content, "supports_enabled", True)
        # Consider making slicer_settings an argument or derive dynamically
        slicer_settings = extract_config_value(config_content, "slicer_settings",
                                             "0.2mm layer, 20% infill, supports=auto", as_string=True)

        # Create temporary config files in a dedicated temp dir
        # Using system temp might be more robust in some environments
        temp_dir = os.path.join(output_dir, "temp_configs")
        os.makedirs(temp_dir, exist_ok=True)

        config_with_supports = os.path.join(temp_dir, "config_with_supports.ini")
        config_without_supports = os.path.join(temp_dir, "config_without_supports.ini")

        # Create config with supports (Ensure setting exists or add it)
        with open(config_with_supports, "w") as f:
            modified_content = re.sub(r'^support_material\s*=\s*\d', 'support_material = 1', config_content, flags=re.MULTILINE)
            if 'support_material = 1' not in modified_content:
                 # Check if the line was modified; if not, append it. Handles cases where the key exists but has a different value, or doesn't exist.
                 if not re.search(r'^support_material\s*=', modified_content, flags=re.MULTILINE):
                    modified_content += '\nsupport_material = 1'
            f.write(modified_content)

        # Create config without supports (Ensure setting exists or add it)
        with open(config_without_supports, "w") as f:
            modified_content = re.sub(r'^support_material\s*=\s*\d', 'support_material = 0', config_content, flags=re.MULTILINE)
            if 'support_material = 0' not in modified_content:
                if not re.search(r'^support_material\s*=', modified_content, flags=re.MULTILINE):
                    modified_content += '\nsupport_material = 0'
            f.write(modified_content)

        # Initialize metrics dictionary
        metrics = {
            "part_name": base_name,
            "dimensions_mm": None,
            "object_weight_g": 0.0, # Changed default to float
            "supports_weight_g": 0.0,# Changed default to float
            "total_weight_g": 0.0, # Changed default to float
            "print_time": "Unknown",
            "price_egp": 0.0, # Initialize as float
            "print_settings": slicer_settings,
            "object_weight_kg": 0.0,
            "supports_weight_kg": 0.0,
            "total_weight_kg": 0.0
        }

        # --- SLICING ---
        prusa_slicer_cmd = "prusa-slicer-console.exe" # Consider making this configurable

        # First slice: with supports if enabled
        if supports_enabled:
            command = [
                prusa_slicer_cmd, "--export-gcode", "--output", gcode_path_with_supports,
                "--load", config_with_supports, "--info", step_file_path
            ]
            result = subprocess.run(command, capture_output=True, text=True, check=False) # Use check=False and handle error below

            if result.returncode != 0:
                print(f"ERROR: PrusaSlicer (with supports) failed with code {result.returncode}")
                print(f"STDERR: {result.stderr}")
                # Cleanup temp files before returning
                try:
                   os.remove(config_with_supports)
                   os.remove(config_without_supports)
                   os.rmdir(temp_dir)
                except OSError:
                    pass # Ignore cleanup errors on failure
                return None

            with_supports_metrics = extract_metrics(result.stdout + "\n" + result.stderr, gcode_path_with_supports)
            if with_supports_metrics:
                metrics["total_weight_g"] = with_supports_metrics.get("weight_g", 0.0) # Use 0.0
                metrics["print_time"] = with_supports_metrics.get("print_time", "Unknown")
                metrics["dimensions_mm"] = with_supports_metrics.get("dimensions_mm")
            else:
                print("WARNING: Could not extract metrics from 'with supports' slice.")
                metrics["total_weight_g"] = 0.0 # Ensure it's zero if extraction failed


        # Second slice: always without supports
        command = [
            prusa_slicer_cmd, "--export-gcode", "--output", gcode_path_without_supports,
            "--load", config_without_supports, "--info", step_file_path
        ]
        result = subprocess.run(command, capture_output=True, text=True, check=False)

        if result.returncode != 0:
            print(f"ERROR: PrusaSlicer (without supports) failed with code {result.returncode}")
            print(f"STDERR: {result.stderr}")
            # Cleanup temp files before returning
            try:
               os.remove(config_with_supports)
               os.remove(config_without_supports)
               os.rmdir(temp_dir)
            except OSError:
                pass # Ignore cleanup errors on failure
            return None

        without_supports_metrics = extract_metrics(result.stdout + "\n" + result.stderr, gcode_path_without_supports)
        if without_supports_metrics:
            metrics["object_weight_g"] = without_supports_metrics.get("weight_g", 0.0) # Use 0.0
            if not metrics["dimensions_mm"]: # Use dimensions from this slice if not found earlier
                metrics["dimensions_mm"] = without_supports_metrics.get("dimensions_mm")
        else:
             print("WARNING: Could not extract metrics from 'without supports' slice.")
             metrics["object_weight_g"] = 0.0 # Ensure it's zero


        # --- CALCULATIONS ---
        # Calculate support weight by subtraction
        if supports_enabled:
             # Ensure subtraction only happens if both weights are valid numbers >= 0
             if isinstance(metrics["total_weight_g"], (int, float)) and isinstance(metrics["object_weight_g"], (int, float)) and metrics["total_weight_g"] >= 0 and metrics["object_weight_g"] >= 0:
                 metrics["supports_weight_g"] = max(0.0, metrics["total_weight_g"] - metrics["object_weight_g"]) # Use 0.0
             else:
                 metrics["supports_weight_g"] = 0.0 # Default to 0 if weights are invalid
                 print("WARNING: Could not calculate support weight due to invalid total/object weight.")
        else:
            # If supports aren't enabled, ensure total weight equals object weight
            metrics["total_weight_g"] = metrics["object_weight_g"]
            metrics["supports_weight_g"] = 0.0

        # Ensure weights are numbers before converting/calculating price
        total_weight_g = round(metrics["total_weight_g"] if isinstance(metrics["total_weight_g"], (int, float)) else 0.0, 1)
        object_weight_g = round(metrics["object_weight_g"] if isinstance(metrics["object_weight_g"], (int, float)) else 0.0, 1)
        supports_weight_g = round(metrics["supports_weight_g"] if isinstance(metrics["supports_weight_g"], (int, float)) else 0.0, 1)


        # Convert to kg and calculate price
        metrics["total_weight_kg"] = total_weight_g / 1000.0
        metrics["object_weight_kg"] = object_weight_g / 1000.0
        metrics["supports_weight_kg"] = supports_weight_g / 1000.0

        if isinstance(filament_cost_per_kg, (int, float)) and filament_cost_per_kg > 0:
             # Calculate price based on total weight to 1 decimal place
             metrics["price_egp"] = round(metrics["total_weight_kg"] * filament_cost_per_kg, 1)
        else:
             metrics["price_egp"] = 0.0
             print(f"WARNING: Invalid filament cost per kg ({filament_cost_per_kg}). Price set to 0.")


        # Write metrics to JSON file
        try:
            with open(stats_file_path, 'w') as f:
                json.dump(metrics, f, indent=4)
            # Optional: Keep this print if you want the metrics output in logs
            # print(f"Final metrics: {json.dumps(metrics, indent=2)}")
        except IOError as e:
             print(f"ERROR: Could not write metrics to {stats_file_path}: {e}")
             # Decide if this is a fatal error or not


        # Clean up temporary config files and dir
        try:
            os.remove(config_with_supports)
            os.remove(config_without_supports)
            os.rmdir(temp_dir)
        except OSError as e:
            # Warning instead of error, as slicing might have succeeded
             print(f"Warning: Could not clean up temporary config files in {temp_dir}: {e}")

        return metrics

    except Exception as e:
        print(f"ERROR in slice_with_prusa: {str(e)}")
        print("Detailed traceback:")
        traceback.print_exc()
        # Attempt cleanup even on general exceptions
        try:
            if 'temp_dir' in locals() and os.path.exists(temp_dir):
                 if 'config_with_supports' in locals() and os.path.exists(config_with_supports): os.remove(config_with_supports)
                 if 'config_without_supports' in locals() and os.path.exists(config_without_supports): os.remove(config_without_supports)
                 os.rmdir(temp_dir)
        except OSError:
             pass # Ignore cleanup errors during exception handling
        return None


def extract_config_value(config_content, key, default_value, as_string=False):
    """
    Extract a value from the config content using regex. Handles comments.
    Looks for 'key = value' at the beginning of a line.
    """
    pattern = rf"^\s*{key}\s*=\s*([^#\n]+)" # Match key at start of line
    match = re.search(pattern, config_content, re.IGNORECASE | re.MULTILINE)

    if not match:
        return default_value

    value = match.group(1).strip()

    if as_string:
        # Remove potential surrounding quotes if any
        if (value.startswith('"') and value.endswith('"')) or \
           (value.startswith("'") and value.endswith("'")):
            return value[1:-1]
        return value

    # Try boolean conversion
    if value.lower() in ('true', 'yes', '1', 'on'):
        return True
    elif value.lower() in ('false', 'no', '0', 'off'):
        return False

    # Try float conversion
    try:
        return float(value)
    except ValueError:
        # Return as string if float conversion fails
        return value


def extract_metrics(combined_output, gcode_path):
    """
    Extract metrics from slicer output and gcode file comments.
    Focuses on reliable extraction from G-code comments first.
    Removes volume-based weight calculation as requested.
    """
    metrics = {
        "dimensions_mm": None,
        "weight_g": 0.0, # Default to float
        "print_time": "Unknown"
    }
    weight_match = None # Initialize weight_match to check later

    # --- G-code Parsing (More reliable for weight/time) ---
    if os.path.exists(gcode_path):
        try:
            with open(gcode_path, 'r') as gcode_file:
                gcode_content = gcode_file.read()

                # Preferred: Extract weight directly in grams if available
                weight_match = re.search(r"(?:;|^)\s*total filament used\s*\[g\]\s*=\s*(\d+\.?\d*)",
                                           gcode_content, re.IGNORECASE | re.MULTILINE)
                metrics["weight_g"] = float(weight_match.group(1))

                # Extract estimated print time from G-code
                # Format: HHh MMm SSs or similar variations
                time_match_gcode = re.search(r"(?:;|^)\s*estimated printing time.*=\s*(?:(\d+)h\s*)?(?:(\d+)m\s*)?(?:(\d+)s\s*)?$",
                                            gcode_content, re.IGNORECASE | re.MULTILINE)
                if time_match_gcode:
                    h = int(time_match_gcode.group(1) or 0)
                    m = int(time_match_gcode.group(2) or 0)
                    s = int(time_match_gcode.group(3) or 0)
                    if h > 0 or m > 0 or s > 0:
                         time_parts = []
                         if h > 0: time_parts.append(f"{h}h")
                         if m > 0: time_parts.append(f"{m}m")
                         if s > 0: time_parts.append(f"{s}s") # Optionally include seconds
                         metrics["print_time"] = " ".join(time_parts)

        except Exception as gcode_error:
            print(f"Warning: Error parsing G-code file {gcode_path}: {str(gcode_error)}")

    # --- Slicer Output Parsing (Mainly for dimensions) ---
    # Use slicer output primarily for dimensions as G-code doesn't usually have it.
    size_pattern = r"size\s*\(mm\):\s*(\d+\.?\d*)\s*x\s*(\d+\.?\d*)\s*x\s*(\d+\.?\d*)" # More specific pattern
    size_match = re.search(size_pattern, combined_output, re.IGNORECASE | re.MULTILINE)
    if size_match:
        x, y, z = map(float, [size_match.group(1), size_match.group(2), size_match.group(3)])
        metrics["dimensions_mm"] = f"{x:.2f} x {y:.2f} x {z:.2f}"

    # Fallback for time if not found in G-code
    if metrics["print_time"] == "Unknown":
        time_pattern_output = r"estimated printing time\D*(?:(\d+)h\s*)?(?:(\d+)m\s*)?(?:(\d+)s\s*)?" # Similar pattern for output
        time_match_output = re.search(time_pattern_output, combined_output, re.IGNORECASE)
        if time_match_output:
             h = int(time_match_output.group(1) or 0)
             m = int(time_match_output.group(2) or 0)
             s = int(time_match_output.group(3) or 0) # Include seconds if present
             if h > 0 or m > 0 or s > 0:
                 time_parts = []
                 if h > 0: time_parts.append(f"{h}h")
                 if m > 0: time_parts.append(f"{m}m")
                 if s > 0: time_parts.append(f"{s}s") # Optionally include seconds
                 metrics["print_time"] = " ".join(time_parts)

    # Final check for zero weight if slicing supposedly succeeded
    # Note: This warning might trigger more often now if the g-code comment is missing
    if metrics["weight_g"] == 0.0 and weight_match is None and os.path.exists(gcode_path): # Only warn if weight is 0 *because* the comment was missing and gcode existed
         print(f"Warning: Extracted weight is zero for {gcode_path} because '; total filament used [g]' comment was not found in the G-code file.")

    return metrics


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python slice_with_prusa.py <step_file_path>")
        sys.exit(1)

    step_file_path_arg = os.path.abspath(sys.argv[1])
    if not os.path.exists(step_file_path_arg):
        print(f"ERROR: File not found: {step_file_path_arg}")
        sys.exit(1)

    print(f"Processing: {step_file_path_arg}")
    result_metrics = slice_with_prusa(step_file_path_arg)

    if result_metrics is None:
        print("--- Slicing process failed ---")
        sys.exit(1)
    else:
        print("--- Slicing process completed ---")
        print("Metrics Summary:")
        print(json.dumps(result_metrics, indent=2))
        sys.exit(0)