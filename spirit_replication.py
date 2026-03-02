import sys
import os
import clr
import gc
import logging
from datetime import datetime

# Setup basic logging to ensure we capture all successes and failures
log_filename = f"spirit_replication_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def configure_tekla_pythonnet():
    """
    Configures PythonNet to load the correct .dll files from the Tekla 2025 installation.
    """
    # Standard installation path for Tekla Structures 2025
    tekla_bin_path = r"C:\Program Files\Tekla Structures\2025.0\bin"
    if os.path.exists(tekla_bin_path) and tekla_bin_path not in sys.path:
        sys.path.append(tekla_bin_path)

    try:
        clr.AddReference("Tekla.Structures")
        clr.AddReference("Tekla.Structures.Model")
        clr.AddReference("Tekla.Structures.Drawing")
    except Exception as e:
        logging.error(f"Failed to load Tekla libraries. Please ensure Tekla 2025 is installed. Error: {e}")
        print("Failed to load Tekla libraries. Check log for details.")
        sys.exit(1)

def main():
    configure_tekla_pythonnet()

    from Tekla.Structures.Model import Model
    from Tekla.Structures.Drawing import DrawingHandler

    # Initialize Model and DrawingHandler
    model = Model()
    if not model.GetConnectionStatus():
        logging.error("Tekla Structures is not running or model is not opened.")
        print("Tekla Structures is not running.")
        sys.exit(1)

    dh = DrawingHandler()
    if not dh.GetConnectionStatus():
        logging.error("Tekla Drawing Handler is not connected.")
        print("Tekla Drawing Handler is not connected.")
        sys.exit(1)

    # 1. Reference Selection: Masterpiece (Spirit Source)
    drawing_selector = dh.GetDrawingSelector()
    selected_drawings = drawing_selector.GetSelected()

    if selected_drawings.GetSize() == 0:
        msg = "No drawing selected in the Drawing List. Please select a Masterpiece drawing."
        logging.error(msg)
        print(msg)
        sys.exit(1)

    selected_drawings.MoveNext()
    master_drawing = selected_drawings.Current

    if master_drawing is None:
        msg = "Failed to retrieve the selected master drawing."
        logging.error(msg)
        print(msg)
        sys.exit(1)

    logging.info(f"Master drawing selected successfully: {master_drawing.Name} / {master_drawing.Title1}")

    # 2. Target Selection: Assemblies in 3D Model
    selected_objects = model.GetSelectedObjects()
    if selected_objects.GetSize() == 0:
        msg = "No assemblies selected in the 3D Model. Please select Target assemblies."
        logging.error(msg)
        print(msg)
        sys.exit(1)

    # Convert selection to a list for batch processing
    assemblies_to_clone = []
    while selected_objects.MoveNext():
        assemblies_to_clone.append(selected_objects.Current)

    logging.info(f"Found {len(assemblies_to_clone)} selected objects in the model.")

    # 3. Cloning Process
    success_count = 0
    failure_count = 0

    for i, assembly in enumerate(assemblies_to_clone, start=1):
        try:
            # CloneDrawing using the masterpiece drawing as the source
            # This preserves View Attributes, OLS, and Dimensioning Styles
            cloned_drawing = dh.CloneDrawing(master_drawing, assembly)

            if cloned_drawing is not None:
                # 4. Post-Processing: Trigger 'View Arrangement' update to avoid dimension overlaps
                cloned_drawing.PlaceViews()

                # Commit changes for this drawing
                cloned_drawing.Modify()

                logging.info(f"Successfully cloned drawing for Assembly ID: {assembly.Identifier.ID()}")
                success_count += 1
            else:
                logging.warning(f"Failed to clone drawing for Assembly ID: {assembly.Identifier.ID()}")
                failure_count += 1

        except Exception as e:
            logging.error(f"Error during cloning for Assembly ID {assembly.Identifier.ID()}: {e}")
            failure_count += 1

        # Memory management: Force garbage collection periodically
        # This is critical to ensure the script does not crash when processing 2000+ tons of steel
        if i % 50 == 0:
            gc.collect()

    print(f"Cloning completed. Success: {success_count}, Failed: {failure_count}.")
    logging.info(f"Cloning batch completed. Total processed: {len(assemblies_to_clone)}. Success: {success_count}, Failed: {failure_count}.")

if __name__ == "__main__":
    main()
