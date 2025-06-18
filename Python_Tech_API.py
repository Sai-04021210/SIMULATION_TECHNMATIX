import win32com.client

# --- CONFIGURABLE SECTION ---
model_path = r"C:\Users\makkapat\Documents\simulation_technmatix\html_start1.spp"
object_path = ".Models.Model.Station1"     #  Change to Station1, Drain, etc.
attribute_name = "ProcTime"                  #  Attribute you want to modify
# ----------------------------

# Connect to Plant Simulation
ps = win32com.client.Dispatch("Tecnomatix.PlantSimulation.RemoteControl")
ps.SetTrustModels(True)
ps.LoadModel(model_path)
ps.SetVisible(True)
print(" Model loaded.")

# Build full path to attribute
attr_path = f"{object_path}.{attribute_name}"

# Read current value
try:
    current_val = ps.GetValue(attr_path)
    val_type = type(current_val).__name__
    print(f" {attribute_name} on {object_path} = {current_val} ({val_type})")
except Exception as e:
    print(f" Failed to read attribute: {e}")
    exit()

# Generate new value
if isinstance(current_val, float):
    new_val = round(current_val + 5.5, 2)
elif isinstance(current_val, int):
    new_val = current_val + 3
elif isinstance(current_val, str):
    new_val = current_val + "_updated"
elif isinstance(current_val, bool):
    new_val = not current_val
else:
    raise TypeError(f"Unsupported data type: {val_type}")

# Set the new value
try:
    ps.SetValue(attr_path, new_val)
    print(f"  Updated {attr_path} → {new_val}")
except Exception as e:
    print(f" Failed to update value: {e}")
    exit()

# Save model
try:
    ps.SaveModel(model_path)
    print(" Model saved.")
except Exception as e:
    print(f" Save failed: {e}")

# Confirm update
try:
    confirmed_val = ps.GetValue(attr_path)
    print(f" Confirmed {attribute_name} = {confirmed_val}")
except Exception as e:
    print(f" Could not confirm: {e}")
