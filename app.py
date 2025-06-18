from flask import Flask, render_template, request, redirect, url_for, flash
import win32com.client
import pythoncom  # Required to initialize COM in Flask
import os

app = Flask(__name__)
app.secret_key = os.urandom(24)

# Path to your simulation model
MODEL_PATH = r"C:\Users\makkapat\Documents\simulation_technmatix\html_start1.spp"

def connect_to_plant_simulation(model_path):
    pythoncom.CoInitialize()  # 🔧 FIX: Required to avoid COM error in Flask thread
    ps = win32com.client.Dispatch("Tecnomatix.PlantSimulation.RemoteControl")
    ps.SetTrustModels(True)
    ps.LoadModel(model_path)
    ps.SetVisible(True)
    return ps

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        object_path = request.form["object_path"].strip()
        attribute_name = request.form["attribute_name"].strip()
        new_value_str = request.form["new_value"].strip()

        return redirect(url_for(
            "update_attribute",
            object_path=object_path,
            attribute_name=attribute_name,
            new_value=new_value_str
        ))

    return render_template("index.html")

@app.route("/update")
def update_attribute():
    object_path = request.args.get("object_path")
    attribute_name = request.args.get("attribute_name")
    new_value_str = request.args.get("new_value")
    full_attr_path = f"{object_path}.{attribute_name}"

    try:
        ps = connect_to_plant_simulation(MODEL_PATH)
        current_val = ps.GetValue(full_attr_path)
        val_type = type(current_val).__name__

        # Convert new value to the same type
        if isinstance(current_val, float):
            new_val = float(new_value_str)
        elif isinstance(current_val, int):
            new_val = int(new_value_str)
        elif isinstance(current_val, str):
            new_val = new_value_str
        elif isinstance(current_val, bool):
            new_val = new_value_str.lower() in ['true', '1', 'yes']
        else:
            return render_template("result.html", error=f"Unsupported type: {val_type}")

        # Set the new value
        ps.SetValue(full_attr_path, new_val)

        # Save the model
        ps.SaveModel(MODEL_PATH)

        # Confirm
        confirmed_val = ps.GetValue(full_attr_path)
        return render_template("result.html",
                               object_path=object_path,
                               attribute_name=attribute_name,
                               current_val=current_val,
                               new_val=new_val,
                               confirmed_val=confirmed_val)

    except Exception as e:
        return render_template("result.html", error=str(e))

# Run the server
if __name__ == "__main__":
    app.run(debug=True)
