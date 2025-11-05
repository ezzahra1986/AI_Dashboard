import pandas as pd
import numpy as np
from flask import Flask, jsonify
from flask_cors import CORS
import os
import json

app = Flask(__name__)
CORS(app)


def load_data():
    excel_path = "data/projects.xlsx"
    
    if not os.path.exists(excel_path):
        print(f"❌ ERROR: File not found at {excel_path}")
        return pd.DataFrame()
    
    try:
        df = pd.read_excel(excel_path)
        print(f"✅ Loaded {len(df)} rows from Excel")
        
        # IMPORTANT : Remplacer TOUS les NaN/NaT/None par des valeurs valides
        # Pour les colonnes numériques : remplacer par 0
        numeric_cols = ["Total_Cost", "Total_Saving", "ROI", "Payback", "PxTxD"]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Pour les colonnes datetime : remplacer par string vide
        datetime_cols = ["Start_Date", "End_Date"]
        for col in datetime_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else "")
        
        # Pour les colonnes string : remplacer par string vide
        string_cols = ["Project", "Tech_Maturity", "Data_Maturity", "Milestone", "PxTxD"]
        for col in string_cols:
            if col in df.columns:
                df[col] = df[col].fillna("")
        
        # CRITIQUE : Remplacer TOUT NaN restant par None (devient null en JSON)
        df = df.where(pd.notna(df), None)
        
        return df
        
    except Exception as e:
        print(f"❌ ERROR loading Excel: {e}")
        return pd.DataFrame()


@app.route("/api/projects")
def get_projects():
    df = load_data()
    return jsonify(df.to_dict(orient="records"))


@app.route("/api/roi")
def get_roi():
    df = load_data()

    if "ROI" in df.columns and "Project" in df.columns:
        #  Clean text: remove % signs, fix commas
        df["ROI"] = (
            df["ROI"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )

        # Convert to numeric
        df["ROI"] = pd.to_numeric(df["ROI"], errors="coerce")

        # If values look like fractions (e.g. 0.5 = 50%), scale them up
        # Check median — if most values are between -1 and 1, they’re likely decimals
        median_val = df["ROI"].median()
        if -1 < median_val < 1:
            df["ROI"] = df["ROI"] * 100

        # Drop invalid entries
        df = df.dropna(subset=["ROI"])

        #  Sort from lowest to highest
        df = df.sort_values(by="ROI", ascending=True)

        #  Round
        df["ROI"] = df["ROI"].round(2)

        return jsonify(df[["Project", "ROI"]].to_dict(orient="records"))

    return jsonify([])


# 🟢 Cost & saving per project
@app.route("/api/cost-saving")
def get_cost_saving():
    print("📊 /api/cost-saving called")
    df = load_data()
    
    cols = ["Project", "Total_Cost", "Total_Saving"]
    missing = [c for c in cols if c not in df.columns]
    
    if missing:
        print(f"⚠️ Missing columns: {missing}")
        return jsonify([])
    
    result_df = df[cols].copy()
    result_df["Total_Cost"] = pd.to_numeric(result_df["Total_Cost"], errors='coerce').fillna(0)
    result_df["Total_Saving"] = pd.to_numeric(result_df["Total_Saving"], errors='coerce').fillna(0)
    result_df["Project"] = result_df["Project"].fillna("")
    
    result_df = result_df.replace([np.inf, -np.inf], 0)
    result_df = result_df.where(pd.notna(result_df), None)
    
    result = result_df.to_dict(orient="records")
    print(f"✅ Returning {len(result)} cost-saving records")
    
    return jsonify(result)


@app.route("/api/kpis")
def get_kpis():
    df = load_data()
    total_cost = float(df["Total_Cost"].sum()) if "Total_Cost" in df.columns else 0.0
    total_saving = float(df["Total_Saving"].sum()) if "Total_Saving" in df.columns else 0.0
    avg_roi = float(round(df["ROI"].mean(), 2)) if "ROI" in df.columns else 0.0
    return jsonify({
        "totalCost": total_cost,
        "totalSaving": total_saving,
        "avgROI": avg_roi
    })


@app.route("/api/timeline")
def get_timeline():
    df = load_data()
    if "Start_Date" in df.columns and "End_Date" in df.columns:
        return jsonify(df[["Project", "Start_Date", "End_Date"]].to_dict(orient="records"))
    return jsonify([])


@app.route("/api/maturity")
def get_maturity():
    df = load_data()
    if "Tech_Maturity" in df.columns:
        maturity_counts = df["Tech_Maturity"].value_counts().reset_index()
        maturity_counts.columns = ["name", "value"]
        return jsonify(maturity_counts.to_dict(orient="records"))
    return jsonify([])


@app.route("/api/ptd")
def get_ptd():
    df = load_data()
    required_cols = ["Project", "Payback", "Tech_Maturity", "Data_Maturity"]

    if set(required_cols).issubset(df.columns):
        result = df[required_cols].to_dict(orient="records")
        return jsonify(result)

    return jsonify([])


@app.route("/api/milestones")
def get_milestone_distribution():
    df = load_data()  # your Excel data

    if "Milestone" in df.columns and "Project" in df.columns:
        # Group by milestone to get project lists
        milestone_groups = (
            df.groupby("Milestone")["Project"]
            .apply(list)
            .reset_index(name="projects")
        )

        # Count projects per milestone
        milestone_counts = (
            df["Milestone"].value_counts().reset_index()
        )
        milestone_counts.columns = ["name", "count"]

        # Merge with project lists
        merged = milestone_counts.merge(
            milestone_groups, left_on="name", right_on="Milestone", how="left"
        ).drop(columns=["Milestone"])

        # Calculate total projects and percentage
        total = merged["count"].sum()
        merged["value"] = (merged["count"] / total * 100).round(2)

        return jsonify(merged.to_dict(orient="records"))

    return jsonify([])

@app.route("/api/px_tx_d")
def px_tx_d():
    df = load_data()
    if "Project" in df.columns and "PxTxD" in df.columns:
        # S'assurer que PxTxD est numérique
        df["PxTxD"] = pd.to_numeric(df["PxTxD"], errors="coerce")
        df = df.dropna(subset=["PxTxD"])
        # Retourner uniquement Project et PxTxD
        return jsonify(df[["Project", "PxTxD"]].to_dict(orient="records"))
    return jsonify([])

#🟢 Geometry 4.0 specific endpoints

@app.route("/api/future-savings/geometry4")
def get_geometry4_savings():
    print("📊 /api/future-savings/geometry4 called")
    df = load_data()

    project_name = "Geometry 4.0"
    years = ["2026", "2027", "2028"]
    cols = ["Project"] + [f"{year}_Saving" for year in years] + ["Total_Cost"]  # Ajout de Total_Cost

    # Vérifier les colonnes
    missing = [c for c in cols if c not in df.columns]
    if missing:
        print(f"⚠️ Missing columns: {missing}")
        return jsonify({"error": f"Missing columns: {missing}"}), 400

    # Filtrer le projet Geometry 4.0
    project_df = df[df["Project"].astype(str).str.strip().str.lower() == project_name.lower()]

    if project_df.empty:
        print(f"⚠️ No data found for project '{project_name}'")
        return jsonify({"error": f"No data found for project '{project_name}'"}), 404

    # Nettoyer et convertir les colonnes numériques
    for year in years:
        col = f"{year}_Saving"
        project_df[col] = pd.to_numeric(project_df[col], errors="coerce").fillna(0)

    # Total_Cost
    project_df["Total_Cost"] = pd.to_numeric(project_df["Total_Cost"], errors="coerce").fillna(0)

    project_df["Project"] = project_df["Project"].fillna("")
    project_df = project_df.replace([np.inf, -np.inf], 0)
    project_df = project_df.where(pd.notna(project_df), None)

    # Convertir en dictionnaire (un seul enregistrement)
    result = project_df[cols].to_dict(orient="records")[0]

    print(f"✅ Returning savings and cost for '{project_name}': {result}")
    return jsonify(result)

@app.route("/api/roi/geometry4")
def get_geometry4_roi():
    print("📊 /api/roi/geometry4 called")
    df = load_data()
    project_name = "Geometry 4.0"

    if "ROI" in df.columns and "Project" in df.columns:
        # Clean ROI text
        df["ROI"] = (
            df["ROI"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )

        # Convert to numeric
        df["ROI"] = pd.to_numeric(df["ROI"], errors="coerce")

        # Detect fractional ROIs (like 0.5 = 50%)
        median_val = df["ROI"].median()
        if median_val and -1 < median_val < 1:
            df["ROI"] = df["ROI"] * 100

        # Drop invalid entries
        df = df.dropna(subset=["ROI"])

        # Filter only Geometry 4.0
        project_df = df[df["Project"].astype(str).str.strip().str.lower() == project_name.lower()]

        if project_df.empty:
            print(f"⚠️ No ROI data found for '{project_name}'")
            return jsonify({"error": f"No ROI data found for '{project_name}'"}), 404

        # Round for clarity
        project_df["ROI"] = project_df["ROI"].round(2)

        result = project_df[["Project", "ROI"]].to_dict(orient="records")[0]
        print(f"✅ Returning ROI for '{project_name}': {result}")
        return jsonify(result)

    print("⚠️ Missing required columns in dataset")
    return jsonify({"error": "Missing columns 'Project' or 'ROI'"}), 400


@app.route("/api/projects/geometry4")
def get_geometry4_project():
    print("📊 /api/projects/geometry4 called")
    df = load_data()

    project_name = "Geometry 4.0"

    # Filter only the desired project
    project_df = df[df["Project"].astype(str).str.strip().str.lower() == project_name.lower()]

    if project_df.empty:
        print(f"⚠️ No data found for project '{project_name}'")
        return jsonify({"error": f"No data found for project '{project_name}'"}), 404

    # Return the single project as JSON (one object, not array)
    result = project_df.to_dict(orient="records")[0]
    print(f"✅ Returning project info for '{project_name}'")
    return jsonify(result)


@app.route("/api/kpis/<project_name>")
def get_kpis_for_project(project_name):
    print(f"📊 /api/kpis/{project_name} called")
    df = load_data()
    
    # Filter by project name (case-insensitive)
    project_df = df[
        df["Project"].astype(str).str.strip().str.lower() == project_name.lower()
    ]

    if project_df.empty:
        return jsonify({"error": f"No data found for project '{project_name}'"}), 404

    # Extract values from first matching row
    row = project_df.iloc[0]

    total_cost = float(row["Total_Cost"]) if "Total_Cost" in row and pd.notna(row["Total_Cost"]) else 0.0
    total_saving = float(row["Total_Saving"]) if "Total_Saving" in row and pd.notna(row["Total_Saving"]) else 0.0
    roi = float(row["ROI"]) if "ROI" in row and pd.notna(row["ROI"]) else 0.0
    
    if roi < 10:  # handle fractional ROI like 1.45 → 145
        roi *= 100

    print(f"✅ Returning KPIs for '{project_name}': Cost={total_cost}, Saving={total_saving}, ROI={roi}")
    return jsonify({
        "project": project_name,
        "totalCost": total_cost,
        "totalSaving": total_saving,
        "ROI": round(roi, 2)
    })

def get_geometry4_ptd():
    print("📊 /api/ptd/geometry4 called")
    df = load_data()

    project_df = df[df["Project"].astype(str).str.strip().str.lower() == "geometry 4.0"]
    if project_df.empty:
        return jsonify({"error": "No data found for Geometry 4.0"}), 404

    ptd_cols = ["Payback", "Tech_Maturity", "Data_Maturity"]
    row = project_df.iloc[0]

    result = {c: float(row[c]) if c in row and not pd.isna(row[c]) else 0 for c in ptd_cols}
    print(f"✅ Returning PTD for 'Geometry 4.0': {result}")
    return jsonify(result)

@app.route("/api/timeline/geometry4")
def get_geometry4_timeline():
    print("📊 /api/timeline/geometry4 called")
    df = load_data()

    # Check for required columns
    required_columns = ["Project", "Start_Date", "End_Date", 
                        "Ideation", "Framing & scoping", 
                        "Development & industrialization", 
                        "Roll-out & deployment"]
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        return jsonify({"error": f"Missing columns: {', '.join(missing_cols)}"}), 400

    # Filter Geometry 4.0
    project_df = df[df["Project"].astype(str).str.strip().str.lower() == "geometry 4.0"]
    if project_df.empty:
        return jsonify({"error": "No data found for Geometry 4.0"}), 404

    row = project_df.iloc[0]

    # Define phases
    phases = ["Ideation", "Framing & scoping", "Development & industrialization", "Roll-out & deployment"]

    timeline = []
    for i, phase in enumerate(phases):
        start_date = row[phase]
        if pd.isna(start_date) or start_date == "":
            continue

        try:
            start_date = pd.to_datetime(start_date).strftime("%Y-%m-%d")
        except Exception:
            start_date = str(start_date)

        # End date = next phase's start date, or overall End_Date if last phase
        if i + 1 < len(phases):
            next_phase = phases[i + 1]
            next_date = row[next_phase]
            if not pd.isna(next_date) and next_date != "":
                try:
                    end_date = pd.to_datetime(next_date).strftime("%Y-%m-%d")
                except Exception:
                    end_date = str(next_date)
            else:
                end_date = None
        else:
            # Last phase uses project End_Date
            end_date = row.get("End_Date")
            if not pd.isna(end_date) and end_date != "":
                try:
                    end_date = pd.to_datetime(end_date).strftime("%Y-%m-%d")
                except Exception:
                    end_date = str(end_date)
            else:
                end_date = None

        timeline.append({
            "phase": phase,
            "start": start_date,
            "end": end_date
        })

    print(f"✅ Returning timeline with {len(timeline)} phases")
    return jsonify(timeline)

if __name__ == "__main__":
    from waitress import serve
    print("🚀 Starting Geometry Flask backend on Render...")
    serve(app, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

