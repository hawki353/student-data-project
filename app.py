import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os

# Load Excel files
file1 = "Discovery Applications by Academic Year (07.01.2010 - Present).xlsx"
file2 = "Funding - Student Outside & Competitions.xlsx"

# Read and merge all sheets
df1 = pd.concat(pd.read_excel(file1, sheet_name=None).values(), ignore_index=True)
df2 = pd.concat(pd.read_excel(file2, sheet_name=None).values(), ignore_index=True)

# Merge both datasets
combined_df = pd.merge(df1, df2, how="outer")

# Clean data
combined_df.drop_duplicates(inplace=True)
combined_df.fillna("", inplace=True)

# Save cleaned data
combined_df.to_excel("combined_cleaned_data.xlsx", index=False)

# Set up Flask web app
app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return """
    <h2>Student Data Dashboard</h2>
    <p>Use these endpoints:</p>
    <ul>
        <li><b>Search for a student:</b> /search?query=NAME_OR_ID</li>
        <li><b>Download full dataset:</b> /download</li>
        <li><b>Generate a report:</b> /report?year=YEAR</li>
        <li><b>View data analysis charts:</b> /charts</li>
    </ul>
    """

@app.route("/search", methods=["GET"])
def search_student():
    """Search for a student by name or ID."""
    query = request.args.get("query", "").strip().lower()

    if not query:
        return jsonify({"error": "Please provide a name or ID"}), 400

    results = combined_df[combined_df.apply(lambda row: row.astype(str).str.lower().str.contains(query).any(), axis=1)]

    if results.empty:
        return jsonify({"message": "No results found"}), 404

    return results.to_json(orient="records")

@app.route("/download", methods=["GET"])
def download_data():
    """Download the full cleaned dataset."""
    return send_file("combined_cleaned_data.xlsx", as_attachment=True)

@app.route("/report", methods=["GET"])
def generate_report():
    """Generate reports for a specific year."""
    year = request.args.get("year", "").strip()

    if not year.isdigit():
        return jsonify({"error": "Provide a valid year"}), 400

    report_df = combined_df[combined_df.apply(lambda row: row.astype(str).str.contains(year).any(), axis=1)]

    if report_df.empty:
        return jsonify({"message": "No data for this year"}), 404

    report_file = f"report_{year}.xlsx"
    report_df.to_excel(report_file, index=False)

    return send_file(report_file, as_attachment=True)

@app.route("/charts", methods=["GET"])
def generate_charts():
    """Generate and save data charts for visualization."""
    plt.figure(figsize=(10, 6))
    
    # Example: Plot student applications by year
    if "Year" in combined_df.columns:
        sns.countplot(data=combined_df, x="Year")
        plt.xticks(rotation=45)
        plt.title("Student Applications by Year")
        plt.savefig("student_applications_chart.png")
    
    return send_file("student_applications_chart.png", as_attachment=True)

if __name__ == "__main__":
    app.run(port=5000)




from flask import Flask

app = Flask(__name__)

@app.route('/')
def home():
    return "Hello, this is your Flask app running!"

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)  # Render uses a specific port

