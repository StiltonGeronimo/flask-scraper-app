from flask import Flask, render_template, request, send_file
import subprocess
import os
import time  # Added for waiting functionality

app = Flask(__name__)

# Dictionary mapping think tanks to their respective scraper scripts
SCRAPER_SCRIPTS = {
    "Atlantic Council": "scraperACwebsite.py",
    "AEI": "scraperAEIwebsite.py",
    "Baker Institute": "scraperBakerwebsite.py",
    "Belfer Center": "scraperBelferwebsite.py",
    "Brookings Institution": "scraperBrookingswebsite.py",
    "Carnegie Endowment": "scraperCEIPwebsite.py",
    "Chicago Council": "scraperChicagowebsite.py",
    "CSIS": "scraperCSISwebsite.py",
    "FDD": "scraperFDDwebsite.py",
    "GMF": "scraperGMFwebsite.py",
    "Heritage Foundation": "scraperHeritagewebsite.py",
    "Hudson Institute": "scraperHudsonwebsite.py",
    "MEI": "scraperMEIwebsite.py",
    "Pew Research Center": "scraperPewwebsite.py",
    "PIIE": "scraperPIIEwebsite.py",
    "Quincy Institute": "scraperQuincywebsite.py",
    "Stimson Center": "scraperStimsonwebsite.py",
    "USIP": "scraperUSIPwebsite.py",
    "Wilson Center": "scraperWilsonwebsite.py",
    "WINEP": "scraperWINEPwebsite.py"
}

OUTPUT_FOLDER = "output_files"

@app.route("/")
def home():
    return render_template("index.html", think_tanks=SCRAPER_SCRIPTS.keys())

@app.route("/run_scraper", methods=["POST"])
def run_scraper():
    think_tank = request.form["think_tank"]
    script_path = SCRAPER_SCRIPTS.get(think_tank)

    if script_path:
        print(f"Running script: {script_path}")  # Debugging output
        subprocess.run(["python3", script_path], check=True)

        # Define the expected Word file path
        word_file = os.path.join(OUTPUT_FOLDER, f"{think_tank}.docx")
        print(f"Checking for file: {word_file}")  # Debugging output

        # Wait up to 10 seconds for the file to appear
        for _ in range(10):
            if os.path.exists(word_file):
                return f"<a href='/download/{think_tank}'>Download {think_tank} report</a>"
            time.sleep(1)  # Wait for 1 second before checking again

        print("Error: Word file not found after waiting.")  # Debugging output
        return "Error: Word file not found.", 500

    return "Invalid request", 400

@app.route("/download/<think_tank>")
def download(think_tank):
    file_path = os.path.join(OUTPUT_FOLDER, f"{think_tank}.docx")
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "File not found", 404

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
