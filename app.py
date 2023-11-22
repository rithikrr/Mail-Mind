from flask import Flask, render_template, request
import subprocess

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('api_hp.html')

@app.route('/run_script', methods=['POST'])
def run_script():
    try:
        # Run the Python script using subprocess
        subprocess.run(['python', 'Credit Risk API.py'], check=True)
        return "Script executed successfully!"
    except subprocess.CalledProcessError:
        return "Error executing the script."

if __name__ == '__main__':
    app.run(debug=True)
