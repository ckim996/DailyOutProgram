from flask import Flask, render_template, jsonify, send_file
import src.main as main
import os
from src.shipping.shipping_ops import shipping_label_algo
import traceback

app = Flask(__name__)

@app.route("/")
def index():

    awaiting_count = main.count_awaiting_shipments()

    return render_template("index.html", awaiting_count=awaiting_count)

@app.route("/run/extract", methods=["POST"])
def run_extract():
    try:
        main.progress_status['percent'] = 0
        result = main.extract_todays_shipments()
        return jsonify({
            "status": "success",
            "duration": f"{result.get("duration", 0)} minutes"
            })
    except Exception as e:
        # print to console
        print("Error in run_extract:", e)
        import traceback
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500
    
@app.route("/run/algo_debug", methods=["POST"]) # Changed 'method' to 'methods'
def run_algo_debug():
    try:
        # Simply call the function you already wrote in main.py
        main.run_debug_list_algorithm()
        
        return jsonify({
            "status": "success", 
            "message": "List Algorithm updated successfully from the Daily Sheet!"
        })
    
    except Exception as e:
        # This will catch file-in-use errors or sheet-not-found errors
        return jsonify({
            "status": "error", 
            "message": f"Algorithm failed: {str(e)}"
        }), 500
    
@app.route("/run/shipping_algo", methods=["POST"])
def run_shipping_algo_route():
    try:
        SHEET_NAME = "Decision Log" 
        
        result_pdf_path = shipping_label_algo(SHEET_NAME)
        
        if result_pdf_path and os.path.exists(result_pdf_path):
            # We send the file using a response object to ensure headers are clean
            response = send_file(
                result_pdf_path,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=os.path.basename(result_pdf_path)
            )
            # Disable caching to prevent the browser from trying to open a partially downloaded file
            response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
            response.headers["Pragma"] = "no-cache"
            response.headers["Expires"] = "0"
            return response
        else:
            return jsonify({"error": "Batch failed. Check terminal."}), 400
            
    except Exception as e:
        print("--- CRITICAL ERROR IN SHIPPING ALGO ---")
        print(traceback.format_exc()) 
        return jsonify({"error": str(e)}), 500

@app.route('/download-test-pdf')
def download_test_pdf():
    # This route allows the browser to actually download the file we just created
    pdf_path = os.path.join(os.getcwd(), "test_batch.pdf")
    if os.path.exists(pdf_path):
        return send_file(pdf_path, as_attachment=True)
    return "File not found", 404

@app.route('/progress')
def get_progress():

    data = dict(main.progress_status)

    for key, value in data.items():
        if isinstance(value,set):
            data[key] = list(value)

    return jsonify(data)

if __name__ == "__main__":
    app.run(debug=True)