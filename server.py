from flask import Flask, request, jsonify
from main import rate_card_generator
import threading

app = Flask(__name__)
lock = threading.Lock()
is_generating = False


@app.route('/generate_rate_card', methods=['GET'])
def generate_rate_card():
    global is_generating
    if lock.locked():
        return jsonify({"id": None, "message": "Rate card generation already in progress"}), 429
    with lock:
        is_generating = True
        try:
            rate_card_id = rate_card_generator()
            response_message = "Rate card has been successfully generated" if rate_card_id else "Error occured while generating Rate Card"
            results = {"id": rate_card_id, "message": response_message}
        except Exception as e:
            results = {"id": None, "message": str(e)}
        finally:
            is_generating = False
    return jsonify(results)


app.run(host='0.0.0.0', port=7620)
