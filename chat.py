from flask import render_template_string, Flask, request, jsonify, render_template

app = Flask(__name__)


@app.route('/ai', methods=['GET'])
def chat():

    return render_template('chat.html')  # 确保 'index.html' 在


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)