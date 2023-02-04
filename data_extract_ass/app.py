import io
from flask import Flask, request, jsonify, make_response
from flask import Response
from extract_data_xml import extract_data
app = Flask(__name__)

@app.route('/get_data', methods=['POST'])
def extract():
    try:
        file_path = request.json['file_path']
        all_output, df = extract_data(file_path)
        buffer = io.BytesIO()
        df.to_excel(buffer)
        headers = {
            'Content-Disposition': 'attachment; filename=output.xlsx',
            'Content-type': 'application/vnd.ms-excel'
        }
        return Response(buffer.getvalue(), mimetype='application/vnd.ms-excel', headers=headers)

    except Exception as e:
        print(e)
        return send_response({
            "status": False,
            'message': "Something seems to have failed. Please try again.",
        }, 500)


def send_response(data, code):
    response = make_response(
        jsonify(
            data
        ),
        code,
    )
    response.headers["Content-Type"] = "application/json"
    return response

@app.route('/')
def hello_world():
   return "Hello World"

if __name__ == '__main__':
   app.run(debug=True)


