import flask

app = flask.Flask(__name__)

@app.route("/")
def hello_banana():
    return "<p>Hello, Banana!</p>"