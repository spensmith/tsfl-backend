import flask

app = flask.Flask(__name__)


@app.route("/")
def hello_banana():
    return "<p>Hello, Banana!</p>"


@app.route("/picks")
def picks():
    week_range = range(1, 18)
    return '<br/>'.join([f"<a href='www.google.com'>Week {week}</a>" for week in week_range])
