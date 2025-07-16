from flask import Flask, render_template, request
from excel_search import fast_search_in_excel

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    result = None
    if request.method == "POST":
        keyword = request.form["keyword"]
        result = fast_search_in_excel("excel_file.xlsx", "Sheet1", keyword)
        print("🔍 result", result)
    return render_template("index.html", result=result)

if __name__ == "__main__":
    app.run(debug=True)