from flask import Flask, render_template
app = Flask(__name__)

@app.route('/')
def inicio():
    return render_template('inicio.html')

@app.route('/Ofimatica')
def ofimatica():
    return render_template('ofimatica.html')

@app.route('/VisualStudioCode')
def visualStudioCode():
    return render_template('visualstudiocode.html')

@app.route('/HTML_CSS')
def htmlCss():
    return render_template('htmlcss.htm')

@app.route('/Javascript')
def javaScript():
    return render_template('javascript.html')

@app.route('/Git_Github')
def gitGitHub():
    return render_template('gitgithub.html')

@app.route('/frameworkflask')
def frameworkFlask():
    return render_template('frameworkflask.html')

@app.route('/arduino_uno')
def arduinoUno():
    return render_template('arduino_uno.html')

if __name__=='__main__':
    app.run(debug=True)
    """ app.run(host="0.0.0.0", port=10000) se debe aplicar para subirlo al servidor RENDER y comentar la línea de arriba """


