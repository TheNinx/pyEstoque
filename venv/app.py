from flask import Flask, render_template,request,url_for,Response
from flask_mysqldb import MySQL

import io
import xlwt
import pymysql


app = Flask(__name__)

app.config['MYSQL_HOST'] = "localhost"
app.config['MYSQL_USER'] = "root"
app.config['MYSQL_PASSWORD'] = ""
app.config['MYSQL_DB'] = "estoque"

mysql = MySQL(app);

@app.route("/")
def table():
 return render_template("table.html")

@app.route("/Estoque", methods=['GET','POST'])
def estoque():
    if request.method == 'POST':
        codigo = request.form['codigo']
        descricao = request.form['descricao']
        desconto = request.form['tipo']
        loja = request.form['loja']

        cur = mysql.connection.cursor()
        cur.execute("INSERT INTO tb_estoque(codigo, descricao, tipo, loja) VALUES (%s,%s,%s,%s)",(codigo,descricao,desconto,loja))
        mysql.connection.commit()
        cur.close()
        return "Produto cadastrado com sucesso"

    return render_template("estoque.html")

@app.route("/Produtos")
def produtos():
    cur = mysql.connection.cursor()
    produtos = cur.execute("SELECT * FROM tb_estoque")
    if produtos>0:
        detalesProdutos = cur.fetchall()
        return render_template("produtos.html", detalesProdutos=detalesProdutos)


@app.route("/Produtos")
def relatorioxls():
    conn = mysql.connect()
    cursor = conn.cursor(pymysql.cursors.DictCursor)

    cursor.execute("SELECT codigo,descricao,tipo,loja FROM tb_estoque ")
    resultado = cursor.fetchall()


    output = io.BytesIO()
    workbook = xlwt.Workbook()
    sh = workbook.add_sheet('Employee Report')

    # add headers
    sh.write(0, 0, 'CODIGO PRODUTO')
    sh.write(0, 1, 'DESCRICAO')
    sh.write(0, 2, 'TIPO DESCONTO')
    sh.write(0, 3, 'LOJA')

    idx = 0
    for row in resultado:
        sh.write(idx + 1, 0, str(row['codigo']))
        sh.write(idx + 1, 1, row['descricao'])
        sh.write(idx + 1, 2, row['tipo'])
        sh.write(idx + 1, 3, row['loja'])
        idx += 1

    workbook.save(output)
    output.seek(0)

    return Response(output, mimetype="application/ms-excel",headers={"Content-Disposition": "attachment;filename=RELATORIO.xls"})






