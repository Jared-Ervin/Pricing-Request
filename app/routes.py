from flask import render_template, request, url_for
from app import app
import os
import win32com.client
from datetime import datetime


@app.context_processor
def override_url_for():
    return dict(url_for=dated_url_for)


def dated_url_for(endpoint, **values):
    if endpoint == 'static':
        filename = values.get('filename', None)
        if filename:
            file_path = os.path.join(app.root_path,
                                     endpoint, filename)
            values['q'] = int(os.stat(file_path).st_mtime)
    return url_for(endpoint, **values)


@app.route('/')
@app.route('/index')
def index():
    return render_template("index.html", title="Home")


@app.route('/handle_data', methods=["POST"])
def handle_data():
    data = request.form.to_dict()
    price_options = ["Review Only", "High Margin",
                     "Mid Tier", "Agressive", "Floor"]
    selected_prices = []
    flag = ''
    for key in data:
        if key in price_options:
            selected_prices.append(key)
    data2 = data.copy()
    data2["pricepoint"] = selected_prices
    try:
        data2["time"] = datetime.strptime(
            data2["time"], "%H:%M").strftime("%I:%M %p")
    except Exception:
        data2["time"] = ''
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(int(0))
    mail.To = ("j.ervin@supply.com;")
    try:
        if data2["timing"] != '':
            mail.Importance = 2
            flag = "URGENT - "
    except Exception:
        mail.Importance = 1
    mail.Subject = flag + "Pricing Request - " + str(data2["number"].upper())
    mail.HTMLBody = render_template(
        "email.html",
        # name=data2["fullname"].title(),
        number=data2["number"].upper(),
        ffa=data2["ffa"].title(),
        job=data2["job"].title(),
        need=data2["time"],
        price=', '.join(data2["pricepoint"]),
        floor=data2["floor"].title(),
        pay=data2["paytype"].title(),
        ship=data2["shipments"].title(),
        notes=data2["description"])
    mail.Send()

    return render_template("submit.html")
