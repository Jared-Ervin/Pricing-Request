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
    return render_template("index2.html", title="Home")


@app.route('/handle_data', methods=["POST"])
def handle_data():
    data = request.form.to_dict()
    try:
        data["time"] = datetime.strptime(
            data["time"], "%H:%M").strftime("%I:%M %p")
    except Exception:
        data["time"] = ''

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(int(0))
    mail.To = ("j.ervin@supply.com;")
    try:
        if data["timing"] != '':
            mail.Importance = 2
    except Exception:
        mail.Importance = 1
    mail.Subject = "Pricing Request - " + str(data["number"].upper())
    mail.Body = f"""
        Sales Rep: 
        {data["fullname"].title()}

        Estimate Number: 
        {data["number"].upper()}

        FFA Maufacturers: 
        {data["ffa"].title()}

        Job Pricing: 
        {data["job"].title()}

        Due By: 
        {data["time"]}

        Notes: 
        {data["description"]}
    """
    mail.Send()

    return render_template("complete.html", title="Submitted")
