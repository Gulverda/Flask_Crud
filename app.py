from flask import Flask, request, render_template, url_for, redirect, send_file
from flask_pymongo import PyMongo
from bson.objectid import ObjectId
import pandas as pd
import io

app = Flask(__name__)
app.config["MONGO_URI"] = "mongodb://localhost:27017/test_db"
mongo = PyMongo(app)

@app.route('/')
def index():
    users = mongo.db.users.find()
    return render_template('index.html', users=users)

@app.route('/add', methods=['POST'])
def add():
    name = request.form.get('name')
    email = request.form.get('email')
    password = request.form.get('password')
    mongo.db.users.insert_one({'name': name, 'email': email, 'password': password})
    return redirect(url_for('index'))

@app.route('/edit/<id>', methods=['GET', 'POST'])
def edit_user(id):
    user = mongo.db.users.find_one({'_id': ObjectId(id)})
    if request.method == 'POST':
        name = request.form.get('name')
        email = request.form.get('email')
        password = request.form.get('password')
        mongo.db.users.update_one({'_id': ObjectId(id)}, {'$set': {'name': name, 'email': email, 'password': password}})
        return redirect(url_for('index'))
    return render_template('edit_user.html', user=user)

@app.route('/delete/<id>', methods=['GET'])
def delete_user(id):
    mongo.db.users.delete_one({'_id': ObjectId(id)})
    return redirect(url_for('index'))

@app.route('/export')
def export_users():
    users = list(mongo.db.users.find())
    df = pd.DataFrame(users)
    df.drop('_id', axis=1, inplace=True)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Users')
    
    output.seek(0)
    return send_file(output, attachment_filename='users.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
