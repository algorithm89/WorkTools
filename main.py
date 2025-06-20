from flask import Flask, request, jsonify
from flask_sqlalchemy import SQLAlchemy


app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///travel.db'


db = SQLAlchemy(app)

#Create Database

class Destination(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    destination = db.Column(db.String(50), nullable=False)
    country = db.Column(db.String(50), nullable=False)
    rating = db.Column(db.Float, nullable=False)

    def to_dict(self):
        return {"id": self.id,
                "destination": self.destination,
                "country": self.country,
                "rating": self.rating }

    with app.app_context():
        db.create_all()

#Create Routes
@app.route('/')
def home():
    return jsonify({"message": "Welcome to the Travel API!"})


@app.route('/destination', methods=['GET'])
def get_destinations():
    destinations = Destination.query.all()
    return jsonify([destination.to_dict()] for destination in destinations)


@app.route('/destination/<int:destination_id>', methods=['GET'])
def get_destination(destination_id):
    destination=Destination.query.get(destination_id)
    if destination:
        return jsonify(destination.to_dict())
    else:
        return jsonify({"error": "Destination not found"}), 404

#POST

@app.route('/destination', methods=['POST'])
def add_destination():
    data = request.get_json()

    new_destination = Destination(destination=data['destination'],
                                  country=data['country'],
                                  rating=data['rating'])

    db.session.add(new_destination)
    db.session.commit()
    return jsonify(new_destination.to_dict())

if __name__ == '__main__':
    app.run(debug=True)