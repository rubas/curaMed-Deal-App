from flask import Flask

app = Flask(__name__, instance_relative_config=True)

# Load the views
from app import views
from app import hubspot

# Load the config file
app.config.from_object('config')
