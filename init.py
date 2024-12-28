from flask import Flask
from routes import main
import logging

def create_app():
    app = Flask(__name__)
    # logging.basicConfig(filename='record.log', level=logging.DEBUG, format=f'%(asctime)s %(levelname)s %(name)s %(threadName)s : %(message)s')
    logging.basicConfig(level=logging.INFO, format=f'%(asctime)s %(levelname)s %(name)s : %(message)s')
    app.config.from_object('config.Config')
    app.logger.debug(f"upload_folder= {app.config['UPLOAD_FOLDER']}")
    app.register_blueprint(main)
    return app
