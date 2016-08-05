from flask import Flask, g, abort
import configparser
import rethinkdb as r
from rethinkdb.errors import RqlDriverError
import os

app = Flask(__name__)
app_settings = configparser.ConfigParser()

# If config file does not exist, run setup routine
if not os.path.isfile('./config.ini'):
    print("Config file not found. Please run 'python setup.py install'")
    exit()
app_settings.read('./config.ini')


@app.before_request
def before_request():
    try:
        g.rdb_conn = r.connect(
            host=app_settings['rethink']['db_host'],
            port=app_settings['rethink']['db_port'],
            db=app_settings['rethink']['db_name']
            )
    except RqlDriverError:
        abort(503, "No database connection could be established.")


@app.teardown_request
def teardown_request(exception):
    try:
        g.rdb_conn.close()
    except AttributeError:
        pass


if __name__ == "__main__":
    app.run(port=app_settings['web']['port'])
