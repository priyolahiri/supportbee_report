## SupportBee reporting application
Build in progress. Do not fork/pull currently

### Requirements
Python >= 3.4

RethinkDB >= 2.0

Virtualenv

Virtualenvwrapper (for syntactic sugar)

### Installation instructions
Clone the repository

Create a virtualenv

Edit the activate script of the virtualenv to insert this line at the end 'export FLASK_APP=supportbee.py'

Deactivate and re-activate the virtualenv

Run 'pip3 install -r requirements.txt'

Run python cli.py install to follow the instructions

Run python cli.py --help for further instructions
