import configparser
import rethinkdb as r
from rethinkdb.errors import RqlRuntimeError, RqlDriverError
import os
import requests
from lib import SupportBee
from cement.core.foundation import CementApp
from cement.core.controller import CementBaseController, expose
import click


class InstallController(CementBaseController):

    class Meta:
        label = 'installcontroller'
        description = "Install the software"

    @expose(help="This will setup the application")
    def install(self):
        if os.path.isfile('./config.ini'):
            try:
                click.confirm("Config file already exists. Do you want to continue overwriting the current file?",
                              abort=True)
            except click.exceptions.Abort:
                print("Exiting...")
                exit()
        app_settings = configparser.ConfigParser()
        app_settings['rethink'] = {}
        app_settings['supportbee'] = {}
        app_settings['web'] = {}
        app_settings['rethink']['db_name'] = click.prompt(
            "Database Name", default="supportbee")
        app_settings['rethink']['db_host'] = click.prompt(
            "Database Host", default="localhost")
        app_settings['rethink']['db_port'] = click.prompt(
            "Database Port", default="28015")
        try:
            r_conn = r.connect(host=app_settings['rethink'][
                               'db_host'], port=int(app_settings['rethink']['db_port']))
        except RqlDriverError:
            r_conn = False
            print("Could not connect to database")
            exit()
        try:
            r.db_create(app_settings['rethink']['db_name']).run(r_conn)
        except RqlRuntimeError as error:
            print(error)
            try:
                click.confirm(
                    "Database already exists. Do you want to continue?", abort=True)
            except click.exceptions.Abort:
                print("Exiting...")
                exit()
        try:
            r.db(app_settings['rethink']['db_name']
                 ).table_create("tickets").run(r_conn)
            r.db(app_settings['rethink']['db_name']
                 ).table_create("replies").run(r_conn)
            r.db(app_settings['rethink']['db_name']
                 ).table_create("teams").run(r_conn)
            r.db(app_settings['rethink']['db_name']
                 ).table_create("users").run(r_conn)
            r.db(app_settings['rethink']['db_name']
                 ).table_create("labels").run(r_conn)
        except RqlRuntimeError as error:
            print(error)
            try:
                click.confirm(
                    "Tickets table already exists. Do you want to continue?", abort=True)
            except click.exceptions.Abort:
                print("Exiting...")
                exit()
        app_settings['web']['port'] = click.prompt(
            "On which port do you want to run this app?")
        app_settings['supportbee']['company'] = click.prompt(
            "SupportBee Company (?.supportbee.com)")
        app_settings['supportbee'][
            'apikey'] = click.prompt("SupportBee API Key")
        headers = {'Content-Type': 'application/json',
                   'Accept': 'application/json'}
        base_url = "https://" + \
            app_settings['supportbee']['company'] + ".supportbee.com"
        url_params = {"auth_token": app_settings['supportbee']['apikey']}
        check_request = requests.get(
            base_url + "/users", params=url_params, headers=headers)
        if check_request.status_code == 200:
            print("Authenticated.")
        else:
            print(check_request.status_code)
            print(
                "Non-200 status received from SupportBee. Please check your API key and company name...")
            exit()
        app_settings['web']['timezone'] = click.prompt("Enter timezone (from TZ column at "
                                                       "https://en.wikipedia.org/wiki/List_of_tz_database_time_zones)",
                                                       default="Asia/Kolkata")
        with open('config.ini', 'w') as configfile:
            app_settings.write(configfile)
        print("Congrats! Setup is complete. Please refer to readme for further steps...")
        exit()


class ExcelController(CementBaseController):

    class Meta:
        label = 'excelcontroller'
        description = "Excel Report Generation"
        arguments = [
            (['-s', '--since'],
             dict(action='store', help='date in YYYY-MM-DD format of starting date (for excel report)')),
            (['-u', '--until'],
             dict(action='store', help='date in YYYY-MM-DD format of ending date (for excel report)')),
        ]

    def default(self):
        print("Initializing...")
        if not os.path.isfile('./config.ini'):
            print(
                "The settings file does not exists. Please run 'python support_cli.py install'")
        else:
            print("Initializing...")
            app_settings = configparser.ConfigParser()
            app_settings.read('./config.ini')
            self.app_settings = app_settings

    @expose(help="This will generate excel report", )
    def excel_report(self):
        filename = click.prompt(
            "What do you want to name the file (before the .xlsx)?", default="myfile")
        print(self.app.pargs)
        since = False
        until = False
        self.default()
        app_settings = self.app_settings
        supportbee = SupportBee(app_settings=app_settings)
        if self.app.pargs.since:
            since = self.app.pargs.since
        if self.app.pargs.until:
            until = self.app.pargs.until
        supportbee.excel(since=since, until=until, filename=filename)


class SupportBaseController(CementBaseController):

    class Meta:
        label = 'base'
        description = 'Support Bee CLI tool'

    @expose(hide=True)
    def default(self):
        print("Initializing...")
        if not os.path.isfile('./config.ini'):
            print(
                "The settings file does not exists. Please run 'python support_cli.py install'")
        else:
            app_settings = configparser.ConfigParser()
            app_settings.read('./config.ini')
            self.app_settings = app_settings

    @expose(help="This will synchronize tickets")
    def sync_tickets(self):
        self.default()
        app_settings = self.app_settings
        print("Getting tickets")
        supportbee = SupportBee(app_settings=app_settings)
        ticket_response = supportbee.get_tickets(archived=True)
        print(str(len(ticket_response['tickets'])))
        c = 1
        for ticket in ticket_response['tickets']:
            print("Writing ticket " + str(c))
            supportbee.write_ticket(ticket)
            print("Getting replies for ticket " + str(c))
            replies = supportbee.get_replies(ticket['id'])
            print("Found " + str(len(replies)) +
                  " replies. Writing the replies")
            supportbee.write_replies(replies)
            c += 1
        if ticket_response['total_pages'] > 1:
            for x in range(2, ticket_response['total_pages']+1):
                ticket_response = supportbee.get_tickets(archived=True, page=x)
                print(str(len(ticket_response['tickets'])))
                for ticket in ticket_response['tickets']:
                    print("Writing ticket " + str(c))
                    supportbee.write_ticket(ticket)
                    print("Getting replies for ticket " + str(c))
                    replies = supportbee.get_replies(ticket['id'])
                    print("Found " + str(len(replies)) +
                          " replies. Writing the replies")
                    supportbee.write_replies(replies)
                    c += 1


class SupportCLI(CementApp):

    class Meta:
        label = 'support_cli'
        base_controller = 'base'
        handlers = [SupportBaseController, InstallController, ExcelController]


with SupportCLI() as app:
    app.run()
