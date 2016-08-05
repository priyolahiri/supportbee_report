import rethinkdb as r
import requests
import arrow
import simplejson
from openpyxl import Workbook
from operator import itemgetter


class SupportBee:

    def __init__(self, app_settings):
        self.app_settings = app_settings
        self.headers = {'Content-Type': 'application/json',
                        'Accept': 'application/json'}
        self.base_url = "https://" + \
            app_settings['supportbee']['company'] + ".supportbee.com"
        self.url_params = {"auth_token": app_settings['supportbee']['apikey']}
        self.r_conn = r.connect(host=app_settings['rethink']['db_host'], port=int(app_settings['rethink']['db_port']),
                                db=app_settings['rethink']['db_name'])

    def get_tickets(self, page=1, per_page=99, archived=False, assigned_user=False, assigned_team=False, label=False,
                    since=False, until=False, sort_by="last_actvity", requester_emails=[], replies=False):
        print("starting")
        url_params = self.url_params
        if archived:
            url_params['archived'] = "any"
        if assigned_user:
            url_params['assigned_user'] = assigned_user
        if assigned_team:
            url_params['assigned_team'] = assigned_team
        if label:
            url_params['label'] = "label"
        if since:
            url_params['since'] = arrow.get(since).format('YYYY-MM-DD')
        if until:
            url_params['since'] = arrow.get(since).format('YYYY-MM-DD')
        url_params['sort_by'] = sort_by
        url_params['page'] = page
        url_params['per_page'] = per_page
        if requester_emails:
            url_params['requester_emails'] = ','.join(requester_emails)
        get_tickets = requests.get(
            self.base_url + "/tickets", params=url_params, headers=self.headers)
        if get_tickets.status_code != 200:
            raise SupportBeeException(error_type="Get Ticket Error", message="Wrong status code received - " +
                                                                             str(get_tickets.status_code))
        return simplejson.loads(get_tickets.text)

    def write_ticket(self, ticket=dict):
        ticket_id = ticket['id']
        write_ticket = dict()
        write_ticket['id'] = ticket_id
        write_ticket['source_type'] = list(ticket['source'].keys())[0]
        write_ticket['source'] = list(ticket['source'].values())[0]
        write_ticket['labels'] = []
        for label in ticket['labels']:
            write_ticket['labels'].append(label['name'])
        write_ticket['subject'] = ticket['subject']
        write_ticket['replies_count'] = ticket['replies_count']
        write_ticket['agent_replies_count'] = ticket['agent_replies_count']
        write_ticket['comments_count'] = ticket['comments_count']
        write_ticket['created_at'] = arrow.get(ticket['created_at']).datetime
        write_ticket['last_activity_at'] = arrow.get(
            ticket['last_activity_at']).datetime
        write_ticket['unanswered'] = ticket['unanswered']
        write_ticket['closed'] = ticket['archived']
        write_ticket['private'] = ticket['private']
        write_ticket['trash'] = ticket['trash']
        write_ticket['draft'] = ticket['draft']
        try:
            write_ticket['current_team_assignee_id'] = ticket[
                'current_team_assignee']['team']['id']
            write_ticket['current_team_assignee_name'] = ticket[
                'current_team_assignee']['team']['name']
        except KeyError:
            write_ticket['current_team_assignee_id'] = None
            write_ticket['current_team_assignee_name'] = None
        try:
            write_ticket['current_user_assignee_id'] = ticket[
                'current_user_assignee']['user']['id']
            write_ticket['current_user_assignee_name'] = ticket[
                'current_user_assignee']['user']['name']
        except KeyError:
            write_ticket['current_user_assignee_id'] = None
            write_ticket['current_user_assignee_name'] = None
        write_ticket['starred'] = ticket['starred']
        write_ticket['cc'] = []
        for cc in ticket['cc']:
            write_ticket['cc'].append(
                {'id': cc['id'], 'name': cc['name'], 'email': cc['email']}
            )
        write_ticket['requester_id'] = ticket['requester']['id']
        write_ticket['requester_name'] = ticket['requester']['name']
        write_ticket['requester_email'] = ticket['requester']['email']
        r.table('tickets').insert(write_ticket,
                                  conflict="replace").run(self.r_conn)

    def get_replies(self, ticket_id):
        url_params = self.url_params
        get_replies = requests.get(self.base_url + "/tickets/" + str(ticket_id) + "/replies", params=url_params,
                                   headers=self.headers)
        if get_replies.status_code != 200:
            raise SupportBeeException(error_type="Get Replies Error", message="Wrong status code received - " +
                                      str(get_replies.status_code))
        replies_data = simplejson.loads(get_replies.text)
        reply_return = []
        for reply in replies_data['replies']:
            reply_push = dict()
            reply_push['id'] = reply['id']
            reply_push['created_at'] = arrow.get(reply['created_at']).datetime
            reply_push['replier_id'] = reply['replier']['id']
            reply_push['replier_email'] = reply['replier']['email']
            reply_push['replier_name'] = reply['replier']['name']
            reply_push['replier_agent'] = reply['replier']['agent']
            reply_push['ticket_id'] = ticket_id
            reply_return.append(reply_push)
        return reply_return

    def write_replies(self, replies):
        for reply in replies:
            r.table('replies').insert(
                reply, conflict="replace").run(self.r_conn)

    def get_replies_db(self, ticket_id):
        get_replies = r.table('replies').filter(
            {'ticket_id': ticket_id}).run(self.r_conn)
        replies_return = []
        for reply in get_replies:
            replies_return.append(reply)
        return replies_return

    def excel(self, since=False, until=False, filename="filename"):
        tickets = r.table('tickets')
        if since or until:
            if since:
                tickets = tickets.filter(
                    r.row["created_at"] >= arrow.get(since).to('utc').datetime)
            if until:
                tickets = tickets.filter(
                    r.row["created_at"] <= arrow.get(until).to('utc').datetime)
        if tickets.count().run(self.r_conn):
            print(tickets.count().run(self.r_conn), " tickets found")
            tickets = tickets.run(self.r_conn)
        else:
            print("No tickets found")
            return False
        wb = Workbook()
        ws_second = wb.active
        ws_second.title = "Support Tickets"
        c = 1
        ws_second["A"+str(c)] = "Subject"
        ws_second["B" + str(c)] = "Team Assigned"
        ws_second["C" + str(c)] = "User Assigned"
        ws_second["D" + str(c)] = "Created"
        ws_second["E" + str(c)] = "First Response"
        ws_second["F" + str(c)] = "Last Response"
        ws_second["G" + str(c)] = "Closed"
        ws_second["H" + str(c)] = "Labels"
        ws_second["I" + str(c)] = "FRT"
        ws_second["J" + str(c)] = "CT"
        ws_second["K" + str(c)] = "Requester"
        for ticket in tickets:
            c += 1
            ws_second["A"+str(c)] = ticket['subject']
            ws_second["B"+str(c)] = ticket['current_team_assignee_name']
            ws_second["C" + str(c)] = ticket['current_user_assignee_name']
            created_at = arrow.get(ticket['created_at']).to(
                self.app_settings['web']['timezone']).datetime
            ws_second["D" + str(c)] = created_at
            get_replies = self.get_replies_db(ticket['id'])
            replies_count = len(get_replies)
            get_replies.sort(key=itemgetter('created_at'), reverse=True)
            replies = get_replies
            if c == 2:
                print(replies)
            first_response = None
            last_response = None
            frt = None
            if replies_count == 1:
                first_response = arrow.get(replies[0]['created_at']).to(
                    self.app_settings['web']['timezone']).datetime
                last_response = arrow.get(replies[0]['created_at']).to(
                    self.app_settings['web']['timezone']).datetime
                frt = (first_response - created_at).seconds/60
            if replies_count > 1:
                first_response = arrow.get(replies[replies_count-1]['created_at']).to(
                    self.app_settings['web']['timezone']).datetime
                last_response = arrow.get(replies[0]['created_at']).to(
                    self.app_settings['web']['timezone']).datetime
                frt = (first_response - created_at).seconds/60
            ws_second["E" + str(c)] = first_response
            ws_second["F" + str(c)] = last_response
            if ticket['closed']:
                ws_second["G" + str(c)] = "Y"
                if not first_response:
                    ct = None
                else:
                    ct = (last_response - created_at).seconds/60
                    if frt > ct:
                        ct = frt
            else:
                ws_second["G" + str(c)] = "N"
                ct = None
            if len(ticket['labels']):
                ws_second["H" + str(c)] = ",".join(ticket['labels'])
            else:
                ws_second["H" + str(c)] = None
            ws_second["I" + str(c)] = frt
            ws_second["J" + str(c)] = ct
            ws_second["K" + str(c)] = ticket['requester_name'] + \
                " <" + ticket['requester_email'] + ">"
        wb.save("xlsx/" + filename + ".xlsx")


class SupportBeeException:

    def __init__(self, error_type, message):
        self.error_type = error_type
        self.message = message

    def __str__(self):
        return repr("Error occurred of Type " + self.error_type + " with message " + self.message)
