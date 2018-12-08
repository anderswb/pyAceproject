import requests
import xml.etree.ElementTree as ET
import urllib.parse
from datetime import datetime, timedelta
import argparse

VERSION = '0.1'

verbose = False


def getetree(function, parameters):
    parameters['fct'] = function
    if 'format' not in parameters:
        parameters['format'] = 'xml'
    url = 'http://api.aceproject.com/?{}'.format(urllib.parse.urlencode(parameters))
    r = requests.get(url).content
    if verbose:
        print('Parameters sent:')
        for k, v in parameters.items():
            if k == 'password':
                v = '***'
            print(' - {: <15} {}'.format(k+':', v))
    return ET.fromstring(r)


def login(account, username, password):
    print('Logging into account: \"{}\" using username: \"{}\"'.format(account, username))
    param_dict = {
    'accountid': account,
    'username': username,
    'password': password}
    root = getetree('login', param_dict)
    guid = root.find('row').get('GUID')
    if verbose:
        print('Got guid: {}'.format(guid))
    return guid


def saveworkitem(guid, date, hours, comment, projectid, taskid, debug_mode=False):
    print('Adding time to worksheet...')

    weekday = date.weekday()+1 if date.weekday() < 6 else 0
    weekstart = datetime.strftime(date - timedelta(days=weekday), '%Y-%m-%d')

    hoursday = dict()
    hoursday['sun'] = 0
    hoursday['mon'] = 0
    hoursday['tue'] = 0
    hoursday['wed'] = 0
    hoursday['thu'] = 0
    hoursday['fri'] = 0
    hoursday['sat'] = 0

    if weekday == 0: # sunday
        hoursday['sun'] = hours
    elif weekday == 1: # monday
        hoursday['mon'] = hours
    elif weekday == 2:
        hoursday['tue'] = hours
    elif weekday == 3:
        hoursday['wed'] = hours
    elif weekday == 4:
        hoursday['thu'] = hours
    elif weekday == 5:
        hoursday['fri'] = hours
    elif weekday == 6:
        hoursday['sat'] = hours

    param_dict = {
    'guid': guid,
    'weekstart': weekstart,
    'projectid': projectid,
    'timetypeid': 1, # TODO: make dynamic
    'hoursday1': hoursday['sun'],
    'hoursday2': hoursday['mon'],
    'hoursday3': hoursday['tue'],
    'hoursday4': hoursday['wed'],
    'hoursday5': hoursday['thu'],
    'hoursday6': hoursday['fri'],
    'hoursday7': hoursday['sat'],
    'comments': comment}

    if taskid:
        param_dict['taskid'] = taskid

    if not debug_mode:
        getetree('saveworkitem', param_dict)
    else:
        print("Debug mode enabled, command not sent.")


def getuserid(guid, username):
    print('Getting userid for {}'.format(username))
    param_dict = {
    'guid': guid,
    'FilterUserName': username}
    root = getetree('getusers', param_dict)
    userid = root.find('row').get('USER_ID')
    if verbose:
        print('Got userid {}'.format(userid))
    return int(userid)


def listprojects(guid, username):
    userid = getuserid(guid, username)
    print('Getting all active projects for user {} with id {}'.format(username, userid))
    param_dict = {
    'guid': guid,
    'Filterassigneduserid': userid,
    'Filtercompletedproject': 'False',
    'SortOrder': 'PROJECT_ID'}
    root = getetree('getprojects', param_dict)
    for child in root:
        id = child.attrib.get('PROJECT_ID', '')
        name = child.attrib.get('PROJECT_NAME', '')
        print('{:<5} | {}'.format(id, name))


def listtasks(guid, projectid):
    print('Listing tasks for projectid {}'.format(projectid))
    param_dict = {
    'guid': guid,
    'projectid': projectid,
    'forcombo': 'true'}
    root = getetree('gettasks', param_dict)
    for child in root:
        id = child.attrib.get('TASK_ID', '')
        resume = child.attrib.get('TASK_RESUME', '')
        print('{:<5} | {}'.format(id, resume))


def gettimeentries(guid, username, days=30):
    userid = getuserid(guid, username)
    print('Getting time entries for {} days.'.format(days))
    param_dict = {
    'guid': guid,
    'View': 1,
    'FilterMyWorkItems': 'False',
    'FilterTimeCreatorUserId': userid,
    'FilterDateFrom': datetime.strftime(datetime.today() - timedelta(days=days), '%Y-%m-%d'),
    'FilterDateTo': datetime.strftime(datetime.today(), '%Y-%m-%d')}
    root = getetree('GetTimeReport', param_dict)
    print('-----------+------------+-----------------------+------------+------+--------------------------------------------------')
    print('Date       | Client     | Project               | Task       | Hour | Comment')
    print('-----------+------------+-----------------------+------------+------+--------------------------------------------------')
    for child in root:
        date = child.attrib.get('DATE_WORKED', '----------')[0:10]
        client = child.attrib.get('CLIENT_NAME', '')
        project = child.attrib.get('PROJECT_NAME', '')
        task = child.attrib.get('TASK_RESUME', '')
        hours = child.attrib.get('TOTAL', '')
        comment = child.attrib.get('COMMENT', '')
        print('{} | {:<10.10} | {:<21.21} | {:<10.10} | {:<4.4} | {:<49.49}'.format(date, client, project, task, hours, comment))
    print('-----------+------------+-----------------------+------------+------+--------------------------------------------------')


class ValidateAddHours(argparse.Action):
    def __call__(self, parser, args, values, option_string=None):
        projectid, taskid, date, time, comment = values
        projectid = int(projectid)
        if taskid.upper() == "NA":
            taskid = None
        else:
            try:
                taskid = int(taskid)
            except ValueError:
                raise argparse.ArgumentError(self, 'taskid is not a number or "NA"')
        
        try:
            date = datetime.strptime(date, '%d-%m-%Y')
        except ValueError:
            raise argparse.ArgumentError(self, 'Date is not of the format DD-MM-YYYY, or the day does not exist')

        try:
            time = float(time)
        except ValueError:
            raise argparse.ArgumentError(self, 'Time is not a number')

        if comment == '':
            raise argparse.ArgumentError(self, 'The comment field is empty')
        values = {'projectid': projectid, 'taskid': taskid, 'date': date, 'time': time, 'comment': comment}
        setattr(args, self.dest, values)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Aceproject command line interface v" + VERSION + " by Anders Winther Brandt 2018")
    parser.add_argument('-g', '--debug', help='Do not store any values', action="store_true")
    parser.add_argument('-v', '--verbose', help='Print more information', action="store_true")
    group = parser.add_mutually_exclusive_group(required = True)
    group.add_argument('-a', '--addhours', nargs=5, metavar=('PROJECTID', 'TASKID', 'DATE', 'TIME', 'COMMENT'), action=ValidateAddHours,
    help='Add a new time entry. projectid: ID of the project to add the hours to. taskid: The ID of the task to add the hours to, set to NA to not assign a task. data: The date in the format dd-mm-yyyy. Comment: The comment line')
    group.add_argument('-p', '--projects', nargs=1, type=str, metavar=('USERNAME'), help="Get a list of active project for the given username")
    group.add_argument('-t', '--tasks', nargs=1, type=int, metavar=('PROJECTID'), help="Get a list of all tasks for a given project ID")
    group.add_argument('-l', '--log', nargs=2, metavar=('USERNAME', 'DAYS'), help='Get all time entries for the specified number of days for the specified username')
    args = parser.parse_args()

    verbose = args.verbose

    print('Reading settings from config.txt...')
    with open('.\\config.txt') as f:
        account = f.readline().rstrip('\n') 
        user = f.readline().rstrip('\n') 
        password = f.readline().rstrip('\n') 

    ## LOGIN
    guid = login(account, user, password)

    if args.addhours:
        saveworkitem(guid, args.addhours['date'], args.addhours['time'], args.addhours['comment'], args.addhours['projectid'], args.addhours['taskid'], args.debug)
    elif args.projects:
        listprojects(guid, args.projects[0])
    elif args.tasks:
        listtasks(guid, args.tasks[0])
    elif args.log:
        gettimeentries(guid, args.log[0], int(args.log[1]))

    print('Done')
