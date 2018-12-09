import requests
import xml.etree.ElementTree as ET
import urllib.parse
from datetime import datetime, timedelta
import argparse
import textwrap

VERSION = '0.1'

verbose = False


def getetree(function, parameters):
    parameters['fct'] = function
    if 'format' not in parameters:
        parameters['format'] = 'xml'
    url = 'http://api.aceproject.com/?{}'.format(urllib.parse.urlencode(parameters))
    try:
        r = requests.get(url)
        r.raise_for_status()
    except requests.exceptions.RequestException as e:
        print('Connectivity error. Returned exception:')
        print(e)
        print("Exiting!")
        exit(1)
    if verbose:
        print('Parameters sent:')
        for k, v in parameters.items():
            if k == 'password':
                v = '***'
            print(' - {: <15} {}'.format(k+':', v))
    return ET.fromstring(r.content)


def login(account, username, password):
    print('Logging into account: \"{}\" using username: \"{}\"'.format(account, username))
    param_dict = {
    'accountid': account,
    'username': username,
    'password': password}
    root = getetree('login', param_dict)
    guid = root.find('row').get('GUID')
    if not guid:
        print('Failed to log in, exiting!')
        exit(1)
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
        root = getetree('saveworkitem', param_dict)
        error_description = root.find('row').get('ErrorDescription')
        if error_description:
            print('Something went wrong when adding the time item, the following error message was returned from the server:\n"{}"'.format(error_description))
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
    print('+--------+------------------------------------------------------------------------------------------------------------+')
    print('| ID     | Project name                                                                                               |')
    print('+--------+------------------------------------------------------------------------------------------------------------+')
    for child in root:
        id = child.attrib.get('PROJECT_ID', '')
        name = child.attrib.get('PROJECT_NAME', '')
        print('| {:<6.6} | {:<106} |'.format(id, name))
    print('+--------+------------------------------------------------------------------------------------------------------------+')


def listtasks(guid, projectid):
    print('Listing tasks for projectid {}'.format(projectid))
    param_dict = {
    'guid': guid,
    'projectid': projectid,
    'forcombo': 'true'}
    root = getetree('gettasks', param_dict)
    print('+--------+------------------------------------------------------------------------------------------------------------+')
    print('| ID     | Task name                                                                                                  |')
    print('+--------+------------------------------------------------------------------------------------------------------------+')
    for child in root:
        id = child.attrib.get('TASK_ID', '')
        resume = child.attrib.get('TASK_RESUME', '')
        print('| {:<6.6} | {:<106} |'.format(id, resume))
    print('+--------+------------------------------------------------------------------------------------------------------------+')


def gettimeentries(guid, username, days=30):
    userid = getuserid(guid, username)
    print('Getting time entries for {} days.'.format(days))
    param_dict = {
    'guid': guid,
    'View': 1,
    'FilterMyWorkItems': 'False',
    'FilterTimeCreatorUserId': userid,
    'FilterDateFrom': datetime.strftime(datetime.today() - timedelta(days=days), '%Y-%m-%d'),
    'FilterDateTo': datetime.strftime(datetime.today() + timedelta(days=10*356), '%Y-%m-%d')}
    root = getetree('GetTimeReport', param_dict)
    print('+--------+------------+--------------------------+-----------+------+-------------------------------------------------+')
    print('| Date   | Client     | Project                  | Task      | Hour | Comment                                         |')
    print('+--------+------------+--------------------------+-----------+------+-------------------------------------------------+')
    wrapper = textwrap.TextWrapper(width=47)
    for child in root:
        date_str = child.attrib.get('DATE_WORKED', '')[0:10]
        datetime_obj = datetime.strptime(date_str, '%Y-%m-%d')
        date = datetime.strftime(datetime_obj, '%y%m%d')
        client = child.attrib.get('CLIENT_NAME', '')
        project = child.attrib.get('PROJECT_NAME', '')
        task = child.attrib.get('TASK_RESUME', '')
        hours = child.attrib.get('TOTAL', '')
        comment = child.attrib.get('COMMENT', '-')
        comment_wrapped = wrapper.wrap(comment)
        first_line = True
        for comment_line in comment_wrapped:
            if not first_line:
                date, client, project, task, hours = ('', '', '', '', '')    
            print('| {:<6.6} | {:<10.10} | {:<24.24} | {:<9.9} | {:<4.4} | {:<47.47} |'.format(date, client, project, task, hours, comment_line))
            first_line = False
    print('+--------+------------+--------------------------+-----------+------+-------------------------------------------------+')


def loadconfig():
    print('Reading settings from config.txt...')
    try:
        f = open('.\\config.txt')
    except IOError:
        print('Unable to open config.txt. This file is required. It must be located in the working path and must contain 3 lines: 1st company account name, 2nd username, 3rd password.')
        print('Exiting!')
        exit(1)
    else:
        with f:
            account = f.readline().rstrip('\n') 
            user = f.readline().rstrip('\n') 
            password = f.readline().rstrip('\n')
    
    if not account or not user or not password:
        print('config.txt must contain 3 lines: 1st company account name, 2nd username, 3rd password. Exiting!')
        exit(1)
    
    return account, user, password


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
            date = datetime.strptime(date, '%y%m%d')
        except ValueError:
            raise argparse.ArgumentError(self, 'Date is not of the format YYMMDD, or the day does not exist')

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
    help='Add a new time entry. projectid: ID of the project to add the hours to. taskid: The ID of the task to add the hours to, set to NA to not assign a task. data: The date in the format YYMMDD. Comment: The comment line')
    group.add_argument('-p', '--projects', nargs=1, type=str, metavar=('USERNAME'), help="Get a list of active project for the given username")
    group.add_argument('-t', '--tasks', nargs=1, type=int, metavar=('PROJECTID'), help="Get a list of all tasks for a given project ID")
    group.add_argument('-l', '--log', nargs=2, metavar=('USERNAME', 'DAYS'), help='Get all time entries for all future entries and DAYS in the past, for the specified username. Eg. DAYS=10 will get all future entries and for the past 10 days.')
    args = parser.parse_args()

    verbose = args.verbose

    account, user, password = loadconfig() # load configuration file
    guid = login(account, user, password) # log in and get the guid used in subsequent API calls

    if args.addhours:
        saveworkitem(guid, args.addhours['date'], args.addhours['time'], args.addhours['comment'], args.addhours['projectid'], args.addhours['taskid'], args.debug)
    elif args.projects:
        listprojects(guid, args.projects[0])
    elif args.tasks:
        listtasks(guid, args.tasks[0])
    elif args.log:
        gettimeentries(guid, args.log[0], int(args.log[1]))

    print('Done')
