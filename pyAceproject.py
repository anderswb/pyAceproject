import requests
import xml.etree.ElementTree as ET
import urllib.parse
from datetime import datetime, timedelta
import argparse
import textwrap
import xml.dom.minidom

VERSION = '0.1'

verbose = False
debug_mode = False

debug_filename = 'debug_{}_{}.xml'


def print_parameters(parameters):
    for k, v in parameters.items():
        if k == 'password':
            v = '***'
        print(' - {: <18} {}'.format(k + ':', v))


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
        print_parameters(parameters)
    if debug_mode:
        xml_str = xml.dom.minidom.parseString(r.content.decode('utf-8'))
        pretty_xml_as_string = xml_str.toprettyxml()
        filename = debug_filename.format(parameters['fct'], datetime.strftime(datetime.now(), '%y%m%d-%H%M%S%f')).lower()
        with open(filename, 'w') as f:
            print('Saving received xml data to {}'.format(filename))
            f.write(pretty_xml_as_string)

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


def saveworkitem(guid, date, hours, comment, projectid, taskid, line_id=None):
    if not line_id:
        print('Adding time to worksheet...')
    else:
        print('Editing time entry...')

    # get the weekday of the passed date and convert it to sunday=0 and friday=6
    weekday = date.weekday() + 1 if date.weekday() < 6 else 0

    # get the start date of the week as a aceproject timestamp
    weekstart = datetime.strftime(date - timedelta(days=weekday), '%Y-%m-%d')

    # add the hours to the correct weekday
    hours_per_day = [0, 0, 0, 0, 0, 0, 0, 0]
    hours_per_day[weekday] = hours

    param_dict = {
        'guid': guid,
        'weekstart': weekstart,
        'projectid': projectid,
        'timetypeid': 1,  # TODO: make dynamic
        'hoursday1': hours_per_day[0],
        'hoursday2': hours_per_day[1],
        'hoursday3': hours_per_day[2],
        'hoursday4': hours_per_day[3],
        'hoursday5': hours_per_day[4],
        'hoursday6': hours_per_day[5],
        'hoursday7': hours_per_day[6],
        'comments': comment}

    if taskid:
        param_dict['taskid'] = taskid

    if line_id:
        param_dict['TimesheetLineId'] = line_id

    if not debug_mode:
        root = getetree('saveworkitem', param_dict)
        error_description = root.find('row').get('ErrorDescription')
        if error_description:
            print('Something went wrong when adding the time item, the following error message was returned from the server:\n"{}"'.format(error_description))
    else:
        print('Debug mode enabled, command not sent.')
        if verbose:
            print('Would have sent the following parameters:')
            print_parameters(param_dict)


def getuserid(guid, username):
    print('Getting userid for {}'.format(username))
    param_dict = {
        'guid': guid,
        'FilterUserName': username}
    root = getetree('getusers', param_dict)
    if not root:
        print('Unable to find a userid for {}. Check the username and try again. Exiting!'.format(username))
        exit(1)
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
    if len(root) == 0:
        print('No active projects found')
    else:
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
    if len(root) == 0:
        print('No tasks found for the given projectid')
    else:
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
        'FilterDateTo': datetime.strftime(datetime.today() + timedelta(days=10 * 356), '%Y-%m-%d')}
    root = getetree('GetTimeReport', param_dict)
    print('+-----+------+----------+-------------------------+---------+----+----------------------------------------------------+')
    print('| ID  | Date | Client   | Project                 | Task    | T  | Comment                                            |')
    print('+-----+------+----------+-------------------------+---------+----+----------------------------------------------------+')
    wrapper = textwrap.TextWrapper(width=52)
    for child in root:
        line_id = child.attrib.get('TIMESHEET_LINE_ID', '')
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
                line_id, date, client, project, task, hours = ('', '', '', '', '', '')
            print('|{:<5.5}|{:<6.6}|{:<10.10}|{:<25.25}|{:<9.9}|{:<4.4}|{:<52.52}|'.format(line_id, date, client, project, task, hours, comment_line))
            first_line = False
    print('+-----+------+----------+-------------------------+---------+----+----------------------------------------------------+')


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
        if self.dest == 'addhours':
            projectid, taskid, date, time, comment = values
            lineid = None
        else:
            lineid, projectid, taskid, date, time, comment = values

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

        if lineid:
            try:
                lineid = int(lineid)
            except ValueError:
                raise argparse.ArgumentError(self, 'LINEID is not a number')

        values = {'projectid': projectid, 'taskid': taskid, 'date': date, 'time': time, 'comment': comment, 'lineid': lineid}
        setattr(args, self.dest, values)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Aceproject command line interface v" + VERSION + " by Anders Winther Brandt 2018")
    parser.add_argument('-g', '--debug', help='Do not store any values', action="store_true")
    parser.add_argument('-v', '--verbose', help='Print more information', action="store_true")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-a', '--addhours', nargs=5, metavar=('PROJECTID', 'TASKID', 'DATE', 'TIME', 'COMMENT'), action=ValidateAddHours,
                       help='Add a new time entry. projectid: ID of the project to add the hours to. taskid: The ID of the task to add the hours to, set to NA to not assign a task. data: The date in the format YYMMDD. Comment: The comment line')
    group.add_argument('-e', '--edithours', nargs=6, metavar=('LINEID', 'PROJECTID', 'TASKID', 'DATE', 'TIME', 'COMMENT'), action=ValidateAddHours,
                       help='Edit an existing time entry. Same parameters as for --addhours, but with the addition LINEID parameters, as can be found in the log')
    group.add_argument('-p', '--projects', nargs=1, type=str, metavar=('USERNAME'), help="Get a list of active project for the given username")
    group.add_argument('-t', '--tasks', nargs=1, type=int, metavar=('PROJECTID'), help="Get a list of all tasks for a given project ID")
    group.add_argument('-l', '--log', nargs=2, metavar=('USERNAME', 'DAYS'), help='Get all time entries for all future entries and DAYS in the past, for the specified username. Eg. DAYS=10 will get all future entries and for the past 10 days.')
    args = parser.parse_args()

    verbose = args.verbose
    debug_mode = args.debug

    account, user, password = loadconfig()  # load configuration file
    guid = login(account, user, password)  # log in and get the guid used in subsequent API calls

    if args.addhours:
        saveworkitem(guid, args.addhours['date'], args.addhours['time'], args.addhours['comment'], args.addhours['projectid'], args.addhours['taskid'])
    if args.edithours:
        saveworkitem(guid, args.edithours['date'], args.edithours['time'], args.edithours['comment'], args.edithours['projectid'], args.edithours['taskid'], line_id=args.edithours['lineid'])
    if args.projects:
        listprojects(guid, args.projects[0])
    if args.tasks:
        listtasks(guid, args.tasks[0])
    if args.log:
        gettimeentries(guid, args.log[0], int(args.log[1]))

    print('Done')
