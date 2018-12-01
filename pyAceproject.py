from requests_xml import XMLSession
import urllib.parse
from datetime import datetime, timedelta
import argparse

VERSION = '0.1'

def login(account, username, password):
    print('Logging into account: \"{}\" using username: \"{}\"'.format(account, username))
    url = 'http://api.aceproject.com/?fct=login&accountid={}&username={}&password={}&browserinfo=NULL&language=NULL&format=ds'.format(account, username, password)
    session = XMLSession()
    r = session.get(url)
    item = r.xml.xpath('//GUID', first=True)
    guid = item.text
    return guid


def gettimetypes(guid):
    params = urllib.parse.urlencode({'fct': 'gettimetypes',
    'guid': guid,
    'timetypeid': 'NULL',
    'sortorder': 'NULL',
    'format': 'DS'})
    session = XMLSession()
    url = 'http://api.aceproject.com/?%s' % params
    r = session.get(url)
    #print(r.text)


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

    param_dict = {'fct': 'saveworkitem',
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
    'comments': comment,
    'format': 'ds'}

    if taskid:
        param_dict['taskid'] = taskid

    print('Parameters sent:')
    for k,v in param_dict.items():
        print(' - {: <13} {}'.format(k+':', v))

    params = urllib.parse.urlencode(param_dict)

    url = 'http://api.aceproject.com/?%s' % params

    if not debug_mode:
        session = XMLSession()
        r = session.get(url)
    else:
        print("Debug mode enabled, command not sent.")

def getuserid(guid, username):
    print('Getting userid of user {}'.format(username))
    param_dict = {'fct': 'getusers',
    'guid': guid,
    'FilterUserName': username,
    'format': 'ds'}
    session = XMLSession()
    params = urllib.parse.urlencode(param_dict)
    url = 'http://api.aceproject.com/?%s' % params
    r = session.get(url)
    userid = r.xml.xpath('//USER_ID', first=True)
    return int(userid.text)

def listprojects(guid, username):
    userid = getuserid(guid, username)
    print('Getting all active projects for user "{}" with id {}'.format(username, userid))
    param_dict = {'fct': 'getprojects',
    'guid': guid,
    'Filterassigneduserid': userid,
    'Filtercompletedproject': 'False',
    'SortOrder': 'PROJECT_ID',
    'format': 'ds'}
    session = XMLSession()
    params = urllib.parse.urlencode(param_dict)
    url = 'http://api.aceproject.com/?%s' % params
    r = session.get(url)
    #print(r.text)
    ids = r.xml.xpath('//PROJECT_ID')
    names = r.xml.xpath('//PROJECT_NAME')
    projects = list()
    for id, name in zip(ids, names):
        projects.append({'id': id.text, 'project_name': name.text})
    
    for project in projects:
        print('{:<5}: {}'.format(project['id'], project['project_name']))
    #print(projects)

class ValidateAddTime(argparse.Action):
    def __call__(self, parser, args, values, option_string=None):
        projectid, taskid, date, time, comment = values
        projectid = int(projectid)
        if taskid == "NA":
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
    group = parser.add_mutually_exclusive_group(required = True)
    group.add_argument('-t', '--addtime', nargs=5, metavar=('projectid', 'taskid', 'date', 'time', 'comment'), action=ValidateAddTime,
    help='Add a new time entry. projectid: ID of the project to add the hours to. taskid: The ID of the task to add the hours to, set to NA to not assign a task. data: The date in the format dd-mm-yyyy. Comment: The comment line')
    group.add_argument('-p', '--projects', nargs=1, metavar=('username'), help="Get a list a off active project for the given username")
    args = parser.parse_args()

    print('Reading settings from config.txt...')
    with open('.\\config.txt') as f:
        account = f.readline().rstrip('\n') 
        user = f.readline().rstrip('\n') 
        password = f.readline().rstrip('\n') 

    ## LOGIN
    guid = login(account, user, password)

    if args.addtime:
        hours = args.addtime['time']
        projectid = args.addtime['projectid']
        comment = args.addtime['comment']
        date = args.addtime['date']
        taskid = args.addtime['taskid']
        

        ## GET TIME TYPES
        #gettimetypes(guid)
        
        ## ADD TIME
        saveworkitem(guid, date, hours, comment, projectid, taskid, args.debug)
    elif args.projects:
        listprojects(guid, args.projects[0])

    print('Done')
