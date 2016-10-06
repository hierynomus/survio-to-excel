import argparse
import requests
from to_excel import parse_respondents_to_excel

parser = argparse.ArgumentParser()
parser.add_argument('-u', metavar='username')
parser.add_argument('-p', metavar='password')
parser.add_argument('-s', metavar='survey_id')
parser.add_argument('-o', metavar="output_file")

args = parser.parse_args()

form = {'username': args.u, 'password': args.p}

session = requests.Session()
login = session.post("https://my.survio.com/login", data=form)
if login.status_code < 200 and login.status_code > 399:
    print("Got status code %s, login unsuccesful" % login.status_code)
    exit(1)

r = session.get("https://my.survio.com/%s/data/view?get_data=1" % args.s, headers={'Accept': 'application/json', 'X-Requested-With': 'XMLHttpRequest'})
if r.status_code > 399:
    print("Survey with id %s not found (HTTP Status code %s)." % (args.s, r.status_code))
    exit(1)

parse_respondents_to_excel(r.json(), args.o)
