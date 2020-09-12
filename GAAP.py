from __future__ import print_function
import pickle
import os.path
import sys
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import cvxpy as cp
import numpy as np

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
FIRSTNAME_RANGE_NAME = 'I3:BI3'
LASTNAME_RANGE_NAME = 'I2:BI2'
STUDENTTYPE_RANGE_NAME = 'I1:BI1'
GAAP_RANGE_NAME = 'D4:D16'
RANKING_RANGE_NAME = 'I4:BI16'
GAAP_COUNT = 13
STUDENT_COUNT = 53

EXTERNES_EXPERIMENTES = [5, 41, 49] # Index des externes ayant plus de 3 ans d'expérience pro
WOMEN = [0, 1, 2, 3, 4, 5, 8, 13, 16, 17, 18, 19, 20, 21, 23, 28, 31, 36, 38, 40, 42, 43, 44, 47, 49] # Index des femmes dans le tableau
MEN = [6, 7, 9, 10, 11, 12, 14, 15, 22, 24, 25, 26, 27, 29, 30, 32, 33, 34, 35, 37, 39, 41, 45, 46, 48, 50, 51, 52] # Index des hommes dans le tableau

def solve(firstnames, lastnames, studenttypes, gaaps, rankings):
    for i in range(GAAP_COUNT):
        for j in range(STUDENT_COUNT):
            if len(rankings[i]) <= j:
                rankings[i].append(1000)
            elif rankings[i][j] == "":
                rankings[i][j] = 1000
            elif rankings[i][j] != "1" and studenttypes[j] == "exterieur": # Ajout d'un coefficient 15 sur les choix secondaires des externes pour prioriser leur premier choix
                rankings[i][j] = 15*int(rankings[i][j])
            else:
                rankings[i][j] = int(rankings[i][j])

    rankings = np.array(rankings)
    groupmax = np.array([5]*GAAP_COUNT)
    groupmin = np.array([4]*GAAP_COUNT)
    externemax = np.array([2]*GAAP_COUNT)
    experimentesmin = np.array([1]*GAAP_COUNT)
    womenmin = np.array([1]*GAAP_COUNT)
    menmin = np.array([1]*GAAP_COUNT)

    selection = cp.Variable(shape=rankings.shape,boolean=True)

    group_constraint_1 = cp.sum(selection,axis=1) <= groupmax

    group_constraint_2 = cp.sum(selection,axis=1) >= groupmin

    externe_filter = np.zeros((STUDENT_COUNT, STUDENT_COUNT))
    for i in range(STUDENT_COUNT):
        if studenttypes[i] == "exterieur":
            externe_filter[i][i] = 1
    externe_constraint = cp.sum((selection @ externe_filter),axis=1) <= externemax

    experimentes_filter = np.zeros((STUDENT_COUNT, STUDENT_COUNT))
    for i in range(STUDENT_COUNT):
        if studenttypes[i] == "IVP" or "Interne" in studenttypes[i] or i in EXTERNES_EXPERIMENTES:
            experimentes_filter[i][i] = 1
    experimentes_constraint = cp.sum((selection @ experimentes_filter),axis=1) >= experimentesmin

    women_filter = np.zeros((STUDENT_COUNT, STUDENT_COUNT))
    for i in range(STUDENT_COUNT):
        if i in WOMEN:
            women_filter[i][i] = 1
    women_constraint = cp.sum((selection @ women_filter),axis=1) >= womenmin

    men_filter = np.zeros((STUDENT_COUNT, STUDENT_COUNT))
    for i in range(STUDENT_COUNT):
        if i in MEN:
            men_filter[i][i] = 1
    men_constraint = cp.sum((selection @ men_filter),axis=1) >= menmin

    assignment_constraint = cp.sum(selection,axis=0) == 1

    cost = cp.sum(cp.multiply(rankings,selection))

    constraints = [group_constraint_1, group_constraint_2, assignment_constraint, externe_constraint, experimentes_constraint, women_constraint, men_constraint]

    assign_prob = cp.Problem(cp.Minimize(cost),constraints)

    assign_prob.solve(solver=cp.GLPK_MI)

    for i in range(GAAP_COUNT):
        print("GAAP " + str(i+1) + " - " + gaaps[i][0])
        for j in range(STUDENT_COUNT):
            if selection.value[i][j] == 1:
                print(firstnames[j] + " " + lastnames[j])

        print()

    selected_rankings = np.multiply(rankings, selection.value)
    unique, counts = np.unique(selected_rankings, return_counts=True)
    occurences = dict(zip(unique, counts))
    for i in range(1, GAAP_COUNT + 1):
        print("Nombre de choix #" + str(i) + " satisfaits : " + str(occurences.get(i, 0) + occurences.get(i*15, 0)))

    print("Nombre de choix non satisfaits : " + str(occurences.get(1000, 0)))


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    spreadsheet = sys.argv[1]

    firstname_result = sheet.values().get(spreadsheetId=spreadsheet,
                                range=FIRSTNAME_RANGE_NAME).execute()
    firstnames = firstname_result.get('values', [])

    lastname_result = sheet.values().get(spreadsheetId=spreadsheet,
                                range=LASTNAME_RANGE_NAME).execute()
    lastnames = lastname_result.get('values', [])

    studenttype_result = sheet.values().get(spreadsheetId=spreadsheet,
                                range=STUDENTTYPE_RANGE_NAME).execute()
    studenttypes = studenttype_result.get('values', [])

    gaap_result = sheet.values().get(spreadsheetId=spreadsheet,
                                range=GAAP_RANGE_NAME).execute()
    gaaps = gaap_result.get('values', [])

    ranking_result = sheet.values().get(spreadsheetId=spreadsheet,
                                range=RANKING_RANGE_NAME).execute()
    rankings = ranking_result.get('values', [])

    solve(firstnames[0], lastnames[0], studenttypes[0], gaaps, rankings)


if __name__ == '__main__':
    main()
