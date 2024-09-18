#!/usr/bin/env python
# -*- coding: utf-8 -*-

# -------------------------------------------------------------------------------------------------------------------------------------------------------- #
#   Main routines for managing Magazine library
# -------------------------------------------------------------------------------------------------------------------------------------------------------- #

from argparse import ArgumentParser, Namespace
import argparse
import re
from database import Database
from report_table import Report
from datetime import datetime, timedelta
import csv
import calendar

oDatabase: Database


def parseCommandLine() -> Namespace:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   Parse the options specified
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    oParserCmdLine: ArgumentParser = argparse.ArgumentParser(description='Magazine Collection Organiser')

    # switches
    oParserCmdLine.add_argument('--update', action='store_const', const=True, default=False, help='Update a record type')
    oParserCmdLine.add_argument('--exclude', action='store_const', const=True, default=False, help='Exclude from reporting')
    oParserCmdLine.add_argument('--summary', action='store_const', const=True, default=False, help='Print the summary tables')

    # data load
    oParserCmdLine.add_argument('-table', type=str, help='Table to load.')
    oParserCmdLine.add_argument('-file', type=str, help='File to load.')

    # updates
    oParserCmdLine.add_argument('-incident', type=str, help='Incident to update.')
    oParserCmdLine.add_argument('-change', type=str, help='Change to update.')
    oParserCmdLine.add_argument('-request', type=str, help='Request to update.')
    oParserCmdLine.add_argument('-notes', type=str, help='Notes to add.')
    oParserCmdLine.add_argument('-service', type=str, help='Incident to update.')

    # reporting
    oParserCmdLine.add_argument('-month', type=int, help='Month to process for report.')
    oParserCmdLine.add_argument('-report', type=int, help='Type of report to produce.')

    return oParserCmdLine.parse_args()


def parseCommentsAndWorkNotes(sText: str) -> list[tuple[datetime, str, str, str]]:
    # Define the regex pattern to match the date, name, update type, and the update content
    rPattern = re.compile(
        r"(?P<datetime>\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2}) - "
        r"(?P<name>[\w\s]+) \((?P<update_type>[\w\s]+)\)\n"
        r"(?P<update>[\s\S]+?)(?=\n\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2} -|$)"
    )

    matches = rPattern.finditer(sText)
    updates = []

    for match in matches:
        # Parse the datetime string into a datetime object
        dt = datetime.strptime(match.group("datetime"), "%d-%m-%Y %H:%M:%S")

        # Get the other fields
        name: str = match.group("name")
        update_type: str = match.group("update_type")
        update: str = match.group("update").strip()

        # Append to the list as a tuple
        updates.append((dt, name, update_type, update))

    # reverse the order of the list
    updates.reverse()

    return updates


lPublicHolidays = [datetime(2024, 4, 25).date(), datetime(2024, 6, 10).date(), datetime(2024, 8, 5).date()]


def calculateBusinessTimeDuration(dtStart: datetime, sDaysHoursMinutes: str) -> datetime:

    # Define business hours
    tdBusinessStart = timedelta(hours=7)
    tdBusinessEnd = timedelta(hours=18)

    sParts = sDaysHoursMinutes.split(':')
    iSLADays = int(sParts[0])
    iSLAHours = int(sParts[1])
    iSLAMinutes = int(sParts[2])

    # if this day is not a business day then it just needs to be advanced to the next business day at start of business day
    if not (dtStart.weekday() < 5 and dtStart.date() not in lPublicHolidays):
        while not (dtStart.weekday() < 5 and dtStart.date() not in lPublicHolidays):  # Monday to Friday, not a public holiday
            dtStart += timedelta(days=1)
        dtStart -= (timedelta(hours=dtStart.hour, minutes=dtStart.minute) - tdBusinessStart)

    # now set the end date
    dtEnd = dtStart

    # if the start time is prior to business hours start then move it forward
    if timedelta(hours=dtStart.hour, minutes=dtStart.minute) < tdBusinessStart:
        dtEnd += (tdBusinessStart - timedelta(hours=dtStart.hour, minutes=dtStart.minute))

    # if the open time is already past the end of business day, then move it forward one day and set the start time
    if timedelta(hours=dtStart.hour, minutes=dtStart.minute) > tdBusinessEnd:
        iSLADays += 1
        dtEnd -= (timedelta(hours=dtStart.hour, minutes=dtStart.minute) - tdBusinessStart)

    # add the minutes and hours
    dtEnd += timedelta(hours=iSLAHours, minutes=iSLAMinutes)

    # now if the end time is before business hours it's because it's wrapped into the next day
    if timedelta(hours=dtEnd.hour, minutes=dtEnd.minute) < tdBusinessStart:
        dtEnd = dtEnd + tdBusinessStart + timedelta(hours=6)

    # now if the end time is past business hours then move it to the next day
    if timedelta(hours=dtEnd.hour, minutes=dtEnd.minute) > tdBusinessEnd:
        iSLADays += 1
        dtEnd -= (tdBusinessEnd - tdBusinessStart)

    # Process days
    while iSLADays > 0:
        dtEnd += timedelta(days=1)
        if dtEnd.weekday() < 5 and dtEnd.date() not in lPublicHolidays:  # Monday to Friday, not a public holiday
            iSLADays -= 1

    return dtEnd


def calculateActualTimeDiff(dtStart: datetime, dtEnd: datetime) -> tuple[int, int, int]:

    tdTotalActualTime = dtEnd - dtStart

    # Calculate days, hours, and minutes
    days = tdTotalActualTime.days
    hours, remainder = divmod(tdTotalActualTime.seconds, 3600)
    minutes = remainder // 60

    return days, hours, minutes


def subtractDuration(tTotal, tSubtract) -> tuple[int, int, int]:
    d1, h1, m1 = tTotal
    d2, h2, m2 = tSubtract

    if m1 < m2:
        m1 += 60
        h1 -= 1

    if h1 < h2:
        h1 += 11
        d1 -= 1

    return d1-d2, h1-h2, m1-m2


def calculateBusinessTimeDiff(dtStart: datetime, dtEnd: datetime) -> tuple[int, int, int]:

    # Define business hours
    tdBusinessStart = timedelta(hours=7)
    tdBusinessEnd = timedelta(hours=18)

    # Initialize total business time difference
    tdTotalBusinessTime = timedelta()

    # Handle the case where the start and end dates are the same
    if dtStart.date() == dtEnd.date():
        if dtStart.weekday() < 5 and dtStart.date() not in lPublicHolidays:  # Monday to Friday
            tdStartTime = max(tdBusinessStart, timedelta(hours=dtStart.hour, minutes=dtStart.minute))
            tdEndTime = min(tdBusinessEnd, timedelta(hours=dtEnd.hour, minutes=dtEnd.minute))
            if tdStartTime < tdEndTime:
                tdTotalBusinessTime = tdEndTime - tdStartTime
    else:
        # Process the first day
        if dtStart.weekday() < 5 and dtStart.date() not in lPublicHolidays:  # Monday to Friday
            tdStartTime = max(tdBusinessStart, timedelta(hours=dtStart.hour, minutes=dtStart.minute))
            tdEndTime = min(tdBusinessEnd, timedelta(hours=18))
            if tdStartTime < tdEndTime:
                tdTotalBusinessTime += tdEndTime - tdStartTime

        # Process intermediate days
        dtNextDay = dtStart + timedelta(days=1)
        while dtNextDay.date() < dtEnd.date():
            if dtNextDay.weekday() < 5 and dtNextDay.date() not in lPublicHolidays:  # Monday to Friday
                # print(f"checking date {dtNextDay} which is day {dtNextDay.weekday()} and date {dtNextDay.date()} public holidays {lPublicHolidays}")
                tdTotalBusinessTime += (tdBusinessEnd - tdBusinessStart)
            dtNextDay += timedelta(days=1)

        # Process the last day
        if dtEnd.weekday() < 5 and dtEnd.date() not in lPublicHolidays:  # Monday to Friday
            tdEndTime = min(tdBusinessEnd, timedelta(hours=dtEnd.hour, minutes=dtEnd.minute))
            if tdBusinessStart < tdEndTime:
                tdTotalBusinessTime += tdEndTime - tdBusinessStart

    # Calculate days, hours, and minutes
    days = tdTotalBusinessTime.days
    hours, remainder = divmod(tdTotalBusinessTime.seconds, 3600)
    minutes = remainder // 60

    rHours = days * 24 + hours + minutes/60
    days = 0
    while rHours > 11:
        days += 1
        rHours -= 11
    hours = int(rHours)
    minutes = round((rHours - hours) * 60)

    return days, int(hours), int(minutes)


def readCSVFile(filePath, encoding='utf-8'):
    # List to store each row as a list
    data_list = []

    try:
        # Open the CSV file with the specified encoding
        with open(filePath, mode='r', newline='', encoding=encoding) as csv_file:
            # Create a CSV reader object
            csv_reader = csv.reader(csv_file)

            # Iterate over each row in the CSV file
            for row in csv_reader:
                # Add it to the list
                data_list.append(row)

    except UnicodeDecodeError as e:
        print(f"Error decoding file: {e}")
        print("Try a different encoding, such as 'ISO-8859-1' or 'latin1'.")
        return []

    return data_list


def convertField(sField: str) -> str:
    datePattern = r"\b(0[1-9]|[12][0-9]|3[01])-(0[1-9]|1[0-2])-(\d{4})\s([01][0-9]|2[0-3]):([0-5][0-9]):([0-5][0-9])\b"
    sExcelDateTimeFormat = '%d-%m-%Y %H:%M:%S'
    if len(sField) == 19 and re.match(datePattern, sField):
        dtField = datetime.strptime(sField, sExcelDateTimeFormat)
        return dtField.strftime('%Y-%m-%d %H:%M:%S')

    datePattern = r"\b([1-9]|[12][0-9]|3[01])\/([1-9]|1[0-2])\/(\d{4})\s([0-9]|1[0-9]|2[0-3]):([0-5][0-9])\b"
    sExcelDateTimeFormat = '%d/%m/%Y %H:%M'
    if len(sField) <= 16 and re.match(datePattern, sField):
        dtField = datetime.strptime(sField, sExcelDateTimeFormat)
        return dtField.strftime('%Y-%m-%d %H:%M:%S')

    return sField


def loadTableToDatabase(sTableName, lDataList):

    dFieldTranslations = {'Incident': dict(number='Number', opened_at='Opened', short_description='ShortDescription', caller_id='Caller',
                                           priority='Priority', description='Description', comments_and_work_notes='CommentsAndWorkNotes', state='State',
                                           category='Category', assignment_group='AssignmentGroup', assigned_to='AssignedTo', sys_updated_on='Updated',
                                           closed_at='Closed', cmdb_ci='ConfigurationItem', sys_created_on='Created', resolved_at='Resolved',
                                           subcategory='Subcategory', business_service='Service', close_code='ResolutionCode', close_notes='ResolutionNotes'),
                          'IncidentSLA': dict(inc_number='Number', taskslatable_sla='SLADefinition', taskslatable_stage='Stage',
                                              taskslatable_start_time='StartTime', taskslatable_end_time='StopTime',
                                              taskslatable_planned_end_time='BreachTime', inc_made_sla='MadeSLA'),
                          'Request': dict(request_item='RequestItem', number='Number', request_item_cat_item='Item', request_item_request_u_requested_by='RequestedBy',
                                          request_item_requested_for='RequestedFor', state='State', priority='Priority', short_description='ShortDescription',
                                          assignment_group='AssignmentGroup', assigned_to='AssignedTo', comments_and_work_notes='CommentsAndWorkNotes', opened_at='Opened',
                                          closed_at='Closed', sys_updated_on='Updated'),
                          'Change': dict(number='Number', type='Type', short_description='ShortDescription', state='State', start_date='PlannedStartDate', end_date='PlannedEndDate',
                                         u_approval_stage='ApprovalStage', assigned_to='AssignedTo', comments='AdditionalComments', assignment_group='AssignmentGroup',
                                         backout_plan='BackoutPlan', category='Category', close_code='CloseCode', closed_at='Closed', comments_and_work_notes='CommentsAndWorkNotes',
                                         sys_created_on='Created', implementation_plan='ImplementationPlan', justification='Justification', risk_impact_analysis='RiskAndImpactAnalysis',
                                         work_start='ActualStartDate', work_end='ActualEndDate', cmdb_ci='ConfigurationItem', business_service='BusinessService', u_environment='Environment',
                                         sys_created_by='CreatedBy')}

    sInsertSQL = f'INSERT INTO {sTableName} ('
    sUpdateSQL = f'UPDATE {sTableName} SET '
    lInsertFields = []
    lUpdateFields = []

    lKeyPredicates = []
    lTableKeyFields = []
    lKeyColumnNumbers = []
    sUpdateField = None
    iUpdateColumnFieldNumber = None

    if sTableName == 'Incident':
        lTableKeyFields = ['Number']
        sUpdateField = 'Updated'
    elif sTableName == 'IncidentSLA':
        lTableKeyFields = ['Number', 'SLADefinition', 'Stage']
        sUpdateField = 'StopTime'
    elif sTableName == 'Request':
        lTableKeyFields = ['Number']
        sUpdateField = 'Updated'
    elif sTableName == 'Change':
        lTableKeyFields = ['Number']
        sUpdateField = ''

    tFieldNames = lDataList[0]
    for iField, sField in enumerate(tFieldNames):
        sDBFieldName = dFieldTranslations[sTableName][sField.replace('.', '_')]
        lInsertFields.append(sDBFieldName)
        if sDBFieldName in lTableKeyFields:
            lKeyColumnNumbers.append(iField)
            lKeyPredicates.append(f'{sDBFieldName} = ?')
        else:
            lUpdateFields.append(f'{sDBFieldName} = ?')

        if sDBFieldName == sUpdateField:
            iUpdateColumnFieldNumber = iField

    if iUpdateColumnFieldNumber:
        sFindSQL = f"SELECT {sUpdateField} FROM {sTableName} WHERE {' AND '.join(lKeyPredicates)}"
    else:
        sFindSQL = f"SELECT {', '.join(lTableKeyFields)} FROM {sTableName} WHERE {' AND '.join(lKeyPredicates)}"

    sCountSQL = f"SELECT COUNT(*) FROM {sTableName} WHERE {' AND '.join(lKeyPredicates)}"
    sInsertSQL += (', '.join(lInsertFields) + ') VALUES (' + '?, ' * (len(lInsertFields) - 1) + '?)')

    # print(sFindSQL)
    # print(sInsertSQL)

    sUpdateSQL += (', '.join(lUpdateFields) + ' WHERE ' + ' AND '.join(lKeyPredicates))
    # print(sUpdateSQL)

    iInsertCount = 0
    iUpdateCount = 0
    iIgnoreCount = 0

    oUpdateField = None

    for iRow, lData in enumerate(lDataList):
        # ignore the first row as it's the field names
        if iRow > 0:
            # convert all the date fields in the list
            lFields = list(convertField(field) for field in lData)

            tKeyFields = tuple(field for iCtr, field in enumerate(lFields) if iCtr in lKeyColumnNumbers)

            iCount = oDatabase.fetchValue(sCountSQL, tKeyFields)
            oField = oDatabase.fetchValue(sFindSQL, tKeyFields)

            if iUpdateColumnFieldNumber:
                oUpdateField = lFields[iUpdateColumnFieldNumber]

            if iCount == 0:
                # create the tuple for the insert statement
                tInsertFields = tuple(lFields)
                # print(f"Field value in row {iRow} is {tKeyFields}, needs to be inserted into database.")

                oDatabase.execute(sInsertSQL, tInsertFields)
                iInsertCount += 1
            elif oField == oUpdateField:
                # print(f"Field value in row {iRow} is {tKeyFields}, data is the same")
                iIgnoreCount += 1
            else:
                # print(f"Field value in row {iRow} is {tKeyFields}, value is {oField}, new value is {oUpdateField}.")
                tUpdateFields = (tuple(field for iCtr, field in enumerate(lFields) if iCtr not in lKeyColumnNumbers))
                tUpdateFields += tKeyFields

                oDatabase.execute(sUpdateSQL, tUpdateFields)
                iUpdateCount += 1

    print(f'Rows inserted - {iInsertCount}')
    print(f'Rows updated  - {iUpdateCount}')
    print(f'Rows ignored  - {iIgnoreCount}')

    if oArgs.table in ['Incident', 'Request']:
        sSQL = f"UPDATE {oArgs.table} SET ReportPriority = Priority WHERE ReportPriority IS NULL"
        oDatabase.execute(sSQL, ())


def mapService(sService) -> str:
    dServices = {'Compute Infrastructure': "Compute", "Storage Network": "Network", "Data Storage": "Storage", "Data Protection": "Protect", "Unknown": "[yellow]??", "Other": "Other"}
    return dServices.get(sService, f"[orange3]{sService}?")


def mapPerson(sPerson: str) -> str:
    if sPerson:
        return sPerson.split(' ')[0]
    return '-'


def mapLongField(sText, iMax) -> str:
    if len(sText) > iMax:
        return sText[0:iMax-2] + '…'
    return sText


def stripColours(sText) -> str:
    return sText.replace('[green]', '').replace('[red]', '').replace('[orange3]', '').replace('[yellow]', '').replace('[white]', '')


def ServiceNOWReports():
    dtStartOfMonth = dtEndOfMonth = None
    lTeam_HV = ['Shagufta Anjum Shaik', 'Marwa Elshawy', 'Siddhartha Dutta', 'Mathieu Doumerc', "Keith D'Souza", 'Wilson Lee', 'Wayne Moss']
    lTeam_RBA = ['Antonio Salmeri', 'Max Sekula']
    sDateTimeFormat = '%Y-%m-%d %H:%M:%S'

    if oArgs.month:
        iLastDay = calendar.monthrange(2024, oArgs.month)[1]
        dtStartOfMonth = datetime.strptime(f"2024-{oArgs.month}-01 00:00:00", sDateTimeFormat)
        dtEndOfMonth = datetime.strptime(f"2024-{oArgs.month}-{iLastDay} 23:59:59", sDateTimeFormat)

    if oArgs.report == 1:
        lRespondedCount = [0, 0, 0, 0, 0]
        lResolvedCount = [0, 0, 0, 0, 0]

        # ---------- Incident response and resolution------------
        sSQL = ("SELECT ReportingService, Number, State, Priority, ReportPriority, AssignedTo, ShortDescription, Opened, Resolved, CommentsAndWorkNotes, Notes, Caller "
                "FROM Incident WHERE ReportingService <> 'Other' AND Exclude = ? AND State <> 'Cancelled' AND Updated > ? ORDER BY ReportPriority, Number, ReportingService")

        oReport = Report(f'Incident Response and Resolution')
        oReport.addColumn('#', sJust='right')
        oReport.addColumn('Service', sJust='left')
        oReport.addColumn('Reference', sJust='left')
        oReport.addColumn('P', sJust='centre', bBreak=True)
        oReport.addColumn('Caller', sJust='left')
        oReport.addColumn('Assigned To', sJust='left')
        oReport.addColumn('State', sJust='left', bShow=False)
        oReport.addColumn('Description', sJust='left')
        oReport.addColumn('Opened', sJust='left')
        oReport.addColumn('Received', sJust='left')
        oReport.addColumn('Target', sJust='left')
        oReport.addColumn('Responded', sJust='left')
        oReport.addColumn('Resp. Dur', sJust='left')
        oReport.addColumn('Target', sJust='left')
        oReport.addColumn('Resolved', sJust='left')
        oReport.addColumn('Rslv. Dur', sJust='left')
        oReport.addColumn('Notes', sJust='left')

        oReportXL = Report(f'Excel Report for Incident Response & Resolution')
        oReportXL.addColumn('Priority', sJust='centre', bBreak=True)
        oReportXL.addColumn('Service', sJust='left')
        oReportXL.addColumn('Reference', sJust='left')
        oReportXL.addColumn('Description', sJust='left')
        oReportXL.addColumn('Opened', sJust='left')
        oReportXL.addColumn('Received', sJust='left')
        oReportXL.addColumn('Response Target', sJust='left')
        oReportXL.addColumn('Responded', sJust='left')
        oReportXL.addColumn('Response Duration d:hh:mm', sJust='left')
        oReportXL.addColumn('Resolve Target', sJust='left')
        oReportXL.addColumn('Resolved', sJust='left')
        oReportXL.addColumn('Resolve Duration d:hh:mm', sJust='left')
        oReportXL.addColumn('Notes', sJust='left')

        tData = (1 if oArgs.exclude else 0, dtStartOfMonth)
        lResults = oDatabase.fetchList(sSQL, tData)
        iCtr = 0

        lIncidentTargetResponseTimes = ['', '0:00:10', '0:00:30', '0:01:00', '0:04:00', '0:08:00']
        lIncidentTargetResolveTimes = ['', '0:01:00', '0:04:00', '0:10:00', '5:00:00', '20:00:00']

        sPrevPriority = ' '

        for tResult in lResults:
            sService, sNumber, sState, sPriority, sReportPriority, sAssignedTo, sDescription, sOpened, sResolved, sCommentsAndWorkNotes, sNotes, sCaller = tResult

            if not sNotes:
                sNotes = ''

            dtOpened = datetime.strptime(sOpened, sDateTimeFormat)
            dtResolved = datetime.strptime(sResolved, sDateTimeFormat) if sResolved else None
            sResolveTime = None
            sResolveSuffix = ''

            # don't bother with any incident that was resolved prior to the month start
            if dtResolved and dtResolved < dtStartOfMonth:
                continue

            iOrigPriority = None

            iCtr = 1 if sReportPriority[0] != sPrevPriority[0] else iCtr + 1

            sOverrideSQL = "SELECT StartTime, EndTime FROM SLAOverride WHERE Number = ? AND SLA = 'Response'"
            sResponseSQL = "SELECT StartTime, StopTime FROM IncidentSLA WHERE Number = ? AND SLADefinition = ?"
            sSQLAnyResponse = "SELECT SLADefinition, StartTime, StopTime FROM IncidentSLA WHERE Number = ? AND SLADefinition LIKE '%Response'"
            sSQLAnyResolve = "SELECT StartTime, StopTime FROM IncidentSLA WHERE Number = ? AND SLADefinition LIKE '%Resolution%'"

            sResponded = "[yellow]      N/A"
            dtResponded = None
            sRespondSuffix = ''

            sRespondDuration = sResolveDuration = "[yellow]   N/A"

            # Work out the first touch time by someone in the team, and save the time before that it was touched, as that is when the SLA should start
            dtFirstTeamTouch = None
            dtReceiveNonTeamTouch = None

            lOtherTeamTransfers = []
            dtTeamTouch = dtOtherTeamTouch = None
            for dtUpdate, sPerson, _sUpdateType, _sUpdateContent in parseCommentsAndWorkNotes(sCommentsAndWorkNotes):
                if sPerson in lTeam_HV:
                    if dtFirstTeamTouch is None:
                        # this is the timestamp at which we have received the incident
                        dtFirstTeamTouch = dtUpdate

                    if dtOtherTeamTouch:
                        # another team has been working on the incident, this time needs to be excluded
                        lOtherTeamTransfers.append((dtTeamTouch, dtOtherTeamTouch))
                        dtOtherTeamTouch = None
                    dtTeamTouch = dtUpdate

                elif sPerson not in lTeam_RBA:
                    if dtFirstTeamTouch is None:
                        # this is potentially the time as which we have received the incident, i.e. if helpdesk are updating it
                        dtReceiveNonTeamTouch = dtUpdate
                    else:
                        # ITIMS HV team have updated the incident, now it's being updated by other teams, this incident has been transferred
                        dtOtherTeamTouch = dtUpdate

            # check to see whether a team member created the ticket, if so then it's an instance response and SLA met
            if sCaller in lTeam_HV:
                dtReceived = dtResponded = dtOpened
                sReceivedSuffix = sRespondSuffix = '[white] (Open)'
            elif dtReceiveNonTeamTouch:
                # check if someone outside of the ITIMS team touched the ticket and use that timestamp as the receive and SLA start time
                dtReceived = dtReceiveNonTeamTouch
                sReceivedSuffix = '[purple] (Assigned)'
                # in which case the time in which we touched the case next is the SLA stop time
                if dtFirstTeamTouch:
                    dtResponded = dtFirstTeamTouch
                    sRespondSuffix = '[orange3] (Touch)'

            else:
                # get the actual SLA record as reported by ServiceNOW
                lSLAResults = oDatabase.fetchList(sResponseSQL, (sNumber, f"P{sPriority[0]} Response"))
                if lSLAResults:
                    sSLAStartTime, sSLAStopTime = lSLAResults[0]
                    sReceivedSuffix = sRespondSuffix = ' (SLA)'
                else:
                    # maybe there is a response under the original severity
                    lSLAResults = oDatabase.fetchList(sSQLAnyResponse, (sNumber,))
                    if lSLAResults:
                        sSLADefinition, sSLAStartTime, sSLAStopTime = lSLAResults[0]
                        sReceivedSuffix = sRespondSuffix = f'[orange3] ({sSLADefinition}) '
                        iOrigPriority = int(sSLADefinition[1])
                    else:
                        # get the SLA start time for the restore metric, this will most likely be the same time
                        lSLAResults = oDatabase.fetchList(sSQLAnyResolve, (sNumber,))
                        if lSLAResults:
                            sSLAStartTime, sSLAStopTime = lSLAResults[0]
                            sReceivedSuffix = '[orange3] (Resolve Start)'
                        else:
                            # last resort, use the open time, and the restore time, these might be good enough
                            sSLAStartTime = sOpened
                            sSLAStopTime = sResolved
                            sReceivedSuffix = '[yellow] (Open)'
                            sRespondSuffix = '[yellow] (Resolved)'
                dtReceived = datetime.strptime(sSLAStartTime, sDateTimeFormat)
                if sSLAStopTime:
                    dtResponded = datetime.strptime(sSLAStopTime, sDateTimeFormat)

            # check if there's an SLA override
            lSLAResults = oDatabase.fetchList(sOverrideSQL, (sNumber,))
            if lSLAResults:
                sStart, sEnd = lSLAResults[0]
                if sStart:
                    dtReceived = datetime.strptime(sStart, sDateTimeFormat)
                    sReceivedSuffix = '[pink3] (o/ride)'
                if sEnd:
                    dtResponded = datetime.strptime(sEnd, sDateTimeFormat)
                    sRespondSuffix = '[pink3] (o/ride)'

            sResponseTarget = lIncidentTargetResponseTimes[iOrigPriority if iOrigPriority else int(sReportPriority[0])]

            # if the ticket has been received after the end of the month, then don't include
            if dtReceived and dtReceived > dtEndOfMonth:
                continue

            sReceived = dtReceived.strftime('%a %-d/%m %H:%M') + sReceivedSuffix

            if dtResponded:
                sResponded = dtResponded.strftime('%a %-d/%m %H:%M') + sRespondSuffix

                iDays, iHours, iMinutes = calculateBusinessTimeDiff(dtReceived, dtResponded)
                sRespondTime = f"{iDays:2}:{iHours:02}:{iMinutes:02}"

                dtSLABreachTime = calculateBusinessTimeDuration(dtReceived, sResponseTarget)

                if dtResponded > dtSLABreachTime:
                    sRespondDuration = '[red]' + sRespondTime
                else:
                    sRespondDuration = '[green]' + sRespondTime

            if sCaller in lTeam_HV:
                sCaller = f"[yellow]{sCaller.split(' ')[0]}"

            if sAssignedTo in lTeam_RBA:
                sAssignedTo = f"[yellow]{sAssignedTo}"

            # check if any are not applicable to be reported in the month
            if dtResponded and dtResponded < dtStartOfMonth:
                sResponseTarget = "[yellow] N/A"
                sRespondDuration = "[yellow]  N/A"

            # ----- now work out the resolve time --------
            sResolveTarget = lIncidentTargetResolveTimes[int(sReportPriority[0])]

            sOverrideSQL = "SELECT EndTime FROM SLAOverride WHERE Number = ? AND SLA = 'Resolve'"

            # check if there's an SLA override
            lSLAResults = oDatabase.fetchList(sOverrideSQL, (sNumber,))
            if lSLAResults:
                sEnd = lSLAResults[0][0]
                print(f"Incident {sNumber} - sEnd {sEnd}")
                dtResolved = datetime.strptime(sEnd, sDateTimeFormat)
                sResolveSuffix = '[pink3] (o/ride)'

            if dtResolved and dtResolved < dtEndOfMonth:
                iDays, iHours, iMinutes = calculateBusinessTimeDiff(dtReceived, dtResolved)
                sResolveTime = f"{iDays:2}:{iHours:02}:{iMinutes:02}"

                dtSLABreachTime = calculateBusinessTimeDuration(dtReceived, sResolveTarget)

                if dtResolved > dtSLABreachTime:
                    if lOtherTeamTransfers:
                        sResolveDuration = '[orange3]' + sResolveTime
                    else:
                        sResolveDuration = '[red]' + sResolveTime
                else:
                    sResolveDuration = '[green]' + sResolveTime

            if dtResolved:
                sResolved = dtResolved.strftime('%a %-d/%m %H:%M')
                if dtResolved > dtEndOfMonth:
                    sResolved = f"[yellow]{sResolved}"
            else:
                sResolved = '[yellow]   N/A'

            if sResolveDuration == '[yellow]   N/A':
                sResolveTarget = '[yellow]   N/A'

            sTransfers = ''
            # ------------- If there are multiple transfers then this won't work properly
            if lOtherTeamTransfers and dtResolved:
                tDuration = calculateBusinessTimeDiff(dtReceived, dtResolved)
                sTransferTime = None
                tTimeDiff = None
                for tTransfer in lOtherTeamTransfers:
                    sTransfers += f"{tTransfer[0].strftime('%a %-d/%m %H:%M')} to {tTransfer[1].strftime('%a %-d/%m %H:%M')}"
                    tTimeDiff = calculateBusinessTimeDiff(tTransfer[0], tTransfer[1])
                    sTransferTime = f"{tTimeDiff[0]:2}:{tTimeDiff[1]:02}:{tTimeDiff[2]:02}"

                tActualTeamDuration = subtractDuration(tDuration, tTimeDiff)
                sActualTime = f"{tActualTeamDuration[0]:2}:{tActualTeamDuration[1]:02}:{tActualTeamDuration[2]:02}"
                # sResolveDuration = f"[orange3]{sResolveTime} - {sTransferTime} = {sActualTime}"
                sTransfers = f"[white]{sTransfers}[cyan] "

            oReport.addCells(iCtr, mapService(sService), sNumber, sReportPriority[0], sCaller, mapPerson(sAssignedTo), sState, mapLongField(sDescription, 80),
                             dtOpened.strftime('%a %-d/%m %H:%M'), sReceived, sResponseTarget, sResponded, sRespondDuration,
                             sResolveTarget, sResolved + sResolveSuffix, sResolveDuration, sTransfers + sNotes)
            oReport.addRow()

            oReportXL.addCells(sReportPriority[0], sService, sNumber, sDescription, dtOpened, dtReceived, stripColours(sResponseTarget), dtResponded, stripColours(sRespondDuration),
                               stripColours(sResolveTarget), dtResolved, stripColours(sResolveDuration), stripColours(sTransfers + sNotes))
            oReportXL.addRow()

            sPrevPriority = sReportPriority

            iIndex = int(sReportPriority[0]) - 1
            if stripColours(sRespondDuration).strip() != 'N/A':
                lRespondedCount[iIndex] += 1
            if stripColours(sResolveDuration).strip() != 'N/A':
                lResolvedCount[iIndex] += 1

        oReport.printReport()
        print('')
        if oArgs.summary:
            oReportXL.printReport()
            oReportXL.sendToClipboard()

        print(f"Respond counts {lRespondedCount}")
        print(f"Resolve counts {lResolvedCount}")

    if oArgs.report == 2:
        lRespondedCount = [0, 0, 0, 0, 0]
        lResolvedCount = [0, 0, 0, 0, 0]

        # ---------- Request response and resolution------------
        sSQL = ("SELECT Service, Number, RequestItem, Priority, ReportPriority, AssignedTo, ShortDescription, Opened, Closed, CommentsAndWorkNotes, Notes, RequestedBy "
                "FROM Request WHERE Service <> 'Other' AND Exclude = ? AND Updated > ? ORDER BY ReportPriority, Number, Service")

        oReport = Report(f'Request Response and Resolution')
        oReport.addColumn('#', sJust='right')
        oReport.addColumn('Service', sJust='left')
        oReport.addColumn('Reference', sJust='left')
        oReport.addColumn('P', sJust='centre', bBreak=True)
        oReport.addColumn('Requested By', sJust='left')
        oReport.addColumn('Assigned To', sJust='left')
        oReport.addColumn('Description', sJust='left')
        oReport.addColumn('Opened', sJust='left')
        oReport.addColumn('Received', sJust='left')
        oReport.addColumn('Target', sJust='left')
        oReport.addColumn('Responded', sJust='left')
        oReport.addColumn('Resp. Dur', sJust='left')
        oReport.addColumn('Target', sJust='left')
        oReport.addColumn('Resolved', sJust='left')
        oReport.addColumn('Rslv. Dur', sJust='left')
        oReport.addColumn('Notes', sJust='left')

        oReportXL = Report(f'Excel Request Response and Fulfillment')
        oReportXL.addColumn('Priority', sJust='centre', bBreak=True)
        oReportXL.addColumn('Service', sJust='left')
        oReportXL.addColumn('Reference', sJust='left')
        oReportXL.addColumn('Description', sJust='left')
        oReportXL.addColumn('Opened', sJust='left')
        oReportXL.addColumn('Received', sJust='left')
        oReportXL.addColumn('Response Target', sJust='left')
        oReportXL.addColumn('Responded', sJust='left')
        oReportXL.addColumn('Respond Duration d:hh:mm', sJust='left')
        oReportXL.addColumn('Fulfill Target', sJust='left')
        oReportXL.addColumn('Fulfilled', sJust='left')
        oReportXL.addColumn('Fulfillment Duration d:hh:mm', sJust='left')

        tData = (1 if oArgs.exclude else 0, dtStartOfMonth)
        lResults = oDatabase.fetchList(sSQL, tData)
        iCtr = 0

        lRequestTargetResponseTimes = ['', '0:00:10', '0:00:30', '0:01:00', '0:04:00', '0:08:00']
        lRequestTargetResolveTimes = ['', '0:02:00', '1:00:00', '5:00:00', '10:00:00', '260:00:00']

        sPrevPriority = ' '

        for tResult in lResults:
            sService, sNumber, sReqItem, sPriority, sReportPriority, sAssignedTo, sDescription, sOpened, sResolved, sCommentsAndWorkNotes, sNotes, sCaller = tResult

            if not sNotes:
                sNotes = ''

            dtOpened = datetime.strptime(sOpened, sDateTimeFormat)
            dtResolved = datetime.strptime(sResolved, sDateTimeFormat) if sResolved else None

            # don't bother with any incident that was resolved prior to the month start
            if dtResolved and dtResolved < dtStartOfMonth:
                continue

            iOrigPriority = None

            iCtr = 1 if sReportPriority[0] != sPrevPriority[0] else iCtr + 1

            sOverrideSQL = "SELECT StartTime, EndTime FROM SLAOverride WHERE Number = ? AND SLA = 'Response'"

            sResponded = "[yellow]      N/A"
            dtResponded = None
            sRespondSuffix = ''

            dtReceived = None
            sReceivedSuffix = ''

            sRespondDuration = sResolveDuration = "[yellow]   N/A"

            # Work out the first touch time by someone in the team, and save the time before that it was touched, as that is when the SLA should start
            dtFirstTeamTouch = None
            dtReceiveNonTeamTouch = None

            lOtherTeamTransfers = []
            dtTeamTouch = dtOtherTeamTouch = None
            for dtUpdate, sPerson, _sUpdateType, _sUpdateContent in parseCommentsAndWorkNotes(sCommentsAndWorkNotes):
                if sPerson in lTeam_HV:
                    if dtFirstTeamTouch is None:
                        # this is the timestamp at which we have received the incident
                        dtFirstTeamTouch = dtUpdate

                    if dtOtherTeamTouch:
                        # another team has been working on the incident, this time needs to be excluded
                        lOtherTeamTransfers.append((dtTeamTouch, dtOtherTeamTouch))
                        dtOtherTeamTouch = None
                    dtTeamTouch = dtUpdate

                elif sPerson not in lTeam_RBA:
                    if dtFirstTeamTouch is None:
                        # this is potentially the time as which we have received the incident, i.e. if helpdesk are updating it
                        dtReceiveNonTeamTouch = dtUpdate
                    else:
                        # ITIMS HV team have updated the incident, now it's being updated by other teams, this incident has been transferred
                        dtOtherTeamTouch = dtUpdate

            # check to see whether a team member created the ticket, if so then it's an instance response and SLA met
            if sCaller in lTeam_HV:
                dtReceived = dtResponded = dtOpened
                sReceivedSuffix = sRespondSuffix = '[white] (Open)'
            elif dtReceiveNonTeamTouch:
                # check if someone outside of the ITIMS team touched the ticket and use that timestamp as the receive and SLA start time
                dtReceived = dtReceiveNonTeamTouch
                sReceivedSuffix = '[purple] (Assigned)'
                # in which case the time in which we touched the case next is the SLA stop time
                if dtFirstTeamTouch:
                    dtResponded = dtFirstTeamTouch
                    sRespondSuffix = '[orange3] (Touch)'

            if dtReceived is None:
                dtReceived = dtOpened
                sReceivedSuffix = '[white] (Open)'
                dtResponded = dtFirstTeamTouch
                sRespondSuffix = '[orange3] (Touch)'

            # check if there's an SLA override
            lSLAResults = oDatabase.fetchList(sOverrideSQL, (sNumber,))
            if lSLAResults:
                sStart, sEnd = lSLAResults[0]
                if sStart:
                    dtReceived = datetime.strptime(sStart, sDateTimeFormat)
                    sReceivedSuffix = '[pink3] (o/ride)'
                if sEnd:
                    dtResponded = datetime.strptime(sEnd, sDateTimeFormat)
                    sRespondSuffix = '[pink3] (o/ride)'

            sResponseTarget = lRequestTargetResponseTimes[iOrigPriority if iOrigPriority else int(sReportPriority[0])]

            # if the ticket has been received after the end of the month, then don't include
            if dtReceived and dtReceived > dtEndOfMonth:
                continue

            sReceived = dtReceived.strftime('%a %-d/%m %H:%M') + sReceivedSuffix

            if dtResponded:
                sResponded = dtResponded.strftime('%a %-d/%m %H:%M') + sRespondSuffix

                iDays, iHours, iMinutes = calculateBusinessTimeDiff(dtReceived, dtResponded)
                sRespondTime = f"{iDays:2}:{iHours:02}:{iMinutes:02}"

                dtSLABreachTime = calculateBusinessTimeDuration(dtReceived, sResponseTarget)

                if dtResponded > dtSLABreachTime:
                    sRespondDuration = '[red]' + sRespondTime
                else:
                    sRespondDuration = '[green]' + sRespondTime

            if sCaller in lTeam_HV:
                sCaller = f"[yellow]{sCaller.split(' ')[0]}"

            if sAssignedTo in lTeam_RBA:
                sAssignedTo = f"[yellow]{sAssignedTo}"

            # check if any are not applicable to be reported in the month
            if dtReceived and dtReceived < dtStartOfMonth:
                sResponseTarget = "[yellow] N/A"
                sRespondDuration = "[yellow]  N/A"

            if dtReceived < dtStartOfMonth:
                sReceived = f"[yellow]{sReceived}"
                sResponded = f"[yellow]{sResponded}"

            # ----- now work out the resolve time --------
            sResolveTarget = lRequestTargetResolveTimes[int(sReportPriority[0])]

            if dtResolved and dtResolved < dtEndOfMonth:
                iDays, iHours, iMinutes = calculateBusinessTimeDiff(dtReceived, dtResolved)
                sResolveTime = f"{iDays:2}:{iHours:02}:{iMinutes:02}"

                dtSLABreachTime = calculateBusinessTimeDuration(dtReceived, sResolveTarget)

                if dtResolved > dtSLABreachTime:
                    if lOtherTeamTransfers:
                        sResolveDuration = '[orange3]' + sResolveTime
                    else:
                        sResolveDuration = '[red]' + sResolveTime
                else:
                    sResolveDuration = '[green]' + sResolveTime

            if dtResolved:
                sResolved = dtResolved.strftime('%a %-d/%m %H:%M')
                if dtResolved > dtEndOfMonth:
                    sResolved = f"[yellow]{sResolved}"
            else:
                sResolved = '[yellow]   N/A'

            if sResolveDuration == '[yellow]   N/A':
                sResolveTarget = '[yellow]   N/A'

            sTransfers = ''
            if lOtherTeamTransfers:
                for tTransfer in lOtherTeamTransfers:
                    sTransfers += f"{tTransfer[0].strftime('%a %-d/%m %H:%M')} to {tTransfer[1].strftime('%a %-d/%m %H:%M')}"
                sTransfers = f"[white]{sTransfers}[cyan] "

            oReport.addCells(iCtr, mapService(sService), sNumber, sReportPriority[0], sCaller, mapPerson(sAssignedTo), mapLongField(sDescription, 80),
                             dtOpened.strftime('%a %-d/%m %H:%M'), sReceived, sResponseTarget, sResponded, sRespondDuration,
                             sResolveTarget, sResolved, sResolveDuration, sTransfers + sNotes)
            oReport.addRow()

            oReportXL.addCells(sReportPriority[0], sService, sNumber, sDescription, dtOpened, dtReceived, stripColours(sResponseTarget), dtResponded, stripColours(sRespondDuration),
                               stripColours(sResolveTarget), dtResolved, stripColours(sResolveDuration))
            oReportXL.addRow()

            sPrevPriority = sReportPriority

            iIndex = int(sReportPriority[0]) - 1
            if stripColours(sRespondDuration).strip() != 'N/A':
                lRespondedCount[iIndex] += 1
            if stripColours(sResolveDuration).strip() != 'N/A':
                lResolvedCount[iIndex] += 1

        oReport.printReport()
        print("")

        if oArgs.summary:
            oReportXL.printReport()
            print("")
            oReportXL.sendToClipboard()

        print(f"Respond counts {lRespondedCount}")
        print(f"Resolve counts {lResolvedCount}")

    if oArgs.report == 3:
        # ---------- Incidents outstanding ------------
        sSQL = ("SELECT ReportingService, Number, State, ReportPriority, ShortDescription, Opened , Resolved FROM Incident WHERE "
                "ReportingService <> 'Other' AND Exclude = 0 ORDER BY ReportingService, Number")

        oReport = Report(f'Incidents still open at the end of month.')
        oReport.addColumn('#', sJust='right')
        oReport.addColumn('Service', sJust='left')
        oReport.addColumn('Reference', sJust='left')
        oReport.addColumn('Priority', sJust='centre')
        oReport.addColumn('Summary', sJust='left')
        oReport.addColumn('Received', sJust='left')

        lResults = oDatabase.fetchList(sSQL, ())
        iCtr = 0
        for tResult in lResults:
            sService, sNumber, sState, sReportPriority, sDescription, sOpened, sResolved = tResult

            if sState in ('Closed', 'Resolved'):
                dtResolved = datetime.strptime(sResolved, sDateTimeFormat)
            else:
                dtResolved = None

            dtOpened = datetime.strptime(sOpened, sDateTimeFormat)

            if dtStartOfMonth <= dtOpened <= dtEndOfMonth and (dtResolved is None or dtResolved > dtEndOfMonth) and sState != 'Cancelled':
                iCtr += 1

                oReport.addCells(iCtr, sService, sNumber, sReportPriority[0], sDescription, dtOpened.strftime('%-d/%m/%Y %H:%M'))
                oReport.addRow()

        oReport.showColumn(0, False)
        oReport.printReport()
        print("")
        oReport.sendToClipboard()

    if oArgs.report == 4:
        # ---------- Changes delivered ------------
        sSQL = ("SELECT Service, Number, ShortDescription, Category, PlannedStartDate, CloseCode, Closed, AssignedTo FROM Change WHERE "
                "Service <> 'Other' AND Exclude = 0 AND State = 'Closed' ORDER BY Service, Number")

        oReport = Report(f'Changes delivered in the month.')
        oReport.addColumn('#', sJust='right')
        oReport.addColumn('Service', sJust='left', bBreak=True)
        oReport.addColumn('Reference', sJust='left')
        oReport.addColumn('Description', sJust='left')
        oReport.addColumn('Assigned To', sJust='left')
        oReport.addColumn('Type', sJust='left')
        oReport.addColumn('Close Code', sJust='left')
        oReport.addColumn('Scheduled Date', sJust='left')
        oReport.addColumn('Date Completed', sJust='left')

        oReportXL = Report(f'Excel report for changes delivered in the month.')
        oReportXL.addColumn('Service', sJust='left', bBreak=True)
        oReportXL.addColumn('Reference', sJust='left')
        oReportXL.addColumn('Description', sJust='left')
        oReportXL.addColumn('Logical/Physical', sJust='left')
        oReportXL.addColumn('Scheduled Date', sJust='left')
        oReportXL.addColumn('Date Completed', sJust='left')
        oReportXL.addColumn('Completion Status', sJust='left')

        lResults = oDatabase.fetchList(sSQL, ())
        iCtr = 0
        for tResult in lResults:
            sService, sNumber, sDescription, sCategory, sPlannedStart, sClosedCode, sClosed, sAssignedTo = tResult

            dtPlanned = datetime.strptime(sPlannedStart, sDateTimeFormat)
            dtClosed = datetime.strptime(sClosed, sDateTimeFormat)

            if dtStartOfMonth <= dtClosed <= dtEndOfMonth:
                iCtr += 1

                if sCategory in ['Software', 'Service']:
                    sCategory = 'Logical'
                elif sCategory == 'Hardware':
                    sCategory = 'Physical'
                else:
                    sCategory = f"[orange3]{sCategory}"

                oReport.addCells(iCtr, mapService(sService), sNumber, sDescription, mapPerson(sAssignedTo), sCategory, sClosedCode, dtPlanned.strftime('%-d/%m/%Y %H:%M'),
                                 dtClosed.strftime('%-d/%m/%Y %H:%M'))
                oReport.addRow()

                oReportXL.addCells(sService, sNumber, sDescription, sCategory, dtPlanned, dtClosed, 'ü')
                oReportXL.addRow()

        oReport.printReport()
        print('')
        if oArgs.summary:
            oReportXL.printReport()
            oReportXL.sendToClipboard()
            print("")

    if oArgs.report == 7:
        # ---------- History of ticket based on comments and work notes ------------
        if oArgs.incident is None and oArgs.request is None:
            print('No incident or request number was supplied!')
            exit(1)

        if oArgs.incident:
            sSQL = "SELECT Created, Resolved, Closed, Caller, CommentsAndWorkNotes, Description FROM Incident WHERE Number = ?"
            tData = (oArgs.incident, )
        else:
            sSQL = "SELECT Opened, Closed, Closed, RequestedBy, CommentsAndWorkNotes, ShortDescription FROM Request WHERE Number = ?"
            tData = (oArgs.request, )

        oReport = Report(f'Ticket History.')
        oReport.addColumn('TimeStamp', sJust='right')
        oReport.addColumn('Action', sJust='left')
        oReport.addColumn('Person', sJust='left')
        oReport.addColumn('Comment', sJust='left')

        lResults = oDatabase.fetchList(sSQL, tData)

        if len(lResults) != 1:
            print("Incident or request was not found!")
            exit(1)

        sOpened, sResolved, sClosed, sCaller, sCommentsAndWorkNotes, sDescription = lResults[0]

        dtOpened = datetime.strptime(sOpened, sDateTimeFormat)

        lComments = parseCommentsAndWorkNotes(sCommentsAndWorkNotes)

        oReport.addCells(dtOpened.strftime('%a %-d/%m %H:%M'), 'Opened', sCaller, sDescription)
        oReport.addRow(True)
        for tComment in lComments:
            oReport.addCells(tComment[0].strftime('%a %-d/%m %H:%M'), tComment[2], tComment[1], tComment[3])
            oReport.addRow(True)

        if sResolved:
            if oArgs.incident:
                dtResolved = datetime.strptime(sResolved, sDateTimeFormat)
                oReport.addCells(dtResolved.strftime('%a %-d/%m %H:%M'), 'Resolved', '', '')
                oReport.addRow(True)

            if sClosed:
                dtClosed = datetime.strptime(sClosed, sDateTimeFormat)
                oReport.addCells(dtClosed.strftime('%a %-d/%m %H:%M'), 'Closed', '', '')
                oReport.addRow(True)

        oReport.printReport()
        print('')


def updateTableEntry():
    if oArgs.incident:
        if oDatabase.fetchValue("SELECT COUNT(*) FROM Incident WHERE Number = ?", (oArgs.incident,)) == 0:
            print("Incident does not exist!")
            exit(1)

        if oArgs.notes:
            sNotes: str = oDatabase.fetchValue("SELECT Notes FROM Incident WHERE Number = ?", (oArgs.incident,))
            if oArgs.notes.startswith('+'):
                oArgs.notes = f"{sNotes} {oArgs.notes[1:]}"
            oDatabase.execute("UPDATE Incident SET Notes = ? WHERE Number = ?", (oArgs.notes, oArgs.incident))
        if oArgs.service:
            if oArgs.service.lower() in ['ds', 'dp', 'ci', 'sn']:
                sService = dict(ds='Data Storage', dp='Data Protection', ci='Compute Infrastructure', sn='Storage Network')[oArgs.service.lower()]
                oDatabase.execute("UPDATE Incident SET ReportingService = ? WHERE Number = ?", (sService, oArgs.incident))
            else:
                print("Unknown service.")
                exit(1)
        if oArgs.exclude:
            oDatabase.execute('UPDATE Incident SET Exclude = 1 WHERE Number = ?', (oArgs.incident, ))

        if not (oArgs.notes or oArgs.service or oArgs.exclude):
            print("Nothing supplied to update!")
            exit(1)

    if oArgs.request:
        if oDatabase.fetchValue("SELECT COUNT(*) FROM Request WHERE Number = ?", (oArgs.request,)) == 0:
            print("Request does not exist!")
            exit(1)

        if oArgs.notes:
            sNotes: str = oDatabase.fetchValue("SELECT Notes FROM Request WHERE Number = ?", (oArgs.request,))
            if oArgs.notes.startswith('+'):
                oArgs.notes = f"{sNotes} {oArgs.notes[1:]}"
            oDatabase.execute("UPDATE Request SET Notes = ? WHERE Number = ?", (oArgs.notes, oArgs.request))
        if oArgs.service:
            if oArgs.service.lower() in ['ds', 'dp', 'ci', 'sn']:
                sService = dict(ds='Data Storage', dp='Data Protection', ci='Compute Infrastructure', sn='Storage Network')[oArgs.service.lower()]
                oDatabase.execute("UPDATE Request SET Service = ? WHERE Number = ?", (sService, oArgs.request))
            else:
                print("Unknown service.")
                exit(1)
        if oArgs.exclude:
            oDatabase.execute('UPDATE Request SET Exclude = 1 WHERE Number = ?', (oArgs.request, ))

        if not (oArgs.notes or oArgs.service or oArgs.exclude):
            print("Nothing supplied to update!")
            exit(1)

    if oArgs.change:
        if oDatabase.fetchValue("SELECT COUNT(*) FROM Change WHERE Number = ?", (oArgs.change,)) == 0:
            print("Change does not exist!")
            exit(1)
        if oArgs.service:
            if oArgs.service.lower() in ['ds', 'dp', 'ci', 'sn']:
                sService = dict(ds='Data Storage', dp='Data Protection', ci='Compute Infrastructure', sn='Storage Network')[oArgs.service.lower()]
                oDatabase.execute("UPDATE Change SET Service = ? WHERE Number = ?", (sService, oArgs.change))
            else:
                print("Unknown service.")
                exit(1)
        else:
            print("Nothing supplied to update!")
            exit(1)


def loadDatabaseTable():
    lDataList = readCSVFile(oArgs.file, encoding='ISO-8859-1')

    # print(lDataList[0])
    loadTableToDatabase(oArgs.table, lDataList)

    exit()


if __name__ == '__main__':
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   This is the main program. Depending on the command line arguments, do some things
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    oArgs = parseCommandLine()

    oDatabase = Database("/Users/waynemoss/Library/CloudStorage/OneDrive-SharedLibraries-HitachiVantara/MS-ANZ RBA - Reserve Bank of Australia - General/"
                         "Reports/Service Management Report/2024/Database/ServiceNOW.db")

    if oArgs.update and (oArgs.incident or oArgs.request or oArgs.change):
        updateTableEntry()

    if oArgs.table:
        loadDatabaseTable()

    if oArgs.report:
        ServiceNOWReports()
