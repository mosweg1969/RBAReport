# if oArgs.report == 5:
#     # ---------- Service Request response ------------
#     sSQL = ("SELECT Service, Number, RequestItem, State, ReportPriority, AssignedTo, ShortDescription, CommentsAndWorkNotes, Opened, Closed "
#             "FROM Request WHERE Service <> 'Other' AND Exclude = 0 ORDER BY ReportPriority, RequestItem")
#
#     oReport = Report(f'Service Requests opened in the month - Response Time')
#     oReport.addColumn('#', sJust='right')
#     oReport.addColumn('Service', sJust='left')
#     oReport.addColumn('Number', sJust='left')
#     oReport.addColumn('Request Item', sJust='left')
#     oReport.addColumn('Priority', sJust='centre', bBreak=True)
#     oReport.addColumn('AssignedTo', sJust='left')
#     oReport.addColumn('Description', sJust='left')
#     oReport.addColumn('Received', sJust='left')
#     oReport.addColumn('Target', sJust='left')
#     oReport.addColumn('Respond SLA Start', sJust='left')
#     oReport.addColumn('Respond SLA Breach', sJust='left')
#     oReport.addColumn('Respond SLA Stop', sJust='left')
#     oReport.addColumn('Response', sJust='left')
#     oReport.addColumn('Response', sJust='left')
#     oReport.addColumn('Points', sJust='right')
#
#     lResults = oDatabase.fetchList(sSQL, ())
#     iCtr = 0
#
#     lRequestTargetResponseTimes = ['', '0:00:10', '0:00:30', '0:01:00', '0:04:00', '0:08:00']
#     sOverrideSQL = "SELECT StartTime, EndTime FROM SLAOverride WHERE Number = ? AND SLA = ?"
#
#     sPrevPriority = ' '
#
#     for tResult in lResults:
#         sService, sNumber, sRequestItem, sState, sReportPriority, sAssignedTo, sDescription, sCommentsAndWorkNotes, sOpened, sResolved = tResult
#         dtOpened = datetime.strptime(sOpened, sDateTimeFormat)
#
#         if dtStartOfMonth <= dtOpened <= dtEndOfMonth:
#             if sReportPriority[0] != sPrevPriority[0]:
#                 iCtr = 1
#             else:
#                 iCtr += 1
#
#             sRespondSLAStart = sRespondSLAStop = sRespondSLABreach = "[yellow]      N/A"
#             sRespondTimeColour = "[yellow]   N/A"
#             sRespondTime = "   N/A"
#
#             sColourStart = sColourEnd = ''
#             sStartSuffix = sEndSuffix = ''
#             dtSLAStartTime = dtSLAStopTime = None
#
#             # Work out the first touch time by someone in the team, and save the time before that it was touched, as that is when the SLA should start
#             dtFirstTeamTouch = None
#             dtPreviousNonTeamTouch = None
#
#             sSLAStartTime = sOpened
#             sSLAEndTime = sResolved
#
#             for dtUpdate, sPerson, _sUpdateType, _sUpdateContent in parseCommentsAndWorkNotes(sCommentsAndWorkNotes):
#                 if sPerson in ['Shagufta Anjum Shaik', 'Marwa Elshawy', 'Siddhartha Dutta', 'Mathieu Doumerc', "Keith D'Souza", 'Wilson Lee']:
#                     dtFirstTeamTouch = dtUpdate
#                     break
#                 elif sPerson in ['Antonio Salmeri', 'Max Sekula']:
#                     pass
#                 else:
#                     dtPreviousNonTeamTouch = dtUpdate
#
#             # check if there's an SLA override
#             lSLAResults = oDatabase.fetchList(sOverrideSQL, (sNumber, 'Resolve'))
#             if lSLAResults:
#                 sStart, sEnd = lSLAResults[0]
#                 if sStart:
#                     sSLAStartTime = sStart
#                     sColourStart = '[pink3]'
#                     sStartSuffix = ' (Ex)'
#
#             lSLAResults = oDatabase.fetchList(sOverrideSQL, (sNumber, 'Response'))
#             if lSLAResults:
#                 sStart, sEnd = lSLAResults[0]
#                 if sStart:
#                     sSLAStartTime = sStart
#                     sColourStart = '[pink3]'
#                     sStartSuffix = ' (Ex)'
#                 if sEnd:
#                     sSLAEndTime = sEnd
#                     sColourEnd = '[pink3]'
#                     sEndSuffix = ' (Ex)'
#
#             sTarget = lRequestTargetResponseTimes[int(sReportPriority[0])]
#
#             if sSLAStartTime:
#                 dtSLAStartTime = datetime.strptime(sSLAStartTime, sDateTimeFormat)
#                 dtSLAStopTime = None
#
#                 # the SLA should start at the time when a non-team member last updated the case before the team were working on it
#                 if dtPreviousNonTeamTouch and dtPreviousNonTeamTouch > dtSLAStartTime and sStartSuffix == '':
#                     dtSLAStartTime = dtPreviousNonTeamTouch
#                     sColourStart = '[purple]'
#                     sStartSuffix = ' (T!)'
#
#                 sRespondSLAStart = sColourStart + dtSLAStartTime.strftime('%a %-d/%m %H:%M') + sStartSuffix
#
#                 # the SLA should end when the team first touched the ticket
#                 if sSLAEndTime:
#                     dtSLAStopTime = datetime.strptime(sSLAEndTime, sDateTimeFormat)
#                     if dtFirstTeamTouch and dtFirstTeamTouch < dtSLAStopTime and sEndSuffix == '':
#                         dtSLAStopTime = dtFirstTeamTouch
#                 elif dtFirstTeamTouch:
#                     # the SLA should end when the team first touched the ticket
#                     dtSLAStopTime = dtFirstTeamTouch
#
#                 if sSLAEndTime == '':
#                     # incident is not resolved, see if there's a touch
#                     if dtFirstTeamTouch and dtFirstTeamTouch >= dtSLAStartTime:
#                         sSLAEndTime = sSLAStartTime
#
#                 if sSLAEndTime:
#                     dtSLAStopTime = datetime.strptime(sSLAEndTime, sDateTimeFormat)
#
#                     if dtFirstTeamTouch and dtFirstTeamTouch >= dtSLAStartTime and sEndSuffix == '':
#                         dtSLAStopTime = dtFirstTeamTouch
#                         sColourEnd = '[purple]'
#                         sEndSuffix = ' (T!)'
#
#                     sRespondSLAStop = sColourEnd + dtSLAStopTime.strftime('%a %-d/%m %H:%M') + sEndSuffix
#
#                     iDays, iHours, iMinutes = calculateBusinessTimeDiff(dtSLAStartTime, dtSLAStopTime)
#                     sRespondTime = f"{iDays:2}:{iHours:02}:{iMinutes:02}"
#
#                     dtSLABreachTime = calculateBusinessTimeDuration(dtSLAStartTime, sTarget)
#                     sRespondSLABreach = dtSLABreachTime.strftime('%a %-d/%m %H:%M')
#
#                     if dtSLAStopTime > dtSLABreachTime:
#                         sRespondTimeColour = '[red]' + sRespondTime
#                     else:
#                         sRespondTimeColour = '[green]' + sRespondTime
#
#             if sAssignedTo in ['Max Sekula', 'Antonio Salmeri']:
#                 sAssignedTo = f"[yellow]{sAssignedTo}"
#
#             oReport.addCells(iCtr, mapService(sService), sNumber, sRequestItem, sReportPriority[0], mapPerson(sAssignedTo), sDescription, dtOpened.strftime('%-d/%m %H:%M'), sTarget,
#                              sRespondSLAStart, sRespondSLABreach, sRespondSLAStop, sRespondTime, sRespondTimeColour, sPoints)
#             oReport.addRow()
#             sPrevPriority = sReportPriority
#
#     oReport.showColumn(11, False)
#     oReport.printReport()
#     print("")
#
#     if oArgs.summary:
#         oReport.showColumn(0, False)
#         oReport.showColumn(3, False)
#         oReport.showColumn(4, False)
#         oReport.showColumn(7, False)
#         oReport.showColumn(8, False)
#         oReport.showColumn(9, False)
#         oReport.showColumn(10, False)
#         oReport.showColumn(11, True)
#         oReport.showColumn(12, False)
#
#         oReport.printReport()
#         oReport.sendToClipboard()
#         print("")
#
# if oArgs.report == 6:
#     # ---------- Service requests delivered ------------
#     sSQL = ("SELECT Service, Number, RequestItem, ReportPriority, RequestedFor, ShortDescription, Opened, Closed FROM Request WHERE "
#             "Service <> 'Other' AND Exclude = 0 AND State = 'Closed Complete' ORDER BY ReportPriority, RequestItem")
#     sOverrideSQL = "SELECT StartTime, EndTime FROM SLAOverride WHERE Number = ? AND SLA = ?"
#
#     oReport = Report(f'Service requests delivered in the month.')
#     oReport.addColumn('#', sJust='right')
#     oReport.addColumn('Service', sJust='left')
#     oReport.addColumn('Number', sJust='left')
#     oReport.addColumn('Request Item', sJust='left')
#     oReport.addColumn('Priority', sJust='left', bBreak=True)
#     oReport.addColumn('Source', sJust='left')
#     oReport.addColumn('Summary', sJust='left')
#     oReport.addColumn('Received', sJust='left')
#     oReport.addColumn('SLA Start', sJust='left')
#     oReport.addColumn('Target', sJust='left')
#     oReport.addColumn('Fulfilled', sJust='left')
#     oReport.addColumn('Breach Time', sJust='left')
#     oReport.addColumn('Fulfill Time', sJust='left')
#     oReport.addColumn('Points', sJust='right')
#
#     lResults = oDatabase.fetchList(sSQL, ())
#     iCtr = 0
#
#     lRequestTargetResponseTimes = ['', '0:02:00', '1:00:00', '5:00:00', '10:00:00', '260:00:00']
#
#     sPrevPriority = ' '
#
#     for tResult in lResults:
#         sService, sNumber, sRequestItem, sReportPriority, sRequestedFor, sDescription, sOpened, sClosed = tResult
#
#         dtOpened = datetime.strptime(sOpened, sDateTimeFormat)
#         dtClosed = datetime.strptime(sClosed, sDateTimeFormat)
#
#         dtSLAStart = dtOpened
#         dtSLAEnd = dtClosed
#
#         if dtStartOfMonth <= dtClosed <= dtEndOfMonth:
#             if sReportPriority[0] != sPrevPriority[0]:
#                 iCtr = 1
#             else:
#                 iCtr += 1
#
#             sColourStart = sColourEnd = ''
#             sStartSuffix = sEndSuffix = ''
#
#             # check if there's an SLA override
#             lSLAResults = oDatabase.fetchList(sOverrideSQL, (sNumber, 'Resolve'))
#             if lSLAResults:
#                 sStart, sEnd = lSLAResults[0]
#                 if sStart:
#                     dtSLAStart = datetime.strptime(sStart, sDateTimeFormat)
#                     sColourStart = '[pink3]'
#                     sStartSuffix = ' (Ex)'
#                 if sEnd:
#                     dtSLAEnd = datetime.strptime(sEnd, sDateTimeFormat)
#                     sColourEnd = '[pink3]'
#                     sEndSuffix = ' (Ex)'
#
#             iDays, iHours, iMinutes = calculateBusinessTimeDiff(dtSLAStart, dtSLAEnd)
#             sFulfillTime = f"{iDays:2}:{iHours:02}:{iMinutes:02}"
#
#             sTarget = lRequestTargetResponseTimes[int(sReportPriority[0])]
#             dtSLABreachTime = calculateBusinessTimeDuration(dtSLAStart, sTarget)
#
#             if dtClosed > dtSLABreachTime:
#                 sFulfillTime = '[red]' + sFulfillTime
#
#             if sRequestedFor in ("Keith D'Souza", "Mathieu Doumerc", "Marwa Elshawy", "Max Sekula", "Wilson Lee", "Antonio Salmeri",
#                                  "Shagufta Anjum Shaik", "Siddhartha Dutta"):
#                 sRequestedFor = 'ITIMS BAU'
#
#             oReport.addCells(iCtr, sService, sNumber, sRequestItem, sReportPriority[0], sRequestedFor, sDescription,
#                              dtOpened.strftime('%a %-d/%m %H:%M'),
#                              sColourStart + dtSLAStart.strftime('%a %-d/%m %H:%M') + sStartSuffix,
#                              sTarget, sColourEnd + dtSLAEnd.strftime('%a %-d/%m %H:%M') + sEndSuffix,
#                              dtSLABreachTime.strftime('%a %-d/%m %H:%M'), sFulfillTime, sPoints)
#             oReport.addRow()
#             sPrevPriority = sReportPriority
#
#     oReport.printReport()
#     print('')
#
#     if oArgs.summary:
#         oReport.showColumn(0, False)
#         oReport.showColumn(2, False)
#         oReport.showColumn(4, False)
#         oReport.showColumn(8, False)
#         oReport.showColumn(9, False)
#         oReport.showColumn(11, False)
#         oReport.printReport()
#         oReport.sendToClipboard()
#         print("")
#
