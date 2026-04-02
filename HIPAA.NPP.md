# HIPAA.NPP

\* HIPAA.NPP

\* ====================================================================

\* DESC: GENERATE HIPAA NOTICE OF PRIVACY PRACTICES

\*

\* ====================================================================

\* MAINTENANCE HISTORY:

\*

\* 04-14-25 MCH RTC222760 New hardcoded value of CORRESP.INPUT.DEF.ID

\* 01-09-25 SAC RTC219016 Fix FILENO being submitted to CORRESP.TRACKING

\* 09-12-24 ABA RTC209552 Exclude Admin EEIDs

\* 08-14-24 MLZ RTC208830 Restrict multiple sessions

\* 11-01-23 ABA RTC197552 Update for SF groups to be sent to SmartComm

\*                        Some code standard updates. FI groups still

\*                        need to work the original way.

\* 11-10-21 MCH RTC164083 Call new page4 version with fixed typos

\* 03-18-21 DHS RTC152723 Copy FILENAME to REDCARD.PCL before FTP

\* 10-10-20 ABA RTC131357 Fix extra form feeds after 32K lines and use

\*                        FTP.ID: 103069 which uses binary transfer

\*                        mode ONLY.

\* 05-01-20 EFZ RTC101156 During \'FI\' processing, pull HIPAA years
from

\*                        OUTPUT.CONTROL instead of always using 3
years;

\*                        Store all of the PCL files generated in the

\*                        file REDCARD.PCL

\* 03-26-20 3LM RTC131358 Correct the privacy notices flag for deps when

\*                        EE last HIPAA notice sent is less than 3 years

\* 02-21-20 CLJ RTC131331 rollback of HIPAA privacy notices fail at

\*                        redcard

\* 05-10-18 BAJ RTC88114 - added counter for files

\* 04-30-18 BAJ RTC71176 - Modified program so error report apart of

\*                         report on HIPAA.PMENU

\* 12-13-17 JEM RTC53113 - Remove 2000 record limit on a file.

\* 12-08-17 ABA RTC 15965 - Reopened to ensure dependents have coverage

\*                        - to receive hipaa notice.

\* 07-17-17 ABA RTC15965 - Update to use alternate address if different

\*                       - than the address for the employee

\*                       - hipaa.sent is now multi-value

\* 07-07-17 KES RTC20305 Corrected the .pcl file name per RedCard

\*                       information.

\* 04-27-17 KES RTC13794 Add code to FTP the print file

\* 06-23-16 ZPS R82375 HIPAA Privacy not printing for \'W\' divisions

\* 06-29-15 RJH R65258 add MAX.PER.PRINT.FILE to limit size of spool
file

\* 04-13-15 MJF R64885 - Move generic pg 1 09-23-13 eff date placement,

\*                       and update last page for Meritain Amherst

\*                       contact info.

\*                       New page 1 format: HIPAA.NPP1.150415

\* 12-10-13 MJF R49526 - Update generic forms for new HIPAA language.

\*                       Effective date of new text is 9-23-13.

\*                       This does not apply to fully-insured members.

\* 09-23-13 TAS R46665 Add Revised Privacy Notice for the University of

\*                     Notre Dame.

\* 01-04-13 IM  R34361 Remove Privacy Notice for University of Notre
Dame

\* 11-14-12 TMH R30697 Add Privacy Notice for University of Notre Dame

\* 01-27-12 LJH R14565 allow TEXT fields in HIPAA forms

\* 10-26-11 RJH R11924 hard code the PLAN.COMPLIANCE.DATE

\* 02-18-11 SAS 29264.1 Update hardcoded addresses for move on 03/18/11

\* 01-10-11 LJH LONH_110110 fix CERTS.LAYOUT select statement

\* 05-14-10 RJH #28259 subtract 1 from EOIDX if before origeff

\* 07-13-09 LJH 27286 Modify process to handle Extraterritoriality

\* 6-24-09 AKB 27212.1 Change to pull HIPAA forms from certs layout

\*  record for fully insured to make the forms printed here consistent

\*  with the forms printed when the certs are generated

\* 09-09-08 ABS #25145.39 email domain name change project

\* 1-8-08 AKB 23445 Change error message to contact IS Dept

\* 04-30-07 JDB #25184 added ee hipaa.sent logic and changed to print fi

\*  forms

\* 11-01-05 RJH #22794 change from CERTS.TEXT to CERTS.TEXT2

\* 12-SEP-05 AKB 20073 Removed hard coded checks for KEVINS & JOEYE

\* 09-27-04 je JOEYE_040927A someone changed the FORMATS

\*  SECD.HIPAA.NOTICE1 but didn\'t change the program

\* 09-14-04 je 19808 uncommented code that would skip the letter under

\*  certain cirumstances

\* 09-14-04 je 19808 removed code added to skip DO.BILLING for ROBINS

\* 09-10-04 je 19808 don\'t execute DO.BILLING for ROBINS (meant for one

\*  time re-run of notices that were already billed for)

\* 07-26-04 je 18477 changed all references to ADMINBP to BP

\* 04-06-04 je 18783 validation checks not needed here, done in calling

\*  programs

\* 03-30-04 KS  #16499.2 changed HIPAA.NOTICE formats and variables

\* 03-17-04 RJH #18627 change CBSA address

\* 03-15-04 ks  #16499.5 spool HIPAA notices to JRT when called from

\*  HIPAA.PMENU

\* 05-14-03 je - if EE already got notice don\'t resend

\* 05-08-03 RJH #16499.3 make the complaint.address a variable

\* 04-03-03 RJH #16499.1 add DO.BILLING

\* 03-24-03 je 16499

\*

\* ====================================================================

      EQUATE PROGRAM.NAME TO \'HIPAA.NPP\'

\*

      GOSUB CHECK.IF.ALREADY.RUNNING

      IF PROGRAM.ALREADY.RUNNING THEN GOTO EXIT

\*

      OPENIT = \'Y\'

\$INCLUDE LIBBP PCL.CODES.INS

\$INCLUDE BP EMPLOYERS.INS

\$INCLUDE BP EMPLOYEES.INS

\$INCLUDE BP EROPTIONS.INS

\$INCLUDE BP LETTER.COPIES.INS

\$INCLUDE BP ERADJ.INS

\$INCLUDE BP CERTS.TEXT2.INS

\$INCLUDE BP CARRIERS.INS

\$INCLUDE BP COBRA.INS

\$INCLUDE BP COMPANY.INS

\$INCLUDE BP GROUPINFO.INS

\$INCLUDE BP CERTS.LAYOUT.INS

\$INCLUDE BP CERTS.MERGE.INS

\$INCLUDE BP SYSMGMT.INS

\$INCLUDE BP ERAUX2.INS

\$INCLUDE BP FTP.REQUESTS.INS

\$INCLUDE BP FTP.LOG.INS

\$INCLUDE BP DEPENDENTS.INS

\$INCLUDE BP OUTPUT.CONTROL.INS

\*

      CALL \*OPEN.FILE(\'\',\'FORMATS\',FORMATS)

      CALL \*OPEN.FILE(\'\',\'&HOLD&\',HOLD)

\*

\*========================  MAIN PROCESSING ===========================

\*

      GOSUB INIT.VARS

      CALL \*GET.LASER.SPOOLIT(SPOOLIT)

      IF SPOOLIT = \'Q\' THEN GOTO EXIT

      GOSUB SET.FILE.NAME

      GOSUB FILTER.EEIDS

      GOSUB MAIN.LOOP

\*

      IF FI.FILE.STARTED THEN

         GOSUB FINISH.PRINT.FILE

         EXECUTE \'COPY FROM &HOLD& TO REDCARD.PCL \':FILENAME:\'
OVERWRITING DELETING\'

         GOSUB FTP.FILE

      END

      FORMLIST HIPAA.PDONE TO SLUNIT1

\* command is used below so user can see listname

      EXECUTE \'SAVE.LIST HIPAA.PDONE\':UV.LOGNAME:\' FROM \': SLUNIT1

\*

\* save a list of SF members that went to SmartComm

\*

\* command is used below so user can see listname

      MAX.SF.DEPNO = DCOUNT(SF.DEPNO.LIST,@FM)

      IF MAX.SF.DEPNO \> 0 THEN

         FORMLIST SF.DEPNO.LIST TO SLUNIT2

         EXECUTE \'SAVE.LIST HIPAA.PDONE.SC.\':UV.LOGNAME:\' FROM \':
SLUNIT2

      END

      IF ERRORS.FOUND THEN \@USER0 = ERROR.REPORT.NAME

      GOTO EXIT

\*

\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\* SUBROUTINES
\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*

\*

\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*

CHECK.IF.ALREADY.RUNNING:

\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*

      PROGRAM.ALREADY.RUNNING = 0

      LOCK.KEY = PROGRAM.NAME:\'.RUN.TIME.LOCK\'

      CALL \*RUN.TIME.LOCK(LOCK.STATUS,LOCK.KEY,\'SET\')

      IF LOCK.STATUS\<1\> = \'RUNNING\' THEN

         PROGRAM.ALREADY.RUNNING = 1

      END

      RETURN                             ; \*CHECK.IF.ALREADY.RUNNING:

\*

INIT.VARS:

\*

      UV.LOGNAME = \'\'

      UV.OPTION = \'ACTUAL\'

      CALL \*GET.UV.LOGNAME(UV.LOGNAME,UV.OPTION)

      SPOOLIT = \'\'

      \@USER0 = \'\'

      HIPAA.OPTIONS = \'TRANS.LOG,\':PROGRAM.NAME

\*

      REC.CNT=0

      ERRORS.FOUND = 0

      MACRO.NUMBERS = \'\'

\*

      TODAY = DATE()

      YYYYMMDD = CONVERT(\'-\',\'\',OCONV(TODAY,\'D4-YMD\'))

      CALL \*GET.FREE.LISTNO(SLUNIT1)

      FORMLIST \'X\' TO SLUNIT1

      CALL \*GET.FREE.LISTNO(SLUNIT2)

      FORMLIST \'X\' TO SLUNIT2

      CALL \*GET.FREE.LISTNO(SLUNIT3)

      CLEARSELECT SLUNIT1

      CLEARSELECT SLUNIT2

\*

\* store list of depnos that went to SmartComm

      SF.DEPNO.LIST = \'\'

      FI.FILE.STARTED = 0

\*

      RETURN                             ; \* INIT.VARS:

\*

SET.FILE.NAME:

\*

\* command is used below so user can see listname

      EXECUTE \'GET.LIST HIPAA.P\':UV.LOGNAME:\' TO \': SLUNIT1

\*

      EEID.LIST = \'\'

      FILENAME = \'\'

\* creates counter below so no duplicate files for the same day

      CALL
\*READREC(HIPAA.CNT.REC,\'\',SYSMGMT,\'HIPAA.DAILY.COUNTER\',\'WAIT\':@VM\'NO.MAT\')

      IF \@USER.RETURN.CODE = \'\' THEN

         IF HIPAA.CNT.REC\<2\> = TODAY THEN

            IF NUM(HIPAA.CNT.REC\<1\>) AND HIPAA.CNT.REC\<1\> \<\>
\'99\' THEN

               HIPAA.CNT.REC\<1\> += 1

            END

         END ELSE

            HIPAA.CNT.REC\<1\> = 1

            HIPAA.CNT.REC\<2\> = TODAY

         END

      END ELSE

         HIPAA.CNT.REC\<1\> = 1

         HIPAA.CNT.REC\<2\> = TODAY

      END

      CALL
\*WRITEREC(HIPAA.CNT.REC,\'\',SYSMGMT,\'HIPAA.DAILY.COUNTER\',HIPAA.OPTIONS:@VM:\'NO.MAT\')

      FILENAME =
\'1004HP\':YYYYMMDD:\'RH\':HIPAA.CNT.REC\<1\>:\'E1EDD11.PCL\'

\*

      RETURN                             ; \* SET.FILE.NAME:

\*

FILTER.EEIDS:

\* prevents back dating hipaa letters, send new ones based on that date

      START.ALT.ADD = ICONV(\'10-01-17\',\'D2-\')

      NUM.MONTHS = -36

      THREE.YEARS.AGO = SUBR(\'\*BUMP.DATE2\',TODAY,NUM.MONTHS)

      LAST.FILENO = \'UNASSIGNED\'

      LAST.OUTPUT.CONTROL.ID = \'UNASSIGNED\'

\*

      LOOP

         READNEXT EEID FROM SLUNIT1 ELSE EXIT

         GOSUB SET.ADMIN.EE

         IF ADMIN.EE THEN CONTINUE

         CALL \*READREC(\'\',MAT EEREC,EMPLOYEES,EEID,\'\')

         IF \@USER.RETURN.CODE \# \'\' THEN

            PRINT \'Error - EMPLOYEE \[\':EEID:\'\] record missing, it
will be skipped.\'

            CONTINUE

         END

         LOCATE \'E\' IN EERELATIONSHIP\<1,1\> SETTING DEPIDX ELSE
DEPIDX = 1

\*

\* this is the old record on CCLAIM

         IF INDEX(EENEED,\'X\',1) \> 0 THEN CONTINUE

\*

         LOCATE TODAY IN EEEFFDATE\<1,1\> BY \'DR\' SETTING EEIDX ELSE
NULL

         IF EEEFFDATE\<1,EEIDX\> = \'\' AND EEIDX \> 1 THEN EEIDX -= 1

         FILENO = EEFILENO\<1,EEIDX\>

         GOSUB SET.SKIP.FILENO

         IF SKIP.FILENO THEN CONTINUE

         IF FILENO\[1,4\] = \'9999\' THEN CONTINUE

         TERM = SUBR(\'\*TERMDATE.COBRA\',EEID,EETERMDATE,EESTATUS)

\*

         IF (TERM \<\> \'\' AND TERM \< TODAY) OR (TERM \<\> \'\' AND
EETERMDATE \<\> ERTERMDATE) THEN CONTINUE

         IF EESTATUS\<1,DEPIDX\> = \'I\' OR EESTATUS\<1,DEPIDX\> = \'D\'
OR EESTATUS\<1,DEPIDX\> = \'W\' OR EESTATUS\<1,DEPIDX\> = \'R\' THEN
CONTINUE

         IF (EEMEDCOV\<1,EEIDX\> = \'\' AND EEDENCOV\<1,EEIDX\> = \'\'
AND (EEVISCOV\<1,EEIDX\> = \'\' OR EEVISCOV\<1,EEIDX\> = \'N\')) THEN

            \* if bumper line, check if next line above has coverage

            IF EEEFFDATE\<1,EEIDX\> = \'\' AND EEIDX \> 1 THEN EEIDX -=
1

            IF (EEMEDCOV\<1,EEIDX\> = \'\' AND EEDENCOV\<1,EEIDX\> =
\'\' AND (EEVISCOV\<1,EEIDX\> = \'\' OR EEVISCOV\<1,EEIDX\> = \'N\'))
THEN CONTINUE

            FILENO = EEFILENO\<1,EEIDX\>

         END

\*

\* Whenever FILENO changes determine SF or FI, OUTPUT.CONTROL \@ID and

\* HIPAA.CUTOFF.DT

         IF LAST.FILENO \# FILENO THEN

            GOSUB DETERMINE.HIPAA.CUTOFF.DT

         END

         IF ((EEHIPAA.SENT\<1,DEPIDX\> \# \'\') AND
(EEHIPAA.SENT\<1,DEPIDX\> \< HIPAA.CUTOFF.DT)) OR
(EEHIPAA.SENT\<1,DEPIDX\> = \'\') THEN

\* Send new HIPAA notice if the last sent date is greater than the HIPAA

\* cutoff, which is 3 years for SF and a variable period of time for FI

\* line of business. If the last sent date is less than the cutoff do
NOT

\* send the HIPAA notice unless HIPAA.SENT date is null - a null date

\* means that no HIPAA notice has been sent yet.

            EEID.LIST\<-1\> = EEID:\'.\':EEDEPNO\<1,DEPIDX\>

         END

\* get employees starting effdate

         GOSUB SET.EE.START.EFF.DATE

         IF EE.START.EFF.DATE = \'\' THEN

            EE.START.EFF.DATE = EEEFFDATE\<1,EEIDX\>

         END

\* send hipaa letter - to dependents

         MAX.DEP = DCOUNT(EEDEPNO,@VM)

         FOR DEPIDX2=1 TO MAX.DEP

            IF EERELATIONSHIP\<1,DEPIDX2\> \# \'E\' THEN

               ADD.DEPNO = 0

               IF EESTATUS\<1,DEPIDX2\> = \'I\' OR EESTATUS\<1,DEPIDX2\>
= \'D\' OR EESTATUS\<1,DEPIDX2\> = \'W\' OR EESTATUS\<1,DEPIDX2\> =
\'R\' OR EESTATUS\<1,DEPIDX2\> = \'T\' THEN CONTINUE

\* Don\'t send a new privacy notice if one was sent since the cutoff
date

               IF EEHIPAA.SENT\<1,DEPIDX2\> \> HIPAA.CUTOFF.DT THEN
CONTINUE

               CALL \*READREC(\'\',MAT
DPREC,DEPENDENTS,EEID:\'.\':EEDEPNO\<1,DEPIDX2\>,\'\')

               IF \@USER.RETURN.CODE \<\> \'\' THEN

                  PRINT \'Error - Cannot read DEPENDENT \[\':
EEID:\'.\':EEDEPNO\<1,DEPIDX2\>: \'\]\'

                  CONTINUE

               END

               DEP.HAS.COV = \'\'          ; \* must have a coverage

               GOSUB DEP.HAS.COVERAGE

               IF NOT(DEP.HAS.COV) THEN CONTINUE

               DEP.DIFF.ADD = \'\'

               IF TRIM(DPADD1:DPADD2:DPCITY:DPSTATE:DPZIP) \<\> \'\'
THEN

                  DEP.DIFF.ADD = 1

               END

               BEGIN CASE

                  CASE EE.START.EFF.DATE = DPREL.TYPE.EFFDATE\<1,1\> AND
DPREL.TYPE.EFFDATE\<1,1\> \# \'\' AND DPREL.TYPE.EFFDATE\<1,1\> \>=
START.ALT.ADD

\* new employee and dependents with same effdate

\* send only if diff add from employee

                     IF DEP.DIFF.ADD THEN

                        ADD.DEPNO = 1

                     END

                  CASE EE.START.EFF.DATE \# DPREL.TYPE.EFFDATE\<1,1\>
AND DPREL.TYPE.EFFDATE\<1,1\> \# \'\' AND DPREL.TYPE.EFFDATE\<1,1\> \>=
START.ALT.ADD

\* covers recently added dep regardless of address

\* will have different effdate

                     ADD.DEPNO = 1

                  CASE EEHIPAA.SENT\<1,DEPIDX2\> = \'\' AND
(EEHIPAA.SENT\<1,DEPIDX\> \< HIPAA.CUTOFF.DT AND
EEHIPAA.SENT\<1,DEPIDX\> \>= START.ALT.ADD)

\* dependents didn\'t receive a privacy notice and the subscriber/EE has

\* but not in the last N years. need to check dependents for alternate

\* address if there is then a 2nd letter is needed for privacy for

\* dependent.

                     IF DEP.DIFF.ADD THEN

                        ADD.DEPNO = 1

                     END

                  CASE EEHIPAA.SENT\<1,DEPIDX2\> \# \'\' AND
EEHIPAA.SENT\<1,DEPIDX2\> \< HIPAA.CUTOFF.DT

\* going forward, after the variable cutoff date send notification to
the

\* dependent if have a different address

                     IF DEP.DIFF.ADD THEN

                        ADD.DEPNO = 1

                     END

                  CASE 1

                     NULL

               END CASE

               IF ADD.DEPNO THEN

                  EEID.LIST\<-1\> = EEID:\'.\':EEDEPNO\<1,DEPIDX2\>

               END

            END

         NEXT DEPIDX2

      REPEAT

      HIPAA.PDONE = EEID.LIST            ; \* use for in hipaa.pmenu

      WRITELIST EEID.LIST ON \'EEID.LIST.\':UV.LOGNAME

      RETURN                             ; \* FILTER.EEIDS:

\*

SET.ADMIN.EE:

\*

      CE.RULE.IDS = \'ADMIN\'

      CE.RULE.IDS := \@VM:\'DENTAL\'

      CE.RULE.IDS := \@VM:\'DEP.LIFE\'

      CE.RULE.IDS := \@VM:\'DIVISION\'

      CE.OPTIONS = 0

      CE.OPTIONS := \@VM:0

      CE.OPTIONS := \@VM:0

      CE.OPTIONS := \@VM:0

\* Skip EE when the EEID begins with \'S\' or ends in \'X\', \'C\',
\'D\' or \'S\'

      ADMIN.EE = 0

      PASS.FAIL = \'\'

      ERR.MSG = \'\'

      CALL \*CHECK.EEID(PASS.FAIL,EEID,CE.RULE.IDS,ERR.MSG,CE.OPTIONS)

      IF SUM(PASS.FAIL) \> 0 THEN

         ADMIN.EE = 1

         GOSUB ADD.TO.ERR.REPORT

      END

      RETURN                             ; \* SET.ADMIN.EE:

\*

SET.SKIP.FILENO:

\*

      SKIP.FILENO = 0

      CALL \*READREC(\'\',MAT ERREC,EMPLOYERS,FILENO,\'\')

      IF \@USER.RETURN.CODE \<\> \'\' THEN

         PRINT \'Error - EMPLOYERS \[\':FILENO:\'\] record missing, it
will be skipped.\'

         SKIP.FILENO = 1

      END

      IF NOT(SKIP.FILENO) THEN

         CALL \*READREC(\'\',MAT EOREC,EROPTIONS,FILENO,\'\')

         IF \@USER.RETURN.CODE \<\> \'\' THEN

            PRINT \'Error - EROPTIONS \[\':FILENO:\'\] record missing,
it will be skipped.\'

            SKIP.FILENO = 1

         END

      END

      RETURN                             ; \* SET.SKIP.FILENO:

\*

DETERMINE.HIPAA.CUTOFF.DT:

      CALL \*SF.FI(SF.FI,ERTRUST,\'\')

      LAST.FILENO = FILENO

\* When dealing with a FI group determine if OUTPUT.CONTROL ID changes

\* and if so calculate HIPAA.YEARS so a HIPAA.CUTOFF.DT can be
established

      IF SF.FI = \'FI\' THEN

         CALL \*OUTPUT.CONTROL.ID(OUTPUT.CONTROL.ID,FILENO,TODAY)

         IF LAST.OUTPUT.CONTROL.ID \# OUTPUT.CONTROL.ID THEN

            READV HIPAA.YEARS FROM OUTPUT.CONTROL, OUTPUT.CONTROL.ID,
OCHIPAA.YEARS.F ELSE

               HIPAA.YEARS = \'\'

            END

\* If no HIPAA.YEARS on OUTPUT.CONTROL then use default of 3 years

            IF NOT(HIPAA.YEARS MATCH \'1N0N\') THEN

               HIPAA.YEARS = 3

            END

            LAST.OUTPUT.CONTROL.ID = OUTPUT.CONTROL.ID

            NUM.MONTHS = -12 \* HIPAA.YEARS

            HIPAA.CUTOFF.DT = SUBR(\'\*BUMP.DATE2\',TODAY,NUM.MONTHS)

         END

      END ELSE

         HIPAA.CUTOFF.DT = THREE.YEARS.AGO

         LAST.OUTPUT.CONTROL.ID = \'\'

      END

      RETURN                             ; \* DETERMINE.HIPAA.CUTOFF.DT:

\*

SET.EE.START.EFF.DATE:

      EE.START.EFF.DATE = \'\'

      OPTIONS = \'FILENO,\':EEFILENO\<1,1\>

      LOWEST.DATE = \'\'

      CALL
\*DEP.EFFDATES(MED.EFFDATES,TERM,\'\',\'MED\',OPTIONS,EEID,EEDEPNO\<1,DEPIDX\>)

      IF MED.EFFDATES \# \'\' THEN

         LOWEST.DATE = MED.EFFDATES

      END

      CALL
\*DEP.EFFDATES(DEN.EFFDATES,TERM,\'\',\'DEN\',OPTIONS,EEID,EEDEPNO\<1,DEPIDX\>)

      IF DEN.EFFDATES \# \'\' AND LOWEST.DATE = \'\' THEN

         LOWEST.DATE = DEN.EFFDATES

      END ELSE

         IF DEN.EFFDATES \< LOWEST.DATE AND DEN.EFFDATES \# \'\' THEN

            LOWEST.DATE = DEN.EFFDATES

         END

      END

      CALL
\*DEP.EFFDATES(VIS.EFFDATES,TERM,\'\',\'VIS\',OPTIONS,EEID,EEDEPNO\<1,DEPIDX\>)

      IF VIS.EFFDATES \# \'\' AND LOWEST.DATE = \'\' THEN

         LOWEST.DATE = VIS.EFFDATES

      END ELSE

         IF VIS.EFFDATES \< LOWEST.DATE AND VIS.EFFDATES \# \'\' THEN

            LOWEST.DATE = VIS.EFFDATES

         END

      END

\* if at least one coverage, that is valid effective date

      EE.START.EFF.DATE = LOWEST.DATE

      RETURN                             ; \* SET.EE.START.EFF.DATE:

\*

DEP.HAS.COVERAGE:

      OPTIONS = \'FILENO,\':EEFILENO\<1,1\>

      CALL
\*DEP.EFFDATES(MED.EFFDATES,TERM,\'\',\'MED\',OPTIONS,EEID,EEDEPNO\<1,DEPIDX2\>)

      CALL
\*DEP.EFFDATES(DEN.EFFDATES,TERM,\'\',\'DEN\',OPTIONS,EEID,EEDEPNO\<1,DEPIDX2\>)

      CALL
\*DEP.EFFDATES(VIS.EFFDATES,TERM,\'\',\'VIS\',OPTIONS,EEID,EEDEPNO\<1,DEPIDX2\>)

      IF (MED.EFFDATES:DEN.EFFDATES:VIS.EFFDATES) \# \'\' THEN

         DEP.HAS.COV = 1

      END

      RETURN                             ; \* DEP.HAS.COVERAGE:

\*

MAIN.LOOP:

\*

\* command is used below so user can see listname

      EXECUTE \'GET.LIST EEID.LIST.\':UV.LOGNAME:\' TO \':SLUNIT1 : \'
CAPTURING XXX\'

\*

      LOOP

         READNEXT EEID.DEPNO FROM SLUNIT1 ELSE EXIT

         EEID = FIELD(EEID.DEPNO,\'.\',1)

         DEPNO = FIELD(EEID.DEPNO,\'.\',2)

         CALL \*READREC(\'\',MAT EEREC,EMPLOYEES,EEID,\'\')

         IF \@USER.RETURN.CODE \# \'\' THEN

            PRINT \'Error - EMPLOYEE \[\':EEID:\'\] record missing, it
will be skipped.\'

            CONTINUE

         END

         CALL \*READREC(\'\',MAT DPREC,DEPENDENTS,EEID.DEPNO,\'\')

         IF \@USER.RETURN.CODE \<\> \'\' THEN

            PRINT \'Error - Cannot read DEPENDENT \[\': EEID.DEPNO:
\'\]\'

            CONTINUE

         END

\*

         LOCATE DEPNO IN EEDEPNO\<1,1\> SETTING DEPIDX ELSE DEPIDX = 1

\*

         MACRO.NUMBERS = \'\'

         ASOFDATE = \'\'

         IF LOWEST.DATE \> TODAY THEN

            ASOFDATE = LOWEST.DATE

         END ELSE

            ASOFDATE = TODAY

         END

         LOCATE ASOFDATE IN EEEFFDATE\<1,1\> BY \'DR\' SETTING EEIDX
ELSE NULL

         IF EEEFFDATE\<1,EEIDX\> = \'\' AND EEIDX \> 1 THEN EEIDX -= 1

         FILENO = EEFILENO\<1,EEIDX\>

         GOSUB SET.SKIP.FILENO

         IF SKIP.FILENO THEN CONTINUE

         TERM = SUBR(\'\*TERMDATE.COBRA\',EEID,EETERMDATE,EESTATUS)

         GOSUB SET.EE.ADDRESS

         CALL \*SF.FI(SF.FI,ERTRUST,\'\')

         IF SF.FI = \'SF\' THEN

\*

\* SF groups will go to smart comm

\*

            LOCATE TODAY IN EOEFFDATE\<1,1\> BY \'DR\' SETTING EOIDX
ELSE NULL

            IF EOIDX \> 1 AND EOEFFDATE\<1,EOIDX\> = \'\' THEN EOIDX -=
1

            IF EOHIPAAP\<1,EOIDX\> = \'\' THEN CONTINUE

            GOSUB SET.GI.NPP

            IF GI.NPP = \'G\' THEN CONTINUE

            \* there is no spool it options for SF

            GOSUB DO.BILLING

            GOSUB SUBMIT.CORRESP.REQUEST

            CONTINUE

         END

\*

\* at this point only FI groups will use old style

\*

         REC.CNT+=1

         IF MOD(REC.CNT,10)=0 THEN PRINT \'\*\':

         IF REC.CNT = 1 THEN

            CALL \*SETPRINT(\'4\',FILENAME,\'512\',\'32000 0 0 ,NFMT\')

            FI.FILE.STARTED = 1

         END

\*

         UNIT.NO = 5

         PRINT ON 4 CHAR(27):\'E\'

         HUSH ON

         CALL \*SETPRINT(\'5\',\'TEMP\':UV.LOGNAME,\'513\',\'32000 0 0
,NFMT\')

         HUSH OFF

\*

         FORMS.USED.ON.THIS.NOTICE = \'\'

\*   Proceed according to Trust

\*

         GOSUB GET.HIPAA.FORMS

         IF LAYOUT.TO.USE = \'\' THEN

            ERR.MSG = \'MISSING HIPAA CERT LAYOUT FOR \':EEID.DEPNO

            GOSUB ADD.TO.ERR.REPORT

         END ELSE

            CALL \*COMPANY.ID(COMPANY.ID,FILENO,\'\')

            ADDR.LINES = EENAME:@VM:ADDLINE1:@VM:ADDLINE2:@VM:ADDLINE3

            CALL
\*MAILING.PAGE(INFO.OUT,COMPANY.ID,ADDR.LINES,\'EMPLOYEES\',EEID,\'\',\'\',\'TOP\',\'PP\')

            PRINT ON UNIT.NO INFO.OUT

            GOSUB GENERATE.HIPAA.FORMS

         END

         GOSUB PRINT.NOTICE

         GOSUB STORE.LETTER

      REPEAT

      RETURN                             ; \* MAIN.LOOP:

\*

SET.EE.ADDRESS:

\*

      DEP.LAST = TRIM(EELAST\<1,DEPIDX\>)

      DEP.FIRST = TRIM(EEFIRST\<1,DEPIDX\>)

      EENAME = DEP.FIRST:\' \':DEP.LAST

      CALL
\*VALIDATE.ADDRESS(VALID.ADDR,OUT.ADD1,OUT.ADD2,OUT.CITY,OUT.STATE,OUT.ZIP,OUT.COUNTRY,DPADD1,DPADD2,DPCITY,DPSTATE,DPZIP,DPCOUNTRY,\'\')

      IF VALID.ADDR THEN

         DEP.ADDR1 = OUT.ADD1

         DEP.ADDR2 = OUT.ADD2

         DEP.CITY = OUT.CITY

         DEP.STATE = OUT.STATE

         DEP.ZIP = OUT.ZIP

      END ELSE

         DEP.ADDR1 = EEADD1

         DEP.ADDR2 = EEADD2

         DEP.CITY = EECITY

         DEP.STATE = EESTATE

         DEP.ZIP = EEZIP

      END

      IF DEP.ADDR1 \# \'\' THEN

         ADDLINE1 = DEP.ADDR1

         ADDLINE2 = DEP.ADDR2

         ADDLINE3 = DEP.CITY:\' \':DEP.STATE:\' \':DEP.ZIP

      END ELSE

         ADDLINE1 = DEP.ADDR2

         ADDLINE2 = DEP.CITY:\' \':DEP.STATE:\' \':DEP.ZIP

         ADDLINE3 = \'\'

      END

      RETURN                             ; \* SET.EE.ADDRESS:

\*

SET.GI.NPP:

      GI.NPP = \'\'

      GROUPINFO.DATA =
SUBR(\'\*SELECTINDEX\',\'FILENO\',FILENO\[1,5\],\'GROUPINFO\')

      IF GROUPINFO.DATA \<\> \'\' THEN

         MAX.GROUPINFO.DATA = DCOUNT(GROUPINFO.DATA,@VM)

         READV GINPP FROM GROUPINFO,
GROUPINFO.DATA\<1,MAX.GROUPINFO.DATA\>, GINPP.F THEN

            GI.NPP = GINPP

         END

      END

      RETURN                             ; \* SET.GI.NPP:

\*

DO.BILLING:

\*

\* add manual billing entry to ERADJ file if needed

      IF EOHIPAAP\<1,EOIDX\> \> 0 THEN

         BILLPERIOD = CONVERT(\'-\',\'\',OCONV(ERBILLEDTO+1,\'D2-YM\'))

         SUFFIX=\'\'

         EXIT.LOOP = 0

         LOOP

            EAID=FILENO\<1,EEIDX\>:BILLPERIOD:SUFFIX

            SUFFIX+=1

            CALL \*READREC(\'\',MAT EAREC,ERADJ,EAID,\'WAIT\')

            IF \@USER.RETURN.CODE \# \'\' THEN EXIT

            IF INDEX(EADESC,\'HIPAA NOTICE OF PRIVACY PRACTICES\',1) \>
0 THEN

               EXIT.LOOP = 1

            END

         UNTIL EXIT.LOOP

         REPEAT

         LOCATE EEID.DEPNO IN EACOMMENTS\<1,1\> SETTING XXX THEN

            RELEASE ERADJ,EAID

         END ELSE

            INS EEID.DEPNO BEFORE EACOMMENTS\<1,-1\>

            EADESC = DCOUNT(EACOMMENTS,@VM):\' HIPAA NOTICE OF PRIVACY
PRACTICES\'

            COVPERIOD = BILLPERIOD

            EAIDX=1

            \* do not add extra line if condition below is met

            IF EADISTFIELD\<1,EAIDX\>=\'\' OR (EADISTFIELD\<1,EAIDX\> =
\'HIPAA\' AND EACOVPERIOD\<1,EAIDX\> = COVPERIOD) ELSE

               LOOP

                  EAIDX+=1

               UNTIL EADISTFIELD\<1,EAIDX\>=\'\' OR
(EADISTFIELD\<1,EAIDX\> = \'HIPAA\' AND EACOVPERIOD\<1,EAIDX\> =
COVPERIOD)

               REPEAT

            END

            EADISTFIELD\<1,EAIDX\> = \'HIPAA\'

            EADISTAMOUNT\<1,EAIDX\> += EOHIPAAP\<1,EOIDX\>

            EACOVPERIOD\<1,EAIDX\> = COVPERIOD

            EAAMOUNT += EOHIPAAP\<1,EOIDX\>

            CALL \*WRITEREC(\'\',MAT EAREC,ERADJ,EAID,\'\')

         END

      END

      RETURN                             ; \* DO.BILLING:

\*

SUBMIT.CORRESP.REQUEST:

\*

      ERROR.MESSAGES = \'\'

      CORRESP.DATA.REC = \'\'

      CORRESP.DATA.REC\<2\> = FILENO

      CORRESP.INPUT.DEF.ID = \'BASIC.SMARTCOMM.DOC.1\'

      RECIPIENT.TYPE = \'DEPENDENT\'

      RECIPIENT.ID = EEID.DEPNO

      METHOD.REQUESTED = \'LOOKUP\'

      REQUEST.SOURCE = PROGRAM.NAME

      REQUESTOR.TRACKING.ID = \'\'

\* this document name is setup manually by SmartComm admin users

      DOCUMENT.NAME = \'Meritain.HIPAA.Notice\'

      CALL
\*CREATE.CORRESP.TRACKING(ERROR.MESSAGES,CORRESP.DATA.REC,CORRESP.INPUT.DEF.ID,RECIPIENT.TYPE,RECIPIENT.ID,METHOD.REQUESTED,REQUEST.SOURCE,REQUESTOR.TRACKING.ID,DOCUMENT.NAME,\'\',\'\',\'\',\'\')

      IF ERROR.MESSAGES \# \'\' THEN

         ERR.MSG = \'Correspondence request failed for Dependent
\':EEID.DEPNO

         GOSUB ADD.TO.ERR.REPORT

         ERR.MSG = ERROR.MESSAGES\<1,2\>

         GOSUB ADD.TO.ERR.REPORT

      END ELSE

         LOCATE EEID.DEPNO IN SF.DEPNO.LIST\<1\> SETTING SF.DEP.IDX ELSE

            INS EEID.DEPNO BEFORE SF.DEPNO.LIST\<SF.DEP.IDX\>

         END

      END

      RETURN                             ; \* SUBMIT.CORRESP.REQUEST:

\*

ADD.TO.ERR.REPORT:\*

\*

      IF NOT(ERRORS.FOUND) THEN

         ERROR.REPORT.NAME =
PROGRAM.NAME:\'.ERROR.RPT.\':TODAY:\'.\':TIME()

         EXECUTE \'SETPRINT 1 \':ERROR.REPORT.NAME:\' 177\'

         HEADING ON 1 \"Run Date: \'DG\'Page \'SLC\'HIPAA NOTICE ERROR
REPORT\'LLL\'\"

         ERRORS.FOUND = 1

      END

      PRINT ON 1 ERR.MSG

      RETURN                             ; \* ADD.TO.ERR.REPORT:

\*

GET.HIPAA.FORMS:\*

\*

\*  \-\-- For fully-insured only \-\--

\*  Check for certs layout records for the carrier & state and use the

\*  most recent layout with hipaa forms.  These will be the hipaa forms

\*  sent out with the most recent certs for this carrier.  Hipaa forms

\*  are the same for all types of certs with a given carrier so do not

\*  need to check type of coverage or type of cert

\*

      LOCATE TODAY IN EREFFDATE\<1,1\> BY \'DR\' SETTING ERIDX ELSE NULL

      IF ERIDX \> 1 AND EREFFDATE\<1,ERIDX\> = \'\' THEN ERIDX -= 1

      GOSUB GET.PRIMARY.STATE

      FOUND.LAYOUT = 0

      MOST.RECENT = \'\'

      LAYOUT.TO.USE = \'\'

      EXECUTE \'SELECT CERTS.LAYOUT WITH CARRIER.PROD LIKE
\"\':ERINS.CARRIER\<1,ERIDX\>:\'\...\" BY.DSND EFF.DATE TO \':SLUNIT2

      EOR = 0

      LOOP

         READNEXT CERT.ID FROM SLUNIT2 ELSE EOR = 1

      UNTIL EOR OR FOUND.LAYOUT DO

         CALL \*READREC(\'\',MAT CRTLYREC,CERTS.LAYOUT,CERT.ID,\'\')

         IF \@USER.RETURN.CODE \# \'\' THEN MAT CRTLYREC = \'\'

         IF CRTLYTYPE = \'MAS\' THEN CONTINUE

         IF CRTLYSTATE = \'DEF\' OR CRTLYSTATE = PRIMARY.STATE THEN

            IF CRTLYEFF.DATE \< TODAY THEN

               IF CRTLYTERM.DATE = \'\' OR CRTLYTERM.DATE \> TODAY THEN

                  IF CRTLYACTIVATION.DATE \# \'\' THEN

                     LOCATE \'HIPAA\' IN CRTLYFORM.TYPE\<1,1\> SETTING
XXX THEN

                        IF CRTLYSTATE = PRIMARY.STATE THEN

                           FOUND.LAYOUT = 1

                           LAYOUT.TO.USE = CERT.ID

                        END ELSE

                           IF MOST.RECENT = \'\' THEN

                              MOST.RECENT = CRTLYEFF.DATE

                              LAYOUT.TO.USE = CERT.ID

                           END ELSE

                              IF CRTLYEFF.DATE \> MOST.RECENT THEN

                                 MOST.RECENT = CRTLYEFF.DATE

                                 LAYOUT.TO.USE = CERT.ID

                              END

                           END

                        END

                     END

                  END

               END

            END

         END

      REPEAT

      CLEARSELECT SLUNIT2

      RETURN                             ; \* GET.HIPAA.FORMS:

\*

GET.PRIMARY.STATE:

\* HIPAA forms don\'t vary by state so for now this check is not needed

      PRIMARY.STATE = ERACTUAL.STATE

      RETURN                             ; \* GET.PRIMARY.STATE:

\*

GENERATE.HIPAA.FORMS:\*

\*

\*   Fully-insured only \-\--

\*   Pulls pages for HIPAA certs from the certs layout found earlier and

\*   Formats the pages for printing following the certs layout rules

\*

\*  Need to reread the cert layout rec because we checked multiple
layouts

\*  and may not have the one selected read in any more

\*

      CALL \*READREC(\'\',MAT CRTLYREC,CERTS.LAYOUT,LAYOUT.TO.USE,\'\')

      IF \@USER.RETURN.CODE \# \'\' THEN MAT CRTLYREC = \'\'

      MAX.PAGES = DCOUNT(CRTLYFORMAT,@VM)

      FIRST.PAGE = 1

      FOR PAGE.IDX = 1 TO MAX.PAGES

         IF CRTLYFORM.TYPE\<1,PAGE.IDX\> = \'HIPAA\' THEN

            IF CRTLYCONTROL\<1,PAGE.IDX\> \# \'\' THEN

               CONTROL.TEST =
SUBR(\'\*TRANSS\',\'EMPLOYEES\',EEID,CRTLYCONTROL\<1,PAGE.IDX\>)

               IF CONTROL.TEST \# \'Y\' THEN CONTINUE

            END

            GOSUB BUILD.CERT.MACRO

            IF NOT(CREATED.MACRO) THEN EXIT

            INS CRTLYFORMAT\<1,PAGE.IDX\>:@SM:CRTMID BEFORE
FORMS.USED.ON.THIS.NOTICE\<1,-1\>

            GOSUB CHECK.FOR.NEW.PAGE

            GOSUB PRINT.CERT.VARIABLES

            PRINT ON UNIT.NO CHAR(12):CHAR(27):\'&f\':CRTMID:\'y5X\'

         END

      NEXT PAGE.IDX

      RETURN                             ; \* GENERATE.HIPAA.FORMS:

\*

BUILD.CERT.MACRO:\*

\*

\* \-\-- Fully-insured only \-\--

\* gets the macro number and creates a permanent macro from the cert

\* layout

\*

      MAT CRTMREC = \'\'

      CREATED.MACRO = 1

      EXECUTE \'SELECT CERTS.MERGE WITH FORMAT =
\"\':CRTLYFORMAT\<1,PAGE.IDX\>:\'\" TO \':SLUNIT3

      IF \@SELECTED \< 1 THEN

         CRTMID = SUBR(\'\*NEXT.NUMBER\',\'CERTS.MERGE\',\'ID\')

         CRTMFORMAT = CRTLYFORMAT\<1,PAGE.IDX\>

         CALL \*WRITEREC(\'\',MAT CRTMREC,CERTS.MERGE,CRTMID,\'\')

      END ELSE

         READNEXT CRTMID FROM SLUNIT3 ELSE

            ERR.MSG = \'UNABLE TO FIND CERTS.MERGE RECORD FOR
\':CRTLYFORMAT\<1,PAGE.IDX\>

            GOSUB ADD.TO.ERR.REPORT

            CREATED.MACRO = 0

         END

      END

      CLEARSELECT SLUNIT3

      LOCATE CRTMID IN MACRO.NUMBERS\<1\> SETTING XXX ELSE

         FORMNAME = \'\$\':CRTLYFORMAT\<1,PAGE.IDX\>

         READ FORMAT FROM FORMATS,FORMNAME THEN

            INS CRTMID BEFORE MACRO.NUMBERS\<-1\>

            PRINT ON 4 CHAR(27):\'E\':

            CALL \*CALLMACRO.ABI(FORMAT,CRTMID,4,\'\')

            PRINT ON 4 CHAR(27):\'&f10x5X\':         ; \* make permanent

\* Note: sub CHECK.MACRO.IN.CERTS.TEXT2 set back to original position

\* and not brought to standards. Not sure how to test this logic.

\* This will eventually not be needed once FI groups go to SmartComm.

            GOSUB CHECK.MACRO.IN.CERTS.TEXT2

         END ELSE

            ERR.MSG = \'UNABLE TO READ FORMAT \':FORMNAME

            GOSUB ADD.TO.ERR.REPORT

            CREATED.MACRO = 0

         END

      END

      RETURN                             ; \* BUILD.CERT.MACRO:

\*

CHECK.FOR.NEW.PAGE:\*

\*

\* \-\-- Fully-insured only \-\--

\* cert layouts contain flags for forcing page breaks and blank pages

\*  check for a forced page and print accordingly

\*

      BEGIN CASE

         CASE FIRST.PAGE

            PRINT.MACRO.LINE =
PCL.RESET:PCL.DUPLEX.LONG:CHAR(27):\'&f\':CRTMID:\'y4X\':CHAR(27):\'&l63F\'

            FIRST.PAGE = 0

         CASE CRTLYNPAGE\<1,PAGE.IDX\> = \'Y\'

            PRINT.MACRO.LINE =
PCL.DUPLEX.FRONT:CHAR(27):\'&f\':CRTMID:\'y4X\':CHAR(27):\'&l63F\'

         CASE 1

            PRINT.MACRO.LINE =
CHAR(27):\'&f\':CRTMID:\'y4X\':CHAR(27):\'&l63F\'

      END CASE

      PRINT ON UNIT.NO PRINT.MACRO.LINE

      RETURN                             ; \* CHECK.FOR.NEW.PAGE:

\*

PRINT.CERT.VARIABLES:\*

\*

\*  \-\--Fully-insured only \-\--

\*  Cert layouts allow variables to be defined to print on the cert form

\*  calculate and place the variables on the form

\*

      CALL \*READREC(\'\',MAT CRTMREC,CERTS.MERGE,CRTMID,\'\')

      IF \@USER.RETURN.CODE \# \'\' THEN MAT CRTMREC = \'\'

      MAX.VARS = DCOUNT(CRTMFIELD,@VM)

      INFO.OUT = \'\'

      IF MAX.VARS \> 0 THEN

         FOR VAR.IDX = 1 TO MAX.VARS

            IF CRTMFILE.NAME\<1,VAR.IDX\> = \'TEXT\' THEN

               VAR.VALUE1 = CRTMFIELD\<1,VAR.IDX\>

               IF CRTMFONT\<1,VAR.IDX\> = \"\" THEN

                  FONT.NAME = \'I010.10\'

               END ELSE

                  FONT.NAME = CRTMFONT\<1,VAR.IDX\>

               END

               IF INDEX(FONT.NAME,\'.\',1) = 0 THEN FONT.NAME := \'.10\'

               CALL
\*XPOS(CRTMHORIZONTAL\<1,VAR.IDX\>,CRTMVERTICAL\<1,VAR.IDX\>,VAR.VALUE1,FONT.NAME,INFO.OUT)

            END ELSE

               ERR.MSG = \'HIPAA CERT LAYOUT \':CRTMID:\' HAS VARIABLES.
 HIPAA CERTS ARE NOT PROGRAMMED FOR VARIABLES.\'

               GOSUB ADD.TO.ERR.REPORT

            END

         NEXT VAR.IDX

      END

\*

\*   Print footers

\*

      IF CRTMLEFT.FOOTER \<\> \"\" THEN

         TTT.FOOTER = CRTMLEFT.FOOTER

         GOSUB PARSE.FOOTER

         CALL \*XPOS(300,7320,TTT.FOOTER,\'I014.10\',INFO.OUT)

      END

      IF CRTMRIGHT.FOOTER \<\> \"\" THEN

         TTT.FOOTER = CRTMRIGHT.FOOTER

         GOSUB PARSE.FOOTER

         HOR.POS = 5490-SUBR(\'\*FONT.WIDTH\',\'I014.10\',TTT.FOOTER)

         CALL \*XPOS(HOR.POS,7320,TTT.FOOTER,\'I014.10\',INFO.OUT)

      END

      PRINT ON UNIT.NO INFO.OUT

      RETURN                             ; \* PRINT.CERT.VARIABLES:

\*

CHECK.MACRO.IN.CERTS.TEXT2:

\*

      READ TEXT.MACRO FROM CERTS.TEXT2,FORMNAME\[2,99\] THEN

         IF TEXT.MACRO \<\> FORMAT THEN

            PRINT FORMNAME\[2,99\]:\' HAS CHANGED IN FORMATS FROM WHAT
IS STORED IN CERTS.TEXT2.\'

            PRINT \'YOU CANNOT CONTINUE, CALL IS DEPARTMENT.\'

            PRINT \'PRESS RETURN TO CONTINUE\...\'

            INPUT XXX

            GOTO EXIT

         END

      END ELSE

         WRITE FORMAT TO CERTS.TEXT2,FORMNAME\[2,99\]

      END

      RETURN                             ; \*
CHECK.MACRO.IN.CERTS.TEXT2:

\*

PARSE.FOOTER:

\* Footer now allows for inclusion of state-specific text, the state-

\* specific portion is indicated by {} within here each entry is

\* delimited by \|. Usually just the state initials are needed

\* otherwise a \^ indicates the state-specific text string to print.

\* Nested braces will be used to indicate leading or trailing characters

\* that print if the state is found and does not have a specific text

\* string

\*

      IF INDEX(TTT.FOOTER,\'{\',1) \> 0 AND INDEX(TTT.FOOTER,\'}\',1) \>
0 THEN

         \* we found the braces now check for second pair of braces

         IF INDEX(TTT.FOOTER,\'{\',2) \> 0 AND INDEX(TTT.FOOTER,\'}\',2)
\> 0 THEN

            STATE.PREFIX = FIELD(TTT.FOOTER,\'{\',2)

            STATE.SUFFIX = FIELD(TTT.FOOTER,\'}\',2)

            FOOTER.PREFIX = FIELD(TTT.FOOTER,\'{\',1)

            FOOTER.SUFFIX = FIELD(TTT.FOOTER,\'}\',3)

            STATE.STRING = FIELD(TTT.FOOTER,\'{\',3)

            STATE.STRING = \"\|\":FIELD(STATE.STRING,\'}\',1):\"\|\"

         END ELSE

            STATE.PREFIX = \'\'

            STATE.SUFFIX = \'\'

            FOOTER.PREFIX = FIELD(TTT.FOOTER,\'{\',1)

            FOOTER.SUFFIX = FIELD(TTT.FOOTER,\'}\',2)

            STATE.STRING = FIELD(TTT.FOOTER,\'{\',2)

            STATE.STRING = \"\|\":FIELD(STATE.STRING,\'}\',1):\"\|\"

         END

         STATE.SUBSTRING = \"\"

         IF INDEX(STATE.STRING,\"\|\":PRIMARY.STATE,1) \> 0 THEN

            FIELD.COUNT = DCOUNT(STATE.STRING,\'\|\')

            FOR FIELD.IDX = 1 TO FIELD.COUNT

               TTT.STATE = FIELD(STATE.STRING,\'\|\',FIELD.IDX)

               IF TTT.STATE\[1,2\] = PRIMARY.STATE THEN

                  IF TRIM(TTT.STATE) = PRIMARY.STATE THEN

                     STATE.SUBSTRING =
STATE.PREFIX:TRIM(TTT.STATE):STATE.SUFFIX

                  END ELSE

                     STATE.SUBSTRING = FIELD(TTT.STATE,\"\~\",2)

                  END

               END

            NEXT FIELD.IDX

         END

         TTT.FOOTER = TRIM(FOOTER.PREFIX:STATE.SUBSTRING:FOOTER.SUFFIX)

      END

      RETURN                             ; \* PARSE.FOOTER:

\*

PRINT.NOTICE:

\*

      PRINTER CLOSE ON 5

      READ STORE.REC FROM HOLD,\'TEMP\':UV.LOGNAME ELSE NULL

      PRINT ON 4 STORE.REC

      RETURN                             ; \* PRINT.NOTICE:

\*

STORE.LETTER:

\*

      MAT LETCREC = \'\'

      LETCRECORD = EEID

      LETCFILE = \'EMPLOYEES\'

      LETCDEPNO = FIELD(EEID.DEPNO,\'.\',2)

      LETCPROGRAM = PROGRAM.NAME

      LETCTEMPLATE =\'\'

      LETCFORM = FORMS.USED.ON.THIS.NOTICE

      LETCTEXT.FILE = \'CERTS.TEXT2\'

      LETNO = 0

      IF SPOOLIT \<\> \'N\' THEN

         CALL \*ADD.LETTER.COPIES(LETNO,MAT LETCREC,STORE.REC)

         IF NOT(LETNO) THEN PRINT \'THE LETTER FOR \':EEID:\' COULD NOT
BE STORED.\'

      END

      STORE.REC = \'\'

      RETURN                             ; \* STORE.LETTER:

\*

FINISH.PRINT.FILE:

      MACRO.COUNT=0

      DELIM=99

\*

\* delete (make temporary) all macros

\*

      LOOP WHILE DELIM\<\>0

         REMOVE MACRO.NUMBER FROM MACRO.NUMBERS SETTING DELIM

         MACRO.COUNT+=1

         IF MACRO.COUNT = 50 THEN

            PRINT ON 4 CHAR(27):\'&f\':MACRO.NUMBER:\'y9X\'

            MACRO.COUNT = 0

         END ELSE

            PRINT ON 4 CHAR(27):\'&f\':MACRO.NUMBER:\'y9X\':

         END

      REPEAT

      PRINT ON 4 CHAR(27):\'E\'

      PRINTER CLOSE ON 4

      IF SPOOLIT \<\> \'N\' THEN

         EXECUTE \'SPOOL \':FILENAME:\' -ATT \':SPOOLIT

      END

      RETURN                             ; \* FINISH.PRINT.FILE:

\*

FTP.FILE:

\*

      FTP.REQUEST.ID = \'103069\'

      FTP.FILENAME = FILENAME

      FTP.ERROR = \'\'

      FTP.OVERRIDE = \'\'

      FTP.OPTIONS = \'\'

      PRINT

      PRINT \'FTP-ing print file: \':FTP.FILENAME

      FTP.OVERRIDE\<FTP.REQUEST.FILENAME.F\> = LOWER(FTP.FILENAME)

      CALL
\*SUBMIT.FTP.REQUEST(FTP.COMPLETED,FTP.ERROR,FTP.REQUEST.ID,FTP.OVERRIDE,FTP.OPTIONS)

      IF FTP.COMPLETED THEN

         FTP.LOG.ID = FTP.ERROR

         CALL \*READREC(\'\',MAT FTPLREC,FTP.LOG,FTP.LOG.ID,\'\')

         IF \@USER.RETURN.CODE = \'\' THEN

            IF FTPLCOMMENTS = \'\' THEN

               PRINT \'FTP request completed successfully.\'

            END ELSE

               ERR.MSG = \'Error. The FTP request generated the
following errors: \'

               GOSUB ADD.TO.ERR.REPORT

               ERR.MSG = FTPLCOMMENTS

               GOSUB ADD.TO.ERR.REPORT

            END

         END ELSE

            ERR.MSG = \'Error. Unable to read FTP.LOG for
\[\':FTP.LOG.ID:\'\]\'

            GOSUB ADD.TO.ERR.REPORT

         END

      END

      RETURN                             ; \* FTP.FILE:

\*

EXIT:

      IF NOT(PROGRAM.ALREADY.RUNNING) THEN

         CALL \*RUN.TIME.LOCK(LOCK.STATUS,LOCK.KEY,\'RELEASE\')

      END

   END
