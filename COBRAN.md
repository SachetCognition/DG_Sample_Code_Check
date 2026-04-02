# COBRAN

\* COBRAN

\* Process COBRA initial notice - send to SMARTCOMM

\*

\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\* MAINTENANCE HISTORY
\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*

\*

\* 04-14-25 ABA RTC222755 SMARTCOMM changes to use BASIC.SMARTCOMM.DOC.1

\*                        format

\* 02-03-25 ABA RTC221333 Ensure FILENO is correctly sent based on

\*                        effdate used.

\* 10-20-23 MLZ RTC197524 Convert for printing at SmartComm

\* 03-03-21 ABA RTC151399 Add counter to filename so if program

\*                        is ran more than once a day. So that previous

\*                        file is not overriden.

\* 09-12-19 BEG RTC100645 Changed SAVEDLIST name to avoid overwriting.

\* 07-27-17 TH  RTC15944 create COBRA.DEP.NOTICE for Dependents with

\*                  alt address. Call DO.FORM instead of RUN BP FORMS

\* 04-27-17 KES RTC13794 Added code to FTP the print file

\* 07-27-15 RJH R70190 print if status P and from EMPLOYES.COBRA.NOTICE

\* 01-07-13 IM  R34361 Remove flex plan letter to COBRA notice for UND.

\* 12-14-11 TJL R15993 Add flex plan letter to COBRA notice for UND.

\* 08-18-09 RJH #27028.63 add NETC_DUMMY

\* 13-SEP-05 AKB 20073 Removed hard code check for MICHELEB

\* 09-28-04 KS #20077.1 spool to JRT for DATA.CENTER

\* 07-26-04 je 18477 changed all references to ADMINBP to BP

\* 03-31-03 RJH #14702.3 don\'t bill eeid starting with S

\* 03-12-03 MMB per JIMBO, change from LASER31 to LASER23

\* 08-14-02 jb  15224-Added INCLUDE BP ERADJ.INS.

\* 05-21-02 RJH 14702-add DO.BILLING

\* 11-02-00 RJH #9670 remove RBIS list

\* 03-17-00 MMB #7948 - Per Kathy Berg, change from LASER23 to LASER31

\* 02-08-00 MMB #7948 new program

\*

      EQUATE PROGRAM.NAME TO \'COBRAN\'

      EQUATE FALSE TO 0

      EQUATE TRUE TO 1

\*

      OPENIT=\'Y\'

\$INCLUDE BP EMPLOYEES.INS

\$INCLUDE BP EMPLOYERS.INS

\$INCLUDE BP EROPTIONS.INS

\$INCLUDE BP ERADJ.INS

\*

      GOSUB INIT.VARS

      GOSUB MAIN.LOOP

      GOSUB END.REPORT.FILE

\*

      GOTO EXIT

\*

\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\* SUBROUTINES
\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*

\*

INIT.VARS:

      CALL \*GET.UV.LOGNAME(UV.LOGNAME,\'ACTUAL\')

      PREVIOUS.LIST.INPUT = \'\'

      STORE.LIST = \'\'

      DEP.STORE.LIST = \'\'

      SLUNIT = SUBR(\'\*GET.FREE.LISTNO\')

      TODAY = DATE()

      RETURN                             ; \* INIT.VARS:

\*

MAIN.LOOP:

      EXECUTE \'SELECT &SAVEDLISTS& LIKE
\"COBRAN.DONE\':UV.LOGNAME:\'\....\" COUNT.SUP TO \':SLUNIT

      READLIST COBRAN.MATCH.LIST FROM SLUNIT THEN

         GOSUB GET.PREVIOUS.LIST.INPUT

         IF PREVIOUS.LIST.INPUT = \'Y\' THEN

            LOOP

               FORMLIST COBRAN.MATCH.LIST TO SLUNIT

               EXECUTE \'SHOW &SAVEDLISTS& DATE.CHANGED EVAL
\"FMT(@ID\[6\],\"R##:##:##\")\" COL.HDG \"TIME.CHANGED\" FMT \"12R\"
COUNT.SUP FROM \':SLUNIT:\' TO \':SLUNIT

               READLIST SELECTED.RECS FROM SLUNIT ELSE GOTO EXIT

               MAX.RECS = DCOUNT(SELECTED.RECS,@FM)

               IF MAX.RECS = 1 THEN

                  GETLIST SELECTED.RECS TO SLUNIT ELSE NULL

               END

            UNTIL MAX.RECS = 1

               PRINT \"You can select only one list.\"

               PRINT \'Press ENTER to continue\':

               INPUT XXX

            REPEAT

         END ELSE

            GETLIST PROGRAM.NAME:UV.LOGNAME TO SLUNIT ELSE NULL

         END

      END ELSE

         GETLIST PROGRAM.NAME:UV.LOGNAME TO SLUNIT ELSE NULL

      END

      READLIST COBRAN.DONE FROM SLUNIT ELSE GOTO EXIT

      FORMLIST COBRAN.DONE TO SLUNIT

      TEMPLATE.NAME = \'Meritain.COBRA.Intro.Letter\'

      DELETELIST \'STORE.\':UPCASE(TEMPLATE.NAME)

      DEP.TEMPLATE.NAME = \'Meritain.COBRA.Dependent.Notice\'

      DELETELIST \'STORE.\':UPCASE(DEP.TEMPLATE.NAME)

      FILENO = \'\'

      EEID.DEPNO.LIST = \'\'

      CNT = 0

      DEP.COUNT = 0

      LOOP

         READNEXT EEID FROM SLUNIT ELSE EXIT

         CALL \*READREC(\'\',MAT EEREC,EMPLOYEES,EEID,\'\')

         IF \@USER.RETURN.CODE \# \'\' THEN

            PRINT \'EMPLOYEE NOT FOUND \[\':EEID:\'\]\'

            CONTINUE

         END

         IF INDEX(EENEED,\'X\',1) \> 0 THEN CONTINUE

\*

\* find employee position

         LOCATE \'E\' IN EERELATIONSHIP\<1,1\> SETTING DEPIDX ELSE
DEPIDX = 1

\*

         LOCATE TODAY IN EEEFFDATE\<1,1\> BY \'DR\' SETTING EEIDX ELSE
NULL

         IF EEEFFDATE\<1,EEIDX\> = \'\' AND EEIDX \> 1 THEN EEIDX -= 1

         GOSUB SET.FUTURE.EEIDX

         IF SKIP.EE THEN CONTINUE

\*

\* load employer data

         IF EEFILENO\<1,EEIDX\> \<\> FILENO THEN

            FILENO = EEFILENO\<1,EEIDX\>

\*

            CALL \*READREC(\'\',MAT ERREC,EMPLOYERS,FILENO,\'\')

            IF \@USER.RETURN.CODE \# \'\' THEN

               PRINT \'EMPLOYER NOT FOUND \[\':FILENO:\'\]\'

               CONTINUE

            END

            CALL \*READREC(\'\',MAT EOREC,EROPTIONS,FILENO,\'\')

            IF \@USER.RETURN.CODE \# \'\' THEN

               PRINT \'EROPTIONS NOT FOUND \[\':FILENO:\'\]\'

               CONTINUE

            END

            \*

         END

         LOCATE TODAY IN EREFFDATE\<1,1\> BY \'DR\' SETTING ERIDX ELSE
NULL

         IF EREFFDATE\<1,ERIDX\> = \'\' AND ERIDX \> 1 THEN ERIDX -= 1

\*

         IF EETERMDATE\<\>\'\' OR ERCOBRAFEE\<1,ERIDX\>=\'N\' OR
INDEX(\' DWRI\',CONVERT(\'GL\',\'\',EESTATUS\<1,DEPIDX\>),1) \> 1 THEN
CONTINUE

         EMP.DEP.KEY = EEID:\'.\':EEDEPNO\<1,DEPIDX\>

         CNT += 1

         IF MOD(CNT,100) = 0 THEN PRINT \'\*\':

         GOSUB GENERATE.LETTERS          ; \* Employee or Dependent

         GOSUB DO.BILLING

      REPEAT

      RETURN                             ; \* MAIN.LOOP:

\*

GET.PREVIOUS.LIST.INPUT:

\* Get input to run previous process

      DEF.PREVIOUS.LIST.INPUT = \'Q\'

      LOOP

         PRINT \"Do you want to run for a previous list(Y/N) or
Q(Quit)\":DEF.PREVIOUS.LIST.INPUT:@(-9,LEN(DEF.PREVIOUS.LIST.INPUT)+1):

         INPUT PREVIOUS.LIST.INPUT

         PREVIOUS.LIST.INPUT = UPCASE(PREVIOUS.LIST.INPUT)

         IF PREVIOUS.LIST.INPUT = \'\' THEN PREVIOUS.LIST.INPUT =
DEF.PREVIOUS.LIST.INPUT

         IF PREVIOUS.LIST.INPUT = \'Q\' THEN GOTO EXIT

      UNTIL PREVIOUS.LIST.INPUT = \'Y\' OR PREVIOUS.LIST.INPUT = \'N\'

         PRINT \' Please enter a valid input(Y/N) or Q (Quit) \'

      REPEAT

      RETURN                             ; \* GET.PREVIOUS.LIST.INPUT:

\*

SET.FUTURE.EEIDX:

\* determine if covered currently or in future

      SKIP.EE = TRUE

      FOR FUTURE.IDX = EEIDX TO 1 STEP - 1

         IF ((EEMEDCOV\<1,FUTURE.IDX\>:EEDENCOV\<1,FUTURE.IDX\>) = \'\'
AND (EESTATE \<\> \'MN\' OR EELIFECOV\<1,FUTURE.IDX\> \<= 0)) THEN

            IF EESTATUS\<1,DEPIDX\> = \'P\' THEN

               \* If no coverage and NEED is C and status is P it was

               \* entered on screen EMPLOYEES.COBRA.NOTICE, so print

               \* the notice.

               SKIP.EE = FALSE

            END ELSE

               SKIP.EE = TRUE

            END

         END ELSE

            \* If coverage but status is P then skip

            IF EESTATUS\<1,DEPIDX\> = \'P\' THEN

               SKIP.EE = TRUE

            END ELSE

               SKIP.EE = FALSE

            END

         END

         IF NOT(SKIP.EE) THEN

            EEIDX = FUTURE.IDX

            EXIT

         END

      NEXT FUTURE.IDX

      RETURN                             ; \* SET.FUTURE.EEIDX:

\*

GENERATE.LETTERS:

\* COBRA.NOTICE + COBRA.DEP.NOTICE if dep address

      MATBUILD EE.REC FROM EEREC

\*

      SENT.SPOUSE.NOTICE = FALSE

      DEPID.LIST = \'\'                    ; \* list of dependents

      ADDRESS.OPTIONS = \'\'

      CALL
\*GET.DEPENDENT.ADDRESS.IDS(DEPID.LIST,EMP.DEP.KEY,EE.REC,ADDRESS.OPTIONS)

      IF DEPID.LIST\<1\> \# \'\' THEN

\*

\* Generate separate notice if Dependent has different address

\* do not send if child - only spouses

         LETTER.TYPE = \'DP\'

         DIFF.DEP.COUNT = DCOUNT(DEPID.LIST\<1\>,@VM)

         FOR DPNUM = 1 TO DIFF.DEP.COUNT

            DIFF.DEP = DEPID.LIST\<1,DPNUM\>         ; \* get EEID.DEPNO

            DIFF.IDX = FIELD(DIFF.DEP,\'.\',2)       ; \* get DEPIDX

            IF DIFF.IDX LE 0 THEN CONTINUE         ; \* Must be GE 1

            IF EERELATIONSHIP\<1,DIFF.IDX\> = \'C\' THEN CONTINUE

            IF EERELATIONSHIP\<1,DIFF.IDX\> = \'S\' THEN

               SENT.SPOUSE.NOTICE = TRUE

            END

            EEID.DEPNO = EEID:\'.\':EEDEPNO\<1,DIFF.IDX\>

            GOSUB SUBMIT.CORRESP.REQUEST

            DEP.COUNT += 1

            EEID.DEPNO.LIST\<-1\> = EEID.DEPNO       ; \* list of Deps

         NEXT DPNUM

      END

\*

\* Generate notice for Employee

      LETTER.TYPE = \'EE\'

      GOSUB SUBMIT.CORRESP.REQUEST

      RETURN                             ; \* GENERATE.LETTERS:

\*

SUBMIT.CORRESP.REQUEST:

\*

      IF LETTER.TYPE = \'DP\' THEN

\*

         \*\*\*\*\* Send to SmartComm - Separate notice to Spouse
\*\*\*\*\*

\*

         \* Build CORRESP.DATA.REC based on the

         \* CORRESP.INPUT.DEF record.

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

         DOCUMENT.NAME = DEP.TEMPLATE.NAME

         CALL
\*CREATE.CORRESP.TRACKING(ERROR.MESSAGES,CORRESP.DATA.REC,CORRESP.INPUT.DEF.ID,RECIPIENT.TYPE,RECIPIENT.ID,METHOD.REQUESTED,REQUEST.SOURCE,REQUESTOR.TRACKING.ID,DOCUMENT.NAME,\'\',\'\',\'\',\'\')

         IF ERROR.MESSAGES = \'\' THEN

            DEP.STORE.LIST\<-1\> = EEID.DEPNO

         END ELSE

            PRINT \'Correspondence request failed for Dependent
\[\':EEID.DEPNO:\'\]\'

            PRINT ERROR.MESSAGES\<1,2\>

         END

      END ELSE

\*

         \*\*\*\*\* Send to SmartComm  - Employee Notice \*\*\*\*\*

\*

         \* Determine if the Spouse\'s name should be printed

         \* on the letter.

         INCLUDE.SPOUSE.NAME = FALSE

         SPOUSE.DEP.NO = \'\'

         IF NOT(SENT.SPOUSE.NOTICE) AND
INDEX(EEMEDCOV\<1,EEIDX\>:EEDENCOV\<1,EEIDX\>,\'S\',1) \> 0 THEN

            MAX.DEP = DCOUNT(EEDEPNO,@VM)

            FOR DEPIDX = 1 TO MAX.DEP

               IF EERELATIONSHIP\<1,DEPIDX\> = \'S\' AND INDEX(\'
TWDPIR\',CONVERT(\'GL\',\'\',EESTATUS\<1,DEPIDX\>),1) \< 2 THEN

                  INCLUDE.SPOUSE.NAME = TRUE

                  SPOUSE.DEP.NO = EEID:\'.\':EEDEPNO\<1,DEPIDX\>

                  EXIT

               END

            NEXT DEPIDX

         END

         ERROR.MESSAGES = \'\'

         CORRESP.DATA.REC = \'\'

         CORRESP.DATA.REC\<2\> = FILENO

         IF INCLUDE.SPOUSE.NAME THEN

            \* This argument should only be set when sending a notice

            \* to the Employee and separate letter is not being sent

            \* to the spouse due to different address.

\* position below is based on CORRESP.INPUT.DEF Element/Fields

            CORRESP.DATA.REC\<5\> = SPOUSE.DEP.NO    ; \*
AdditionalEEIDDepno

            CORRESP.DATA.REC\<6\> = \'RECIPIENT\'      ; \*
AdditionalDependentPurpose

         END

         CORRESP.INPUT.DEF.ID = \'BASIC.SMARTCOMM.DOC.1\'

         RECIPIENT.TYPE = \'DEPENDENT\'

         RECIPIENT.ID = EMP.DEP.KEY

         METHOD.REQUESTED = \'LOOKUP\'

         REQUEST.SOURCE = PROGRAM.NAME

         REQUESTOR.TRACKING.ID = \'\'

         DOCUMENT.NAME = TEMPLATE.NAME

         CALL
\*CREATE.CORRESP.TRACKING(ERROR.MESSAGES,CORRESP.DATA.REC,CORRESP.INPUT.DEF.ID,RECIPIENT.TYPE,RECIPIENT.ID,METHOD.REQUESTED,REQUEST.SOURCE,REQUESTOR.TRACKING.ID,DOCUMENT.NAME,\'\',\'\',\'\',\'\')

         IF ERROR.MESSAGES = \'\' THEN

            STORE.LIST\<-1\> = EMP.DEP.KEY

         END ELSE

            PRINT \'Correspondence request failed for Employee
\[\':EEID:\'\]\'

            PRINT ERROR.MESSAGES\<1,2\>

         END

      END

      RETURN                             ; \* SUBMIT.CORRESP.REQUEST:

\*

DO.BILLING:

\* add manual billing entry to ERADJ file if needed

\* use top line eeeffdate - this may not bill when a group adds

\* cobra but no other solution

      LOCATE EEEFFDATE\<1,EEIDX\> IN EOEFFDATE\<1,1\> BY \'DR\' SETTING
EOIDX ELSE NULL

      IF EOEFFDATE\<1,EOIDX\> = \'\' AND EOIDX \> 1 THEN EOIDX -= 1

      IF EOCOBRA.NOTICE\<1,EOIDX\> \> 0 AND EEID\[1,1\] \<\> \'S\' THEN

         BILLPERIODO = CONVERT(\'-\',\'\',OCONV(ERBILLEDTO+1,\'D2-YM\'))

         SUFFIX=\'\'

         LOOP

            FILENO.KEY=EEFILENO\<1,EEIDX\>:BILLPERIODO:SUFFIX

            SUFFIX+=1

            CALL \*READREC(\'\',MAT EAREC,ERADJ,FILENO.KEY,\'WAIT\')

            IF \@USER.RETURN.CODE \# \'\' THEN

               RELEASE ERADJ,FILENO.KEY

               EXIT

            END

            IF EADESC\[13\] = \'COBRA NOTICES\' THEN

               EXIT

            END

         REPEAT

         COVPERIODO = \'\'

         LOCATE EEID IN EACOMMENTS\<1,1\> SETTING XXX ELSE

            INS EEID BEFORE EACOMMENTS\<1,-1\>

            EADESC = DCOUNT(EACOMMENTS,@VM):\' COBRA NOTICES\'

            COVPERIODO = BILLPERIODO

            COVPERIODO =
CONVERT(\'-\',\'\',OCONV(EEEFFDATE\<1,EEIDX\>,\'D2-YM\'))

         END

         EAIDX=1

\* kept until at top of loop because it was adding extra blank line
first

         LOOP

         UNTIL EADISTFIELD\<1,EAIDX\>=\'\' OR (EADISTFIELD\<1,EAIDX\> =
\'COBRACHG\' AND EACOVPERIOD\<1,EAIDX\> = COVPERIODO)

            EAIDX+=1

         REPEAT

         EADISTFIELD\<1,EAIDX\> = \'COBRACHG\'

         EADISTAMOUNT\<1,EAIDX\> += EOCOBRA.NOTICE\<1,EOIDX\>

         EACOVPERIOD\<1,EAIDX\> = COVPERIODO

         EAAMOUNT += EOCOBRA.NOTICE\<1,EOIDX\>

\* no trans.log needed

         CALL \*WRITEREC(\'\',MAT EAREC,ERADJ,FILENO.KEY,\'\')

      END

      RETURN                             ; \* DO.BILLING:

\*

END.REPORT.FILE:

\* close letter report file & list of store keys

\*

      IF STORE.LIST \# \'\' THEN

         LIST.NAME = \'STORE.\':UPCASE(TEMPLATE.NAME)

         WRITELIST STORE.LIST ON LIST.NAME

         PRINT \'Saved List \':LIST.NAME:\' created\'

      END

\*

      IF DEP.STORE.LIST \# \'\' THEN

\*

         PRINT \'contains \':DEP.COUNT:\' Dependent Alternate Address
notices\'

\*

         LIST.NAME = \'STORE.\':UPCASE(DEP.TEMPLATE.NAME)

         WRITELIST DEP.STORE.LIST ON LIST.NAME

         PRINT \'Saved List \':LIST.NAME:\' created\'

\*

         LIST.NAME = \'COBRA.DEP.NOTICE.\':UV.LOGNAME

         WRITELIST EEID.DEPNO.LIST ON LIST.NAME

         PRINT \'Saved List \':LIST.NAME:\' created\'

\*

      END

\*

      DATE.DONEO = CONVERT(\'-\',\'\',OCONV(TODAY,\'D4-YMD\'))

      TIME.DONEO = CONVERT(\':\',\'\',OCONV(TIME(),\'MTS\'))

      LIST.NAME =
\'COBRAN.DONE\':UV.LOGNAME:\'.\':DATE.DONEO:\'.\':TIME.DONEO

      WRITELIST COBRAN.DONE ON LIST.NAME

      PRINT \'Saved List \':LIST.NAME:\' created\'

\*

      LIST.NAME = \'COBRAN.DONE\':UV.LOGNAME

      WRITELIST COBRAN.DONE ON LIST.NAME

      PRINT \'Saved List \':LIST.NAME:\' created\'

\*

      RETURN                             ; \* END.REPORT.FILE:

\*

EXIT:

   END
