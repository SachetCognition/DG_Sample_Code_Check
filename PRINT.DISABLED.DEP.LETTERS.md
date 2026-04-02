# PRINT.DISABLED.DEP.LETTERS 

\* PRINT.DISABLED.DEP.LETTERS

\* Program to generate disable dependents letters

\* Dependents

\*

\* MAINTENANCE HISTORY:

\* 04-21-25 ABA RTC222761 Use SmartComm format - BASIC.SMARTCOMM.DOC.1

\* 10-17-24 ABA RTC212587 Do not call sub UPDATE.DEP.NEED when there is

\*                        a corresp tracking error

\* 08-30-24 SAC RTC212587 Move letters from the DG process to using the

\*                        CORRESP.TRACKING for the SmartComm process.

\* 02-21-23 MLZ RTC187698 Custom rule 161 for masterbill custom letters

\* 01-25-22 MCH RTC168107 Custom rule 151 for masterbill custom letters

\* 06-01-21 3LM RTC154558 Used custom rule 136 for alternate

\*                        version of Disabled Dep Denial Letter

\* 01-12-21 PTH RTC149118 Do not include \$ in LETCFORM value

\* 09-09-20 BEG RTC141159 Fixed unassigned variable error message and

\*                        BENTON job to run suceesfully

\* 07-29-20 MCH RTC135172 Adjust DATA calls to generate FORM.LET letters

\* 07-23-20 ABA RTC135172 Adjust DATA calls to generate FORM.LET letters

\* 06-02-20 ABA RTC134584 New Program

\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*

\* Main Program

\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*

      EQUATE PROGRAM.NAME TO \'PRINT.DISABLED.DEP.LETTERS\'

      OPENIT = \"Y\"

\$INCLUDE BP EMPLOYEES.INS

\$INCLUDE BP EMPLOYERS.INS

\$INCLUDE BP DEPENDENTS.INS

      CALL \*OPEN.FILE(\'\',\'&HOLD&\',HOLD)

      GOSUB INIT.VARS

      GOSUB GET.DEP.LIST

      GOSUB GET.BACKG

      GOSUB VALIDATE.DEP.DATA

      MAX.DEP.LIST = DCOUNT(DEP.LIST,@FM)

      IF MAX.DEP.LIST \> 0 THEN

         GOSUB GENERATE.DD.LETTERS

      END ELSE

         ERROR.TEXT = \'NO DEPENDENTS DEP.NEED POPULATED WITH DISABLED
DEPENDENTS CODES\'

         GOSUB ADD.TO.ERROR.LIST

      END

      IF RESTRICTED.ACCESS.GROUP.LIST \<\> \'\' THEN

         GOSUB SEND.RESTRICTED.ACCESS.GROUP.EMAIL

         GOSUB DISPLAY.RESTRICTED.ACCESS.GROUP.MESSAGE

      END

      CLOSESEQ SEQFILE

      IF CNT.ERR \> 0 THEN

         GOSUB SEND.ERROR.EMAIL

      END

      GOTO EXIT

\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*

\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\* Subroutines
\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*

\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*

INIT.VARS:

      TODAY = DATE()

      DATE.STAMP = CONVERT(\'-\',\'\',OCONV(TODAY,\'D2-MDY\'))

      START.TIME = CONVERT(\':\',\'\',OCONV(TIME(),\'MTS\'))

      TRUE = 1

      FALSE = 0

      DEP.LIST = \'\'

      DEP.LIST.CUR.DEP.NEED = \'\'

      RESTRICTED.ACCESS.GROUP.LIST = \'\'

      UV.LOGNAME = \'\'

      CALL \*GET.UV.LOGNAME(UV.LOGNAME,\'ACTUAL\')

      SLUNIT = SUBR(\'\*GET.FREE.LISTNO\')

      DEFFUN RECORD.EXISTS(A,B,C,D) CALLING \"\*VERIFY.RECORD.EXISTS\"

      ERROR.TEXT = \'\'

      CNT.ERR = 0

      BACKGROUND.PROMPTS = \'\'

      ERROR.FILENAME =
\'ERROR.\':PROGRAM.NAME:\'.\':DATE.STAMP:\'.\':START.TIME:\'.txt\'

      OPENSEQ \'&HOLD&\',ERROR.FILENAME TO SEQFILE ELSE NULL

      WEOFSEQ SEQFILE

      LTR.TYPES = \'RR\'

      LTR.TYPES.TEMPLATE = \'Disabled.Dependent.-.Review.Request\'

      LTR.TYPES\<-1\> = \'RC\'

      LTR.TYPES.TEMPLATE\<-1\> =
\'Disabled.Dependent.-.Recertification\'

      LTR.TYPES\<-1\> = \'IA\'

      LTR.TYPES.TEMPLATE\<-1\> =
\'Disabled.Dependent.-.Indefinite.Approval\'

      LTR.TYPES\<-1\> = \'LA\'

      LTR.TYPES.TEMPLATE\<-1\> =
\'Disabled.Dependent.-.Limited.Approval\'

      LTR.TYPES\<-1\> = \'DL\'

      LTR.TYPES.TEMPLATE\<-1\> = \'Disabled.Dependent.-.Denial\'

      LTR.TYPES\<-1\> = \'NR\'

      LTR.TYPES.TEMPLATE\<-1\> = \'Disabled.Dependent.-.No.Response\'

      PRINT @(-1):PROGRAM.NAME

      PRINT \"NOTE: DISABLED DEPENDENTS LETTERS WILL ONLY BE PRINTED IF
DEP.NEED IS POPULATED AND HAS NOT ALREADY BEEN GENERATED FOR TODAY\'S
DATE\"

      PRINT

      RETURN                             ; \* INIT.VARS:

GET.DEP.LIST:

\* DEP.NEED must be populated

\* - input EMPLOYEES ID and print qualifying dependents

\* - input single DEPENDENTS

      DEF.DEP.LIST = \'ALL\'

      IS.VALID = FALSE

      IS.SELECT.LIST = FALSE

      LOOP

         PRINT \'ENTER A DEPENDENT ID, (ALL) DEPENDENTS IDS, OR SELECT
LIST \':DEF.DEP.LIST:@(-9,LEN(DEF.DEP.LIST)+1):

         INPUT DEP.LIST

         DEP.LIST = UPCASE(DEP.LIST)

         IF DEP.LIST = \'\' THEN DEP.LIST = DEF.DEP.LIST

         IF DEP.LIST = \'Q\' THEN GOTO EXIT

         BEGIN CASE

            CASE DEP.LIST = \'ALL\'

               IS.VALID = TRUE

            CASE RECORD.EXISTS(\'DEPENDENTS\',DEP.LIST,DEPENDENTS,\'\')

               IS.VALID = TRUE

            CASE RECORD.EXISTS(\'&SAVEDLISTS&\',DEP.LIST,\'\',\'\')

               IS.VALID = TRUE

               IS.SELECT.LIST = TRUE

            CASE 1

               NULL

         END CASE

      UNTIL IS.VALID

         PRINT \'Enter a Valid DEPENDENT ID, (ALL) DEPENDENTS IDS, or
SAVEDLISTS or (Q)uit\'

      REPEAT

      BACKGROUND.PROMPTS := \'DATA \':DEP.LIST:@FM

      RETURN                             ; \* GET.DEP.LIST:

GET.BACKG:

\*\-\--get answer

      DEF.BACKG=\'Y\'

      LOOP

         PRINT \'PROCESS THIS JOB IN THE BACKGROUND (Y/N)
 \':DEF.BACKG:@(-9,2):

         INPUT BACKG

         BACKG=UPCASE(BACKG\[1,1\])

         IF BACKG=\'Q\' THEN GOTO EXIT

         IF BACKG=\'\' THEN BACKG=DEF.BACKG

      UNTIL BACKG = \'Y\' OR BACKG = \'N\'

         PRINT \'PLEASE ENTER Y, N OR Q\'

      REPEAT

\*\-\--create paragraph and submit to BENTON

      IF BACKG=\'Y\' THEN

         JOBNAME =PROGRAM.NAME

         JOBNAME:=\'.\':UV.LOGNAME

       
 JOBNAME:=\'\_\':CONVERT(\'-\',\'\',OCONV(TODAY,\'D2-YMD\')):\'\_\':CONVERT(\':\',\'\',OCONV(START.TIME,\'MTS\')):\'\_.COMI\'

         PARAGRAPH =\'PA\':@FM

         PARAGRAPH:=\'RUN BP \':PROGRAM.NAME:\' NOPAGE\':@FM

         PARAGRAPH:=BACKGROUND.PROMPTS

         PARAGRAPH:=\'DATA N\':@FM         ; \* background?

         PARAGRAPH:=\'PTIME\'

         CALL \*OPEN.FILE(\'\',\'VOC\',VOC)

         WRITE PARAGRAPH TO VOC,JOBNAME

         EXECUTE \'JOB \':JOBNAME

         GOTO EXIT

      END

      RETURN                             ; \* GET.BACKG:

VALIDATE.DEP.DATA:

      BEGIN CASE

         CASE DEP.LIST = \'ALL\'

            TEMP.DEP.LIST = \'\'

            GOSUB SELECT.ALL.DEP.NEED

            IF TEMP.DEP.LIST \# \'\' THEN

               DEP.LIST = TEMP.DEP.LIST

            END ELSE

               DEP.LIST = \'\'

            END

         CASE IS.SELECT.LIST

            GOSUB VALIDATE.LIST

            IF TEMP.DEP.LIST \# \'\' THEN

               DEP.LIST = TEMP.DEP.LIST

            END ELSE

               DEP.LIST = \'\'

            END

         CASE 1

            \* single dependent

            EEID.DEPNO = DEP.LIST

            GOSUB CHECK.DEP.NEED

            IF EEID.DEPNO \# \'\' THEN

               EEID = FIELD(EEID.DEPNO,\'.\',1)

               GOSUB CHECK.DEP.IS.RESTRICTED

               IF EEID.DEPNO \# \'\' AND MAX.CUR.DEP.NEED \> 0 THEN

                  DEP.LIST.CUR.DEP.NEED = CUR.DEP.NEED

               END ELSE

                  DEP.LIST = \'\'

               END

            END ELSE

               DEP.LIST = \'\'

            END

      END CASE

      RETURN                             ; \* VALIDATE.DEP.DATA:

SELECT.ALL.DEP.NEED:

      CMD = \'SELECT DEPENDENTS WITH DEP.NEED \# \"\" TO \': SLUNIT

      EXECUTE CMD

      IF \@SELECTED \> 0 THEN

         GOSUB VALIDATE.LIST

         IF TEMP.DEP.LIST \# \'\' THEN

            DEP.LIST = TEMP.DEP.LIST

         END ELSE

            DEP.LIST = \'\'

         END

      END ELSE

         DEP.LIST = \'\'

      END

      RETURN                             ; \* SELECT.ALL.DEP.NEED:

CHECK.DEP.NEED:

\* must populate eeid.depno before calling this sub

\*

      CUR.DEP.NEED = \'\'

      MAX.CUR.DEP.NEED = \'\'

      CALL \*READREC(\'\',MAT DPREC,DEPENDENTS,EEID.DEPNO,\'\')

      IF \@USER.RETURN.CODE = \'\' THEN

         CUR.DEP.NEED = DPDEP.NEED

         TEMP.CUR.DEP.NEED = CONVERT(\',\',@FM,CUR.DEP.NEED)

         NEW.CUR.DEP.NEED = \'\'

         IF TEMP.CUR.DEP.NEED \# \'\' THEN

            \* check to ensure only printing letters that have not been

            \* generated for today

            IF TODAY = DPDDL.SENT\<1,1\> THEN

               CUR.DPDDL.LETTERS =
CONVERT(\',\',@FM,DPDDL.LETTERS\<1,1\>)

               MAX.DEP.NEED = DCOUNT(TEMP.CUR.DEP.NEED,@FM)

               FOR DN.IDX=1 TO MAX.DEP.NEED

                  CUR.NEED = TEMP.CUR.DEP.NEED\<DN.IDX\>

                  LOCATE CUR.NEED IN CUR.DPDDL.LETTERS\<1\> SETTING XXX
ELSE

                     \* process only not in list

                     NEW.CUR.DEP.NEED\<-1\> = CUR.NEED

                  END

               NEXT DN.IDX

            END ELSE

               NEW.CUR.DEP.NEED = TEMP.CUR.DEP.NEED

            END

         END

         CUR.DEP.NEED = NEW.CUR.DEP.NEED

         MAX.CUR.DEP.NEED = DCOUNT(CUR.DEP.NEED,@FM)

         CONVERT \@FM TO \',\' IN CUR.DEP.NEED

      END ELSE

         EEID.DEPNO = \'\'

      END

      RETURN                             ; \* CHECK.DEP.NEED:

CHECK.DEP.IS.RESTRICTED:

\* check if user can access this dependent

\* must set EEID before calling this sub

\*

      CALL \*READREC(\'\',MAT EEREC,EMPLOYEES,EEID,\'\')

      IF \@USER.RETURN.CODE = \'\' THEN

         LOCATE TODAY IN EEEFFDATE\<1,1\> BY \'DR\' SETTING EEIDX ELSE
NULL

         IF EEEFFDATE\<1,EEIDX\> = \'\' AND EEIDX \> 1 THEN EEIDX -= 1

         CUR.FILENO = EEFILENO\<1,EEIDX\>

         PRINT.ERROR = \'\'

         CALL \*OUTSIDER.SECURITY(HAS.ACCESS,CUR.FILENO,PRINT.ERROR)

         IF NOT(HAS.ACCESS) THEN

            LOCATE CUR.FILENO IN RESTRICTED.ACCESS.GROUP.LIST\<1,1\> BY
\'AL\' SETTING RAG.IDX ELSE

               INS CUR.FILENO BEFORE
RESTRICTED.ACCESS.GROUP.LIST\<1,RAG.IDX\>

            END

            EEID.DEPNO = \'\'

            EEID = \'\'

         END

      END ELSE

         EEID.DEPNO = \'\'

         EEID = \'\'

      END

      RETURN                             ; \* CHECK.DEP.IS.RESTRICTED:

VALIDATE.LIST:

\* go thru list and filter deps that need letters

\*

      TEMP.DEP.LIST = \'\'

      IF IS.SELECT.LIST THEN

         GETLIST DEP.LIST TO SLUNIT ELSE NULL

      END

      LOOP

         READNEXT LIST.ID FROM SLUNIT ELSE EXIT

         EEID.DEPNO = LIST.ID

         GOSUB CHECK.DEP.NEED

         IF EEID.DEPNO \# \'\' THEN

            EEID = FIELD(EEID.DEPNO,\'.\',1)

            GOSUB CHECK.DEP.IS.RESTRICTED

            IF EEID.DEPNO \# \'\' AND MAX.CUR.DEP.NEED \> 0 THEN

               LOCATE EEID.DEPNO IN TEMP.DEP.LIST\<1\> BY \'AL\' SETTING
ED.IDX ELSE

                  INS EEID.DEPNO BEFORE TEMP.DEP.LIST\<ED.IDX\>

                  INS CUR.DEP.NEED BEFORE
DEP.LIST.CUR.DEP.NEED\<ED.IDX\>

               END

            END

         END

      REPEAT

      RETURN                             ; \* VALIDATE.LIST:

GENERATE.DD.LETTERS:

\*

\* based on values in DEP.LIST.CUR.DEP.NEED

\* each dependent is printed separately but certain letters

\* are printed together

\*

      CNT = 0

      FOR DL.IDX=1 TO MAX.DEP.LIST

         CNT += 1

         IF MOD(CNT,5000) = 0 THEN

            PRINT CNT:

         END ELSE

            IF MOD(CNT,1000) = 0 THEN

               PRINT \'\*\':

            END

         END

         EEID.DEPNO = DEP.LIST\<DL.IDX\>

         CALL \*READREC(\'\',MAT DPREC,DEPENDENTS,EEID.DEPNO,\'SKIP\')

         IF \@USER.RETURN.CODE = \'LOCKED\' THEN

            ERROR.TEXT = \'DEPENDENTS: \[\':EEID.DEPNO:\'\] IS LOCKED.
UNABLE TO SEND DISABLED DEPENDENTS LETTERS\'

            GOSUB ADD.TO.ERROR.LIST

            RELEASE DEPENDENTS,EEID.DEPNO

            CONTINUE

         END

         CUR.DEP.NEED = DEP.LIST.CUR.DEP.NEED\<DL.IDX\>

         CONVERT \',\' TO \@FM IN CUR.DEP.NEED

         MAX.CUR.DEP.NEED = DCOUNT(CUR.DEP.NEED,@FM)

\* Some forms require other data to be present, in order to output

\* that data in the body of the letter text.  Pull and validate

\* any required data now.

         LOCATE \'LA\' IN CUR.DEP.NEED\<1\> SETTING XXX THEN

            IF DPDISABILITY.EXPIRATION.DATE = \'\' THEN

               ERROR.TEXT = \'EEID.DEPNO \[\':EEID.DEPNO:\'\] DISABLED
DEP LETTER (Limited Approval) HAS NO DISABILITY EXPIRATION DATE\'

               GOSUB ADD.TO.ERROR.LIST

               RELEASE DEPENDENTS,EEID.DEPNO

               CONTINUE

            END

         END

         LOCATE \'NR\' IN CUR.DEP.NEED\<1\> SETTING XXX THEN

            GOSUB SET.LAST.RR.RC.SENT.DATEO

            IF LAST.RR.RC.SENT.DATEO = \'\' THEN

               ERROR.TEXT = \'EEID.DEPNO \[\':EEID.DEPNO:\'\] DISABLED
DEP LETTER (No Response) HAS NO LAST SENT DATE\'

               GOSUB ADD.TO.ERROR.LIST

               RELEASE DEPENDENTS,EEID.DEPNO

               CONTINUE

            END

         END

         EEID = FIELD(DEP.LIST\<DL.IDX\>,\'.\',1)

         CALL \*READREC(\'\',MAT EEREC,EMPLOYEES,EEID,\'\')

         IF \@USER.RETURN.CODE \# \'\' THEN

            ERROR.TEXT = \'EMPLOYEES: \[\':EEID:\'\] IS NOT LOADED\'

            GOSUB ADD.TO.ERROR.LIST

            RELEASE DEPENDENTS,EEID.DEPNO

            CONTINUE

         END

         LOCATE \'E\' IN EERELATIONSHIP\<1,1\> SETTING DEPIDX ELSE
DEPIDX = 1

         EMP.DEP.KEY = EEID:\'.\':EEDEPNO\<1,DEPIDX\>

\* set fileno

         LOCATE TODAY IN EEEFFDATE\<1,1\> BY \'DR\' SETTING EEIDX ELSE
NULL

         IF EEEFFDATE\<1,EEIDX\> = \'\' AND EEIDX \> 1 THEN EEIDX -= 1

         FILENO = EEFILENO\<1,EEIDX\>

         CALL \*READREC(\'\',MAT ERREC,EMPLOYERS,FILENO,\'\')

         IF \@USER.RETURN.CODE \# \'\' THEN

            ERROR.TEXT = \'EMPLOYERS: \[\':FILENO:\'\] IS NOT LOADED\'

            GOSUB ADD.TO.ERROR.LIST

            RELEASE DEPENDENTS,EEID.DEPNO

            CONTINUE

         END

\* if any errors for current letter need, do not update that letter

         ERROR.DEP.NEED = \'\'

         FOR CDN.IDX=1 TO MAX.CUR.DEP.NEED

            CUR.NEED = CUR.DEP.NEED\<CDN.IDX\>

            GOSUB SET.SMARTCOMM.TEMPLATE

            IF TEMPLATE.TO.USE \<\> \'\' THEN

               GOSUB SUBMIT.CORRESP.REQUEST

            END ELSE

               ERROR.TEXT = \'CANNOT DETERMINE SMARTCOMM TEMPLATE FOR
LETTER TYPE \[\':CUR.NEED:\'\]\'

               PRINT ERROR.TEXT

               GOSUB ADD.TO.ERROR.LIST

               ERROR.DEP.NEED\<-1\> = CUR.NEED

            END

         NEXT CDN.IDX

         GOSUB WRITE.DEP.REC

      NEXT DL.IDX

      RETURN                             ; \* GENERATE.DD.LETTERS:

ADD.TO.ERROR.LIST:

      WRITESEQ ERROR.TEXT TO SEQFILE ELSE NULL

      PRINT ERROR.TEXT

      CNT.ERR += 1

      RETURN                             ; \* ADD.TO.ERROR.LIST:

SET.SMARTCOMM.TEMPLATE:

      LOCATE CUR.NEED IN LTR.TYPES\<1\> SETTING LTRIDX THEN

         TEMPLATE.TO.USE = LTR.TYPES.TEMPLATE\<LTRIDX\>

      END ELSE

         TEMPLATE.TO.USE = \'\'

      END

      RETURN                             ; \* SET.SMARTCOMM.TEMPLATE:

SUBMIT.CORRESP.REQUEST:

      \* Build CORRESP.DATA.REC based on the

      \* CORRESP.INPUT.DEF of BASIC.SMARTCOMM.DOC.1.

      ERROR.MESSAGES = \'\'

      CORRESP.DATA.REC = \'\'

\* position below is based on CORRESP.INPUT.DEF Element/Fields

      CORRESP.DATA.REC\<2\> = FILENO       ; \* GroupNum

      CORRESP.DATA.REC\<5\> = EEID.DEPNO   ; \* AdditionalEEIDDepno

      CORRESP.DATA.REC\<6\> = \'SUBJECT\'    ; \*
AdditionalDependentPurpose

      CORRESP.INPUT.DEF.ID = \'BASIC.SMARTCOMM.DOC.1\'

      RECIPIENT.TYPE = \'DEPENDENT\'

      RECIPIENT.ID = EMP.DEP.KEY         ; \* employee depno

      METHOD.REQUESTED = \'LOOKUP\'

      REQUEST.SOURCE = PROGRAM.NAME

      REQUESTOR.TRACKING.ID = \'\'

      DOCUMENT.NAME = TEMPLATE.TO.USE

      CALL
\*CREATE.CORRESP.TRACKING(ERROR.MESSAGES,CORRESP.DATA.REC,CORRESP.INPUT.DEF.ID,RECIPIENT.TYPE,RECIPIENT.ID,METHOD.REQUESTED,REQUEST.SOURCE,REQUESTOR.TRACKING.ID,DOCUMENT.NAME,\'\',\'\',\'\',\'\')

      IF ERROR.MESSAGES \<\> \'\' THEN

         PRINT \'Correspondence request failed for Dependent
\[\':EEID.DEPNO:\'\] FOR LETTER TYPE \[\':CUR.NEED:\'\]\'

         ERROR.TEXT = ERROR.MESSAGES\<1,2\>

         PRINT ERROR.TEXT

         GOSUB ADD.TO.ERROR.LIST

         ERROR.DEP.NEED\<-1\> = CUR.NEED

      END ELSE

         GOSUB UPDATE.DEP.NEED

      END

      RETURN                             ; \* SUBMIT.CORRESP.REQUEST:

SET.LAST.RR.RC.SENT.DATEO:

      LAST.RR.RC.SENT.DATEO = \'\'

      MAX.DDL.SENT = DCOUNT(DPDDL.SENT,@VM)

      FOR DDL.IDX=1 TO MAX.DDL.SENT

         IF INDEX(DPDDL.LETTERS\<1,DDL.IDX\>,\'RR\',1) \> 0 THEN

            LAST.RR.RC.SENT.DATEO = DPDDL.SENT\<1,DDL.IDX\>

            EXIT

         END

         IF INDEX(DPDDL.LETTERS\<1,DDL.IDX\>,\'RC\',1) \> 0 THEN

            LAST.RR.RC.SENT.DATEO = DPDDL.SENT\<1,DDL.IDX\>

            EXIT

         END

      NEXT DDL.IDX

      IF LAST.RR.RC.SENT.DATEO \# \'\' THEN

         LAST.RR.RC.SENT.DATEO = OCONV(LAST.RR.RC.SENT.DATEO,\'D4MDY-\')

         CONVERT \'-\' TO \'/\' IN LAST.RR.RC.SENT.DATEO

      END

      RETURN                             ; \* SET.LAST.RR.RC.SENT.DATEO:

UPDATE.DEP.NEED:

      IF DPDDL.SENT\<1,1\> = TODAY THEN

         TEMP.DDL.LETTERS = CONVERT(\',\',@FM,DPDDL.LETTERS\<1,1\>)

         LOCATE CUR.NEED IN TEMP.DDL.LETTERS\<1\> SETTING XXX ELSE

            INS CUR.NEED BEFORE TEMP.DDL.LETTERS\<-1\>

         END

         DPDDL.LETTERS\<1,1\> = CONVERT(@FM,\',\',TEMP.DDL.LETTERS)

      END ELSE

         INS TODAY BEFORE DPDDL.SENT\<1,1\>

         INS CUR.NEED BEFORE DPDDL.LETTERS\<1,1\>

      END

      RETURN                             ; \* UPDATE.DEP.NEED:

WRITE.DEP.REC:

      MAX.ERR.NEED = DCOUNT(ERROR.DEP.NEED,@FM)

      IF MAX.ERR.NEED = 0 THEN

         DPDEP.NEED = \'\'

      END ELSE

\* only update letters that were generated to smartcomm

\* keep the letters that generated errors for later processing

         CONVERT \@FM TO \',\' IN ERROR.DEP.NEED

         DPDEP.NEED = ERROR.DEP.NEED

      END

      CALL \*WRITEREC(\'\',MAT
DPREC,DEPENDENTS,EEID.DEPNO,\'TRANS.LOG,\':PROGRAM.NAME)

      RETURN                             ; \* WRITE.DEP.REC:

SEND.RESTRICTED.ACCESS.GROUP.EMAIL:

      RESTRICTED.TYPE = \'Groups\'

      RESTRICTED.OPTIONS = \'\'

      CALL
\*EMAIL.RESTRICTED.NOTICE(RESTRICTED.ACCESS.GROUP.MESSAGE,RESTRICTED.TYPE,RESTRICTED.ACCESS.GROUP.LIST,PROGRAM.NAME,UV.LOGNAME,RESTRICTED.OPTIONS)

      RETURN                             ; \*
SEND.RESTRICTED.ACCESS.GROUP.EMAIL:

DISPLAY.RESTRICTED.ACCESS.GROUP.MESSAGE:

      PRINT

      MAX.LINE = DCOUNT(RESTRICTED.ACCESS.GROUP.MESSAGE,@VM)

      FOR LINE.CNT = 1 TO MAX.LINE

         PRINT RESTRICTED.ACCESS.GROUP.MESSAGE\<1,LINE.CNT\>

      NEXT LINE.CNT

      RETURN                             ; \*
DISPLAY.RESTRICTED.ACCESS.GROUP.MESSAGE:

SEND.ERROR.EMAIL:

      EMAIL.FROM = UV.LOGNAME

      RE.LINE = \'\*\*\*Errors Detected in Disabled Dependent Letter
Paperwork\*\*\*\'

      DISTRIBUTION = UV.LOGNAME

      EMAIL.MESSAGE = \'\'

      EMAIL.MESSAGE\<-1\> = \'Total number of errors: \' :CNT.ERR

      ATTACHMENTS = \'\'

      ATTACHMENTS\<-1\> = ERROR.FILENAME

      RE.OPTION = \'\'

      EMAIL.ERRMSG = \'\'

      CALL
\*EMAILGROUP.NOTICE(EMAIL.MESSAGE,PROGRAM.NAME,DISTRIBUTION,\'\')

      CALL
\*MAKE.EMAIL(EMAIL.ERRMSG,EMAIL.FROM,DISTRIBUTION,RE.LINE,EMAIL.MESSAGE,\'\',ATTACHMENTS,RE.OPTION)

      RETURN                             ; \* SEND.ERROR.EMAIL:

EXIT:

   END
