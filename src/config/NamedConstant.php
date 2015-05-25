<?php
/**
* NamedConstant
* This is a config class that can list all constants used by CymapGT
* services incase an implementing module would like to declare them
* prior to including the services file.
* @todo How to keep these declarations in sync with declaration in 
* services file???
*/
//MakerChecker service named constants
const MCHECKER_CREATED_PENDING_AUTHORIZATION   = 1;
const MCHECKER_CREATED_AND_AUTHORIZED          = 2;
const MCHECKER_UPDATED_PENDING_AUTHORIZATION   = 3;
const MCHECKER_UPDATED_AND_AUTHORIZED          = 4;
const MCHECKER_CANCELLED_PENDING_AUTHORIZATION = 5;
const MCHECKER_CANCELLED_AND_RETURNED          = 6;
const MCHECKER_CANCELLED_AND_AUTHORIZED        = 7;
const MCHECKER_CREATED_AND_REJECTED            = 8;
const MCHECKER_UPDATED_AND_RETURNED            = 9;
const MCHECKER_UPDATED_AND_REJECTED            = 10;
const MCHECKER_CREATED_AND_RETURNED            = 11;
const MCHECKER_CANCELLED_AND_REJECTED          = 12;
const MCHECKER_TRANSACTION_PAUSED              = 13;

//define MakerChecker audit var markers as constants
const MCHECKERAUDIT_PROFILE   = 1;
const MCHECKERAUDIT_LASTUSR   = 2;
const MCHECKERAUDIT_LASTSTATE = 3;

//define MakerChecker validation error codes as constants
const MCHECKERVALIDATION_NOMAKERUPDATE  = 100;
const MCHECKERVALIDATION_NOOPTIONSFOUND = 101;

/*define maker checker implementation codes for actions. can be AUTH, DENY, PAUS or RETN
AUTH - Authorize the action, state changes to next state as per state change array
DENY - Deny the action, the object is returned to former state, and can be accessed again
PAUS - Do not change object state, or data, but lock it from any other changes  
RETN - Returned transaction is not returned to former state but returned to maker to make some corrections
*/   
const MCHECKERIMPLEMENT_AUTH = 1000;
const MCHECKERIMPLEMENT_DENY = 1001;
const MCHECKERIMPLEMENT_PAUS = 1002;
const MCHECKERIMPLEMENT_RETN = 1003;

//precision of math calculations
const MATHFINANCE_PRECISION         = 1E-6;
//payment types
const MATHFINANCE_PAYEND            = 0;
const MATHFINANCE_PAYBEGIN          = 1;
//day count methods
const MATHFINANCE_COUNTNASD         = 0;
const MATHFINANCE_COUNTACTUALACTUAL = 1;
const MATHFINANCE_COUNTACTUAL360    = 2;
const MATHFINANCE_COUNTACTUAL365    = 3;
const MATHFINANCE_COUNTEUROPEAN     = 4;

const MATHFINANCE_COUNTNEW_GERMAN     = 1;
const MATHFINANCE_COUNTNEW_GERMANSPEC = 2;
const MATHFINANCE_COUNTNEW_ENGLISH    = 3;
const MATHFINANCE_COUNTNEW_FRENCH     = 4;
const MATHFINANCE_COUNTNEW_US         = 5;
const MATHFINANCE_COUNTNEW_ISMAYEAR   = 6;
const MATHFINANCE_COUNTNEW_ISMA99N    = 7;
const MATHFINANCE_COUNTNEW_ISMA99U    = 8;

const MATHFINANCE_ERROR_BADDCM        = 1;
const MATHFINANCE_ERROR_BADDATES      = 2;

//UserCredential constants for user authentication
const USERCREDENTIAL_ACCOUNTSTATE_LOGGEDOUT   = 1;
const USERCREDENTIAL_ACCOUNTSTATE_LOGGEDIN    = 2;
const USERCREDENTIAL_ACCOUNTSTATE_LOCKED1     = 3;
const USERCREDENTIAL_ACCOUNTSTATE_LOCKED2     = 4;
const USERCREDENTIAL_ACCOUNTSTATE_RESET       = 5;
const USERCREDENTIAL_ACCOUNTSTATE_SUSPENDED   = 6;
const USERCREDENTIAL_ACCOUNTSTATE_AUTHFAILED  = 7;
const USERCREDENTIAL_ACCOUNTSTATE_LOGGEDINTEMP= 8;

//UserCredential constants for account policy actions
const USERCREDENTIAL_ACCOUNTPOLICY_VALID         = 1;
const USERCREDENTIAL_ACCOUNTPOLICY_EXPIRED       = 2;
const USERCREDENTIAL_ACCOUNTPOLICY_ATTEMPTLIMIT1 = 3;
const USERCREDENTIAL_ACCOUNTPOLICY_ATTEMPTLIMIT2 = 4;
const USERCREDENTIAL_ACCOUNTPOLICY_REPEATERROR   = 5;

//Multiprocessing constants for multiproc service
const MULTIPROCESSING_DEFAULT_TIMELIMIT = 30;				// Sets the default timeout for the parent and all children
const MULTIPROCESSING_CACHE_METHOD      = "mysql";			// Accepts either mysql or sqlite
const MULTIPROCESSING_DB_NAME           = "fndmgr191a";			// The name of the database for either sqlite or MySQL
const MULTIPROCESSING_SQLITE_DIRECTORY  = "sqlite";			//
const MULTIPROCESSING_DB_USER           = "root";			// The name of the database user
const MULTIPROCESSING_DB_PASSWORD       = "irungu";                     // The database user's password
const MULTIPROCESSING_DB_HOST           = "localhost";			// The database host. Anything but localhost would be a bad idea for PHP-multi process

//Spreadsheet process or constants for PHPExcel wrapper
/** CYMAPGT_PHPExcelWrapperFactory root directory */
const SPREADSHEETPROCESSOR_CACHEGETALL = 0;
const SPREADSHEETPROCESSOR_CACHEGETCUR = 1;
const SPREADSHEETPROCESSOR_CACHEGETVAL = 2;

//Math Finance. Accrued interest constants
const MATHFINANCE_SWX_BOND_AI_GERMAN      = 1;
const MATHFINANCE_SWX_BOND_AI_SPEC_GERMAN = 2;
const MATHFINANCE_SWX_BOND_AI_ENGLISH     = 3;
const MATHFINANCE_SWX_BOND_AI_FRENCH      = 4;
const MATHFINANCE_SWX_BOND_AI_US          = 5;
const MATHFINANCE_SWX_BOND_AI_ISMA_YEAR   = 6;
const MATHFINANCE_SWX_BOND_AI_ISMA_99N    = 7;
const MATHFINANCE_SWX_BOND_AI_ISMA_99U    = 8;
const MATHFINANCE_SWX_BOND_AI_KENYA       = 9;
const MATHFINANCE_SWX_BOND_AI_CBK_KENYA   = 10;

//Job Scheduler. Service control and state constants
const SCHEDULERSERVICE_JOBSTART      = 0;
const SCHEDULERSERVICE_JOBSTOP       = 1;
const SCHEDULERSERVICE_JOBPAUSE      = 2;
const SCHEDULERSERVICE_JOBPAUSEUNTIL = 3;
const SCHEDULERSERVICE_JOBMARKASDONE = 4;
const SCHEDULERSERVICE_JOBCALCULATENEXTRUNTIME = 5;
const SCHEDULERSERVICE_JOBNEXTRUNTIME = 6;
const SCHEDULERSERVICE_JOBNEXTALTERNATIVERUNTIME = 7;
const SCHEDULERSERVICE_JOBDONE = 8;
const SCHEDULERSERVICE_JOBFAILED = 9;

const SCHEDULERSERVICEJOBSTATE_IDLE        = 0;
const SCHEDULERSERVICEJOBSTATE_STARTED     = 1;
const SCHEDULERSERVICEJOBSTATE_RUNNING     = 2;
const SCHEDULERSERVICEJOBSTATE_PAUSED      = 3;
const SCHEDULERSERVICEJOBSTATE_STOPPED     = 5;
const SCHEDULERSERVICEJOBSTATE_FAILED      = 4;
const SCHEDULERSERVICEJOBSTATE_PAUSEDUNTIL = 5;

const LOGGER_STDERR   = 0;
const LOGGER_LEVEL1   = 1;
const LOGGER_LEVEL2   = 2;
const LOGGER_LEVEL3   = 3;
const LOGGER_SECURITY = 4;
