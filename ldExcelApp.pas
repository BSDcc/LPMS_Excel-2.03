//===========================================================================
// Program ID...: LegalDiary
// Author.......: Francois De Bruin Meyer
// Copyright....: BlueCrane Software Development CC
// Date.........: 03 January 2010
//---------------------------------------------------------------------------
// Description..: Excel Interface
//---------------------------------------------------------------------------
// Changes......:
//===========================================================================

unit ldExcelApp;

//---------------------------------------------------------------------------
// Auto definitions
//---------------------------------------------------------------------------
interface

uses
  {Windows,} Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ToolWin, ComCtrls, Buttons, DB, {ADODB, ShellApi,}
  ExtCtrls, ImgList, ActnList, {nExcel,} StrUtils, {xmldom, XMLIntf, msxmldom,
  XMLDoc,} JvExStdCtrls, JvListBox, Registry, Math, JvComponentBase, JvCipher,
  DateUtils, Printers, SMAPI, IdMessage, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase, IdMessageClient,
  IdSMTPBase, IdSMTP, IdAttachmentFile, IdMessageCoder, IdMessageCoderMIME,
  Mask, UbuntuProgress;

type
   TFldExcel = class(TForm)
      Query1: TADOQuery;
      MySQLCon: TADOConnection;
      Query2: TADOQuery;
      Bevel1: TBevel;
      StatusBar: TStatusBar;
      pnlTop: TPanel;
      Image1: TImage;
      btnCancel: TButton;
      txtDocument: TStaticText;
      txtError: TEdit;
      pnlBottom: TPanel;
      XMLDoc: TXMLDocument;
      jvCipher: TJvVigenereCipher;
      emlSession: TSMapiSession;
      emlSend: TSMapiSendMail;
      memSignature: TMemo;
      lbProgress: TListBox;
      idSMTP: TIdSMTP;
      idMessage: TIdMessage;
      prbProgress: TUbuntuProgress;
      stCount: TStaticText;

      procedure FormActivate(Sender: TObject);
      procedure FormCreate(Sender: TObject);
      procedure btnCancelClick(Sender: TObject);
      procedure FormClose(Sender: TObject; var Action: TCloseAction);
      procedure txtErrorChange(Sender: TObject);

type
   LPMS_Amounts = record
      Fees                              : double;
      Disbursements                     : double;
      Expenses                          : double;
      Payment_Received                  : double;
      Business_To_Trust                 : double;
      Credit                            : double;
      Business_Deposit                  : double;
      Trust_Deposit                     : double;
      Trust_Transfer_Business_Fees      : double;
      Trust_Transfer_Disbursements      : double;
      Trust_Transfer_Client             : double;
      Trust_Transfer_Trust              : double;
      Trust_Investment_S86_4            : double;
      Trust_Withdrawal_S86_4            : double;
      Trust_Interest_S86_4              : double;
      Trust_Investment_S86_3            : double;
      Trust_Withdrawal_S86_3            : double;
      Trust_Interest_S86_3              : double;
      Business_Debit                    : double;
      Trust_Debit                       : double;
      Trust_Interest_Withdrawal_S86_3   : double;
      Write_off                         : double;
      Collection_Debit                  : double;
      Collection_Credit                 : double;
      Reserved_Trust                    : double;
      Trust_Transfer_Business_Other     : double;
      Trust_FF_Interest_S86_4           : double;
   end;

   LPMS_SymVars = record
      Scope    : AnsiString;
      Variable : AnsiString;
      Value    : AnsiString;
   end;

   SymVars_Table = record
      SV : array of LPMS_SymVars;
   end;

   LPMS_BI = record
      FileName     : AnsiString;
      Description  : AnsiString;
      StartDate    : AnsiString;
      EndDate      : AnsiString;
      Invoice      : AnsiString;
      Statement    : AnsiString;
      StoreInvoice : AnsiString;
      SendEmail    : AnsiString;
      CreatePDF    : AnsiString;
      Print        : AnsiString;
      ShowRelated  : AnsiString;
      AutoOpen     : AnsiString;
      EmailTo      : AnsiString;
      CC           : AnsiString;
      BCC          : AnsiString;
      Subject      : AnsiString;
      EmailBody    : AnsiString;
      EditEmail    : AnsiString;
      ReadReceipt  : AnsiString;
      AcctType     : AnsiString;
   end;

   BI_Table = record
      BI : array of LPMS_BI;
   end;

   LPMS_Statement = record
      DateTime    : AnsiString;
      Date        : AnsiString;
      Description : AnsiString;
      S864        : double;
      Trust       : double;
      Business    : double;
   end;

type
   SV_VALS = (SV_BACKUPTYPE,SV_BUYER,SV_CLIENT,SV_COMBINED,SV_CPYFILE,
              SV_CPYNAME,SV_CURRDATE,SV_DATE,SV_DAY,SV_DBPREFIX,SV_DESC,
              SV_DIALCODE,SV_DURATION,SV_EARNER,SV_EMAIL,SV_ENDDATE,SV_FILE,
              SV_FOLDER,SV_FROMDATE,SV_HOSTNAME,SV_HOUR,SV_INVOICE,SV_LONGMONTH,
              SV_MAILNOTICE,SV_MINUTE,SV_MONTH,SV_NUMBER,SV_OPPOSE,SV_PREFIX,
              SV_PRESCRIPTD,SV_RATE,SV_RESULT,SV_ROOTF,SV_SELLER,SV_SERIAL,
              SV_SHORTYEAR,SV_SHORTTIME,SV_SUBJECT,SV_SUFFIX,SV_TEXT,SV_TIME,
              SV_TO,SV_TYPEFILE,SV_UNITS,SV_USER,SV_VATNUM,SV_VERSION,SV_YEAR,
              SV_AMP,SV_COUNT);

type
   Process_Types = (PT_NORMAL, PT_BILLING, PT_PROLOG, PT_MIDLOG, PT_OPTLOG,
                    PT_EPILOG);

type
   PageBreak_Types = (PB_TRUSTDETAIL, PB_TRUSTSUMMARY, PB_TRUSTTOTALS,
                      PB_TRUSTS864, PB_STATEMENT, PB_FILENOTES,
                      PB_FILEDETAILS, PB_FEEEARNERS, PB_FEECONS,
                      PB_ACCOUNTANT01, PB_ACCOUNTANT02, PB_ACCOUNTANT03,
                      PB_ALERTS, PB_PHONEBOOK, PB_TRUSTMAN, PB_PAYMENTS01,
                      PB_PAYMENTS02, PB_TASKLIST, PB_QUOTELIST);

type
   Document_Types = (DT_ACCOUNTANT, DT_PAYMENT, DT_INVOICES, DT_STATEMENT);

type
   Cypher_Type = (CYPHER_ENC, CYPHER_DEC);

  private
    SMTPAuthType : integer;            // SMTP Auth type to be used when using buit-in Email interface
    SMTPAuth     : boolean;            // SMTP Server requires authentication when using buit-in Email interface
    FirstRun     : boolean;            // Flag to deal with unsolicited calls to FormActivate
    SecretPhrase : string;             // Used for decoding the DBPrefix
    SMTPServer   : string;             // SMTP Server name when using buit-in Email interface
    SMTPPort     : string;             // SMTP Port when using buit-in Email interface
    SMTPUser     : string;             // SMTP User name when using buit-in Email interface
    SMTPPass     : string;             // SMTP Password when using buit-in Email interface
    Bill_Instr   : BI_Table;           // Used to hold the Billing Instructions

//    procedure VigenereE(var Str: string; const Key: string);
//    procedure VigenereD(var Str: string; const Key: string);
    procedure Generate_Billing();
    procedure Generate_Accounts();
    procedure Generate_DocHeading(xlsSheet: IXLSWorksheet; PageNum: integer; idx1: integer; Row: integer; Col: integer; ThisType: string);
    procedure Generate_DataHeading(xlsSheet : IXLSWorksheet; Row: integer; Col: integer; ThisType: string);
    procedure Generate_Detail_Account(xlsSheet : IXLSWorksheet; PageRows: integer; ThisRows: integer; ThisItem: string; BalanceStr: string);
    procedure Generate_CallRecs();
    procedure Generate_Notes();
    procedure Generate_Invoice();
    procedure Generate_Detail_Invoice(xlsSheet : IXLSWorksheet; PageRows: integer; ThisRows: integer; ThisItem: string; BalanceStr: string);
    procedure Generate_Detail_Quote(xlsSheet : IXLSWorksheet; PageRows: integer; ThisRows: integer; ThisItem: string; BalanceStr: string);
    procedure Generate_Statement();
    procedure Generate_Statement_Summary(xlsSheet : IXLSWorksheet; PageRows: integer; ThisRows: integer);
    procedure Generate_Trust_Summary();
    procedure DoTrust_All();
    procedure DoPageBreak (ThisType: integer; xlsSheet: IXLSWorksheet; var FirstPage: boolean; var RowsPerPage: integer; var row: integer; var Pages: integer; var PageRow: integer; var PageBreak: boolean; ThisFile: string);
    procedure DoTrust_Simple();
    procedure Generate_Detail_Trust(xlsSheet : IXLSWorksheet; PageRows: integer; ThisRows: integer; ThisItem: string; BalanceStr: string; ThisFile: string);
    procedure Generate_S863_Summary();
    procedure Generate_FeeReport();
    procedure Generate_AcctReport();
    procedure AcctReport_Line1(xlsSheet: IXLSWorksheet; ThisFile: string; ThisStr: string; ThisAmt: double; ThisExtra : string; ThisFill: TColor; ThisText: TColor; var FirstPage: boolean; var RowsPerPage: integer; var row: integer; var Pages: integer; var PageRow: integer; var PageBreak: boolean);
    procedure AcctReport_Line2(ThisVariant: integer; xlsSheet: IXLSWorksheet; ThisFile: string; ThisStr: string; ThisBold: boolean; ThisAmt: double; ThisExtra : string;  ThisFill: TColor; ThisText: TColor; var FirstPage: boolean; var RowsPerPage: integer; var row: integer; var Pages: integer; var PageRow: integer; var PageBreak: boolean);
    procedure Generate_Alerts();
    procedure Generate_Phonebook();
    procedure Generate_BillingReport();
    procedure Billng_Prep_PageBreak(xlsSheet: IXLSWorksheet; var PageBreak: boolean; var FirstPage: boolean; SaveEDate: string; MonthC: string; Month1: string; Month2: string; Month3: string; Month4: string; Month5: string; Month6: string; var RowsPerPage: integer; var row: integer; var PageRow: integer; var Pages: integer);
    procedure Generate_ClientDetails();
    procedure Generate_FileDetails();
    procedure Generate_CollectAcct();
    procedure Generate_InvoiceList();
    procedure Generate_PrefixDetails();
    procedure Generate_SafeDetails();
    procedure Generate_DocCntlDetails();
    procedure Generate_UserDetails();
    procedure Generate_Quotation();
    procedure Generate_TaskList();
    procedure Generate_TrustManagement();
    procedure Generate_PeriodicBilling();
    procedure Generate_SingleBilling();
    procedure Generate_Payments();
    procedure Generate_QuoteDetails();
    procedure Generate_LogDetails();
    procedure Generate_PrintSheets();
    procedure Create_PrintSheet(SheetName: string; ThisTemplate: string; Rows: integer; Cols: integer);
    procedure GetLayout(RunTimeTemplate: string);
    procedure Close_Connection();
    procedure GetCpyVAT();
    procedure GetAddress(FileName: string);
    procedure SetDefaultPrinter(NewDefPrinter: string);
    procedure Clear_Amounts();
    procedure GetAmounts(ThisType: integer; adoQry: TADOQuery);
    procedure LogMsg(ThisType: integer; ThisInit: boolean; ShowOptions: boolean; ThisMsg: string); overload;
    procedure LogMsg(ThisMsg: string; DoAdjust: boolean); overload;
    procedure BubbleSort(var RecList: Array of LPMS_Statement);

//    function  GetTrust(FileName: string): boolean;
//    function  DoTrust_PageBreak(xlsSheet1 : IXLSWorksheet) : integer;
//    function  DoSymVars(ThisStr: AnsiString; FileStr: AnsiString): AnsiString;
//    function  Imod(const x: integer; const y: integer): integer;
//    function  VigenereEnc(const Str: string; const Key: string): string;
//    function  VigenereDec(const Str: string; const Key: string): string;
    function  Open_Connection(OpenHost: string): boolean;
    function  Generate_Statement_Detail(xlsSheet : IXLSWorksheet; ThisFile : string) : boolean;
    function  GetBilling(FileName: string; ThisStr: string; ThisType: string): boolean;
    function  GetQuote(FileName: string; QuoteName: string): boolean;
    function  GetQuoteVal(Quote: string; ThisHost: string): double;
    function  GetBalance(FileName: string): boolean;
    function  GetCurrent(FileName: string): boolean;
    function  GetQuoteFile(QuoteName: string): string;
    function  GetQuoteDetails(QuoteName: string): boolean;
    function  GetBillingR(ThisFile: string; StartDate: string; EndDate: string; var Billing: double; var Invoiced: double; var Paid: double; var ThisReserved: double; var ThisTrust: double): boolean;
    function  GetInvoiceData(Invoice: string): string;
    function  GetS863(FileName: string; StartDate: string; EndDate: string): boolean;
    function  GetRecord(S1: string; QryType: integer): boolean;
    function  GetTotals(S1: string): double;
    function  GetNotes(FileName: string): boolean;
    function  GetFeeRecs(FileName: string; ThisStr : string): boolean;
    function  GetFileDesc(FileName: string): string;
    function  GetAddresses(FileName: string): string;
    function  GetClient(FileName : string): string;
    function  GetAlertRecs(ThisUser: string; ThisFilter : string): boolean;
    function  GetVATNum(FileName: string): string;
    function  GetDescription(FileName: string): string;
    function  GetUser(UserID: string): boolean;
    function  GetAllUsers(ThisStr: string): boolean;
    function  GetFileNames(Filter: string): boolean;
    function  GetPhoneRecs(Filter1: string; Filter2: string): boolean;
    function  GetClientDetails(Filter: string; ThisType: integer): boolean;
    function  GetFileDetails(TimeStamp: string; ThisType: integer): string;
    function  SendEmail(FileRef: string; FileName: string; ProcessType: integer): boolean;
    function  StoreInvoice(ThisFile: string; Amount: double; Fees: double; Disburse: double; Expenses: double): boolean;
    function  GetPayments(ThisInvoice: string): double;
    function  GetInvoices(ThisType: integer; ThisFile: string): boolean;
    function  GetNextInvoice(QryType: integer): boolean;
    function  ReplaceQuote(S1: string): string;
    function  ReplaceXML(S1: string): string;
    function  RoundD(x: Extended; d: integer): extended;
    function  GetDefaultPrinter: string;
    function  PDFDocument(FileName: string): boolean; overload;
    function  PDFDocument(FileName: string; FolderName: string): boolean; overload;
    function  PrintDocument(FileName: string): boolean; overload;
    function  PrintDocument(FileName: string; FolderName: string): boolean; overload;
    function  DoS86_PageBreak(xlsSheet : IXLSWorksheet) : integer;
    function  Disassemble(Str: string; ThisDelim: char): TStringList;
    function  Vignere(ThisType: integer; Phrase: string; Key: string): string;

  public
    RunType        : integer;          // Type of report to run
    GetEmailResult : boolean;          // Holds result from ldGetMailP
    EmailStr       : string;           // List of emaill addresses from ldGetEmailP
    EmailBody      : string;           // Default Email Body when sending emails from ldExcelApp
    ThisSubject    : string;           // Subject to be send to ldGetEmailP
    RegString      : string;           // Pointer to the Registry - Provides Multi Company Support

    This_Bal       : LPMS_Amounts;     // Actual amounts for each Billing Class in a Billing Records Set
    This_Abs       : LPMS_Amounts;     // ABS amounts for each Billing Class in a Billing Records Set
    SymVars_LPMS   : SymVars_Table;    // Used to hold System defined Symbolic Variables
    SymVars_Other  : SymVars_Table;    // Used to hold Local and Global Symbolic Variables

    EMailStrings   : TStringList;      // Holds the User's untangled Email signature

    procedure SetUpSymVars();
    function  DoSymVars(ThisStr: AnsiString; FileStr: AnsiString): AnsiString;

  const

    ThisLabels : array[1..30] of string = (
                 'Invoice & Statement', 'Specified Account', 'Call Records',
                 'File Notes', 'Invoice', 'Statement', 'Trust Account Summary',
                 'Section 86(3) Investment Summary', 'Fee Earners Report',
                 'Accountant Report', 'Diary, Daily and Prescription Alerts',
                 'Phonebook Export', 'Billing Preparation Report',
                 'Customer Details Export', 'Opposition Details Export',
                 'File Details Export', 'Collection Account',
                 'Invoice List Export', 'Prefix Details Export',
                 'Safe Keeping Details Export',
                 'Document Control Details Export', 'Log Details',
                 'User Details', 'Quote', 'Reserved', 'Periodic Billing',
                 'Trust Managtement Report', 'Payment Details Export',
                 'Quote Details Export', 'New File Documents');

    ThisDebug  : array[1..27] of string = (
                 'RunType', 'Host Name', 'Start Date', 'End Date',
                 'File Name', 'Calling HWND', 'Parm 07', 'Show Related',
                 'Show Nil Balance', 'Parm 10', 'Open on Complete',
                 'Close On Complete', 'Parm 13', 'LPMS Heading',
                 'Invoice Information', 'Store Invioce','Send By Email',
                 'Email Address List', 'Create PDF', 'Run In Background',
                 'Registry String', 'User Name', 'Print this Sheet',
                 'LPMS Version', 'Print Sheets', 'Number of Files',
                 'First File Name');

end;

//---------------------------------------------------------------------------
// Global definitions
//---------------------------------------------------------------------------
var
//   TrustPageRow       : integer;       // Used for Page break control when doing a Trust recon
//   TrustRowsPerPage   : integer;       // Used for Page break control when doing a Trust recon
//   VATAmount          : double;        // Used for calculating VAT on Fees
//   DispLess           : boolean;       // Controls initial state of the display screen
//   TrustFirstPage     : boolean;       // Used for Page break control when doing a Trust recon
//   TrustPageBreak     : boolean;       // Used for Page break control when doing a Trust recon
//   Opposition         : string;        // Current Opposition
//   AmtTrustReserve    : double;        // Holds the reserved Trust Deposits
//   AmtInv782Sa        : double;        // Holds the S78(2)(a) Investment totals
//   AmtDrw782Sa        : double;        // Holds the S78(2)(a) Withdrawal totals
//   AmtInt782Sa        : double;        // Holds the S78(2)(a) Accrued Interest totals
//   AmtIntDrw782Sa     : double;        // Holds the S78(2)(a) Interest Withdrawal totals
//   AmtWriteOff        : double;        // Holds the Write-off total
//   BalFees            : double;        // Holds the Fees opening balance
//   BalDisburse        : double;        // Holds the Disbursements opening balance
//   BalExpense         : double;        // Holds the Expenses opening balance
//   BalPayment         : double;        // Holds the Payments Received opening balance
//   BalBusToTrust      : double;        // Holds the Payments Made opening balance
//   BalCredit          : double;        // Holds the Credit granted opening balance
//   BalBusDeposit      : double;        // Holds the Deposits to Business opening balance
//   BalTrustDeposit    : double;        // Holds the Deposits to Trust opening balance
//   BalTrustReserve    : double;        // Holds the Reserved Trust opening balance
//   BalXfrFees         : double;        // Holds the transfers from Trust to Fees opening balance
//   BalXfrDisburse     : double;        // Holds the transfers from Trust to Disbursements opening balance
//   BalXfrClient       : double;        // Holds the transfers from Trust to Clients opening balance
//   BalXfrTrust        : double;        // Holds the transfers from Trust to Trust opening balance
//   BalInv782BA        : double;        // Holds the S78(2A) Investment opening balance
//   BalDrw782BA        : double;        // Holds the S78(2A) Withdrawal opening balance
//   BalInt782BA        : double;        // Holds the S78(2A) Accrued Interest opening balance
//   BalBusDebit        : double;        // Holds the Business Debits opening balance
//   BalTrustDebit      : double;        // Holds the Trust Debits opening balance

   NumParms           : integer;       // Number of Parameters passed
   NumFiles           : integer;       // Holds number of Files passed
   CallHWND           : integer;       // Holds Handle of calling Form
   ShowVAT            : integer;       // Controls whether VAT is added to Fees on Invoices
   AccountType        : integer;       // Controls the type of billing records that will be used for invoices
   InvoiceNum         : integer;       // Holds the next sequential Invoice number
   PDFInterval        : integer;       // Holds Interval time for PDF Creation
   PDFRetry           : integer;       // Holds the number of times moving the PDF will be retried
   TrustPages         : integer;       // Used for Page break control when doing a Trust recon
   S86Pages           : integer;       // Used for Page break control when doing a S86(3) recon
   S86PageRow         : integer;       // Used for Page break control when doing a S86(3) recon
   S86RowsPerPage     : integer;       // Used for Page break control when doing a S86(3) recon
   PDFMergeSel        : integer;       // Enumerates the default PDF Merge Utility
   ThisCount          : integer;       // Used to display the progress of prbProgress in numbers

   ColAB1F            : TColor;        // Fill Colour for the 1st Alternative Block
   ColAB1T            : TColor;        // Text Colour for the 1st Alternative Block
   ColAB2F            : TColor;        // Fill Colour for the 2nd Alternative Block
   ColAB2T            : TColor;        // Text Colour for the 2nd Alternative Block
   ColDHF             : TColor;        // Fill Colour for the Data Headings
   ColDHT             : TColor;        // Text Colour for the Data Headings

   VATRate            : double;        // VAT Rate

   Debug              : boolean;       // Used for debugging and testing
   AutoOpen           : boolean;       // Controls whether export file is opened or not
   CloseOnComplete    : boolean;       // Controls auto close at completion
   CreateInvoice      : boolean;       // Controls whether an Invoice is also created
   CreateStatement    : boolean;       // Controls whether Invoice and Statement are grouped in one email
   NegativeRed        : boolean;       // Controls whether negative numbers are shown in Red or not
   ShowRelated        : boolean;       // Controls whether related billing or actual billing is shown
   IncludeTrust       : boolean;       // Controls whether the Trust balance is included in the Age Analysis
   ExcludeReserve     : boolean;       // Controls whether the Reserved Trust balance is excluded when the Trust balance is included in the Age Analysis
   StoreInv           : boolean;       // Controls whether the Invoice is allocated a sequential Invoice number and stored
   CreatePDF          : boolean;       // Controls whether a PDF is created from the resulting Excel file
   PDFExists          : boolean;       // If true then the PDF file was successfully created
   PDFPrefer          : boolean;       // If true then PDF files are sent via email
   DoPrint            : boolean;       // Controls whether the generated document is also printed on the Default Printer
   EditEmail          : boolean;       // Controls whether the email is edited before sending or not
   ReadReceipt        : boolean;       // Controls whether a Read Receipt is requested
   GroupAttach        : boolean;       // Send Invoice, Account & Statement in 1 email
   ThisWrapText       : boolean;       // Determines whether long text is wrapped
   IntEmail           : boolean;       // If True then the Simple Mapi Api (WinSoft component) is not used - Indy is useed instead
   S86FirstPage       : boolean;       // Used for Page break control when doing a S86(3) recon
   S86PageBreak       : boolean;       // Used for Page break control when doing a S86(3) recon

   CpyName            : string;        // Company Name (From the Database)
   HostName           : string;        // Host name for Source Data
   SDate              : string;        // Start Date for invoices
   EDate              : string;        // End Date for invoices
   FileName           : string;        // Name of the output file
   ErrMsg             : string;        // Last error message
   Customer           : string;        // Current Customer
   Descrip            : string;        // Current File Description
   FldExcelStr        : string;        // Caption for form and messagebox
   CSMySQL            : string;        // Connection string parameters
   Template_A         : string;        // Excel Template for Specified Accounts
   Template_I         : string;        // Excel Template for Invoices
   Template_S         : string;        // Excel Template for Statements
   Template_T         : string;        // Excel Template for Trust Recons
   Template_Q         : string;        // Excel Template for Quotes
   Template_FC        : string;        // Excel Template for File Covers
   Template_FN        : string;        // Excel Template for File Notes
   Template_BS        : string;        // Excel Template for Billing Sheet
   Template_CV        : string;        // Excel Template for Conveyancing File Covers
   Layout_A           : string;        // Layout Code for Specified Accounts
   Layout_I           : string;        // Layout Code for Invoices
   Layout_S           : string;        // Layout Code for Statements
   Layout_T           : string;        // Layout Code for Trust Recons
   Layout_Q           : string;        // Layout Code for Quotes
   Header_A           : string;        // Header for Specified Accounts
   Header_I           : string;        // Header for Invoices
   Header_S           : string;        // Header for Statements
   Header_T           : string;        // Header for Trust Recons
   Header_Q           : string;        // Header for Quotes
   Header_X           : string;        // Transformed Header with symbolic variables expanded
   VATNumber          : string;        // Holds VAT number if regestered for VAT
   Address1           : string;        // Address details for the current Customer
   Address2           : string;
   Address3           : string;
   Address4           : string;
   Address5           : string;
   LastDate           : string;        // Used to ensure the 'Carried Down' date is correct
   NegativeStr        : string;        // Contains code to print negative numbers in red
   CpyFile            : string;        // Holds the default S86(3) File Name
   DBPrefix           : string;        // Holds the decrypted Data Base Pefix
   LPMSHeading        : string;        // Holds the correct LPMS nomenclature
   InvoicePref        : string;        // Holds the Invoice Prefix for stored invoices
   InvoiceStr         : string;        // Holds the formatted Invoice Prefix
   PDFPrinter         : string;        // Holds the name of the PDF Printer
   PDFFolder          : string;        // Temporarily holds generated PDF File
   DefPrinter         : string;        // Holds the name of the Default Printer
   FYMonth            : string;        // Holds the start month for the Fincial Year
   FYYear             : string;        // Holds the start year for the Fincial Year
   UserName           : string;        // UserName of lgged-in user
   SendByEmail        : string;        // Controls whether export file is sent by Email
   Billing_To         : string;        // To field when doing Periodic Billing
   Billing_CC         : string;        // CC field when doing Periodic Billing
   Billing_BCC        : string;        // BCC field when doing Periodic Billing
   Billing_Subject    : string;        // Subject field when doing Periodic Billing
   Billing_Email      : string;        // Email Body when doing Periodic Billing
   NilBalance         : string;        // Flag to allow generation of statements with Nil Balances or Invoice number
   InvoiceInfo        : string;        // Holds info about an Invocie e.g. Billing or Collect and existing invoice number
   RunInBackground    : string;        // Controls whether LPMS_Excel runs in the background or not
   FirstFile          : string;        // Holds the name of the First File or a pointer to a list of Files (depends on FileNum)
   ThisAddrList       : string;        // Holds list of Email addressees
   PDFMergeBullzip    : string;        // Holds the Path to the Bullzip PDF MErge Utility
   PDFMergePDFtk      : string;        // Holds the Path to the PDFtk PDF Merge Utility
   FeeEarnerName      : string;        // Name of the current Fee Earner - Fee Earner Report
   FeeEarnerEmail     : string;        // Email Address of the current Fee Earner - Fee Earner Report
   UserEmail          : string;        // Email Adress of the user that is using FldExcel
   PrintSheets        : string;        // Composite string containg the detail for generatig the default docments for a new File
   Parm07             : string;        // Temporarily store parameter 07 before it is decided how it will be used (depends on RunType)
   Parm10             : string;        // Temporarily store parameter 10 before it is decided how it will be used (depends on RunType)
   Parm13             : string;        // Temporarily store parameter 13 before it is decided how it will be used (depends on RunType)
   VersionNum         : string;        // LPMS Version Number
   ThisMax            : string;        // Used to display the progress of prbProgress in numbers
   Bank_Details       : string;        // Holds the banking details string
   QDisclaimerDet     : string;        // Holds the Quote Disclaimer string
   EmailDetails       : string;        // Holds the Email Signature string
   TypeFile           : string;        // Type of File to be printed on the File Cover
   PrescriptD         : string;        // Prescription Date to be printed on the File Cover
   DialCode           : string;        // Dialing Code to be printed on the File Cover
   SaveRelated        : string;        // Used by Billing Preparation report to manage 'Show Related'
   SaveOwner          : string;        // Used by Billing Preparation report to manage 'Show Related'
   S86Date            : string;        // Holds the Start Daet for S86(4)

   FileArray          : array of string;       // Holds File Names
   TotTotals          : array[1..7] of double; // Holds Trust report totals

   BankStrings        : TStringList;   // Holds the untangled Banking Details
   QDisclaimerStr     : TStringList;   // Holds the untangled Quote Disclaimer details
   PrinterList        : TStringList;   // List of defined Printers
   EmailLst           : TStringList;   // Holds Email details as passed
   AttachList         : TStringList;   // Holds generated files to be emailed
   DocStrings         : TStringList;   // Holds List of parameters passed in PrintSheets for printing a File Cover, File Notes and Bliing sheet

//--- Global Variables representing the Layout Codes for Invoices, Accounts and Statements

   lcShowHeader       : boolean;       // Controls whether the header is shown or not
   lcHeaderPageOne    : boolean;       // Controls whether header is displayed on page one only or all pages
   lcShowAddress      : boolean;       // Controls whether the customer's address is shown or not
   lcShowInstruct     : boolean;       // Controls whether the Instruction details are shown or not
   lcShowSummary      : boolean;       // Controls whether the Summary is shown or not
   lcShowBanking      : boolean;       // Controls whether the Banking details are shown or not
   lcShowAge          : boolean;       // Controls whether Age Analysis details are shown or not
   lcRepeatHeader     : boolean;       // Controls whether the header is repeated on general reports

   lcHSR              : integer;       // Header Start Row
   lcHER              : integer;       // Header End Row
   lcHSC              : integer;       // Header Start Column
   lcHEC              : integer;       // Header End Column
   lcASR              : integer;       // Address Start Row
   lcAER              : integer;       // Address End Row
   lcASC              : integer;       // Address Start Column
   lcISR              : integer;       // Instruction Start Row
   lcISCL             : integer;       // Instruction Start Column - Label
   lcISCD             : integer;       // Instruction Start Column - Data
   lcXSR              : integer;       // Summary Start Row
   lcXSCL             : integer;       // Summary Start Column - Label
   lcXSCD             : integer;       // Summary Start Column - Data
   lcPSR              : integer;       // Page 1 Start Row - Detail (incl heading)
   lcPRows            : integer;       // Page 1 Max Rows - Detail (incl heading)
   lcBSR              : integer;       // Banking Details Start Row
   lcBSC              : integer;       // Banking Details Start Column
   lcSSR              : integer;       // Subsequent Pages - Start Row of 2nd page
   lcSRows            : integer;       // Subsequent Pages - Detail Max Rows (incl heading)
   lcSMaxRows         : integer;       // Subsequent Pages - Max Rows per page
   lcSMaxCols         : integer;       // Subsequent Pages - Max Columns on page
   lcAASR             : integer;       // Age Analysis Start Row
   lcAASC             : integer;       // Age Analysis Start Column
   lcSCB              : integer;       // Start Column for Business Description on Statement
   lcSCBD             : integer;       // Start Column for Business Data on Statement
   lcSCT              : integer;       // Start Column for Trust Description on Statement
   lcSCTD             : integer;       // Start Column for Trust Data on Statement
   lcGRows            : integer;       // Rows per page for general reports
   lcGMRWidth         : integer;       // Maximum Row width for general reports
   lcRowsFC           : integer;       // Rows to search for File Cover
   lcColsFC           : integer;       // Columns to search for File Cover
   lcRowsFN           : integer;       // Rows to search for File Notes
   lcColsFN           : integer;       // Columns to search for File Notes
   lcRowsBS           : integer;       // Rows to search for Billing Sheet
   lcColsBS           : integer;       // Columns to search for Billing Sheet
   lcRowsCV           : integer;       // Rows to search for Conveyancing File Covers
   lcColsCV           : integer;       // Columns to search for Conveyancing File Covers
   SaveFileType       : integer;       // Used to exclude Collections File from Billing Prep Report

//--- Variables to hold the amounts for various billing classes

{
   AmtFees            : double;        // Holds the Fees total
   AmtDisburse        : double;        // Holds the Disbursements total
   AmtExpense         : double;        // Holds the Expenses total
   AmtPayment         : double;        // Holds the Payments Received total
   AmtBusToTrust      : double;        // Holds the Payments Made total
   AmtCredit          : double;        // Holds the Credit granted total
   AmtBusDeposit      : double;        // Holds the Deposits to Business total
   AmtTrustDeposit    : double;        // Holds the Deposits to Trust total
   AmtXfrFees         : double;        // Holds the transfers from Trust to Fees total
   AmtXfrDisburse     : double;        // Holds the transfers from Trust to Disbursements total
   AmtXfrClient       : double;        // Holds the transfers from Trust to Clients total
   AmtXfrTrust        : double;        // Holds the transfers from Trust to Trust total
   AmtInv864          : double;        // Holds the S78(2A)/S86(4) Investment totals
   AmtDrw864          : double;        // Holds the S78(2A)/S86(4) Withdrawal totals
   AmtInt864          : double;        // Holds the S78(2A)/S86(4) Accrued Interest totals
   AmtFF864           : double;        // Holds the S86(4) Accrued Fidelity Fund Interest totals
   AmtBusDebit        : double;        // Holds the Business Debits
   AmtTrustDebit      : double;        // Holds the Trust Debits
}

   BalInv863          : double;        // Holds the S86(3) Investment opening balance
   BalDrw863          : double;        // Holds the S86(3) Withdrawal opening balance
   BalIntDrw863       : double;        // Holds the S86(3) Interest Withdrawal opening balance
   BalInt863          : double;        // Holds the S86(3) Accrued Interest opening balance

   OpenBalFees        : double;        // Opening balance for fees :- Fees + Disbursements + Expenses - Credits
   OpenBalTrust       : double;        // Opening Balance for Trust :- Trust Deposits - Trust Transfers (Client, Fees, Disbursements, Trust) + S86(4) Interest
   OpenBalReserve     : double;        // Opening balance for Reserved Trust deposits
   OpenBal863Inv      : double;        // Opening Balance for S86(3) Trust Investments
   OpenBal863Int      : double;        // Opening Balance for S86(3) Interest
   OpenBal863Drw      : double;        // Opening Balance for S86(3) Trust Investment Withdrawals
   OpenBal863IntDrw   : double;        // Opening Balance for S86(3) Trust Investment Interest Withdrawals
   OpenBal864         : double;        // Opening Balance for S78(2A)/S86(4) Trust Investments
   OpenBal864Int      : double;        // Opening Balance for S78(2A)/S86(4) Trust Interest
   OpenBal864Inv      : double;        // Opening Balance for S78(2A)/S86(4) Trust Investments
   OpenBal864Drw      : double;        // Opening Balance for S78(2A)/S86(4) Trust Withdrawals
   OpenBal864FF       : double;        // Opening Balance for S86(4) Fidelity Fund Interest
   OpenBalVAT         : double;        // Opening Balance for VAT Amounts if registered for VAT
   OpenBalAmount      : double;        // Opening Balance for Fees + VAT if registered for VAT

   SummaryFees        : double;        // Total (Fees - Credits) for the period
   SummaryExpenses    : double;        // Total Expenses for the period
   SummaryDisburse    : double;        // Total Disbursements for the period
   SummaryVAT         : double;        // Total VAT for the period

   StatementFees      : double;        // Total Fees for Statement
   StatementDisburse  : double;        // Total Disbursements for Statement
   StatementExpenses  : double;        // Total Expenses for Statement
   StatementDeposits  : double;        // Total Business Desposits for Statement
   StatementTrustDep  : double;        // Total Deposits to Trust for Statement
   StatementReserve   : double;        // Total Reserved Deposits for Statement
   StatementTrustInt  : double;        // Total S78(2A)/S86(4) interest for Statement
   StatementTrustPay  : double;        // Total Payments from Trust for Statement
   StatementTrustDis  : double;        // Total Disbursements from Trust for Statement
   StatementBusPay    : double;        // total payments made from Trust

   AgeCurrent         : double;        // Age Analysis - Current balance
   Age30Days          : double;        // Age Analysis - 30 Days balance
   Age60Days          : double;        // Age Analysis - 60 Days balance
   Age90Days          : double;        // Age Analysis - 90 Days and over balance

   SaveAmount         : double;        // Holds Invoice amount for storing invoices
   PrevPaid           : double;        // All previous payments for an Invoice (Payments report)

   PayAmount          : double;        // Used when doing Payments report - Amount paid on this Invoice
   PayFees            : double;        // Used when doing Payments report - Fees for this Invoice
   PayDisburse        : double;        // Used when doing Payments report - Disbursements for this Invoice
   PayExpenses        : double;        // Used when doing Payments report - Expenses for this Invoice

   FldExcel: TFldExcel;                // System generated

implementation

uses ldGetEmailP;

{$R *.dfm}

//===========================================================================
//===========================================================================
//===                                                                     ===
//=== Housekeeping functions and procedures                               ===
//===                                                                     ===
//===========================================================================
//===========================================================================

//---------------------------------------------------------------------------
// Executed before the form is created
//---------------------------------------------------------------------------
procedure TFldExcel.FormCreate(Sender: TObject);
begin

//***
//***
   Debug := false;
//***
//***

//--- Deal with unsolicited calls to FormActivate

   FirstRun := True;

//--- Set the Symbolic Variables array size

{$INCLUDE 'SetUpSymVarsLenD.inc'}

   SecretPhrase := 'BlueCrane Software Development CC';

//--- Set the start date for S86(3) and S86(4) investments

   S86Date := '2018/10/01';

//--- Check whether the correct number of parameters were passed

//---  Parameter  1  = Run Type
//---  Parameter  2  = Host Name  / XML File name
//---  Parameter  3  = Start Date
//---  Parameter  4  = End Date
//---  Parameter  5  = Output File Folder
//---  Parameter  6  = Calling Form's handle - used to hide and show
//---  Parameter  7  = Account Type / Alert Type / Phonebook Filter1 / Client or File Details Filter / Quote Number
//---  Parameter  8  = Show Related
//---  Parameter  9  = Print Nil Balance Statements
//---  Parameter 10  = Create Statement / Alert Filter / Phonebook Filter2 / Client Details File|DBPrefix
//---  Parameter 11  = Open Output File?
//---  Parameter 12  = Close on Complete?
//---  Parameter 13  = Create Invoice / Exclude Reserved Deposits from Age Analysis?
//---  Parameter 14  = LPMS Message
//---  Parameter 15  = Read Invoice information from the Collections File / Existing Invoice Number
//---  Paramerer 16  = Store Invoice?
//---  Parameter 17  = Send Workbook via Email?
//---  Parameter 18  = String containing Email Addresses, Subject Addressee and File Reference
//---  Parameter 19  = Create a PDF from the resulting Excel file
//---  Parameter 20  = Run this report in the background
//---  Parameter 21  = Registry String
//---  Parameter 22  = Logged in User
//---  Parameter 23  = Print the generated Spreadsheet
//---  Parameter 24  = LPMS Version Number
//---  Parameter 25  = PrintSheets
//---  Parameter 26  [26]  = Number of Files to follow
//---  Parameter 27  [27]  = First File Name
//---  Parameter 28+ [28+] = More File Names

   NumParms := ParamCount;

//--- Verify the number of parameters

   if (NumParms < 26) then begin
      MessageDlg('Invalid use of ldExcel - RC = ' + IntToStr(NumParms), mtError, [mbOK], 0);
      Application.Terminate;
      Exit;
   end;

//--- Initialise Global variables that must have an initial value

   CreateStatement := False;
   CreateInvoice   := False;

   RunType         := StrToInt(ParamStr(1));
   HostName        := ParamStr(2);
   SDate           := ParamStr(3);
   EDate           := ParamStr(4);
   FileName        := ParamStr(5);
   CallHWND        := StrToInt(ParamStr(6));
   Parm07          := ParamStr(7);
   ShowRelated     := StrToBool(ParamStr(8));
   NilBalance      := ParamStr(9);
   Parm10          := ParamStr(10);
   AutoOpen        := StrToBool(ParamStr(11));
   CloseOnComplete := StrToBool(ParamStr(12));
   Parm13          := ParamStr(13);
   LPMSHeading     := ParamStr(14);
   InvoiceInfo     := ParamStr(15);
   StoreInv        := StrToBool(ParamStr(16));
   SendByEmail     := ParamStr(17);
   ThisAddrList    := ParamStr(18);
   CreatePDF       := StrToBool(ParamStr(19));
   RunInBackground := ParamStr(20);
   RegString       := ParamStr(21);
   UserName        := ParamStr(22);
   DoPrint         := StrToBool(ParamStr(23));

   if (AnsiContainsStr(ParamStr(24), '[DEBUG]') = True) then
      VersionNum := 'DEBUG'
   else
      VersionNum := ParamStr(24);

   PrintSheets     := ParamStr(25);
   NumFiles        := StrToInt(ParamStr(26));
   FirstFile       := ParamStr(27);

//--- Determine whether Parameter 7 (string) should be cast as an integer

   if (RunType in [11,13..15,22,28]) then
      Parm07 := Parm07
   else
      AccountType := StrToInt(Parm07);

//--- Determine whether Parameter 10 (string) should be cast as a boolean

   if (RunType in [25,27]) then
      CreateStatement := StrToBool(Parm10);

//--- Determine whether Parameter 13 (string) should be cast as a boolean

   if (RunType in [25]) then
      CreateInvoice := StrToBool(Parm13);

//--- Check whether the number of parameters received are correct

   if (ParamCount <> (NumFiles + 26)) then begin
      MessageDlg('Invalid use of ldExcel - RC = ' + IntToStr(ParamCount), mtError, [mbOK], 0);
      Application.Terminate;
      Exit;
   end;

   case RunType of
       1: FldExcelStr := LPMSHeading + ' - Specified Accounts';
       2: FldExcelStr := LPMSHeading + ' - Asterisk Interface';
       3: FldExcelStr := LPMSHeading + ' - Notes';
       4: FldExcelStr := LPMSHeading + ' - Invoices';
       5: FldExcelStr := LPMSHeading + ' - Statements';
       6: FldExcelStr := LPMSHeading + ' - Trust Details';
       7: FldExcelStr := LPMSHeading + ' - S86(3) Details';
       8: FldExcelStr := LPMSHeading + ' - Fee Earner Report';
       9: FldExcelStr := LPMSHeading + ' - Accountant Report';
      10: FldExcelStr := LPMSHeading + ' - Alerts Report';
      11: FldExcelStr := LPMSHeading + ' - Phonebook Export';
      12: FldExcelStr := LPMSHeading + ' - Billing Preparation Report';
      13: FldExcelStr := LPMSHeading + ' - Client Details';
      14: FldExcelStr := LPMSHeading + ' - Opposition Details';
      15: FldExcelStr := LPMSHeading + ' - File Details';
      16: FldExcelStr := LPMSHeading + ' - Collections Account';
      17: FldExcelStr := LPMSHeading + ' - Invoice List';
      18: FldExcelStr := LPMSHeading + ' - Prefix Details';
      19: FldExcelStr := LPMSHeading + ' - Safe Keeping Details';
      20: FldExcelStr := LPMSHeading + ' - Document Control Details';
      21: FldExcelStr := LPMSHeading + ' - Log Details';
      22: FldExcelStr := LPMSHeading + ' - User Details';
      23: FldExcelStr := LPMSHeading + ' - Quotation';
      24: FldExcelStr := LPMSHeading + ' - Task List';
      25: FldExcelStr := LPMSHeading + ' - Generate Periodic Billing';
      26: FldExcelStr := LPMSHeading + ' - Trust Management Report';
      27: FldExcelStr := LPMSHeading + ' - Payments Report';
      28: FldExcelStr := LPMSHeading + ' - Quote List';
      29: FldExcelStr := LPMSHeading + ' - Generate File Documents for a new File';
   end;

   GroupAttach  := false;

end;

//---------------------------------------------------------------------------
// Executed before the form is displayed
//---------------------------------------------------------------------------
procedure TFldExcel.FormActivate(Sender: TObject);
var
   idx1, idx2, idx3, Adjust, PaymentRecs  : integer;
   FoundPrinter                           : boolean;
   Device, Driver, Port                   : PChar;
   ThisStr, FilterStr, TestStr, SaveStat  : string;
   RegIni                                 : TRegistry;
   ExportSet, RecordSet                   : IXMLNode;
   HDeviceMode                            : THandle;
   ThisPrinter                            : TPrinter;

begin

//--- Deal with unsolicited calls to FormActivate

   if (FirstRun = False) then
      Exit
   else
      FirstRun := False;

//--- Check whether Invoice and Statement must be grouped in one email

   if (CreateInvoice = true) and (CreateStatement = true) and (SendByEmail = '1') then begin
      GroupAttach := true;
      RunType := 0;
      FldExcelStr := LPMSHeading + ' - Billing';
   end else
      GroupAttach := false;

   PDFExists := False;
   FldExcel.Caption := FldExcelStr;

//--- Hide the calling Form if the Form passed a valid Handle. If "Run in Background"
//--- is set then hide this form until the request is completed

   if (RunInBackground <> '1') then begin

      if (CallHWND <> 0) then
         ShowWindow(CallHWND,SW_HIDE);

   end else
      FldExcel.WindowState := wsMinimized;

//--- Set up the Registry string. The implications of MultiCompany is taken
//--- care of in the calling program and already reflected in RegString

   RegString := RegString + '\Preferences';

//--- Extract various variables from the Registry

   RegIni := TRegistry.Create;
   RegIni.RootKey := HKEY_CURRENT_USER;
   RegIni.OpenKey(RegString,false);

   lcColsBS        := RegIni.ReadInteger('Cols_BS');
   lcColsCV        := RegIni.ReadInteger('Cols_CV');
   lcColsFC        := RegIni.ReadInteger('Cols_FC');
   lcColsFN        := RegIni.ReadInteger('Cols_FN');
   lcGMRWidth      := RegIni.ReadInteger('MaxRowWidth');
   lcGRows         := RegIni.ReadInteger('LinesPerPage');
   lcRowsBS        := RegIni.ReadInteger('Rows_BS');
   lcRowsCV        := RegIni.ReadInteger('Rows_CV');
   lcRowsFC        := RegIni.ReadInteger('Rows_FC');
   lcRowsFN        := RegIni.ReadInteger('Rows_FN');
   PDFInterval     := RegIni.ReadInteger('PDFInterval');
   PDFMergeSel     := RegIni.ReadInteger('PDFMergeSel');
   PDFRetry        := RegIni.ReadInteger('PDFRetry');
   SMTPAuthType    := RegIni.ReadInteger('SMTPAuthType');
   Bank_Details    := RegIni.ReadString('BankDetails');
   DBPrefix        := RegIni.ReadString('DBPrefix');
   DefPrinter      := RegIni.ReadString('DefPrinter');
   EmailBody       := RegIni.ReadString('EmailBody2');
   EmailDetails    := RegIni.ReadString('EmailSignature');
   FYMonth         := RegIni.ReadString('FYMonth');
   Header_A        := RegIni.ReadString('Heading_A');
   Header_I        := RegIni.ReadString('Heading_I');
   Header_Q        := RegIni.ReadString('Heading_Q');
   Header_S        := RegIni.ReadString('Heading_S');
   Header_T        := RegIni.ReadString('Heading_T');
   Layout_A        := RegIni.ReadString('Layout_A');
   Layout_I        := RegIni.ReadString('Layout_I');
   Layout_Q        := RegIni.ReadString('Layout_Q');
   Layout_S        := RegIni.ReadString('Layout_S');
   Layout_T        := RegIni.ReadString('Layout_T');
   PDFFolder       := RegIni.ReadString('PDFFolder');
   PDFMergeBullzip := RegIni.ReadString('PDFMergeBullzip');
   PDFMergePDFtk   := RegIni.ReadString('PDFMergePDFtk');
   PDFPrinter      := RegIni.ReadString('PDFPrinter');
   QDisclaimerDet  := RegIni.ReadString('QDisclaimer');
   SMTPPort        := RegIni.ReadString('SMTPPort');
   SMTPServer      := RegIni.ReadString('SMTPServer');
   SMTPUser        := RegIni.ReadString('SMTPUser');
   Template_A      := RegIni.ReadString('Template_A');
   Template_BS     := RegIni.ReadString('Template_BS');
   Template_CV     := RegIni.ReadString('Template_CV');
   Template_FC     := RegIni.ReadString('Template_FC');
   Template_FN     := RegIni.ReadString('Template_FN');
   Template_I      := RegIni.ReadString('Template_I');
   Template_Q      := RegIni.ReadString('Template_Q');
   Template_S      := RegIni.ReadString('Template_S');
   Template_T      := RegIni.ReadString('Template_T');
   ColAB1F         := TColor(RegIni.ReadInteger('ColAB1F'));
   ColAB1T         := TColor(RegIni.ReadInteger('ColAB1T'));
   ColAB2F         := TColor(RegIni.ReadInteger('ColAB2F'));
   ColAB2T         := TColor(RegIni.ReadInteger('ColAB2T'));
   ColDHF          := TColor(RegIni.ReadInteger('ColDHF'));
   ColDHT          := TColor(RegIni.ReadInteger('ColDHT'));
   ExcludeReserve  := StrToBool(RegIni.ReadString('ExcludeReserve'));
   IncludeTrust    := StrToBool(RegIni.ReadString('IncludeTrust'));
   IntEmail        := StrToBool(RegIni.ReadString('IntEmail'));
   lcRepeatHeader  := StrToBool(RegIni.ReadString('RepeatHeader'));
   NegativeRed     := StrToBool(RegIni.ReadString('NegativeRed'));
   PDFPrefer       := StrToBool(RegIni.ReadString('PDFPrefer'));
   SMTPAuth        := StrToBool(RegIni.ReadString('SMTPAuth'));
   ThisWrapText    := StrToBool(RegIni.ReadString('WrapText'));
   SMTPPass        := Vignere(ord(CYPHER_DEC),RegIni.ReadString('SMTPPass'),SecretPhrase);

   DBPrefix := Copy(DBPrefix,1,6);

   RegIni.CloseKey;
   RegIni.Free;

//--- Untangle the Banking Details and Email Signature into strings

   BankStrings := TStringList.Create;
   BankStrings := Disassemble(Bank_Details,'~');

   EmailStrings := TStringList.Create;
   EmailStrings := Disassemble(EmailDetails,'~');

   QDisclaimerStr := TStringList.Create;
   QDisclaimerStr := Disassemble(QDisclaimerDet,'~');

   EmailLst   := TStringList.Create;
   AttachList := TStringList.Create;

//--- Manage color of negative numbers

   if (NegativeRed = True) then
      NegativeStr := '[Red]'
   else
      NegativeStr := '';

//--- Check if the PDF Printer exists and set PDFPrinter to "Not Found" if not.
//--- The value of PDFPrinter is also used to determine whether the information
//--- message for PDF printing is printed or not

   FoundPrinter         := false;
   Printer.PrinterIndex := -1;

   GetMem(Device, 255);
   GetMem(Driver, 255);
   GetMem(Port,   255);

   ThisPrinter := TPrinter.Create;

//--- Look for the PDF Printer

   try
      for idx3 := 0 to Printer.Printers.Count - 1 do begin
         if Printer.Printers[idx3] = PDFPrinter then begin
            ThisPrinter.PrinterIndex := idx3;
            ThisPrinter.Getprinter(Device, Driver, Port, HDeviceMode);

            StrCat(Device, ',');
            StrCat(Device, Driver);
            StrCat(Device, Port);

            FoundPrinter := true;
            PDFPrinter   := Device;
         end;
      end;
   finally
      ThisPrinter.Free;
   end;

   FreeMem(Device, 255);
   FreeMem(Driver, 255);
   FreeMem(Port,   255);

   if (FoundPrinter = False) then
      PDFPrinter := 'Not Found';

//--- Check whether the default Printer exists and set DefPrinter to "Not Found"
//--- if not. The value of DefPrinter is also used to determine whether the
//--- information message for document printing is printed or not

   FoundPrinter         := false;
   Printer.PrinterIndex := -1;

   GetMem(Device, 255);
   GetMem(Driver, 255);
   GetMem(Port,   255);

   ThisPrinter := TPrinter.Create;

//--- Look for the Default Printer

   try
      for idx3 := 0 to Printer.Printers.Count - 1 do begin
         if Printer.Printers[idx3] = DefPrinter then begin
            ThisPrinter.PrinterIndex := idx3;
            ThisPrinter.Getprinter(Device, Driver, Port, HDeviceMode);

            StrCat(Device, ',');
            StrCat(Device, Driver);
            StrCat(Device, Port);

            FoundPrinter := true;
            DefPrinter   := Device;
         end;
      end;
   finally
      ThisPrinter.Free;
   end;

   FreeMem(Device, 255);
   FreeMem(Driver, 255);
   FreeMem(Port,   255);

   if (FoundPrinter = False) then
      DefPrinter := 'Not Found';

//--- Start the process to export the data contained in the Export Group

   DBPrefix := jvCipher.DecodeString(SecretPhrase,DBPrefix);

(*
   TestStr := Vignere(ord(CYPHER_ENC),'BlueCrane Software','BlueCrane Software Development CC');
   TestStr := Vignere(ord(CYPHER_DEC),TestStr,'BlueCrane Software Development CC');
*)

   CSMySQL  := 'Provider=MSDASQL.1;Persist Security Info=False;Data Source=BSD_MySQL;User ID=' + DBPrefix + '_LD;Password=LD01;Database=' + DBPRefix + '_LPMS;Server=';

   StatusBar.Panels.Items[0].Text := ' ' + HostName + ' [' + DBPrefix + ']';

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~//
// Extract the File Names                                                     //
//                                                                            //
// FileName will be '**ALL**' if the intention is to process all records in   //
//    a collection.                                                           //
//                                                                            //
// FileName will be '##xx##' if a filter, represented by 'xx' is to be        //
//    applied when processing records.                                        //
//                                                                            //
// FileName will be '**NULL**' when Phonebook entries are processed.          //
//                                                                            //
// FileName will be '**XML**' if Runtype is one of the following (In all      //
//    instances the name of the XML file is passed in Parameter 2 and         //
//    contains the Hostname and number of records included):                  //
//                                                                            //
//     2 - Call Records                                                       //
//    15 - File Export                                                        //
//    16 - Collections Records                                                //
//    17 - Invoice List                                                       //
//                                                                            //
// FileName will be '**DOC**' if the intention is to produce the File Cover   //
//    File Notes and Billing Sheet for a newly created File                   //
//                                                                            //
// In all other instances FileNames contains an actual record name.           //
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~//

//--- Open a connection to the datastore named in HostName

   if (FirstFile <> '**XML**') then begin
      if ((Open_Connection(HostName)) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

//--- Get Company specific information

      GetCpyVAT();
   end;

//--- Process the passed parameter

   TestStr := Copy(FirstFile,1,2);

   if (((FirstFile = '**ALL**') or (TestStr = '##')) and (NumFiles = 1)) then begin
      FilterStr := Copy(FirstFile,3,Length(FirstFile) - 4);

//--- Reset FilterStr if NumFiles = '**ALL**'

      if (FirstFile = '**ALL**') then
         FilterStr := '';

      if (RunType in [0,1,3..7,9,12,26]) then begin
         if (GetFileNames(FilterStr + '%') = false) then begin
            LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

            CloseOnComplete  := False;
            AutoOpen         := False;
            txtError.Text    := 'There are errors...';
            Exit;
         end;

         SetLength(FileArray, Query2.RecordCount);
         NumFiles := Query2.RecordCount;
         idx2 := 0;
         Adjust := 0;

         for idx1 := 0 to NumFiles - 1 do begin
            if ((Query2.FieldByName('Tracking_Name').AsString = CpyFile) and (RunType <> 9)) then begin
               Query2.Next;
               inc(Adjust);
            end else begin
               FileArray[idx2] := Query2.FieldByName('Tracking_Name').AsString;
               Query2.Next;
               inc(idx2);
             end;
         end;

         NumFiles := NumFiles - Adjust;

      end else if (RunType in [8,10]) then begin

         if(RunType = 8) then
            ThisStr := ' WHERE Control_FeeEarner = 1'
         else
            ThisStr := '';

         if (GetAllUsers(ThisStr) = false) then begin
            LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

            CloseOnComplete  := False;
            AutoOpen         := False;
            txtError.Text    := 'There are errors...';
            Exit;
         end;

         SetLength(FileArray, Query2.RecordCount);
         NumFiles := Query2.RecordCount;

         for idx1 := 0 to NumFiles - 1 do begin
            FileArray[idx1] := Query2.FieldByName('Control_UserID').AsString;
            Query2.Next;
         end;

      end else if (RunType in [27]) then begin

         if (GetInvoices(ord(DT_INVOICES),'') = false) then begin
            LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

            CloseOnComplete  := False;
            AutoOpen         := False;
            txtError.Text    := 'There are errors...';
            Exit;
         end;

         SetLength(FileArray, Query1.RecordCount);
         NumFiles := Query1.RecordCount;
         PaymentRecs := 0;

         SaveStat := StatusBar.Panels.Items[0].Text;

         for idx1 := 0 to NumFiles - 1 do begin

            if (GetPayments(Query1.FieldByName('Inv_Invoice').AsString) > 0) then begin

               StatusBar.Panels.Items[0].Text := 'Pre-processing Payments for Invoice ''' + Query1.FieldByName('Inv_Invoice').AsString + '''';
               StatusBar.Refresh;

               FileArray[PaymentRecs] := Query1.FieldByName('Inv_Invoice').AsString;
               inc(PaymentRecs);

            end;

            Query1.Next;

         end;

         StatusBar.Panels.Items[0].Text := SaveStat;

         SetLength(FileArray, PaymentRecs);
         NumFiles := PaymentRecs;

      end;
   end else if ((FirstFile = '**XML**') and (NumFiles = 1)) then begin

//--- Open the XML file and process the content of the XML file.

      if (RunType in [0,1,3..14,25..27]) then begin
         XMLDoc.Active := false;
         XMLDoc.LoadFromFile(HostName);
         XMLDoc.Active := true;

         DeleteFile(HostName);

         ExportSet  := XMLDoc.DocumentElement;
         HostName   := ExportSet.Attributes['Host'];

         if (RunType = 25) then
            NumFiles   := ExportSet.Attributes['Count'];

         RecordSet := ExportSet.ChildNodes.First;

//--- Get Company specific information

         if ((Open_Connection(HostName)) = false) then begin
            LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

            CloseOnComplete  := False;
            AutoOpen         := False;
            txtError.Text    := 'There are errors...';
            Exit;
         end;

         GetCpyVAT();

//--- Now process the lines contained in the XML file

         if (RunType = 25) then begin
            SetLength(Bill_Instr.BI,NumFiles);
            for idx1 := 0 to NumFiles - 1 do begin
               Bill_Instr.BI[idx1].FileName     := ReplaceXML(RecordSet.ChildValues['File']);
               Bill_Instr.BI[idx1].Description  := ReplaceXML(RecordSet.ChildValues['Description']);
               Bill_Instr.BI[idx1].StartDate    := ReplaceXML(RecordSet.ChildValues['StartDate']);
               Bill_Instr.BI[idx1].EndDate      := ReplaceXML(RecordSet.ChildValues['EndDate']);
               Bill_Instr.BI[idx1].Invoice      := ReplaceXML(RecordSet.ChildValues['Invoice']);
               Bill_Instr.BI[idx1].Statement    := ReplaceXML(RecordSet.ChildValues['Statement']);
               Bill_Instr.BI[idx1].StoreInvoice := ReplaceXML(RecordSet.ChildValues['StoreInvoice']);
               Bill_Instr.BI[idx1].SendEmail    := ReplaceXML(RecordSet.ChildValues['SendEmail']);
               Bill_Instr.BI[idx1].CreatePDF    := ReplaceXML(RecordSet.ChildValues['CreatePDF']);
               Bill_Instr.BI[idx1].Print        := ReplaceXML(RecordSet.ChildValues['Print']);
               Bill_Instr.BI[idx1].ShowRelated  := ReplaceXML(RecordSet.ChildValues['ShowRelated']);
               Bill_Instr.BI[idx1].AutoOpen     := ReplaceXML(RecordSet.ChildValues['AutoOpen']);
               Bill_Instr.BI[idx1].EmailTo      := ReplaceXML(RecordSet.ChildValues['To']);
               Bill_Instr.BI[idx1].CC           := ReplaceXML(RecordSet.ChildValues['CC']);
               Bill_Instr.BI[idx1].BCC          := ReplaceXML(RecordSet.ChildValues['BCC']);
               Bill_Instr.BI[idx1].Subject      := ReplaceXML(RecordSet.ChildValues['Subject']);
               Bill_Instr.BI[idx1].EmailBody    := ReplaceXML(RecordSet.ChildValues['EmailBody']);
//               Bill_Instr.BI[idx1].EditEmail    := ReplaceXML(RecordSet.ChildValues['EditEmail']);
               Bill_Instr.BI[idx1].ReadReceipt  := ReplaceXML(RecordSet.ChildValues['ReadReceipt']);
               Bill_Instr.BI[idx1].AcctType     := ReplaceXML(RecordSet.ChildValues['AcctType']);

               RecordSet := RecordSet.NextSibling;
            end;
         end else begin
            SetLength(FileArray,NumFiles);
            idx2 := 0;
            Adjust := 0;

            for idx1 := 0 to NumFiles - 1 do begin
               if ((RecordSet.ChildValues['Exp_Item'] = CpyFile) and (RunType <> 9)) then begin
                  RecordSet := Recordset.NextSibling;
                  inc(Adjust);
               end else begin
                  FileArray[idx2] := RecordSet.ChildValues['Exp_Item'];
                  RecordSet := RecordSet.NextSibling;
                  inc(idx2);
               end;
            end;

            NumFiles := NumFiles - Adjust;
         end;

      end else begin

         if ((Open_Connection(NilBalance)) = false) then begin
            LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

            CloseOnComplete  := False;
            AutoOpen         := False;
            txtError.Text    := 'There are errors...';
            Exit;
         end;

         GetCpyVAT();
      end;
   end else if ((FirstFile = '**DOC**') and (NumFiles = 1)) then begin
      DocStrings := TStringList.Create;
      DocStrings := Disassemble(PrintSheets,'|');
   end else begin
      idx1 := 0;
      SetLength(FileArray, NumFiles);

      for idx2 := 0 to (NumFiles - 1) do begin
         FileArray[idx1] := ParamStr(27 + idx1);
         idx1 := idx1 + 1;
      end;
   end;

//--- Set up the Symbolic Variables

   SetUpSymVars;

//--- Get the email adress for the User that invoked FldExcel

   if (GetUser(UserName) = True) then
      UserEmail := Query2.FieldByName('Control_Email').AsString
   else
      UserEmail := '';

//--- Close the DB connection for now

   Close_Connection();

//--- Hide the ProgressBar - it is only displayed when necessary

   prbProgress.Hide;
   stCount.Visible := False;

//--- Now execute the request

   case RunType of
       0: Generate_Billing();
       1: Generate_Accounts();
       2: Generate_CallRecs();
       3: Generate_Notes();
       4: Generate_SingleBilling();
       5: Generate_SingleBilling();
       6: Generate_Trust_Summary();
       7: Generate_S863_Summary();
       8: Generate_FeeReport();
       9: Generate_AcctReport();
      10: Generate_Alerts();
      11: Generate_Phonebook();
      12: Generate_BillingReport();
      13: Generate_ClientDetails();
      14: Generate_ClientDetails();
      15: Generate_FileDetails();
      16: Generate_CollectAcct();
      17: Generate_InvoiceList();
      18: Generate_PrefixDetails();
      19: Generate_SafeDetails();
      20: Generate_DocCntlDetails();
      21: Generate_LogDetails();
      22: Generate_UserDetails();
      23: Generate_Quotation();
      24: Generate_TaskList();
      25: Generate_PeriodicBilling();
      26: Generate_TrustManagement();
      27: Generate_Payments();
      28: Generate_QuoteDetails();
      29: Generate_PrintSheets();
   end;

   FldExcel.WindowState := wsNormal;

//--- Clean up

   if (CloseOnComplete = false) then begin
      btnCancel.Caption := 'Return';

      case RunType of
          0: txtError.Text := 'Billing generation completed';
          1: txtError.Text := 'Specified Account generation completed';
          2: txtError.Text := 'Call Record generation completed';
          3: txtError.Text := 'Notes Details generation completed';
          4: txtError.Text := 'Invoice generation completed';
          5: txtError.Text := 'Statement generation completed';
          6: txtError.Text := 'Trust Account Detail generation completed';
          7: txtError.Text := 'S86(3) Detail generation completed';
          8: txtError.Text := 'Fee Earner Report generation completed';
          9: txtError.Text := 'Accountant Report generation completed';
         10: txtError.Text := 'Alerts Report generation completed';
         11: txtError.Text := 'Phonebook Details generation completed';
         12: txtError.Text := 'Billing Report generation completed';
         13: txtError.Text := 'Client Details generation completed';
         14: txtError.Text := 'Opposition Details generation completed';
         15: txtError.Text := 'File Details generation completed';
         16: txtError.Text := 'Collection Account generation completed';
         17: txtError.Text := 'Invoice List generation completed';
         18: txtError.Text := 'Prefix Details generation completed';
         19: txtError.Text := 'Safe Keeping Details generation completed';
         20: txtError.Text := 'Document Control Details generation completed';
         21: txtError.Text := 'Billing Items Details generation completed';
         22: txtError.Text := 'User Details generation completed';
         23: txtError.Text := 'Quotation generation completed';
         24: txtError.Text := 'Task List generation completed';
         25: txtError.Text := 'Periodic Billing generation completed';
         26: txtError.Text := 'Trust Management Report generation completed';
         27: txtError.Text := 'Payments Report generation completed';
         28: txtError.Text := 'Quotation List export completed';
         29: txtError.Text := 'File Document generation completed';
      end;
      lbProgress.SetFocus;
      lbProgress.Selected[lbProgress.Items.Count - 1] := true;
    end else
      btnCancelClick(Sender);

end;

//---------------------------------------------------------------------------
// Executed when the form is closed
//---------------------------------------------------------------------------
procedure TFldExcel.FormClose(Sender: TObject; var Action: TCloseAction);
begin

//--- Make sure the form is visible

   FldExcel.WindowState := wsNormal;

//--- Clean up

   try BankStrings.Destroy;    except end;
   try EmailStrings.Destroy;   except end;
   try QDisclaimerStr.Destroy; except end;
   try PrinterList.Destroy;    except end;
   try EmailLst.Destroy;       except end;
   try AttachList.Destroy;     except end;
   try DocStrings.Destroy;     except end;

//--- Show the calling Form if it was previosuly hidden. Don't do this if the
//--- report ran in the background

   if (RunInBackground <> '1') then begin
      if (CallHWND <> 0) then
         ShowWindow(CallHWND,SW_SHOW);
   end;

   Application.Terminate;

end;

//---------------------------------------------------------------------------
// Executed whenever the Error/Information message changes
//---------------------------------------------------------------------------
procedure TFldExcel.txtErrorChange(Sender: TObject);
begin
   txtError.Refresh;
end;

//---------------------------------------------------------------------------
// User clicked on the Exit button
//---------------------------------------------------------------------------
procedure TFldExcel.btnCancelClick(Sender: TObject);
begin
   Close;
end;

//---------------------------------------------------------------------------
// Function to open a connection to the datastore
//---------------------------------------------------------------------------
function TFldExcel.Open_Connection(OpenHost: string): boolean;
begin

   try
      Query1.Close;
      Query2.Close;

      MySQLCon.Close;
      MySQLCon.ConnectionString := CSMySQL + OpenHost;
      Query1.Connection := MySQLCon;
      Query2.Connection := MySQLCon;
      MySQLCon.Open;

   except
      ErrMsg := '''Unable to connect to ' + OpenHost + '''';
      Result := false;
      Exit;
   end;

   Result := true;
end;

//---------------------------------------------------------------------------
// Procedure to close the datastore connection
//---------------------------------------------------------------------------
procedure TFldExcel.Close_Connection();
begin

   Query1.Close;
   Query2.Close;

   MySQLCon.Close;

end;

//===========================================================================
//===========================================================================
//===                                                                     ===
//=== Primary Report functions and procedures                             ===
//===                                                                     ===
//===========================================================================
//===========================================================================

//---------------------------------------------------------------------------
// Generate Periodic Billing - One or more of an Invoice and a Statement
// was requested. Can also be multiple requestes for Invoice and Statement
// pairs
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_PeriodicBilling();
var
   idx1 : integer;

begin

   RunType  := 0;
   NumFiles := 1;
   SetLength(FileArray,1);

   LogMsg(ord(PT_PROLOG),True,False,'Billing');

//--- Step through the list of Billing Instructions

   for idx1 := 0 to Length(Bill_Instr.BI) - 1 do begin
      FileArray[0]    := Bill_Instr.BI[idx1].FileName;
      SDate           := Bill_Instr.BI[idx1].StartDate;
      EDate           := Bill_Instr.BI[idx1].EndDate;
      CreateInvoice   := StrToBool(Bill_Instr.BI[idx1].Invoice);
      CreateStatement := StrToBool(Bill_Instr.BI[idx1].Statement);
      StoreInv        := StrToBool(Bill_Instr.BI[idx1].StoreInvoice);
      SendByEmail     := Bill_Instr.BI[idx1].SendEmail;
      CreatePDF       := StrToBool(Bill_Instr.BI[idx1].CreatePDF);
      DoPrint         := StrToBool(Bill_Instr.BI[idx1].Print);
      ShowRelated     := StrToBool(Bill_Instr.BI[idx1].ShowRelated);
      AutoOpen        := StrToBool(Bill_Instr.BI[idx1].AutoOpen);
      Billing_To      := Bill_Instr.BI[idx1].EmailTo;
      Billing_CC      := Bill_Instr.BI[idx1].CC;
      Billing_BCC     := Bill_Instr.BI[idx1].BCC;
      Billing_Subject := Bill_Instr.BI[idx1].Subject;
      Billing_Email   := Disassemble(Bill_Instr.BI[idx1].EmailBody,'|').Text;
//      EditEmail       := StrToBool(Bill_Instr.BI[idx1].EditEmail);
      ReadReceipt     := StrToBool(Bill_Instr.BI[idx1].ReadReceipt);
      AccountType     := StrToInt(Bill_Instr.BI[idx1].AcctType);

//--- We always set GroupAttach to true so that we can control how the email is
//--- constructed from here

      GroupAttach := true;

//--- Do the Invoice and or Statement

      if (CreateInvoice = true) then begin
         LogMsg(ord(PT_MIDLOG),False,False,'Processing Invoice for File: "' + FileArray[0] +'"');
         LogMsg(ord(PT_OPTLOG),False,False,'Invoice');
         Generate_Invoice();
      end;

      if (CreateStatement = true) then begin
         LogMsg(ord(PT_MIDLOG),False,False,'Processing Statement for File: "' + FileArray[0] +'"');
         LogMsg(ord(PT_OPTLOG),False,False,'Statement');
         Generate_Statement();
      end;

//--- If Send by Email is selected and GroupAttach is true then send the
//--- Invoice and Statement in one email

      if ((SendByEmail = '1') and (GroupAttach = true)) then begin
         LogMsg(ord(PT_MIDLOG),False,False,'Processing global instructions');

         if (SendEmail(Bill_Instr.BI[idx1].FileName,'',ord(PT_BILLING)) = true) then
            LogMsg('  Request to send Billing by Email submitted...',True)
         else
            LogMsg('  Request to send Billing by Email not submitted...',True);

         LogMsg(' ',True);
      end;
   end;

   LogMsg(ord(PT_EPILOG),False,False,'Billing');

end;

//---------------------------------------------------------------------------
// Generate Single Billing - either of an Invocie or a Statement was
// requested as a single document to be generated
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_SingleBilling();
begin

//--- If RunType is 4 then an Invoice was requested else a Statement

   ThisCount := 1;
   ThisMax   := IntToStr(NumFiles);

   prbProgress.Max := NumFiles;

   if (RunType = 4) then begin
//      prbProgress.Max := NumFiles;
      LogMsg(ord(PT_PROLOG),True,True,'Invoice');
      Generate_Invoice;
      LogMsg(ord(PT_EPILOG),False,False,'Invoice');
   end else begin
//      prbProgress.Max := NumFiles;
      LogMsg(ord(PT_PROLOG),False,True,'Statement');
      Generate_Statement;
      LogMsg(ord(PT_EPILOG),False,False,'Statement');
   end;
end;

//---------------------------------------------------------------------------
// Generate Billing (Invoice & Statement)
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Billing();
var
   idx1, ThisFiles : integer;
   ThisArray       : array of string;

begin

//--- Transfer the true value of NumFiles to ThisFiles and of FileArray to
//--- ThisArray to ensure that Billing is done one file at a time

   ThisFiles := NumFiles;

   SetLength(ThisArray,NumFiles);
   for idx1 := 0 to NumFiles - 1 do
      ThisArray[idx1] := FileArray[idx1];

   NumFiles := 1;
   SetLength(FileArray,1);

//--- Step through the list of Files

   LogMsg(ord(PT_PROLOG),True,True,'Billing');

   for idx1 := 0 to ThisFiles - 1 do begin
      FileArray[0] := ThisArray[idx1];

      Generate_Invoice();
      Generate_Statement();

//--- If Send by Email is selected then send the Invoice and Statement in one
//--- email

      if (SendByEmail = '1') then begin
         LogMsg('',False);
         LogMsg('***',False);
         LogMsg('',False);

         if (SendEmail(ThisArray[idx1],'',ord(PT_NORMAL)) = True) then
            LogMsg('Request to send Billing by Email submitted...',False)
         else
            LogMsg('Request to send Billing by Email not submitted...',False);

         LogMsg('',False);
         LogMsg('---',False);
         LogMsg('',True);
      end;
   end;

end;

//---------------------------------------------------------------------------
// Generate Specified Accounts
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Accounts();
var
   PageNum, ThisRows, ThisCol, ThisClass, Row, NumPages : integer;
   idx1, idx2, idx3, RemainRows                         : integer;
   ThisAmount                                           : double;
   DoLine                                               : boolean;
   ThisItem, ThisFile, ThisStr                          : string;
   ThisDate                                             : TDateTime;
   xlsBook                                              : IXLSWorkbook;
   xlsSheet                                             : IXLSWorksheet;

begin

   ThisCount := 1;
   ThisMax   := IntToStr(NumFiles);

   prbProgress.Max := NumFiles;
   LogMsg(ord(PT_PROLOG),True,True,'Specified Account');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Database error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Get the Heading and Layout information

   GetLayout('Specified Account');

//--- Read the Billing records from the datastore

   ShortDateFormat := 'yyyy/MM/dd';
   DateSeparator   := '/';

   for idx1 := 0 to NumFiles - 1 do begin
      PageNum := 1;
      ThisDate := StrToDate(EDate);
      ThisFile := FormatDateTime('yyyyMMdd',ThisDate) + ' - Specified Account (' + FileArray[idx1] + ').xls';
      txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;
      txtDocument.Refresh;

//--- Initialise the Total Balance Fields

      OpenBalFees      := 0;
      OpenBalTrust     := 0;
      OpenBalReserve   := 0;
      OpenBal864       := 0;
      OpenBal864Int    := 0;
      OpenBal864Inv    := 0;
      OpenBal864Drw    := 0;
      OpenBalVAT       := 0;
      SummaryVAT       := 0;

//--- Build the unique part of the SQL statement

      ThisStr := ' AND ((B_Class >= 0 AND B_Class <= 14) OR (B_Class >= 18 AND B_Class <= 19) OR (B_Class = 21) OR (B_Class = 24)) AND B_AccountType = 0';

      if ((GetBilling(FileArray[idx1],ThisStr,'Specified Account')) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      txtError.Text := 'Processing: ' + FileArray[idx1];
      txtError.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

      prbProgress.StepIt;
      prbProgress.Refresh;

//--- Only process files that have records

      if (Query1.RecordCount > 0) then begin

//--- Open the Excel workbook template

         xlsBook := TXLSWorkbook.Create;
         xlsBook.Open(Template_A);
         xlsSheet := xlsBook.ActiveSheet;
         xlsSheet.Name := 'Specified Account (' + FileArray[idx1] + ')';

//--- Clear everything but keep the Header information if lcShowHeader is set

         if (lcShowHeader = true) then
            xlsSheet.RCRange[lcHER + 1, 1, 999, lcHEC].Clear
         else
            xlsSheet.RCRange[1, 1, 999, lcSMaxCols].Clear;

//--- Insert the Page 1 Heading

         Generate_DocHeading(xlsSheet,PageNum,idx1,lcHER + 1,lcHEC,'Specified Account');

//--- Insert the Customer Information

         if (lcShowAddress = true) then begin

            GetAddress(FileArray[idx1]);

            with xlsSheet.RCRange[lcASR,1,lcAER,lcHEC] do begin
               Item[1,lcASC].Value := Customer;
               Item[2,lcASC].Value := Address1;
               Item[3,lcASC].Value := Address2;
               Item[4,lcASC].Value := Address3;
               Item[5,lcASC].Value := Address4;
               Item[6,lcASC].Value := Address5;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;

//--- Insert the instruction information

         if (lcShowInstruct = true) then begin
            with xlsSheet.RCRange[lcISR,1,lcISR + 1,lcHEC] do begin
               Item[1,lcISCL].Value := 'Client:';
               Item[2,lcISCL].Value := 'Instruction:';
               Item[1,lcISCD].Value := Customer;
               Item[2,lcISCD].Value := Descrip;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;

//--- Insert the Banking details

         if (lcShowBanking = true) then begin
            for idx3 := 0 to BankStrings.Count - 1 do begin
               with xlsSheet.RCRange[lcBSR + idx3,1,lcBSR + idx3,lcHEC] do begin
                  Item[1,1].Value := DoSymVars(BankStrings.Strings[idx3],FileArray[idx1]);
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;
            end;
         end;

//--- Write the Data Heading Information

         Generate_DataHeading(xlsSheet,lcPSR,lcHEC,'Specified Account');

//--- Page 1 data

         if (Query1.RecordCount <= (lcPRows - 3)) then begin
            ThisRows := Query1.RecordCount + 2;
            ThisItem := 'Closing Balance';
         end else begin
            ThisRows := lcPRows - 1;
            ThisItem := 'Carried Over';
         end;

         LastDate := SDate;
         Generate_Detail_Account(xlsSheet, lcPSR, ThisRows, ThisItem, 'Opening Balance');
      end;

//--- Data on subsequent pages - compensate for Document Heading (3 rows)

      if (Query1.RecordCount > (lcPRows - 3)) then begin
         RemainRows := (Query1.RecordCount - (lcProws - 3));
         NumPages := ((Query1.RecordCount - (lcPRows - 3)) div (lcSRows - 3)) + 1;

//--- Compensate for cases where we have an exact page size

         if ((Query1.RecordCount - (lcPRows - 3)) mod (lcSRows - 3) = 0) then
            NumPages := NumPages - 1;

         Row := lcSSR;
         PageNum := PageNum + 1;

         for idx3 := 0 to NumPages -1 do begin
            if (lcHeaderPageOne = false) then begin
               xlsSheet.RCRange[lcHSR,lcHSC,lcHER,lcHEC].Copy(xlsSheet.RCRange[lcSSR,lcHSC,lcSSR + lcHER,lcHEC]);

               Row := Row + lcHER + 1;
            end;

            Generate_DocHeading(xlsSheet,PageNum,idx1,Row,lcHEC,'Specified Account');
            Row := Row + 4;
            Generate_DataHeading(xlsSheet,Row,lcHEC,'Specified Account');

            if (RemainRows <= (lcSRows - 3)) then begin
               ThisRows := RemainRows + 2;
               ThisItem := 'Closing Balance';
            end else begin
               ThisRows := lcSRows - 1;
               ThisItem := 'Carried Over';
            end;

            Generate_Detail_Account(xlsSheet, Row, ThisRows, ThisItem, 'Carried Down');

            PageNum := PageNum + 1;
            RemainRows := RemainRows - lcSRows + 3;
            Row := Row + lcSMaxRows - 4;
         end;
      end else begin
         Row := lcSSR;
      end;

//--- Write the Summary and the standard copyright notice - Note we do not
//--- write this if no records were found

      if (Query1.RecordCount > 0) then begin
         dec(Row);

//--- Insert the Summary Information

         if (lcShowSummary = true) then begin
            with xlsSheet.RCRange[lcXSR,lcXSCL,lcXSR + 6,lcXSCD] do begin
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;

               Item[1,1].Value := 'Business Balance:';
               Item[1,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[1,2].Borders[xlEdgeTop].Weight := xlThin;
               Item[2,1].Value := 'Trust Balance:';
               Item[2,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[2,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[2,2].Borders[xlEdgeTop].Weight := xlThin;

               if (EDate < S86Date) then
                  Item[3,1].Value := 'S78(2A) Balance:'
               else
                  Item[3,1].Value := 'S86(4) Balance:';

               Item[3,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[3,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[3,2].Borders[xlEdgeTop].Weight := xlThin;
               Item[4,1].Value := 'Current Balance:';
               Item[4,1].Font.Bold := True;
               Item[4,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[4,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[4,1].Borders[xlEdgeBottom].Weight := xlThin;
               Item[4,2].Borders[xlEdgeTop].Weight := xlThin;
               Item[4,2].Borders[xlEdgeBottom].Weight := xlThin;

               Item[1,(lcXSCD - lcXSCL) + 1].Value := RoundD(OpenBalFees,2);
               Item[1,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[2,(lcXSCD - lcXSCL) + 1].Value := RoundD(OpenBalTrust,2);
               Item[2,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[3,(lcXSCD - lcXSCL) + 1].Value := RoundD(OpenBal864,2);
               Item[3,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[4,(lcXSCD - lcXSCL) + 1].Value := RoundD(OpenBalFees,2) + RoundD(OpenBal864,2) + RoundD(OpenBalTrust,2);
               Item[4,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[4,(lcXSCD - lcXSCL) + 1].Font.Bold := True;

               Item[6,(lcXSCD - lcXSCL) + 1].Formula := '=ABS(' + FloatToStr(RoundD(OpenBalFees,2) + RoundD(OpenBal864,2) + RoundD(OpenBalTrust,2)) + ')';
               Item[6,(lcXSCD - lcXSCL) + 1].Font.Bold := True;

               if ((RoundD(OpenBalFees,2) + RoundD(OpenBal864,2) + RoundD(OpenBalTrust,2)) > 0) then
                  Item[6,1].Value := 'Owed to Client:'
               else
                  Item[6,1].Value := 'Owed by Client:';

               Item[6,1].Font.Bold := True;
            end;

            with xlsSheet.RCRange[lcXSR,lcXSCD,lcXSR + 5,lcXSCD] do begin
               NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
            end;
         end;

//--- Insert the copyright notice

         with xlsSheet.RCRange[Row,1,Row,1] do begin
            Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 8;
         end;

//--- Remove the Gridlines which are added by default and set the Page orientation

         xlsSheet.PageSetup.Orientation    := xlLandscape;
         xlsSheet.PageSetup.FitToPagesWide := 1;
         xlsSheet.PageSetup.FitToPagesTall := PageNum - 1;
         xlsSheet.PageSetup.PaperSize      := xlPaperA4;
         xlsSheet.DisplayGridLines         := false;
         xlsSheet.PageSetup.CenterFooter   := 'Page &P of &N';

//--- Write the Excel file to disk

         xlsBook.SaveAs(FileName + ThisFile);
         xlsBook.Close;

         LogMsg('  File ''' + FileArray[idx1] + ''' successfully processed...',True);
         LogMsg(' ',True);

         DoLine := False;

//--- Print the generated document on the Default Printer if requested

         if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin

            if (PrintDocument(ThisFile, FileName) = True) then
               LogMsg('  Document submitted for printing...',True)
            else
               LogMsg('  Printing of document failed...',True);

            DoLine := True;
         end;

//--- Create a PDF file if requested

         if ((CreatePDF = true) and (PDFPrinter <> 'Not Found')) then begin
            PDFExists  := PDFDocument(ThisFile, FileName);

            if (PDFExists = True) then
               LogMsg('  PDF file creation was successfull...',True)
            else
               LogMsg('  PDF file creation failed...',True);

            DoLine := True;
         end;

//--- Send the Excel file via email if requested

         if (SendByEmail = '1') then begin
            if (GroupAttach = true) then
               AttachList.Add(FileName + ThisFile)
            else begin

               if (SendEmail(FileArray[idx1], FileName + ThisFile,ord(PT_NORMAL)) = true) then
                  LogMsg('  Request to send generated Specified Account by Email submitted...',True)
               else
                  LogMsg('  Request to send generated Specified Account by Email not submitted...',True);

               DoLine := True;
            end;
         end;

//--- Now open the Specified Account if requested

         if (AutoOpen = True) then begin
            ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);

            LogMsg('  Request to open Specified Account for ''' + FileArray[idx1] + ''' submitted...',True);
            DoLine := True;
         end;

      end else begin
         LogMsg('  No billing data found for ''' + FileArray[idx1] + '''',True);
         DoLine := True;
      end;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Specified Account');

   Close_Connection;
end;

//---------------------------------------------------------------------------
// Generate the Detail for Specified Accounts
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Detail_Account(xlsSheet : IXLSWorksheet; PageRows: integer; ThisRows: integer; ThisItem: string; BalanceStr: string);
var
   idx1, ThisClass   : integer;
   ThisAmount        : double;
   ThisRelated       : string;

begin

//--- Page data

   with xlsSheet.RCRange[PageRows + 1,1,PageRows + ThisRows,lcHEC] do begin
      Borders[xlAround].LineStyle := xlContinuous;
      Borders[xlAround].Weight := xlThin;
   end;

   with xlsSheet.RCRange[PageRows + 1,lcHEC - 2, PageRows + ThisRows, lcHEC] do begin
      NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

//-- Balance Line

   with xlsSheet.RCRange[PageRows + 1,1,PageRows + 1,lcHEC] do begin
      Borders[xlEdgeBottom].LineStyle := xlContinuous;
      Borders[xlEdgeBottom].Weight := xlThin;

      Item[1,1].Value := LastDate;
      Item[1,2].Value := BalanceStr;
      Item[1,lcHEC - 2].Value := OpenBal864;
      Item[1,lcHEC - 1].Value := OpenBalTrust;

      if ((BalanceStr = 'Opening Balance') and (VATRate > 0)) then
         OpenBalFees := OpenBalFees + OpenBalVAT;

      Item[1,lcHEC].Value := OpenBalFees;
      Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;

      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

//--- Page Detail Lines

   for idx1 := 2 to ThisRows - 1 do begin
      with xlsSheet.RCRange[PageRows + idx1,1,PageRows + idx1,lcHEC] do begin
         Borders[xlEdgeBottom].LineStyle := xlContinuous;
         Borders[xlEdgeBottom].Weight := xlThin;
         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
         Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
         Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;

         LastDate := Query1.FieldByName('B_Date').AsString;
         Item[1,1].Value := LastDate;
         ThisAmount := Query1.FieldByName('B_Amount').AsFloat;
         ThisClass := Query1.FieldByName('B_Class').AsInteger;

         if (ShowRelated = True) then
            ThisRelated := '[' + ReplaceQuote(Query1.FieldByName('B_Owner').AsString) + '] '
         else
            ThisRelated := '';

//---
//--- Process the Fees Column
//---

         if (ThisClass in [0..6,8,18,21,24]) then begin

            if (ThisClass in [8,24]) then begin
               ThisAmount := ThisAmount;
            end else if (ThisClass in [18]) then begin
               ThisAmount := ThisAmount * -1;
            end else if (Query1.FieldByName('B_DrCr').AsInteger = 2) then begin
               ThisAmount := ThisAmount;
            end else begin
               ThisAmount := ThisAmount * -1;
            end;

            if (ThisClass = 1) then begin
               Item[1,2].Value := ThisRelated + '[**Disbursement] ' + ReplaceQuote(Query1.FieldByName('B_Description').AsString);
            end else if (ThisClass = 2) then begin
               Item[1,2].Value := ThisRelated + '[**Expense] ' + ReplaceQuote(Query1.FieldByName('B_Description').AsString);
            end else begin
               Item[1,2].Value := ThisRelated + ReplaceQuote(Query1.FieldByName('B_Description').AsString);
            end;

            Item[1,lcHEC].Value := ThisAmount;

//--- Insert a zero into the Trust and S86(4) columns

            if (ThisClass in [0..3,5..6,18,21]) then
               Item[1,lcHEC - 1].Value := 0;

            Item[1,lcHEC - 2].Value := 0;
            OpenBalFees := OpenBalFees + ThisAmount;
         end;

//---
//--- Process the Trust Column
//---

//--- Reload ThisAmount as it may have been changed in the processing for Fees

         ThisAmount := Query1.FieldByName('B_Amount').AsFloat;

         if (ThisClass in [4,7..13,19,24]) then begin

            if (ThisClass = 4) then begin
               ThisAmount := ThisAmount;
            end else if (ThisClass = 12) then begin
               ThisAmount := ThisAmount * -1;
            end else if (ThisClass = 13) then begin
               ThisAmount := ThisAmount;
            end else if (Query1.FieldByName('B_DrCr').AsInteger = 1) then begin
               ThisAmount := ThisAmount * -1;
            end;

            Item[1,2].Value := ThisRelated + ReplaceQuote(Query1.FieldByName('B_Description').AsString);
            Item[1,lcHEC - 1].Value := ThisAmount;

//--- Insert a zero into the Client and S78(2A) columns

            if (ThisClass in [7, 9..11,19]) then
               Item[1,lcHEC].Value := 0;

            Item[1,lcHEC - 2].Value := 0;
            OpenBalTrust := OpenBalTrust + ThisAmount;
         end;

//---
//--- Process the S86(4) Column
//---

//--- Reload ThisAmount as it may have been changed in the processing for Trust

         ThisAmount := Query1.FieldByName('B_Amount').AsFloat;

         if (ThisClass in [12..14]) then begin

            if (ThisClass in [12,14]) then
               ThisAmount := ThisAmount
            else
               ThisAmount := ThisAmount * -1;

            Item[1,2].Value := ThisRelated + ReplaceQuote(Query1.FieldByName('B_Description').AsString);
            Item[1,lcHEC - 2].Value := ThisAmount;

//--- Insert a zero into the Trust and Client columns

            if (ThisClass = 14) then
               Item[1,lcHEC - 1].Value := 0;

            Item[1,lcHEC].Value := 0;

            OpenBal864 := OpenBal864 + ThisAmount;
         end;

         Query1.Next;

         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

//--- Closing Balance Line

//--- We need to force the additon of another line when VAT processing is active

   if ((ThisItem = 'Closing Balance') and (ShowVAT = 1)) then begin
      with xlsSheet.RCRange[PageRows + ThisRows,1,PageRows + ThisRows,lcHEC] do begin
         Borders[xlEdgeBottom].LineStyle := xlContinuous;
         Borders[xlEdgeBottom].Weight := xlThin;

         Item[1,1].Value := EDate;
         Item[1,2].Value := Format('** Plus VAT on %m',[Abs(RoundD((SummaryVAT / (VATRate / 100)),2))]);
         Item[1,lcHEC].Value := SummaryVAT;
         OpenBalFees := OpenBalFees + SummaryVAT;

         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
         Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
         Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;

         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;

      with xlsSheet.RCRange[PageRows + ThisRows + 1,1,PageRows + ThisRows + 1,lcHEC] do begin
         Borders[xlAround].LineStyle := xlContinuous;
         Borders[xlAround].Weight := xlThin;
      end;

      with xlsSheet.RCRange[PageRows + ThisRows + 1,lcHEC - 2, PageRows + ThisRows + 1, lcHEC] do begin
         NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(ThisRows);
   end;

//--- Now do the final line for this page

   with xlsSheet.RCRange[PageRows + ThisRows,1,PageRows + ThisRows,lcHEC] do begin
      Borders[xlEdgeBottom].LineStyle := xlContinuous;
      Borders[xlEdgeBottom].Weight := xlThin;

      if (ThisItem = 'Closing Balance') then
         Item[1,1].Value := EDate
      else
         Item[1,1].Formula := LastDate;

      Item[1,2].Value         := ThisItem;
      Item[1,lcHEC - 2].Value := RoundD(OpenBal864,2);
      Item[1,lcHEC - 1].Value := RoundD(OpenBalTrust,2);
      Item[1,lcHEC].Value     := RoundD(OpenBalFees,2);

      Item[1,1].Borders[xlEdgeRight].Weight        := xlThin;
      Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,lcHEC].Borders[xlEdgeLeft].Weight     := xlThin;

      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;
end;

//---------------------------------------------------------------------------
// Generate Call Records
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_CallRecs();
var
   idx1, row, Pages : integer;
   RecCount         : integer;
   DoLine           : boolean;
   xlsBook          : IXLSWorkbook;
   xlsSheet         : IXLSWorksheet;

   ExportSet, SectionSet, RecordSet : IXMLNode;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Call Records');

//--- Get Company specific information

   GetCpyVAT();

//--- Open the XML file and process the content

   XMLDoc.Active := false;
   XMLDoc.LoadFromFile(HostName);
   XMLDoc.Active := true;

   DeleteFile(HostName);

   ExportSet  := XMLDoc.DocumentElement;

//--- Create the Excel spreadsheet

   xlsBook       := TXLSWorkbook.Create;
   xlsSheet      := xlsBook.WorkSheets.Add;
   xlsSheet.Name := 'Call Records';

//--- Write the Report Heading

   with xlsSheet.Range['A1', 'H1'] do begin
      Item[1,1].Value := CpyName + ': Asterisk Interface Call Record Log for the Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
      Borders[xlAround].Weight := xlThin;
      Interior.Color := ColDHF;
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 11;
   end;

//--- Write the section heading for Encoded Outgoing Records

   with xlsSheet.Range['A3','H3'] do begin
      Item[1,1].Value := 'Encoded Outgoing Records:';
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 11;
   end;

   with xlsSheet.Range['A4','H4'] do begin
      Item[1,1].Value := 'Date';
      Item[1,1].ColumnWidth := 14;
      Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,2].Value := 'Caller';
      Item[1,2].ColumnWidth := 35;
      Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,3].Value := 'Number';
      Item[1,3].ColumnWidth := 14;
      Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,4].Value := 'File';
      Item[1,4].ColumnWidth := 14;
      Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].Value := 'Duration (Sec)';
      Item[1,5].ColumnWidth := 14;
      Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].HorizontalAlignment := xlHAlignRight;
      Item[1,6].Value := 'Billing (Min)';
      Item[1,6].ColumnWidth := 14;
      Item[1,6].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,6].HorizontalAlignment := xlHAlignRight;
      Item[1,7].Value := 'Description';
      Item[1,7].ColumnWidth := 45;
      Item[1,7].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,8].Value := 'Client Name';
      Item[1,8].ColumnWidth := 45;
      Borders[xlAround].Weight := xlThin;
      Interior.Color := integer(ColDHF);
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   row := 5;

//--- Process the Outgoing Encoded records

   SectionSet := ExportSet.ChildNodes.First;
   RecordSet  := SectionSet.ChildNodes.First;

//--- Now process the records in this section if any

   RecCount := StrToInt(SectionSet.ChildValues['count']);

   if RecCount > 0 then begin
      RecordSet := RecordSet.NextSibling;
      for idx1 := 1 to RecCount do begin
         with xlsSheet.RCRange[row,1,row,8] do begin
            Item[1,1].Value := ReplaceXML(RecordSet.ChildValues['date']);
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := ReplaceXML(RecordSet.ChildValues['caller']);
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := ReplaceXML(RecordSet.ChildValues['number']);
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Value := ReplaceXML(RecordSet.ChildValues['file']);
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].Value := StrToInt(ReplaceXML(RecordSet.ChildValues['seconds']));
            Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].NumberFormat := '#,##0.00_)';
            Item[1,6].Value := StrToInt(ReplaceXML(RecordSet.ChildValues['billing']));
            Item[1,6].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,6].NumberFormat := '#,##0.00_)';
            Item[1,7].Value := ReplaceXML(RecordSet.ChildValues['description']);
            Item[1,7].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,8].Value := ReplaceXML(RecordSet.ChildValues['client']);
            Borders[xlAround].Weight := xlThin;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         RecordSet := RecordSet.NextSibling;
         inc(row);
      end;
   end else begin
      with xlsSheet.RCRange[row,1,row,8] do begin
         Item[1,1].Value := 'No records found';
         Borders[xlAround].Weight := xlThin;
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row);
   end;

   row := row + 2;

//--- Process Outgoing Possible Records

//--- Write the section heading for Outgoing Possible Records

   with xlsSheet.Range['A' + IntToStr(row),'H' + IntToStr(row)] do begin
      Item[1,1].Value := 'Possible Outgoing Records:';
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 11;
   end;

   inc(row);

   with xlsSheet.Range['A' + IntToStr(row),'H' + IntToStr(row)] do begin
      Item[1,1].Value := 'Date';
      Item[1,1].ColumnWidth := 14;
      Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,2].Value := 'Caller';
      Item[1,2].ColumnWidth := 35;
      Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,3].Value := 'Number';
      Item[1,3].ColumnWidth := 14;
      Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,4].Value := 'File';
      Item[1,4].ColumnWidth := 14;
      Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].Value := 'Duration (Sec)';
      Item[1,5].ColumnWidth := 14;
      Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].HorizontalAlignment := xlHAlignRight;
      Item[1,6].Value := 'Billing (Min)';
      Item[1,6].ColumnWidth := 14;
      Item[1,6].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,6].HorizontalAlignment := xlHAlignRight;
      Item[1,7].Value := 'Description';
      Item[1,7].ColumnWidth := 45;
      Item[1,7].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,8].Value := 'Client Name';
      Item[1,8].ColumnWidth := 45;
      Borders[xlAround].Weight := xlThin;
      Interior.Color := integer(ColDHF);
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   inc(row);

//--- Process the Outgoing Possible record

   SectionSet := SectionSet.NextSibling;
   RecordSet  := SectionSet.ChildNodes.First;

//--- Now process the records in this section if any

   RecCount := StrToInt(SectionSet.ChildValues['count']);

   if RecCount > 0 then begin
      RecordSet := RecordSet.NextSibling;
      for idx1 := 1 to RecCount do begin
         with xlsSheet.RCRange[row,1,row,8] do begin
            Item[1,1].Value := ReplaceXML(RecordSet.ChildValues['date']);
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := ReplaceXML(RecordSet.ChildValues['caller']);
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := ReplaceXML(RecordSet.ChildValues['number']);
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Value := ReplaceXML(RecordSet.ChildValues['file']);
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].Value := StrToInt(ReplaceXML(RecordSet.ChildValues['seconds']));
            Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].NumberFormat := '#,##0.00_)';
            Item[1,6].Value := StrToInt(ReplaceXML(RecordSet.ChildValues['billing']));
            Item[1,6].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,6].NumberFormat := '#,##0.00_)';
            Item[1,7].Value := ReplaceXML(RecordSet.ChildValues['description']);
            Item[1,7].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,8].Value := ReplaceXML(RecordSet.ChildValues['client']);
            Borders[xlAround].Weight := xlThin;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         RecordSet := RecordSet.NextSibling;
         inc(row);
      end;
   end else begin
      with xlsSheet.RCRange[row,1,row,8] do begin
         Item[1,1].Value := 'No records found';
         Borders[xlAround].Weight := xlThin;
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row);
   end;

   row := row + 2;

//--- Process the Incomming Possible record

//--- Write the section heading for Incomming Possible Records

   with xlsSheet.Range['A' + IntToStr(row),'H' + IntToStr(row)] do begin
      Item[1,1].Value := 'Possible Incomming Records:';
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 11;
   end;

   inc(row);

   with xlsSheet.Range['A' + IntToStr(row),'H' + IntToStr(row)] do begin
      Item[1,1].Value := 'Date';
      Item[1,1].ColumnWidth := 14;
      Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,2].Value := 'Caller';
      Item[1,2].ColumnWidth := 35;
      Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,3].Value := 'Number';
      Item[1,3].ColumnWidth := 14;
      Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,4].Value := 'File';
      Item[1,4].ColumnWidth := 14;
      Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].Value := 'Duration (Sec)';
      Item[1,5].ColumnWidth := 14;
      Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].HorizontalAlignment := xlHAlignRight;
      Item[1,6].Value := 'Billing (Min)';
      Item[1,6].ColumnWidth := 14;
      Item[1,6].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,6].HorizontalAlignment := xlHAlignRight;
      Item[1,7].Value := 'Description';
      Item[1,7].ColumnWidth := 45;
      Item[1,7].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,8].Value := 'Client Name';
      Item[1,8].ColumnWidth := 45;
      Borders[xlAround].Weight := xlThin;
      Interior.Color := integer(ColDHF);
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   inc(row);

//--- Process the Incomming Possible record

   SectionSet := SectionSet.NextSibling;
   RecordSet  := SectionSet.ChildNodes.First;

//--- Now process the records in this section if any

   RecCount := StrToInt(SectionSet.ChildValues['count']);

   if RecCount > 0 then begin
      RecordSet := RecordSet.NextSibling;
      for idx1 := 1 to RecCount do begin
         with xlsSheet.RCRange[row,1,row,8] do begin
            Item[1,1].Value := ReplaceXML(RecordSet.ChildValues['date']);
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := ReplaceXML(RecordSet.ChildValues['caller']);
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := ReplaceXML(RecordSet.ChildValues['number']);
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Value := ReplaceXML(RecordSet.ChildValues['file']);
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].Value := StrToInt(ReplaceXML(RecordSet.ChildValues['seconds']));
            Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].NumberFormat := '#,##0.00_)';
            Item[1,6].Value := StrToInt(ReplaceXML(RecordSet.ChildValues['billing']));
            Item[1,6].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,6].NumberFormat := '#,##0.00_)';
            Item[1,7].Value := ReplaceXML(RecordSet.ChildValues['description']);
            Item[1,7].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,8].Value := ReplaceXML(RecordSet.ChildValues['client']);
            Borders[xlAround].Weight := xlThin;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         RecordSet := RecordSet.NextSibling;
         inc(row);
      end;
   end else begin
      with xlsSheet.RCRange[row,1,row,8] do begin
         Item[1,1].Value := 'No records found';
         Borders[xlAround].Weight := xlThin;
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

   row := (Pages * lcGRows);

   with xlsSheet.RCRange[row,1,row,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet.PageSetup.Orientation := xlLandscape;
   xlsSheet.PageSetup.FitToPagesWide := 1;
   xlsSheet.PageSetup.FitToPagesTall := Pages;
   xlsSheet.PageSetup.PaperSize := xlPaperA4;
   xlsSheet.DisplayGridLines := false;
   xlsSheet.PageSetup.CenterFooter := 'Page &P of &N';

//--- Save the workbook

   xlsBook.SaveAs(FileName);
   XMLDoc.Active := false;

   LogMsg('  Call Records successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = true) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Call Record log by Email submitted...',True)
      else
         LogMsg('  Request to send generated Call Record log by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the Call Log if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName),nil,nil,SW_SHOWNORMAL);
      LogMsg('  Request to open Call Record log submitted...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Call Records');

   DeleteFile(HostName);

end;

//---------------------------------------------------------------------------
// Export Notes for a File
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Notes();
var
   idx1, row, PageRow, Pages, RowsPerPage, sema  : integer;
   FileCount                                     : integer;
   PageBreak, FirstPage, DoLine                  : boolean;
   ThisFile, ThisName                            : string;
   ThisAB1F, ThisAB1T, ThisAB2F, ThisAB2T        : TColor;
   ThisFill, ThisText                            : TColor;
   xlsBook                                       : IXLSWorkbook;
   xlsSheet                                      : IXLSWorksheet;

begin

   ThisCount := 1;
   ThisMax   := IntToStr(NumFiles);

   prbProgress.Max := NumFiles;
   LogMsg(ord(PT_PROLOG),True,True,'Notes Export');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Process all the records in FileArray

   FileCount   := 0;

   for idx1 := 0 to NumFiles - 1 do begin

      ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - Exported Notes (' + FileArray[idx1] + ').xls';
      txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;
      txtDocument.Refresh;

//--- Read the Notes records from the datastore

      if ((GetNotes(FileArray[idx1])) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      txtError.Text := 'Processing: ' + FileArray[idx1];
      txtError.Refresh;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

//--- Only process if there are Notes records

      if (Query1.RecordCount > 0) then begin

         Pages     := 0;
         FirstPage := true;
         PageBreak := true;

         GroupAttach := true;
         inc(FileCount);
         ThisName := FileArray[idx1];

//--- Open the Excel workbook template that will contain the Specified Account

         xlsBook       := TXLSWorkbook.Create;
         xlsSheet      := xlsBook.WorkSheets.Add;
         xlsSheet.Name := 'Notes (' + FileArray[idx1] + ')';
         sema := 0;

//--- Set up to use alternate color blocks

         ThisAB1F := ColAB1F;
         ThisAB1T := ColAB1T;
         ThisAB2F := ColAB2F;
         ThisAB2T := ColAB2T;

//--- Now step through each Notes record

         Query1.First;

         while Query1.Eof = False do begin

//--- Perform a Page break if necessary

            if (PageBreak = True) then
               DoPageBreak(ord(PB_FILENOTES),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,FileArray[idx1]);

            if (sema = 0) then begin;
               ThisFill := ThisAB1F;
               ThisText := ThisAB1T;
               sema := 1;
            end else begin
               ThisFill := ThisAB2F;
               ThisText := ThisAB2T;
               sema := 0;
            end;

            with xlsSheet.RCRange[row,1,row,4] do begin
               Item[1,1].Value := Query1.FieldByName('Notes_Date').AsString;
               Item[1,1].VerticalAlignment := xlVAlignTop;
               Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,2].Value := Query1.FieldByName('Notes_Time').AsString;
               Item[1,2].VerticalAlignment := xlVAlignTop;
               Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,3].Value := ReplaceQuote(Query1.FieldByName('Notes_Note').AsString);
               Item[1,3].WrapText := ThisWrapText;
               Item[1,3].VerticalAlignment := xlVAlignTop;
               Item[1,3].Borders[xlEdgeRight].Weight := xlThin;

               if(Query1.FieldByName('Notes_User').AsString = '') then
                  Item[1,4].Value := ' '
               else
                  Item[1,4].Value := ReplaceQuote(Query1.FieldByName('Notes_User').AsString);

               Item[1,4].VerticalAlignment := xlVAlignTop;
               Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
               Interior.Color := ThisFill;
               Font.Color := ThisText;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
               Borders[xlAround].Weight := xlThin;
            end;

            Query1.Next;

            inc(row);
            inc(PageRow);

//--- If we've reached the maximum rows per page then it is PageBreak time.
//--- UNLESS WarpText is true in which case we do not do Pagebreak other than
//--- on the fist page. We do the Pagebreak here so that the Copyright notice
//--- will be handled correctly

            if (ThisWrapText = False) then begin
               if (PageRow >= RowsPerPage) then begin
                  PageBreak := True;
                  DoPageBreak(ord(PB_FILENOTES),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,FileArray[idx1]);
               end;
            end;
         end;

         inc(row);

         with xlsSheet.RCRange[row,1,row,1] do begin
            Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 8;
         end;

//--- Remove the Gridlines which are added by default and set the Page orientation

         xlsSheet.PageSetup.Orientation := xlLandscape;
         xlsSheet.PageSetup.PaperSize := xlPaperA4;
         xlsSheet.DisplayGridLines := false;
         xlsSheet.PageSetup.CenterFooter := 'Page &P of &N';
         xlsSheet.PageSetup.FitToPagesWide := 1;

         if (ThisWrapText = False) then
            xlsSheet.PageSetup.FitToPagesTall := Pages;

//--- Write the Excel file to disk

         xlsBook.SaveAs(FileName + ThisFile);
         LogMsg('  File ''' + FileArray[idx1] + ''' successfully processed...',True);
         LogMsg(' ',True);

         DoLine := False;

//--- Print the generated document on the Default Printer if requested

         if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
            if (PrintDocument(ThisFile, FileName) = True) then
               LogMsg('  Document submitted for printing...',True)
            else
               LogMsg('  Printing of document failed...',True);

            DoLine := True;
         end;

//--- Create a PDF file if requested

         if ((CreatePDF = true) and (PDFPrinter <> 'Not Found')) then begin
            PDFExists := PDFDocument(ThisFile, FileName);

            if (PDFExists = True) then
               LogMsg('  PDF file creation was successfull...',True)
            else
               LogMsg('  PDF file creation failed...',True);

            DoLine := True;
         end;

//--- Add the File to the list of files to be attached

         if (SendByEmail = '1') then
            AttachList.Add(FileName + ThisFile);

//--- Now open the Exported Notes if requested

         if (AutoOpen = True) then begin
            ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);

            LogMsg('  Request to open Exported Notes for ' + FileArray[idx1] + ' submitted...',True);

            DoLine := True;
         end;
      end else begin
         LogMsg('  No Notes records found for ' + FileArray[idx1],True);
         DoLine := True;
      end;
   end;

//--- Send the Excel files via email if requested and if there are any

   if (FileCount > 1) then ThisName := '';

   if (GroupAttach = true) and (SendByEmail = '1') then begin
      if (SendEmail(ThisName,'',ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Expoted Notes by Email submitted...',True)
      else
         LogMsg('  Request to send generated Expoted Notes by Email failed...',True);

      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Notes Export');

end;

//---------------------------------------------------------------------------
// Generate Invoices
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Invoice();
var
   PageNum, ThisRows, ThisCol, ThisClass, Row, NumPages   : integer;
   idx1, idx2, idx3, RemainRows                           : integer;
   ThisAmount                                             : double;
   DoLine                                                 : boolean;
   ThisMsg, ThisVal, ThisItem, ThisFile, ThisStr, ThisVat : string;
   ThisInvoice, CurrentFile, S1                           : string;
   ThisDate                                               : TDateTime;
   xlsBook                                                : IXLSWorkbook;
   xlsSheet                                               : IXLSWorksheet;

begin

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Get the Heading and Layout information

   GetLayout('Invoice');

//--- Read the Billing records from the datastore

   for idx1 := 0 to NumFiles - 1 do begin

//--- If this is an invoice to be re-generated then the File Array contains
//--- Invoices and not Files. We need to get the File, AcctType, SDate,
//--- EDate and InvoiceNum in order to process this request

      if ((InvoiceInfo <> '0') and (InvoiceInfo <> '1')) then begin
         CurrentFile := GetInvoiceData(FileArray[idx1]);
         ThisInvoice := FileArray[idx1];
      end else begin
         CurrentFile := FileArray[idx1];
      end;

      ThisVat := GetVATNum(CurrentFile);
      PageNum := 1;
      ThisDate := StrToDate(EDate);
      ThisFile := FormatDateTime('yyyyMMdd',ThisDate) + ' - Invoice (' + CurrentFile + ').xls';
      txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;
      txtDocument.Refresh;

//--- Initialise the Summary Fields

      SummaryFees        := 0;
      SummaryDisburse    := 0;
      SummaryExpenses    := 0;
      SummaryVAT         := 0;

//--- Now process this invoice

      if (InvoiceInfo = '1') then
         S1 := 'Collect'
      else
         S1 := 'B';

      ThisStr := ' AND ((' + S1 + '_Class >= 0 AND ' + S1 + '_Class <= 2) OR (' + S1 + '_Class = 5) OR (' + S1 + '_Class = 21)) AND ' + S1 + '_AccountType = ' + IntToStr(AccountType);

      if ((GetBilling(CurrentFile,ThisStr,'Invoice')) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      txtError.Text := 'Processing: ' + CurrentFile;
      txtError.Refresh;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

//--- Only process files that have records

      if (Query1.RecordCount > 0) then begin

//--- Open the Excel workbook template

         xlsBook := TXLSWorkbook.Create;
         xlsBook.Open(Template_I);
         xlsSheet := xlsBook.ActiveSheet;
         xlsSheet.Name := 'Invoice (' + CurrentFile + ')';

//--- Clear everything but keep the Header information if lcShowHeader is set

         if (lcShowHeader = true) then
            xlsSheet.RCRange[lcHER + 1, 1, 999, lcHEC].Clear
         else
            xlsSheet.RCRange[1, 1, 999, lcSMaxCols].Clear;

//--- Insert the Page 1 Heading

         Generate_DocHeading(xlsSheet,PageNum,idx1,lcHER + 1,lcHEC,'Invoice');

//--- Insert the Customer Information

         if (lcShowAddress = true) then begin

            GetAddress(CurrentFile);

            with xlsSheet.RCRange[lcASR,1,lcAER,lcHEC] do begin
               Item[1,lcASC].Value := Customer;
               Item[2,lcASC].Value := Address1;
               Item[3,lcASC].Value := Address2;
               Item[4,lcASC].Value := Address3;
               Item[5,lcASC].Value := Address4;
               Item[6,lcASC].Value := Address5;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;

//--- Insert the instruction information

         if (lcShowInstruct = true) then begin
            with xlsSheet.RCRange[lcISR,1,lcISR + 2,lcHEC] do begin
               Item[1,1].Value := 'Client:';
               Item[2,1].Value := 'Instruction:';
               Item[3,1].Value := 'VAT Num:';
               Item[1,lcISCD].Value := Customer;
               Item[2,lcISCD].Value := Descrip;
               Item[3,lcISCD].Value := ThisVat;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;

//--- Insert the Summary Information

         if (lcShowSummary = true) then begin
            with xlsSheet.RCRange[lcXSR,lcXSCL,lcXSR + 3,lcXSCD] do begin
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;

               Item[1,1].Value := 'Fees:';
               Item[1,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[1,2].Borders[xlEdgeTop].Weight := xlThin;
               Item[2,1].Value := 'Disbursements:';
               Item[2,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[2,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[2,2].Borders[xlEdgeTop].Weight := xlThin;
               Item[3,1].Value := 'Expenses:';
               Item[3,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[3,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[3,2].Borders[xlEdgeTop].Weight := xlThin;
               Item[4,1].Value := 'Total for this Invoice:';
               Item[4,1].Font.Bold := True;
               Item[4,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[4,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[4,1].Borders[xlEdgeBottom].Weight := xlThin;
               Item[4,2].Borders[xlEdgeTop].Weight := xlThin;
               Item[4,2].Borders[xlEdgeBottom].Weight := xlThin;

               Item[1,(lcXSCD - lcXSCL) + 1].Value := SummaryFees * -1;
               Item[1,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[2,(lcXSCD - lcXSCL) + 1].Value := SummaryDisburse * -1;
               Item[2,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[3,(lcXSCD - lcXSCL) + 1].Value := SummaryExpenses * -1;
               Item[3,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[4,(lcXSCD - lcXSCL) + 1].Value := (SummaryFees + SummaryDisburse + SummaryExpenses) * -1;
               Item[4,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[4,(lcXSCD - lcXSCL) + 1].Font.Bold := True;
            end;

            with xlsSheet.RCRange[lcXSR,lcXSCD,lcXSR + 3,lcXSCD] do begin
               NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
            end;
         end;

//--- Insert the Banking details

         if (lcShowBanking = true) then begin
            for idx3 := 0 to BankStrings.Count - 1 do begin
               with xlsSheet.RCRange[lcBSR + idx3,1,lcBSR + idx3,lcHEC] do begin
                  Item[1,1].Value := DoSymVars(BankStrings.Strings[idx3],CurrentFile);
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;
            end;
         end;

//--- Write the Data Heading Information

         Generate_DataHeading(xlsSheet,lcPSR,lcHEC,'Invoice');

//--- Page 1 data

         if (Query1.RecordCount <= (lcPRows - 3)) then begin
            ThisRows := Query1.RecordCount + 1;
            ThisItem := 'Total for this invoice';
         end else begin
            ThisRows := lcPRows - 1;
            ThisItem := 'Carried Over';
         end;

         OpenBalFees   := 0;
         OpenBalVAT    := 0;
         OpenBalAmount := 0;

         LastDate := Sdate;
         Generate_Detail_Invoice(xlsSheet, lcPSR, ThisRows, ThisItem, '1');
      end;

//--- Data on subsequent pages - compensate for Document Heading (3 rows)

      if (Query1.RecordCount > (lcPRows - 2)) then begin
         RemainRows := (Query1.RecordCount - (lcProws - 2));
         NumPages := ((Query1.RecordCount - (lcPRows - 3)) div (lcSRows - 3)) + 1;

//--- Compensate for cases where we have an exact page size

         if ((Query1.RecordCount - (lcPRows - 3)) mod (lcSRows - 3) = 0) then
            NumPages := NumPages - 1;

         Row := lcSSR;
         PageNum := PageNum + 1;

         for idx3 := 0 to NumPages -1 do begin
            if (lcHeaderPageOne = false) then begin
               xlsSheet.RCRange[lcHSR,lcHSC,lcHER,lcHEC].Copy(xlsSheet.RCRange[lcSSR,lcHSC,lcSSR + lcHER,lcHEC]);

               Row := Row + lcHER + 1;
            end;

            Generate_DocHeading(xlsSheet,PageNum,idx1,Row,lcHEC,'Invoice');
            Row := Row + 4;
            Generate_DataHeading(xlsSheet,Row,lcHEC,'Invoice');

            if (RemainRows <= (lcSRows - 3)) then begin
               ThisRows := RemainRows + 2;
               ThisItem := 'Total for this invoice';
            end else begin
               ThisRows := lcSRows - 1;
               ThisItem := 'Carried Over';
            end;

            Generate_Detail_Invoice(xlsSheet, Row, ThisRows, ThisItem, '2');

            PageNum := PageNum + 1;
            RemainRows := RemainRows - lcSRows + 3;
            Row := Row + lcSMaxRows - 4;
         end;
      end else begin
         Row := lcSSR;
      end;

//--- Write the standard copyright notice - Note we do not write this if no
//---    records were found

      if (Query1.RecordCount > 0) then begin
         dec(Row);

         with xlsSheet.RCRange[Row,1,Row,1] do begin
            Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 8;
         end;

//--- Store the Invoice Information if this was requested. If InvoiceInfo is
//--- not equal to '0' or '1' then this is a request to regenerate an existing
//--- invoice and NilBalance contains the previously allocated Invoice number

         if (StoreInv = true) then begin

            if ((InvoiceInfo <> '0') and (InvoiceInfo <> '1')) then begin
               ThisMsg := '  Invoice information for ''' + ThisInvoice + ''' regenerated...';
            end else begin
               if (GetNextInvoice(1) = false) then
                  ThisMsg := '  Failed to get next sequential Invoice number for ''' + CurrentFile + ''' ...'
               else begin

                  if (StoreInvoice(CurrentFile,SaveAmount * -1, SummaryFees * -1, SummaryDisburse * -1, SummaryExpenses * -1) = false) then begin
                     MessageDlg(ErrMsg, mtWarning, [mbOK], 0);
                     ThisMsg := '  Failed to store Invoice information for ''' + CurrentFile + ''' ...';
                  end else
                     ThisMsg := '  Invoice information for ''' + CurrentFile + ''' stored...';

                  ThisInvoice := InvoicePref + InvoiceStr;
               end;
            end;

            LogMsg(ThisMsg,True);

//--- Put the invoice number on the Invoice

            if (lcShowAge = true) then begin
               with xlsSheet.RCRange[lcAASR,lcAASC,lcAASR,lcAASC + 2] do begin
                  Item[1,1].Value := 'TAX INVOICE NUMBER: ';
                  Item[1,3].Value := ThisInvoice;
                  Item[1,3].HorizontalAlignment := xlHAlignRight;
                  Font.Bold := true;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;
            end;
         end;

//--- Remove the Gridlines which are added by default and set the Page orientation

         xlsSheet.PageSetup.Orientation    := xlLandscape;
         xlsSheet.PageSetup.FitToPagesWide := 1;
         xlsSheet.PageSetup.FitToPagesTall := PageNum - 1;
         xlsSheet.PageSetup.PaperSize      := xlPaperA4;
         xlsSheet.DisplayGridLines         := false;
         xlsSheet.PageSetup.CenterFooter    := 'Page &P of &N';

//--- Write the Excel file to disk

         xlsBook.SaveAs(FileName + ThisFile);
         xlsBook.Close;

         LogMsg('  File ''' + CurrentFile + ''' successfully processed...',True);
         LogMsg(' ',True);

         DoLine := True;

//--- Print the generated document on the Default Printer if requested

         if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
            if (PrintDocument(ThisFile, FileName) = True) then
               LogMsg('  Document submitted for printing...',True)
            else
               LogMsg('  Printing of document failed...',True);

            DoLine := True;
         end;

//--- Create a PDF file if requested

         if ((CreatePDF = true) and (PDFPrinter <> 'Not Found')) then begin
            PDFExists := PDFDocument(ThisFile, FileName);

            if (PDFExists = True) then
               LogMsg('  PDF file creation was successfull...',True)
            else
               LogMsg('  PDF file creation failed...',True);

            DoLine := True;
         end;

//--- Send the Excel file via email if requested. If SendByEmail = '1' then the
//--- value of GroupAttach determines whether we send the invoice now (false) or
//--- later as part of a group (true). The latter happens when this invoice is
//--- part of a 'Billing' group (Invoice and Statement)

         if (SendByEmail = '1') then begin
            if (GroupAttach = true) then
               AttachList.Add(FileName + ThisFile)
            else begin
               if (SendEmail(CurrentFile,FileName + ThisFile,ord(PT_NORMAL)) = true) then
                  LogMsg('  Request to send generated Invoice by Email submitted...',True)
               else
                  LogMsg('  Request to send generated Invoice by Email not submitted...',True);

               DoLine := True;
            end;
         end;

//--- If SendByEmail = '2' then we will send the ivoice later as part of a group

         if (SendByEmail = '2') then begin
            AttachList.Add(FileName + ThisFile);
         end;

//--- Now open the Invoice if requested

         if (AutoOpen = True) then begin
            ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);
            LogMsg('  Request to open Invoice for ''' + CurrentFile + ''' submitted...',True);
            DoLine := True;
         end;

      end else begin
         LogMsg('  No billing data found for ''' + CurrentFile + '''',True);
         DoLine := True;
      end;
      FldExcel.Refresh;
   end;


//--- Send the group of invoices per email if Parameter 17 = '2'

   if (SendByEmail = '2') then begin
      if (SendEmail('','',ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Invoice(s) by Email submitted...',True)
      else
         LogMsg('  Request to send generated Invoice(s) by Email failed...',True);

      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

//--- Finish up

   Close_Connection;

end;

//---------------------------------------------------------------------------
// Generate the Detail for Invoices
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Detail_Invoice(xlsSheet : IXLSWorksheet; PageRows: integer; ThisRows: integer; ThisItem: string; BalanceStr: string);
var
   idx1, ThisClass   : integer;
   ThisAmount        : double;
   ThisRelated, Pref : string;

begin

//--- Page data

   with xlsSheet.RCRange[PageRows + 1,1,PageRows + ThisRows,lcHEC] do begin
      Borders[xlAround].LineStyle := xlContinuous;
      Borders[xlAround].Weight := xlThin;
   end;

   with xlsSheet.RCRange[PageRows + 1,lcHEC - 2, PageRows + ThisRows, lcHEC] do begin
      NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

//-- Opening Balance Line - only used on 2nd Page and onwards

   if (BalanceStr = '2') then begin
      with xlsSheet.RCRange[PageRows + 1,1,PageRows + 1,lcHEC] do begin
         Borders[xlEdgeBottom].LineStyle := xlContinuous;
         Borders[xlEdgeBottom].Weight := xlThin;

         Item[1,1].Value := LastDate;
         Item[1,2].Value := 'Carried Down';

         if (VATRate = 0) then begin
            Item[1,lcHEC].Value := OpenBalFees;
         end else begin
            Item[1,lcHEC - 2].Value := OpenBalFees;
            Item[1,lcHEC - 1].Value := OpenBalVAT;
            Item[1,lcHEC].Value := OpenBalAmount;
            Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
         end;

         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;

         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

//--- Page Detail Lines

   if (InvoiceInfo = '1') then
      Pref := 'Collect'
   else
      Pref := 'B';

   for idx1 := StrToInt(BalanceStr) to ThisRows - 1 do begin
      with xlsSheet.RCRange[PageRows + idx1,1,PageRows + idx1,lcHEC] do begin
         Borders[xlEdgeBottom].LineStyle := xlContinuous;
         Borders[xlEdgeBottom].Weight := xlThin;
         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;

         if (VATRate = 0) then begin
            Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;
         end else begin
            Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;
         end;

         LastDate := Query1.FieldByName(Pref + '_Date').AsString;
         Item[1,1].Value := LastDate;
         ThisAmount := Query1.FieldByName(Pref + '_Amount').AsFloat;
         ThisClass := Query1.FieldByName(Pref + '_Class').AsInteger;

         if (ThisClass in [0..2,5,21]) then begin

            if (Query1.FieldByName(Pref + '_DrCr').AsInteger = 2) then begin
               ThisAmount := ThisAmount * -1;
            end else begin
               ThisAmount := ThisAmount;
            end;

            if (ShowRelated = True) then
               ThisRelated := '[' + ReplaceQuote(Query1.FieldByName(Pref + '_Owner').AsString) + '] '
            else
               ThisRelated := '';

            if (ThisClass = 1) then begin
               Item[1,2].Value := ThisRelated + '[**Disbursement] ' + ReplaceQuote(Query1.FieldByName(Pref + '_Description').AsString);
            end else if (ThisClass = 2) then begin
               Item[1,2].Value := ThisRelated + '[**Expense] ' + ReplaceQuote(Query1.FieldByName(Pref + '_Description').AsString);
            end else begin
               Item[1,2].Value := ThisRelated + ReplaceQuote(Query1.FieldByName(Pref + '_Description').AsString);
            end;

            if (VATRate > 0) then begin
               Item[1,lcHEC - 2].Value := ThisAmount;
               OpenBalFees := OpenBalFees + ThisAmount;

               if (ThisClass in [0,5,21]) then begin
                  Item[1,lcHec - 1].Value := (ThisAmount * VATRate) / 100;
                  Item[1,lcHec].Value := (ThisAmount * (100 + VATRate)) / 100;
                  OpenBalVAT := OpenBalVAT + ((ThisAmount * VATRate) / 100);
                  OpenBalAmount := OpenBalAmount + ((ThisAmount * (100 + VATRate)) / 100);
               end else begin
                  Item[1,lcHEC - 1].Value := 0;
                  Item[1,lcHEC].Value := ThisAmount;
                  OpenBalAmount := OpenBalAmount + ThisAmount;
               end;
            end else begin
               Item[1,lcHEC].Value := ThisAmount;
               OpenBalFees := OpenBalFees + ThisAmount;
            end;
         end;

         Query1.Next;

         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

//--- Closing Balance Line

   with xlsSheet.RCRange[PageRows + ThisRows,1,PageRows + ThisRows,lcHEC] do begin
      Borders[xlEdgeBottom].LineStyle := xlContinuous;
      Borders[xlEdgeBottom].Weight := xlThin;

      if (ThisItem = 'Total for this invoice') then
         Item[1,1].Value := EDate
      else
         Item[1,1].Formula := LastDate;

      Item[1,2].Value := ThisItem;

      if (VATRate = 0) then begin
         Item[1,lcHEC].Formula := OpenBalFees;

         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;

         SaveAmount := OpenBalFees * -1;
      end else begin
         Item[1,lcHEC - 2].Value := OpenBalFees;
         Item[1,lcHEC - 1].Value := OpenBalVAT;
         Item[1,lcHEC].Formula := OpenBalAmount;

         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
         Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
         Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;

         SaveAmount := OpenBalAmount * -1;
      end;

      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;
end;

//---------------------------------------------------------------------------
// Generate Statements
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Statement();
var
   PageNum, idx1, idx2, FileCount, SheetCount   : integer;
   ReturnValue, NumInvoices                     : integer;
   Balance                                      : double;
   Answer, DoLine                               : boolean;
   ThisMsg, ThisFile, ThisStr, TempFile, Delim  : string;
   ThisDate                                     : TDateTime;
   xlsBook                                      : IXLSWorkbook;
   xlsSheet1                                    : IXLSWorksheet;
   xlsSheet2                                    : IXLSWorksheet;

begin

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Database error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get the Heading and Layout information

   GetLayout('Statement');

//--- Read the Billing records from the datastore

   for idx1 := 0 to NumFiles - 1 do begin
      PageNum := 1;
      ThisDate := StrToDate(EDate);
      ThisFile := FormatDateTime('yyyyMMdd',ThisDate) + ' - Statement (' + FileArray[idx1] + ').xls';
      txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;
      txtDocument.Refresh;

//--- We need to know upfront whether there are invoices for this File as no
//--- Age Analysis information will be inserted if there are no invoices

      GetInvoices(ord(DT_STATEMENT),FileArray[idx1]);
      NumInvoices := Query1.RecordCount;

//--- Proceed to get the Billing information for this File

      ThisStr := ' AND B_AccountType = 0';

      if ((GetBilling(FileArray[idx1],ThisStr,'Statement')) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      txtError.Text := 'Processing: ' + FileArray[idx1];
      txtError.Refresh;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

//--- Open the Excel workbook template

      xlsBook := TXLSWorkbook.Create;
      xlsBook.Open(Template_S);
      xlsSheet1      := xlsBook.ActiveSheet;
      xlsSheet1.Name := 'Statement Summary (' + FileArray[idx1] + ')';
      xlsSheet2      := xlsBook.WorkSheets.Add;
      xlsSheet2.Name := 'Statement Detail (' + FileArray[idx1] + ')';

//--- Clear everything but keep the Header information if lcShowHeader is set

      if (lcShowHeader = true) then
         xlsSheet1.RCRange[lcHER + 1, 1, 999, lcHEC].Clear
      else
         xlsSheet1.RCRange[1, 1, 999, lcSMaxCols].Clear;

//--- Insert the Page 1 Heading

      Generate_DocHeading(xlsSheet1,PageNum,idx1,lcHER + 1,lcHEC,'Statement');

//--- Insert the Customer Information

      if (lcShowAddress = True) then begin

         GetAddress(FileArray[idx1]);

         with xlsSheet1.RCRange[lcASR,1,lcAER,lcHEC] do begin
            Item[1,lcASC].Value := Customer;
            Item[2,lcASC].Value := Address1;
            Item[3,lcASC].Value := Address2;
            Item[4,lcASC].Value := Address3;
            Item[5,lcASC].Value := Address4;
            Item[6,lcASC].Value := Address5;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
      end;

//--- Insert the instruction information

      if (lcShowInstruct = True) then begin
         with xlsSheet1.RCRange[lcISR,1,lcISR + 1,lcHEC] do begin
            Item[1,1].Value := 'Client:';
            Item[2,1].Value := 'Instruction:';
            Item[1,lcISCD].Value := Customer;
            Item[2,lcISCD].Value := Descrip;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
      end;

//--- Insert the Age Analysis information

      if (lcShowAge = True) then begin
         if (NumInvoices = 0) then begin
            with xlsSheet1.RCRange[lcAASR,lcAASC,lcAASR,lcAASC + 1] do begin
               Merge(true);
               Item[1,1].Value := 'No Invoices';
               HorizontalAlignment := xlHAlignCenter;
               VerticalAlignment := xlVAlignCenter;
               Borders[xlAround].LineStyle := xlContinuous;
               Borders[xlAround].Weight := xlThin;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
               Font.Color := ColDHT;
               Interior.Color := ColDHF;
            end;

//--- Prevent further Age Analysis processing

            lcShowAge := False;

         end else begin
            with xlsSheet1.RCRange[lcAASR,lcAASC,lcAASR,lcAASC + 1] do begin
               Merge(true);

               if (IncludeTrust = true) then begin
                  if (ExcludeReserve = true) then
                     Item[1,1].Value := 'Ageing (Incl Trust, Excl Res)'
                  else
                     Item[1,1].Value := 'Ageing (Incl Trust)';
               end else
                  Item[1,1].Value := 'Ageing (Excl Trust)';

               HorizontalAlignment := xlHAlignCenter;
               VerticalAlignment := xlVAlignCenter;
               Borders[xlAround].LineStyle := xlContinuous;
               Borders[xlAround].Weight := xlThin;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
               Font.Color := ColDHT;
               Interior.Color := ColDHF;
            end;

            with xlsSheet1.RCRange[lcAASR + 1,lcAASC,lcAASR + 4,lcAASC + 1] do begin
               Item[1,1].Value := 'Current';
               Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,1].Borders[xlEdgeBottom].Weight := xlThin;
               Item[2,1].Value := '30 Days';
               Item[2,1].Borders[xlEdgeRight].Weight := xlThin;
               Item[2,1].Borders[xlEdgeBottom].Weight := xlThin;
               Item[3,1].Value := '60 Days';
               Item[3,1].Borders[xlEdgeRight].Weight := xlThin;
               Item[3,1].Borders[xlEdgeBottom].Weight := xlThin;
               Item[4,1].Value := '90 Days +';
               Item[4,1].Borders[xlEdgeRight].Weight := xlThin;
               Item[4,1].Borders[xlEdgeBottom].Weight := xlThin;
               Borders[xlAround].LineStyle := xlContinuous;
               Borders[xlAround].Weight := xlThin;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;
      end;

//--- Write the Data Heading Information and the Statement data

      Generate_DataHeading(xlsSheet1,lcPSR,lcHEC,'Statement');
      Generate_Statement_Summary(xlsSheet1, lcPSR, lcPRows - 1);
      Answer := Generate_Statement_Detail(xlsSheet2,FileArray[idx1]);

//--- If the result of the Statement detail is False then there is nothing in
//--- Sheet 2 and it is better to delete it

      if (Answer = False) then
         xlsSheet2.Delete;

//--- Insert the detail of the Age Analysis information

      if (lcShowAge = True) then begin
         with xlsSheet1.RCRange[lcAASR + 1,lcAASC + 1,lcAASR + 4,lcAASC + 1] do begin
            Item[1,1].Value := RoundD(AgeCurrent,2);
            Item[1,1].Borders[xlEdgeBottom].Weight := xlThin;
            Item[2,1].Value := RoundD(Age30Days,2);
            Item[2,1].Borders[xlEdgeBottom].Weight := xlThin;
            Item[3,1].Value := RoundD(Age60Days,2);
            Item[3,1].Borders[xlEdgeBottom].Weight := xlThin;
            Item[4,1].Value := RoundD(Age90Days,2);
            NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
      end;

//--- Write the standard copyright notice on the Statement Page

      with xlsSheet1.RCRange[lcPSR + lcPRows + 2,1,lcPSR + lcPRows + 2,1] do begin
         Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 8;
      end;

//--- Remove the Gridlines which are added by default and set the Page
//--- orientation on the Statement page

      xlsSheet1.PageSetup.Orientation    := xlLandscape;
      xlsSheet1.PageSetup.FitToPagesWide := 1;
      xlsSheet1.PageSetup.FitToPagesTall := 1;
      xlsSheet1.PageSetup.PaperSize      := xlPaperA4;
      xlsSheet1.DisplayGridLines         := false;
      xlsSheet1.PageSetup.CenterFooter    := 'Page &P of &N';

//--- Write the Excel file to disk

      xlsBook.SaveAs(FileName + ThisFile);
      xlsBook.Close;

      LogMsg('  File ''' + FileArray[idx1] + ''' successfully processed...',True);
      LogMsg(' ',True);

      DoLine := False;

//--- Print the generated document on the Default Printer if requested

      if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
         FileCount := 0;

         xlsBook.Open(FileName + ThisFile);
         SheetCount := xlsBook.Sheets.Count;

         for idx2 := 1 to SheetCount do begin
            xlsBook.Sheets[idx2].Activate;
            xlsBook.Save;

            if (PrintDocument(ThisFile, FileName) = True) then
               inc(FileCount);
         end;

         xlsBook.Close;

         if (FileCount = Sheetcount) then begin
            LogMsg('  Document submitted for printing...',True);
            DoLine := True;
         end else begin
            LogMsg('  Printing of document failed...',True);
            DoLine := True;
         end;
      end;

//--- Create a PDF file if requested.

      if ((CreatePDF = true) and (PDFPrinter <> 'Not Found')) then begin
         PDFExists  := PDFDocument(ThisFile, FileName);

         if (PDFExists = True) then
            LogMsg('  PDF file creation was successfull...',True)
         else
            LogMsg('  PDF file creation failed...',True);

         DoLine := True;
      end;

//--- Send the Excel file via email if requested

      if (SendByEmail = '1') then begin
         if (GroupAttach = true) then
            AttachList.Add(FileName + ThisFile)
         else begin
            if (SendEmail(FileArray[idx1],FileName + ThisFile,ord(PT_NORMAL)) = true) then
               LogMsg('  Request to send generated Statement by Email submitted...',True)
            else
               LogMsg('  Request to send generated Statement by Email not submitted...',True);

            DoLine := True;
         end;
      end;

//--- Now open the Statement if requested

      if (AutoOpen = True) then begin
         ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);

         ThisMsg := '  Request to open Statement for ''' + FileArray[idx1] + ''' submitted...';

         LogMSg(ThisMsg,True);
         DoLine := True;
      end;

   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   Close_Connection;
end;

//---------------------------------------------------------------------------
// Generate the Summary page for the Statement
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Statement_Summary(xlsSheet : IXLSWorksheet; PageRows: integer; ThisRows: integer);
var
   idx2                           : integer;
   BusinessBal, TrustBal, Balance : double;

begin

//--- Page Detail Lines - Business Account Section

   idx2 := 1;

   with xlsSheet.RCRange[PageRows + idx2,lcSCBD,PageRows + ThisRows,lcSCBD] do begin
      NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCB,PageRows + idx2,lcSCBD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := false;
      Item[1,1].Value := 'Fees';
      Item[1,lcSCBD - lcSCB + 1].Value := RoundD(StatementFees,2);
      Item[1,lcSCBD - lcSCB + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCB,PageRows + idx2,lcSCBD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := false;
      Item[1,1].Value := 'Disbursements';
      Item[1,lcSCBD - lcSCB + 1].Value := RoundD(StatementDisburse,2);
      Item[1,lcSCBD - lcSCB + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCB,PageRows + idx2,lcSCBD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := false;
      Item[1,1].Value := 'Expenses';
      Item[1,lcSCBD - lcSCB + 1].Value := RoundD(StatementExpenses,2);
      Item[1,lcSCBD - lcSCB + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCB,PageRows + idx2,lcSCBD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := false;
      Item[1,1].Value := 'Deposits to Business Account';
      Item[1,lcSCBD - lcSCB + 1].Value := RoundD(StatementDeposits,2);
      Item[1,lcSCBD - lcSCB + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCB,PageRows + idx2,lcSCBD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := false;
      Item[1,1].Value := 'Payments received from Trust';
      Item[1,lcSCBD - lcSCB + 1].Value := RoundD(StatementTrustPay * -1,2);
      Item[1,lcSCBD - lcSCB + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCB,PageRows + idx2,lcSCBD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := false;
      Item[1,1].Value := 'Payments made from Business';
      Item[1,lcSCBD - lcSCB + 1].Value := RoundD(StatementBusPay,2);
      Item[1,lcSCBD - lcSCB + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   BusinessBal := (StatementFees + StatementDisburse + StatementExpenses) + (StatementDeposits + (StatementTrustPay * -1) + StatementBusPay);

   with xlsSheet.RCRange[PageRows + idx2,lcSCB,PageRows + idx2,lcSCBD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := true;
      Item[1,1].Value := 'Business Account Balance';
      Item[1,lcSCBD - lcSCB + 1].Value := RoundD(BusinessBal,2);
      Item[1,lcSCBD - lcSCB + 1].Borders[xlAround].Weight := xlThin;
   end;

//--- Page Detail Lines - Trust Account Section

   idx2 := 1;

   with xlsSheet.RCRange[PageRows + idx2,lcSCTD,PageRows + ThisRows,lcSCTD] do begin
      NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCT,PageRows + idx2,lcSCTD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := false;
      Item[1,1].Value := 'Unreserved Deposits to Trust';
      Item[1,lcSCTD - lcSCT + 1].Value := RoundD(StatementTrustDep - StatementReserve,2);
      Item[1,lcSCTD - lcSCT + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCT,PageRows + idx2,lcSCTD] do begin
      Borders[xlAround].Weight := xlThin;
      Item[1,1].Value := 'Reserved Deposits to Trust';
      Item[1,lcSCTD - lcSCT + 1].Value := RoundD(StatementReserve,2);
      Item[1,lcSCTD - lcSCT + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCT,PageRows + idx2,lcSCTD] do begin
      Borders[xlAround].Weight := xlThin;
      Item[1,1].Value := 'Section 78(2A) Interest';
      Item[1,lcSCTD - lcSCT + 1].Value := RoundD(StatementTrustInt,2);
      Item[1,lcSCTD - lcSCT + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCT,PageRows + idx2,lcSCTD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := false;
      Item[1,1].Value := 'Payments made from Trust';
      Item[1,lcSCTD - lcSCT + 1].Value := RoundD(StatementTrustPay,2);
      Item[1,lcSCTD - lcSCT + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCT,PageRows + idx2,lcSCTD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := false;
      Item[1,1].Value := 'Disbursements from Trust';
      Item[1,lcSCTD - lcSCT + 1].Value := RoundD(StatementTrustDis,2);
      Item[1,lcSCTD - lcSCT + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   with xlsSheet.RCRange[PageRows + idx2,lcSCT,PageRows + idx2,lcSCTD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := false;
      Item[1,1].Value := '';
      Item[1,lcSCTD - lcSCT + 1].Value := '';
      Item[1,lcSCTD - lcSCT + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

   TrustBal := ((StatementTrustPay + StatementTrustDis) + (StatementTrustDep + StatementTrustInt));

   with xlsSheet.RCRange[PageRows + idx2,lcSCT,PageRows + idx2,lcSCTD] do begin
      Borders[xlAround].Weight := xlThin;
      Font.Bold := true;
      Item[1,1].Value := 'Trust Account Balance';
      Item[1,lcSCTD - lcSCT + 1].Value := RoundD(TrustBal,2);
      Item[1,lcSCTD - lcSCT + 1].Borders[xlAround].Weight := xlThin;
      inc(idx2);
   end;

//--- Closing Balance Line

   with xlsSheet.RCRange[PageRows + idx2,1,PageRows + idx2,lcHEC] do begin
      Borders[xlAround].Weight := xlThin;

      Interior.Color := ColDHF;
      Font.Color     := ColDHT;
      Font.Bold      := true;
      Font.Name      := 'Arial';
      Font.Size      := 10;
   end;

   with xlsSheet.RCRange[PageRows + idx2,1,PageRows + idx2,lcHEC] do begin
      Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,lcHEC].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';

      if (IncludeTrust = true) then begin
         if (ExcludeReserve = true) then
            Balance := RoundD(BusinessBal + TrustBal - StatementReserve,2)
         else
            Balance := RoundD(BusinessBal + TrustBal,2);
      end else
         Balance := RoundD(BusinessBal,2);


      Item[1,lcHEC].Formula := '=ABS(' + FloatToStr(Balance) + ')';
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;

      if (Balance >= 0) then begin
         if (ExcludeReserve = true) then
            Item[1,1].Value := 'Amount owed to Client (Excluding Reserved Deposits of R ' + FloatToStrF(StatementReserve,ffNumber,10,2) + ')'
         else
            Item[1,1].Value := 'Amount owed to Client';
      end else begin
         if (ExcludeReserve = true) then
            Item[1,1].Value := 'Amount owed by Client (Excluding Reserved Deposits of R ' + FloatToStrF(StatementReserve,ffNumber,10,2) + ')'
         else
            Item[1,1].Value := 'Amount owed by Client';
      end;
   end;
end;

//---------------------------------------------------------------------------
// Generate the Detail for the Statement
//---------------------------------------------------------------------------
function TFldExcel.Generate_Statement_Detail(xlsSheet : IXLSWorksheet; ThisFile : string) : boolean;
var
   idx1, idx2, idx3, idx4, row, PageRow, Pages, RowsPerPage          : integer;
   ThisClass , ThisYear, ThisMonth                                   : integer;
   ThisS864, ThisTrust, ThisBus, TotS864, TotTrust, TotBus, TotInv   : double;
   BusCurr, Bus30, Bus60, Bus90, Avail, TotReserve, ThisReserve      : double;
   PageBreak, FirstPage                                              : boolean;
   PCurrS, P30S, P60S, P90S, PCurrE, P30E, P60E, P90E                : string;

   ThisRecList  : array of LPMS_Statement;
   LastDay      : array[1..12] of integer;

begin

   idx1       := 0;
   TotS864    := 0;
   TotTrust   := 0;
   TotBus     := 0;
   TotInv     := 0;
   TotReserve := 0;

//--- Only process if there are records

   if (Query1.RecordCount > 0) then begin

      Query1.First;

//--- Step through each record

      while Query1.Eof = False do begin

         ThisClass := Query1.FieldByName('B_Class').AsInteger;

         if (ThisClass in [3..4,6..14,18..19,24]) then begin
            case ThisClass of
               3: begin                     // Payment Received
                  ThisS864     := 0;
                  ThisTrust    := 0;
                  ThisBus      := Query1.FieldByName('B_Amount').AsFloat;
               end;

               4: begin                     // Business To Trust
                  ThisS864     := 0;
                  ThisTrust    := Query1.FieldByName('B_Amount').AsFloat;
                  ThisBus      := ThisTrust * -1;
               end;

               6: begin                     // Business Deposit
                  ThisS864     := 0;
                  ThisTrust    := 0;
                  ThisBus      := Query1.FieldByName('B_Amount').AsFloat;
               end;

               7: begin                     // Trust Deposit
                  ThisS864     := 0;
                  ThisTrust    := Query1.FieldByName('B_Amount').AsFloat;
                  ThisBus      := 0;
                  ThisReserve  := Query1.FieldByName('B_ReserveAmt').AsFloat;
               end;

               8: begin                     // Trust to Business (Fees/Disb/Exp)
                  ThisS864     := 0;
                  ThisTrust    := Query1.FieldByName('B_Amount').AsFloat * -1;
                  ThisBus      := Query1.FieldByName('B_Amount').AsFloat;
               end;

               9: begin                     // Trust Transfer (Disbursement)
                  ThisS864     := 0;
                  ThisTrust    := Query1.FieldByName('B_Amount').AsFloat * -1;
                  ThisBus      := 0;
               end;

               10: begin                     // Trust Transfer (Client)
                  ThisS864     := 0;
                  ThisTrust    := Query1.FieldByName('B_Amount').AsFloat * -1;
                  ThisBus      := 0;
               end;

               11: begin                     // Trust Transfer (Trust)
                  ThisS864     := 0;
                  ThisTrust    := Query1.FieldByName('B_Amount').AsFloat * -1;
                  ThisBus      := 0;
               end;

               12: begin                    // S86(4) Investment
                  ThisS864     := Query1.FieldByName('B_Amount').AsFloat;
                  ThisTrust    := Query1.FieldByName('B_Amount').AsFloat * -1;
                  ThisBus      := 0;
               end;

               13: begin                     // S86(4) Withdrawal
                  ThisS864     := Query1.FieldByName('B_Amount').AsFloat * -1;
                  ThisTrust    := Query1.FieldByName('B_Amount').AsFloat;
                  ThisBus      := 0;
               end;

               14: begin                     // S86(4) Interest
                  ThisS864     := Query1.FieldByName('B_Amount').AsFloat;
                  ThisTrust    := 0;
                  ThisBus      := 0;
               end;

               18: begin                     // Business Debit
                  ThisS864     := 0;
                  ThisTrust    := 0;
                  ThisBus      := Query1.FieldByName('B_Amount').AsFloat * -1;
               end;

               19: begin                     // Trust Debit
                  ThisS864     := 0;
                  ThisTrust    := Query1.FieldByName('B_Amount').AsFloat * -1;
                  ThisBus      := 0;
               end;

               24: begin                     // Trust to Business (Other)
                  ThisS864     := 0;
                  ThisTrust    := Query1.FieldByName('B_Amount').AsFloat * -1;
                  ThisBus      := Query1.FieldByName('B_Amount').AsFloat;
               end;
            end;

            SetLength(ThisRecList,idx1 + 1);

            ThisRecList[idx1].DateTime    := Query1.FieldByName('B_Date').AsString + '00:00:01';
            ThisRecList[idx1].Date        := Query1.FieldByName('B_Date').AsString;
            ThisRecList[idx1].Description := ReplaceQuote(Query1.FieldByName('B_Description').AsString);
            ThisRecList[idx1].S864        := ThisS864;
            ThisRecList[idx1].Trust       := ThisTrust;
            ThisRecList[idx1].Business    := ThisBus;

            TotS864    := TotS864 + ThisS864;
            TotTrust   := TotTrust + ThisTrust;
            TotBus     := TotBus + ThisBus;
            TotReserve := TotReserve + ThisReserve;

            ThisReserve := 0;
            inc(idx1);
         end;

         Query1.Next;
      end;

//--- Add all the invoices for this File to the Record List

      GetInvoices(ord(DT_STATEMENT),ThisFile);

      SetLength(ThisRecList,Query1.RecordCount + idx1);
      Query1.First;

      for idx2 := 0 to Query1.RecordCount - 1 do begin
         ThisRecList[idx1].DateTime    := Query1.FieldByName('Inv_EDate').AsString + '00:00:99';
         ThisRecList[idx1].Date        := Query1.FieldByName('Inv_EDate').AsString;
         ThisRecList[idx1].Description := 'Issue invoice ''' + Query1.FieldByName('Inv_Invoice').AsString + '''';
         ThisRecList[idx1].S864         := 0;
         ThisRecList[idx1].Trust       := 0;
         ThisRecList[idx1].Business    := Query1.FieldByName('Inv_Amount').AsFloat * -1;

         TotInv := TotInv + ThisRecList[idx1].Business;

         Query1.Next;
         inc(idx1);
      end;

//--- Sort the Record List in DateTime order

      BubbleSort(ThisRecList);

//--- Do the Age Analysis. Start by initialising key variables

      BusCurr := 0; Bus30 := 0; Bus60 := 0; Bus90 := 0;
      AgeCurrent := 0; Age30Days := 0; Age60Days := 0; Age90Days := 0;

//--- Set the Dates for the 4 periods. Begin by initialising the LastDay Array

      LastDay[ 1] := 31; LastDay[ 2] := 28; LastDay[ 3] := 31; LastDay[ 4] := 30;
      LastDay[ 5] := 31; LastDay[ 6] := 30; LastDay[ 7] := 31; LastDay[ 8] := 31;
      LastDay[ 9] := 30; LastDay[10] := 31; LastDay[11] := 30; LastDay[12] := 31;

//--- Adjust the number of days for February for Leap Years

      ShortDateFormat := 'yyyy/MM/dd';
      DateSeparator   := '/';

      ThisYear := StrToInt(FormatDateTime('yyyy',StrToDate(EDate)));

      if (ThisYear mod 4 = 0) then
         LastDay[2] := 29;

//--- Set the First and Last Day for the periods

      ThisMonth := StrToInt(FormatDateTime('MM',StrToDate(EDate)));
      PCurrE    := LeftStr(EDate,8) + IntToStr(LastDay[ThisMonth]);
      PCurrS    := LeftStr(PCurrE,8) + '01';

      P30E      := DateToStr(IncMonth(StrToDate(PCurrE), -1));
      ThisMonth := StrToInt(FormatDateTime('MM',StrToDate(P30E)));
      P30E      := LeftStr(P30E,8) + IntToStr(LastDay[ThisMonth]);
      P30S      := LeftStr(P30E,8) + '01';

      P60E      := DateToStr(IncMonth(StrToDate(PCurrE), -2));
      ThisMonth := StrToInt(FormatDateTime('MM',StrToDate(P60E)));
      P60E      := LeftStr(P60E,8) + IntToStr(LastDay[ThisMonth]);
      P60S      := LeftStr(P60E,8) + '01';

      P90E      := DateToStr(IncMonth(StrToDate(PCurrE), -3));
      ThisMonth := StrToInt(FormatDateTime('MM',StrToDate(P90E)));
      P90E      := LeftStr(P90E,8) + IntToStr(LastDay[ThisMonth]);
      P90S      := '1962/03/31';

//--- Now determine the business balance in each of the periods

      for idx3 := 0 to Length(ThisRecList) - 1 do begin
         if ((ThisRecList[idx3].Date >= PCurrS) and (ThisRecList[idx3].Date <= PCurrE)) then
            BusCurr := BusCurr + ThisRecList[idx3].Business
         else if ((ThisRecList[idx3].Date >= P30S) and (ThisRecList[idx3].Date <= P30E)) then
            Bus30 := Bus30 + ThisRecList[idx3].Business
         else if ((ThisRecList[idx3].Date >= P60S) and (ThisRecList[idx3].Date <= P60E)) then
            Bus60 := Bus60 + ThisRecList[idx3].Business
         else if ((ThisRecList[idx3].Date >= P90S) and (ThisRecList[idx3].Date <= P90E)) then
            Bus90 := Bus90 + ThisRecList[idx3].Business;
      end;

//--- Determine the available amount that can be offset against the period
//--- balances - Available on Trust (including S86(4) and with deference to
//--- whether 1) Trust amounts must be included and 2) whether reserved trust
//--- amounts should be included) plus all positive balances for each period

      if (IncludeTrust = True) then begin
         if (ExcludeReserve = True) then
            Avail := (TotTrust + TotS864) - TotReserve
         else
            Avail := TotTrust + TotS864;
      end;

      if (BusCurr > 0) then Avail := Avail + BusCurr;
      if (Bus30 > 0)   then Avail := Avail + Bus30;
      if (Bus60 > 0)   then Avail := Avail + Bus60;
      if (Bus90 > 0)   then Avail := Avail + Bus90;

//--- Step through each period starting with 90+. If the period balance is
//--- negative then add the available amount otherwise ignore. If after the add
//--- the period balance is still negative then set available amount to 0. If
//--- the period amuount is 0 or positive then this becomes the new available
//--- amount and the period balance is set to 0. Repeat for each period.

//--- 90 Days +

      if (Bus90 < 0) then begin
         if (Avail > 0) then begin
            Age90Days := Bus90 + Avail;

            if (Age90Days > 0) then
               Avail := Age90Days
            else
               Avail := 0;
         end else begin
            Age90Days := Bus90;
         end;
      end;

//--- 60 Days

      if (Bus60 < 0) then begin
         if (Avail > 0) then begin
            Age60Days := Bus60 + Avail;

            if (Age60Days > 0) then
               Avail := Age60Days
            else
               Avail := 0;
         end else begin
            Age60Days := Bus60;
         end;
      end;

//--- 30 Days

      if (Bus30 < 0) then begin
         if (Avail > 0) then begin
            Age30Days := Bus30 + Avail;

            if (Age30Days > 0) then
               Avail := Age30Days
            else
               Avail := 0;
         end else begin
            Age30Days := Bus30;
         end;
      end;

//--- Current

      if (BusCurr < 0) then begin
         if (Avail > 0) then begin
            AgeCurrent := BusCurr + Avail;
         end else begin
            AgeCurrent := BusCurr;
         end;
      end;

//--- Finally go through the period balances and set any balance > 0 to 0

      if (Age90Days > 0)  then Age90Days  := 0;
      if (Age60Days > 0)  then Age60Days  := 0;
      if (Age30Days > 0)  then Age30Days  := 0;
      if (AgeCurrent > 0) then AgeCurrent := 0;

//--- If e get here and there is nothin in the Record List then there is
//--- nothing to display - Return False in order to delete the sheet

      if (Length(ThisRecList) = 0) then begin
         Result := False;
         Exit;
      end;

//--- Transfer the content of the Sorted Array to the Excel Spreadsheet

      Pages     := 0;
      FirstPage := true;
      PageBreak := true;

      for idx4 := 0 to Length(ThisRecList) - 1 do begin

//--- Perform a Page break if necessary

         if (PageBreak = True) then
            DoPageBreak(ord(PB_STATEMENT),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,ThisFile);

//--- Insert the detail

         with xlsSheet.RCRange[row,1,row,5] do begin
            Item[1,1].Value := ThisRecList[idx4].Date;
            Item[1,1].VerticalAlignment := xlVAlignTop;
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := ThisRecList[idx4].Description;
            Item[1,2].VerticalAlignment := xlVAlignTop;
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := ThisRecList[idx4].S864;
            Item[1,3].VerticalAlignment := xlVAlignTop;
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Value := ThisRecList[idx4].Trust;
            Item[1,4].VerticalAlignment := xlVAlignTop;
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].Value := ThisRecList[idx4].Business;
            Item[1,5].VerticalAlignment := xlVAlignTop;
            Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
            Borders[xlAround].Weight := xlThin;
         end;

         with xlsSheet.RCRange[row,3,row,5] do begin
            NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;

         inc(row);
         inc(PageRow);

//--- If we've reached the maximum rows per page then it is PageBreak time. We
//--- do the Pagebreak here so that the Total line and Copyright notice will
//--- be handled correctly

         if (PageRow >= RowsPerPage) then begin
            PageBreak := True;
            DoPageBreak(ord(PB_STATEMENT),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,ThisFile);
         end;

      end;

//--- Insert the Total Line

      with xlsSheet.RCRange[row,1,row,5] do begin
         Item[1,1].Value := EDate;
         Item[1,1].VerticalAlignment := xlVAlignTop;
         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,2].Value := 'Totals';
         Item[1,2].VerticalAlignment := xlVAlignTop;
         Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,3].Value := TotS864;
         Item[1,3].VerticalAlignment := xlVAlignTop;
         Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,4].Value := TotTrust;
         Item[1,4].VerticalAlignment := xlVAlignTop;
         Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,5].Value := TotBus + TotInv;
         Item[1,5].VerticalAlignment := xlVAlignTop;
         Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
         Font.Bold := True;
         Font.Name := 'Arial';
         Font.Size := 10;
         Borders[xlAround].Weight := xlThin;
      end;

      with xlsSheet.RCRange[row,3,row,5] do begin
         NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
         Font.Bold := True;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;

//--- Write the standard copyright notice

      row := (Pages * lcGRows);
      with xlsSheet.RCRange[row,1,row,1] do begin
         Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 8;
      end;

//--- Remove the Gridlines which are added by default and set the Page
//--- orientation on the Statement detail page

      xlsSheet.PageSetup.Orientation    := xlLandscape;
      xlsSheet.PageSetup.FitToPagesWide := 1;
      xlsSheet.PageSetup.FitToPagesTall := Pages;
      xlsSheet.PageSetup.PaperSize      := xlPaperA4;
      xlsSheet.DisplayGridLines         := false;
      xlsSheet.PageSetup.CenterFooter    := 'Page &P of &N';
   end;
   Result := True;
end;

//---------------------------------------------------------------------------
// Procedure to manage whether a Trust Reconcilliation or a Simple Trust
// Report for individual Files is generated
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Trust_Summary();
begin

   if (FirstFile = '**ALL**') then
      DoTrust_All()
   else
      DoTrust_Simple();
end;

//---------------------------------------------------------------------------
// Generate Trust Details for a Trust Reconcilliation report
//---------------------------------------------------------------------------
procedure TFldExcel.DoTrust_All();
var
   Sema, FilesFound, Pages1, Pages2,Pages4, SaveRow2, PageRow1    : integer;
   PageRow2, PageRow4, row1, row2, row3, row4, idx1, idx2, idx3   : integer;
   RowsPerPage1, S864Recs, RowsPerPage2, RowsPerPage4, TempYear   : integer;
   ThisAmtTrust, ThisAmt864, ThisTrust, This864, TotalTrust       : double;
   Total864, This863Inv, This863Drw, This863Int, This863IntDrw    : double;
   Total863Int                                                    : double;
   PageBreak1, PageBreak2, PageBreak4, FirstPage1, FirstPage2     : boolean;
   FirstPage4, RecsFound, DoLine                                  : boolean;
   ThisAB1F, ThisAB1T, ThisAB2F, ThisAB2T, ThisFill, ThisText     : TColor;
   S1, ThisStr, ThisFile                                          : string;
   xlsBook                                                        : IXLSWorkbook;
   xlsSheet1, xlsSheet2, xlsSheet3, xlsSheet4                     : IXLSWorksheet;
   ThisDate                                                       : TDateTime;

   CtrlTotals : array[1..6] of double;

const
   CtrlLines : array[1..6] of string = (
               'Section 86(3) Balance',
               'Section 86(3) Interest',
               'Section 86(3) Total',
               'Section 78(2A) Total',
               'Trust Account Balance',
               'Trust Account Total');
begin


   ThisCount := 1;
   ThisMax   := IntToStr(NumFiles);

   prbProgress.Max := NumFiles;
   LogMsg(ord(PT_PROLOG),True,True,'Trust Account Reconciliation');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Open the Excel workbook template

   FilesFound := 0;

   xlsBook        := TXLSWorkbook.Create;
   xlsSheet1      := xlsBook.WorkSheets.Add;
   xlsSheet2      := xlsBook.WorkSheets.Add;
   xlsSheet3      := xlsBook.WorkSheets.Add;
   xlsSheet4      := xlsBook.WorkSheets.Add;
   xlsSheet1.Name := 'Trust Detail';
   xlsSheet2.Name := 'Trust Summary';
   xlsSheet3.Name := 'Trust Control Totals';
   xlsSheet4.Name := 'Section 86(3) Interest';

//--- Determine the full date for the Financial Year

   FYYear := FormatDateTime('yyyy',Now);
   if (FormatDateTime('MM',StrToDate(EDate)) < FYMonth)  then
   begin
      TempYear := StrToInt(FYYear);
      TempYear := TempYear - 1;
      FYYear  := IntToStr(TempYear);
   end;

//--- Construct the output File name

   ThisDate := StrToDate(EDate);
   ThisFile := FormatDateTime('yyyyMMdd',ThisDate) + ' - Trust Account Reconciliation (' + CpyName + ').xls';
   txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;

//--- Set up to Insert the Header (1st and 2nd line) and the Heading for each
//--- worksheet

   Pages1 := 0; Pages2 := 0; Pages4 := 0;
   FirstPage1 := True; FirstPage2 := True; FirstPage4 := True;
   PageBreak1 := True; PageBreak2 := True; PageBreak4 := True;

//--- Set up to use alternate color blocks - if only one File then we use a
//--- default of White background with Black text

   if (NumFiles = 1) then begin
      ThisAB1F := clWhite;
      ThisAB1T := clBlack;
   end else begin
      ThisAB1F := ColAB1F;
      ThisAB1T := ColAB1T;
      ThisAB2F := ColAB2F;
      ThisAB2T := ColAB2T;
   end;

   Sema := 0;

//--- Read the Trust records from the datastore

   TotalTrust  := 0;
   Total864  := 0;

   for idx1 := 0 to NumFiles - 1 do begin

      ThisStr := ' AND ((B_Class = 4) OR (B_Class >= 7 AND B_Class <= 14) OR (B_Class = 19) OR (B_Class = 24)) ';

      if ((GetBilling(FileArray[idx1],ThisStr,'Trust Recon')) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      txtError.Text := 'Processing: ' + FileArray[idx1];
      txtError.Refresh;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

//--- Only process files that have and opening balance and or records

      if (Query1.RecordCount > 0) then
         RecsFound := True
      else
         RecsFound := False;

      if ((Query1.RecordCount > 0) or (RoundD(OpenBalTrust,2) <> 0) or (RoundD(OpenBal864,2) <> 0)) then begin

         inc(FilesFound);

//--- We only do Sheet 1 if there are records for the period

         if (RecsFound = True) then begin

//--- Perform a Page break on the Detail Sheet if necessary

            if (PageBreak1 = True) then
               DoPageBreak(ord(PB_TRUSTDETAIL),xlsSheet1,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1,'');

         end;

//--- Store the opening balances for later use

         TotTotals[1] := This_Bal.Trust_Deposit + This_Bal.Business_To_Trust;
         TotTotals[2] := This_Bal.Trust_Interest_S86_4;
         TotTotals[3] := This_Bal.Trust_Transfer_Business_Fees + This_Bal.Trust_Transfer_Business_Other;
         TotTotals[4] := This_Bal.Trust_Transfer_Disbursements + This_Bal.Trust_Transfer_Client;
         TotTotals[5] := This_Bal.Trust_Transfer_Trust;
         TotTotals[6] := This_Bal.Trust_Debit;

         if (RecsFound = True) then begin

//--- Set this block's colour

            if (Sema = 0) then begin
               Sema := 1;
               ThisFill := ThisAB1F;
               ThisText := ThisAB1T;
            end else begin
               Sema := 0;
               ThisFill := ThisAB2F;
               ThisText := ThisAB2T;
            end;

//--- Write the Opening Balance

            with xlsSheet1.RCRange[row1,1,row1,5] do begin
               Item[1,1].Value := FileArray[idx1];
               Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,2].Value := SDate;
               Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,3].Value := 'Opening Balance (R ' + FloatToStrF((OpenBal864 + OpenBalTrust),ffNumber,10,2) + ')';
               Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,4].Value := RoundD(OpenBal864,2);
               Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,4].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
               Item[1,5].Value := RoundD(OpenBalTrust,2);
               Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,5].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ThisFill;
               Font.Color := ThisText;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;

         ThisTrust := OpenBalTrust;
         This864   := OpenBal864;

         if (RecsFound = True) then begin

            inc(row1);
            inc(PageRow1);

//--- If we've reached the maximum rows per page then it is PageBreak time

            if (PageRow1 >= RowsPerPage1) then begin
               PageBreak1 := True;
               DoPageBreak(ord(PB_TRUSTDETAIL),xlsSheet1,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1,'');
            end;
         end;

//--- Step through each record if there are any

         if (Query1.RecordCount > 0) then begin

            for idx2 := 0 to Query1.RecordCount - 1 do begin

//--- Determine whether the Trust amount is positive or negative

               if (Query1.FieldByName('B_Class').AsInteger in [4,7]) then begin
                  ThisAmtTrust := Query1.FieldByName('B_Amount').AsFloat;
                  ThisAmt864 := 0;
                  TotTotals[1] := TotTotals[1] + ThisAmtTrust;
               end else if (Query1.FieldByName('B_Class').AsInteger in [13]) then begin
                  ThisAmtTrust := Query1.FieldByName('B_Amount').AsFloat;
               end else if (Query1.FieldByName('B_Class').AsInteger in [8..11,19,24]) then begin
                  ThisAmtTrust := (Query1.FieldByName('B_Amount').AsFloat) * -1;
                  ThisAmt864 := 0;

                  if (Query1.FieldByName('B_Class').AsInteger =  8) then
                     TotTotals[3] := TotTotals[3] + ThisAmtTrust;

                  if (Query1.FieldByName('B_Class').AsInteger =  9) then
                     TotTotals[4] := TotTotals[4] + ThisAmtTrust;

                  if (Query1.FieldByName('B_Class').AsInteger = 10) then
                     TotTotals[4] := TotTotals[4] + ThisAmtTrust;

                  if (Query1.FieldByName('B_Class').AsInteger = 11) then
                     TotTotals[5] := TotTotals[5] + ThisAmtTrust;

                  if (Query1.FieldByName('B_Class').AsInteger = 19) then
                     TotTotals[6] := TotTotals[6] + ThisAmtTrust;

                  if (Query1.FieldByName('B_Class').AsInteger = 24) then
                     TotTotals[3] := TotTotals[3] + ThisAmtTrust;

               end else if (Query1.FieldByName('B_Class').AsInteger in [12]) then begin
                  ThisAmtTrust := (Query1.FieldByName('B_Amount').AsFloat) * -1;
               end;

//--- Determine whether the S86(4) amount is positive or negative

               if (Query1.FieldByName('B_Class').AsInteger in [14]) then begin
                  ThisAmt864 := Query1.FieldByName('B_Amount').AsFloat;
                  ThisAmtTrust := 0;
                  TotTotals[2] := TotTotals[2] + ThisAmt864;
               end else if (Query1.FieldByName('B_Class').AsInteger in [12]) then begin
                  ThisAmt864 := Query1.FieldByName('B_Amount').AsFloat;
               end else if (Query1.FieldByName('B_Class').AsInteger in [13]) then begin
                  ThisAmt864 := (Query1.FieldByName('B_Amount').AsFloat) * -1;
               end;

//--- Now write the line - Trust and S86(4)

               with xlsSheet1.RCRange[row1,1,row1,5] do begin
                  Item[1,1].Value := ' ';
                  Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
                  Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
                  Item[1,2].Value := Query1.FieldByName('B_Date').AsString;
                  Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
                  Item[1,3].Value := ReplaceQuote(Query1.FieldByName('B_Description').AsString);
                  Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
                  Item[1,4].Value := RoundD(ThisAmt864,2);
                  Item[1,4].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
                  Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
                  Item[1,5].Value := RoundD(ThisAmtTrust,2);
                  Item[1,5].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
                  Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ThisFill;
                  Font.Color := ThisText;
                  Font.Bold := false;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;

               ThisTrust := ThisTrust + ThisAmtTrust;
               This864 := This864 + ThisAmt864;

               Query1.Next;

               inc(row1);
               inc(PageRow1);

//--- If we've reached the maximum rows per page then it is PageBreak time

               if (PageRow1 >= RowsPerPage1) then begin
                  PageBreak1 := True;
                  DoPageBreak(ord(PB_TRUSTDETAIL),xlsSheet1,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1,'');
               end;
            end;

            if (RecsFound = true) then begin

//--- Write the Closing Balance

               with xlsSheet1.RCRange[row1,1,row1,5] do begin
                  Borders[xlAround].Weight := xlThin;
                  Item[1,1].Value := FileArray[idx1];
                  Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
                  Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
                  Item[1,2].Value := EDate;
                  Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
                  Item[1,3].Value := 'Closing Balance (R ' + FloatToStrF((This864 + ThisTrust),ffNumber,10,2) + ')';
                  Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
                  Item[1,4].Value := RoundD(This864,2);
                  Item[1,4].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
                  Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
                  Item[1,5].Value := RoundD(ThisTrust,2);
                  Item[1,5].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
                  Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
                  Interior.Color := ThisFill;
                  Font.Color := ThisText;
                  Font.Bold := True;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;

               inc(row1);
               inc(PageRow1);

//--- If we've reached the maximum rows per page then it is PageBreak time

               if (PageRow1 >= RowsPerPage1) then begin
                  PageBreak1 := True;
                  DoPageBreak(ord(PB_TRUSTDETAIL),xlsSheet1,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1,'');
               end;
            end;
         end;

//--- Write the Summary Line - Summary. Perform a Page break on the Summary
//--- Sheet if necessary

         if (PageBreak2 = True) then
            DoPageBreak(ord(PB_TRUSTSUMMARY),xlsSheet2,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2,'');

         with xlsSheet2.RCRange[row2,1,row2,6] do begin
            Item[1,1].Value := FileArray[idx1];
            Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := GetDescription(FileArray[idx1]);
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Value := 0;
            Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].Value := RoundD(This864,2);
            Item[1,5].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
            Item[1,6].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,6].Value := RoundD(ThisTrust,2);
            Item[1,6].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
            Borders[xlAround].Weight := xlThin;
            Interior.Color := ThisAB1F;
            Font.Color := ThisAB1T;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;

         TotalTrust := TotalTrust + ThisTrust;
         Total864 := Total864 + This864;

         inc(row2);
         inc(PageRow2);

//--- If we've reached the maximum rows per page then it is PageBreak time

         if (PageRow2 >= RowsPerPage2) then begin
            PageBreak2 := True;
            DoPageBreak(ord(PB_TRUSTSUMMARY),xlsSheet2,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2,'');
         end;
      end;
   end;

   SaveRow2 := row2;
   row1     := 5;
   row2     := 5;

//--- Insert the S86(3) Interest

   with xlsSheet2.RCRange[row2,1,row2,6] do begin
      Item[1,1].Value := CpyFile;
      Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,1].Borders[xlEdgeRight].Weight := xlThin;

      if (EDate < S86Date) then
         Item[1,2].Value := 'Section 78(2)(a) Interest'
      else
         Item[1,2].Value := 'Section 86(3) Interest';

      Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,4].Value := 0;
      Item[1,4].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].Value := 0;
      Item[1,5].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Item[1,6].Value := 0;
      Item[1,6].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Borders[xlAround].Weight := xlThin;
      Interior.Color := ThisAB1F;
      Font.Color := ThisAB1T;
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

//--- Get the Investment/Withdrawal and Interest totals for S86(3)

   if ((GetS863(CpyFile,SDate,EDate)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get the amounts for the current period

   This863Inv    := 0;
   This863Drw    := 0;
   This863Int    := 0;
   This863IntDrw := 0;

   for idx2 := 0 to Query1.RecordCount - 1 do begin
      if (Query1.FieldByName('B_Class').AsInteger in [15]) then begin
         This863Inv := This863Inv + Query1.FieldByName('B_Amount').AsFloat;
      end else if (Query1.FieldByName('B_Class').AsInteger in [16]) then begin
         This863Drw := This863Drw + Query1.FieldByName('B_Amount').AsFloat;
      end else if (Query1.FieldByName('B_Class').AsInteger in [17]) then begin
         This863Int := This863Int + Query1.FieldByName('B_Amount').AsFloat;
      end else if (Query1.FieldByName('B_Class').AsInteger in [20]) then begin
         This863IntDrw := This863IntDrw + Query1.FieldByName('B_Amount').AsFloat;
      end;
      Query1.Next;
   end;

//--- Write the S86(3) Investment Summary Lines

   with xlsSheet1.RCRange[row1,1,row1,5] do begin
      Item[1,1].Value := CpyFile;
      Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,1].Borders[xlEdgeRight].Weight := xlThin;

      if (EDate < S86Date) then
         Item[1,2].Value := 'Section 78(2)(a) Investments'
      else
         Item[1,2].Value := 'Section 86(3) Investments';

      Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].Value := RoundD(OpenBal863Inv + This863Inv,2);
      Item[1,5].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Borders[xlAround].Weight := xlThin;
      Interior.Color := ThisAB2F;
      Font.Color := ThisAB2T;
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   inc(row1);

   with xlsSheet1.RCRange[row1,1,row1,5] do begin
      Item[1,1].Value := CpyFile;
      Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,2].Value := 'Section 86(3) Interest';
      Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].Value := RoundD(OpenBal863Int + This863Int,2);
      Item[1,5].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Borders[xlAround].Weight := xlThin;
      Interior.Color := ThisAB2F;
      Font.Color := ThisAB2T;
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   inc(row1);

   with xlsSheet1.RCRange[row1,1,row1,5] do begin
      Item[1,1].Value := CpyFile;
      Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,2].Value := 'Section 86(3) Witdrawals';
      Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].Value := RoundD((OpenBal863Drw + This863Drw + OpenBal863IntDrw + This863IntDrw) * -1,2);
      Item[1,5].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Borders[xlAround].Weight := xlThin;
      Interior.Color := ThisAB2F;
      Font.Color := ThisAB2T;
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   inc(row1);

//--- Write the S86(3) Total Line - Summary

   with xlsSheet1.RCRange[row1,1,row1,5] do begin
      Borders[xlAround].Weight := xlThin;

      if (EDate < S86Date) then
         Item[1,1].Value := 'Total Section 78(2)(a) Investments, Witdrawals and Interest'
      else
         Item[1,1].Value := 'Total Section 86(3) Investments, Witdrawals and Interest';

      Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].Value := RoundD((OpenBal863Inv + This863Inv) - (OpenBal863Drw + This863Drw + OpenBal863IntDrw + This863IntDrw) + OpenBal863Int + This863Int,2);
      Item[1,5].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Interior.Color := ThisAB2F;
      Font.Color := ThisAB2T;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

//--- Insert the Section 78(2)(a) Interest into the Trust Summary

   row2 := 5;

   with xlsSheet2.RCRange[row2,4,row2,6] do begin
      Item[1,1].Value := RoundD((OpenBal863Int + This863Int - OpenBal863IntDrw - This863IntDrw),2);
      Item[1,2].Value := 0;
      Item[1,3].Value := 0;
   end;

//--- Write the Sub Totals

   with xlsSheet2.RCRange[SaveRow2,1,SaveRow2,6] do begin
      Item[1,1].Value := 'Trust Totals';
      Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,4].Formula := '=SUM(D5:D' + IntToStr(SaveRow2 - 1) + ')';
      Item[1,4].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,5].Formula := '=SUM(E5:E' + IntToStr(SaveRow2 - 1) + ')';
      Item[1,5].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,6].Formula := '=SUM(F5:F' + IntToStr(SaveRow2 - 1) + ')';
      Item[1,6].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Borders[xlAround].Weight := xlThin;
      Interior.Color := ColDHF;
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

//--- Write the Trust Total

   SaveRow2 := SaveRow2 + 2;

   with xlsSheet2.RCRange[SaveRow2,1,SaveRow2,6] do begin
      if (EDate < S86Date) then
         Item[1,1].Value := 'Trust Balance (Including S78(2)(a) Balance and Interest)'
      else
         Item[1,1].Value := 'Trust Balance (Including S86(3) Balance and Interest)';

      Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,6].Formula := '=D' + IntToStr(SaveRow2 - 2) + '+E' + IntToStr(SaveRow2 - 2) + '+F' + IntToStr(SaveRow2 - 2);
      Item[1,6].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Borders[xlAround].Weight := xlThin;
      Interior.Color := ColDHF;
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

//--- Calculate and write Trust related totals for the Trust Totals Sheet

   with xlsSheet3.Range['A1','B1'] do begin
      Item[1,1].Value := CpyName;
      Borders[xlAround].Weight := xlThin;
      Interior.Color := ColDHF;
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 11;
   end;

   with xlsSheet3.Range['A2','B2'] do begin
      Item[1,1].Value := 'Trust Related Control Totals for Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
      Borders[xlAround].Weight := xlThin;
      Interior.Color := ColDHF;
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 11;
   end;

   with xlsSheet3.Range['A4','B4'] do begin
      Item[1,1].Value := 'Description ';
      Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,1].ColumnWidth := 113;
      Item[1,2].Value := 'Amount ';
      Item[1,2].HorizontalAlignment := xlHAlignRight;
      Item[1,2].ColumnWidth := 16;
      Borders[xlAround].Weight := xlThin;
      Interior.Color := integer(ColDHF);
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   row3 := 4;

   for idx3 := 1 to 6 do begin
      with xlsSheet3.RCRange[row3 + idx3,1,row3 + idx3,2] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := CtrlLines[idx3];
         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,2].Value := 0;
         Item[1,2].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
         Interior.Color := ThisAB2F;
         Font.Color := ThisAB2T;
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

   CtrlTotals[1] := (OpenBal863Inv + This863Inv) - (OpenBal863Drw + This863Drw);
   CtrlTotals[2] := (OpenBal863Int + This863Int) - (OpenBal863IntDrw + This863IntDrw);
   CtrlTotals[3] := CtrlTotals[1] + CtrlTotals[2];
   CtrlTotals[4] := Total864;
   CtrlTotals[5] := TotalTrust + CtrlTotals[2] - CtrlTotals[1];
   CtrlTotals[6] := TotalTrust + CtrlTotals[2];

   for idx3 := 1 to 6 do begin
      with xlsSheet3.RCRange[row3 + idx3,2,row3 + idx3,2] do begin
         Item[1,1].Value := RoundD(CtrlTotals[idx3],2);
      end;
   end;

//--- Get and write the S86(3) Interest records for the YTD period

   if ((GetS863(CpyFile,FYYear + '/' + FYMonth + '/01',Edate)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

   S864Recs := Query1.RecordCount;

   if (S864Recs > 0) then begin

//--- Get the amounts for the current period

      This863Int    := 0;

      for idx2 := 0 to Query1.RecordCount - 1 do begin
         if (Query1.FieldByName('B_Class').AsInteger in [17]) then begin

            if (PageBreak4 = True) then
               DoPageBreak(ord(PB_TRUSTS864),xlsSheet4,FirstPage4,RowsPerPage4,row4,Pages4,PageRow4,PageBreak4,'');

            Total863Int := Query1.FieldByName('B_Amount').AsFloat;

            with xlsSheet4.RCRange[row4,1,row4,3] do begin
               Item[1,1].Value := Query1.FieldByName('B_Date').AsString;
               Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,2].Value := Query1.FieldByName('B_Description').AsString;
               Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,3].Value := RoundD(Query1.FieldByName('B_Amount').AsFloat,2);
               Item[1,3].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ThisAB1F;
               Font.Color := ThisAB1T;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row4);
            inc(PageRow4);

//--- If we've reached the maximum rows per page then it is PageBreak time

            if (PageRow4 >= RowsPerPage4) then begin
               PageBreak4 := True;
               DoPageBreak(ord(PB_TRUSTS864),xlsSheet4,FirstPage4,RowsPerPage4,row4,Pages4,PageRow4,PageBreak4,'');
            end;

            This863Int := This863Int + Total863Int;
         end;
         Query1.Next;
      end;

//--- Write the total of the YTD interest

      with xlsSheet4.RCRange[row4,1,row4,3] do begin
         Item[1,1].Value := EDate;
         Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,2].Value := 'Total Section 86(3) interest earned for Year-To-Date';
         Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,3].Value := RoundD(This863Int,2);
         Item[1,3].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ThisAB1F;
         Font.Color := ThisAB1T;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

//--- Write the standard copyright notice

   row1 := (Pages1 * lcGRows);
   with xlsSheet1.RCRange[row1,1,row1,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

   row2 := (Pages2 * lcGRows);
   with xlsSheet2.RCRange[row2,1,row2,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

   row3 := lcGRows;
   with xlsSheet3.RCRange[row3,1,row3,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

   if (S864Recs > 0) then begin
      row4 := (Pages4 * lcGRows);
      with xlsSheet4.RCRange[row4,1,row4,1] do begin
         Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 8;
      end;
   end;

//--- Sheet1 - Remove the Gridlines which are added by default and set the Page
//--- orientation on the Statement page

   xlsSheet1.PageSetup.Orientation    := xlLandscape;
   xlsSheet1.PageSetup.FitToPagesWide := 1;
   xlsSheet1.PageSetup.FitToPagesTall := Pages1;
   xlsSheet1.PageSetup.PaperSize      := xlPaperA4;
   xlsSheet1.DisplayGridLines         := false;
   xlsSheet1.PageSetup.CenterFooter    := 'Page &P of &N';

//--- Sheet2 - Remove the Gridlines which are added by default and set the Page
//--- orientation on the Statement page

   xlsSheet2.PageSetup.Orientation    := xlLandscape;
   xlsSheet2.PageSetup.FitToPagesWide := 1;
   xlsSheet2.PageSetup.FitToPagesTall := Pages2;
   xlsSheet2.PageSetup.PaperSize      := xlPaperA4;
   xlsSheet2.DisplayGridLines         := false;
   xlsSheet2.PageSetup.CenterFooter    := 'Page &P of &N';

//--- Sheet3 - Remove the Gridlines which are added by default and set the Page
//--- orientation on the Statement page

   xlsSheet3.PageSetup.Orientation    := xlLandscape;
   xlsSheet3.PageSetup.FitToPagesWide := 1;
   xlsSheet3.PageSetup.FitToPagesTall := 1;
   xlsSheet3.PageSetup.PaperSize      := xlPaperA4;
   xlsSheet3.DisplayGridLines         := false;
   xlsSheet3.PageSetup.CenterFooter    := 'Page &P of &N';

//--- Sheet4 - Remove the Gridlines which are added by default and set the Page
//--- orientation on the Statement page

   if (S864Recs > 0) then begin
      xlsSheet4.PageSetup.Orientation    := xlLandscape;
      xlsSheet4.PageSetup.FitToPagesWide := 1;
      xlsSheet4.PageSetup.FitToPagesTall := Pages4;
      xlsSheet4.PageSetup.PaperSize      := xlPaperA4;
      xlsSheet4.DisplayGridLines         := false;
      xlsSheet4.PageSetup.CenterFooter    := 'Page &P of &N';
   end else
      xlsSheet4.Delete;

//--- Write the Excel file to disk

   if (FilesFound > 0) then begin
      xlsBook.SaveAs(FileName + ThisFile);
      xlsBook.Close;

      LogMsg('  Trust Account Reconcilliation successfully processed...',True);
      LogMsg(' ',True);

      DoLine := False;

//--- Print the generated document on the Default Printer if requested

      if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
         if (PrintDocument(ThisFile, FileName) = True) then
            LogMsg('  Document submitted for printing...',True)
         else
            LogMsg('  Printing of document failed...',True);

         DoLine := True;
      end;

//--- Create a PDF file if requested

      if ((CreatePDF = true) and (PDFPrinter <> 'Not Found')) then begin
         PDFExists := PDFDocument(ThisFile, FileName);

         if (PDFExists = True) then
            LogMsg('  PDF file creation was successfull...',True)
         else
            LogMsg('  PDF file creation failed...',True);

         DoLine := True;
      end;

//--- Send the Excel file via email if requested

      if (SendByEmail = '1') then begin
         if (SendEmail('',FileName + ThisFile,ord(PT_NORMAL)) = true) then
            LogMsg('  Request to send generated Trust Account Summary by Email submitted...',True)
         else
            LogMsg('  Request to send generated Trust Account Summary by Email not submitted...',True);

         DoLine := True;
      end;

//--- Now open the Trust Account Summary if requested

      if (AutoOpen = True) then begin
         ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);

         LogMsg('  Request to open Trust Account Summary for ''' + PChar(FileName + ThisFile) + ''' submitted...',True);
         DoLine := True;
      end;
   end else begin
      LogMsg('  No Trust records found ...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Trust Account Reconcilliation');

   Close_Connection;
end;

//---------------------------------------------------------------------------
// Generate Trust Detail when a Simple Trust Report was requested
//---------------------------------------------------------------------------
procedure TFldExcel.DoTrust_Simple();
var
   PageNum, ThisRows, Row, NumPages         : integer;
   idx1, idx2, idx3, RemainRows, FilesFound : integer;
   DoLine                                   : boolean;
   ThisItem, ThisFile, ThisStr, ThisName    : string;
   xlsBook                                  : IXLSWorkbook;
   xlsSheet                                 : IXLSWorksheet;
   ThisDate                                 : TDateTime;

const
   TotLines  : array[1..7] of string = (
               'Deposits to Trust',
               'Section 78(2A) Interest',
               'Trust to Business',
               'Trust to Disbursements',
               'Trust to Trust',
               'Trust to Other',
               'Trust Balance');

begin

   ThisCount := 1;
   ThisMax   := IntToStr(NumFiles);

   prbProgress.Max := NumFiles;
   LogMsg(ord(PT_PROLOG),True,True,'Trust Reconciliation');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Get the Heading and Layout information

   GetLayout('Trust');

//--- Read the Trust record from the datastore

   ShortDateFormat := 'yyyy/MM/dd';
   DateSeparator   := '/';

   for idx1 := 0 to NumFiles - 1 do begin
      PageNum := 1;
      ThisDate := StrToDate(EDate);
      ThisFile := FormatDateTime('yyyyMMdd',ThisDate) + ' - Trust Account Reconciliation (' + FileArray[idx1] + ').xls';
      txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;
      txtDocument.Refresh;

      OpenBalTrust := 0;
      OpenBal864   := 0;

//--- Build the unique part of the SQL statement

      ThisStr := ' AND ((B_Class = 4) OR (B_Class >= 7 AND B_Class <= 14) OR (B_Class = 19) OR (B_Class = 24))';
      if ((GetBilling(FileArray[idx1],ThisStr,'Trust Simple')) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      txtError.Text := 'Processing: ' + FileArray[idx1];
      txtError.Refresh;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

//--- Only process files that have records

      if (Query1.RecordCount > 0) then begin

//--- Open the Excel workbook template

         xlsBook := TXLSWorkbook.Create;
         xlsBook.Open(Template_T);
         xlsSheet := xlsBook.ActiveSheet;
         xlsSheet.Name := 'Trust Reconciliation (' + FileArray[idx1] + ')';

//--- Clear everything but keep the Header information if lcShowHeader is set

         if (lcShowHeader = true) then
            xlsSheet.RCRange[lcHER + 1, 1, 999, lcHEC].Clear
         else
            xlsSheet.RCRange[1, 1, 999, lcSMaxCols].Clear;

//--- Insert the Page 1 Heading

         Generate_DocHeading(xlsSheet,TrustPages,idx1,lcHER + 1,lcHEC,'Trust');

//--- Insert the Customer Information

         if (lcShowAddress = true) then begin

            GetAddress(FileArray[idx1]);

            with xlsSheet.RCRange[lcASR,1,lcAER,lcHEC] do begin
               Item[1,lcASC].Value := Customer;
               Item[2,lcASC].Value := Address1;
               Item[3,lcASC].Value := Address2;
               Item[4,lcASC].Value := Address3;
               Item[5,lcASC].Value := Address4;
               Item[6,lcASC].Value := Address5;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;

//--- Insert the instruction information

         if (lcShowInstruct = true) then begin
            with xlsSheet.RCRange[lcISR,1,lcISR + 2,lcHEC] do begin
               Item[1,1].Value := 'Client:';
               Item[2,1].Value := 'Instruction:';
               Item[1,lcISCD].Value := Customer;
               Item[2,lcISCD].Value := Descrip;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;

//--- Write the Data Heading Information

         Generate_DataHeading(xlsSheet,lcPSR,lcHEC,'Trust');

//--- Page 1 data

         if (Query1.RecordCount <= (lcPRows - 3)) then begin
            ThisRows := Query1.RecordCount + 2;
            ThisItem := 'Closing Balance';
         end else begin
            ThisRows := lcPRows - 1;
            ThisItem := 'Carried Over';
         end;

         LastDate := Sdate;
         Generate_Detail_Trust(xlsSheet, lcPSR, ThisRows, ThisItem, 'Opening Balance',FileArray[idx1]);

//--- Data on subsequent pages - compensate for Document Heading (3 rows)

         if (Query1.RecordCount > (lcPRows - 3)) then begin
            RemainRows := (Query1.RecordCount - (lcProws - 3));
            NumPages := ((Query1.RecordCount - (lcPRows - 3)) div (lcSRows - 3)) + 1;

//--- Compensate for cases where we have an exact page size

            if ((Query1.RecordCount - (lcPRows - 3)) mod (lcSRows - 3) = 0) then
               NumPages := NumPages - 1;

            Row := lcSSR;
            PageNum := PageNum + 1;

            for idx3 := 0 to NumPages -1 do begin
               if (lcHeaderPageOne = false) then begin
                  xlsSheet.RCRange[lcHSR,lcHSC,lcHER,lcHEC].Copy(xlsSheet.RCRange[lcSSR,lcHSC,lcSSR + lcHER,lcHEC]);

                  Row := Row + lcHER + 1;
               end;

               Generate_DocHeading(xlsSheet,PageNum,idx1,Row,lcHEC,'Trust');
               Row := Row + 4;
               Generate_DataHeading(xlsSheet,Row,lcHEC,'Trust');

               if (RemainRows <= (lcSRows - 3)) then begin
                  ThisRows := RemainRows + 2;
                  ThisItem := 'Closing Balance';
               end else begin
                  ThisRows := lcSRows - 1;
                  ThisItem := 'Carried Over';
               end;

               Generate_Detail_Trust(xlsSheet, Row, ThisRows, ThisItem, 'Carried Down',FileArray[idx1]);

               PageNum := PageNum + 1;
               RemainRows := RemainRows - lcSRows + 3;
               Row := Row + lcSMaxRows - 4;
            end;
         end else begin
            Row := lcSSR;
         end;

//--- Insert the Summary Information

         TotTotals[7] := TotTotals[1] + TotTotals[2] + TotTotals[3] + TotTotals[4] + TotTotals[5] + TotTotals[6];

         if (lcShowSummary = true) then begin
            for idx2 := 1 to Length(TotTotals) do begin
               with xlsSheet.RCRange[lcXSR + (idx2 - 1),lcXSCL,lcXSR + (idx2 - 1),lcXSCD] do begin
                  Item[1,1].Value := TotLines[idx2];
                  Item[1,(lcXSCD - lcXSCL) + 1].Value := TotTotals[idx2];
                  Item[1,(lcXSCD - lcXSCL) + 1].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
                  Item[1,(lcXSCD - lcXSCL) + 1].Borders[xlEdgeLeft].Weight := xlThin;

                  Borders[xlAround].Weight := xlThin;
                  Font.Bold := false;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;
               xlsSheet.RCRange[lcXSR + 6,lcXSCL,lcXSR + 6,lcXSCD].Font.Bold := true;
            end;
         end;

//--- Write the standard copyright notice

         dec(Row);

         with xlsSheet.RCRange[Row,1,Row,1] do begin
            Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 8;
         end;

//--- Remove the Gridlines which are added by default and set the Page orientation

         xlsSheet.PageSetup.Orientation    := xlLandscape;
         xlsSheet.PageSetup.FitToPagesWide := 1;
         xlsSheet.PageSetup.FitToPagesTall := TrustPages;
         xlsSheet.PageSetup.PaperSize      := xlPaperA4;
         xlsSheet.DisplayGridLines         := false;
         xlsSheet.PageSetup.CenterFooter    := 'Page &P of &N';

//--- Save the Generated Workbook

         xlsBook.SaveAs(FileName + ThisFile);
         xlsBook.Close;

         LogMsg('  Trust Reconcilliation successfully processed...',True);
         LogMsg(' ',True);

         DoLine := False;

//--- Print the generated document on the Default Printer if requested

         if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
            if (PrintDocument(ThisFile, FileName) = True) then
               LogMsg('  Document submitted for printing...',True)
            else
               LogMsg('  Printing of document failed...',True);

            DoLine := True;
         end;

//--- Create a PDF file if requested

         if ((CreatePDF = true) and (PDFPrinter <> 'Not Found')) then begin
            PDFExists := PDFDocument(ThisFile, FileName);

            if (PDFExists = True) then
               LogMsg('  PDF file creation was successfull...',True)
            else
               LogMsg('  PDF file creation failed...',True);

            DoLine := True;
         end;

//--- Send the Excel file via email if requested

         if (NumFiles > 1) then ThisName := '';

         if (SendByEmail = '1') then begin
            if (SendEmail(ThisName,FileName + ThisFile,ord(PT_NORMAL)) = true) then
               LogMsg('  Request to send generated Trust Account Reconciliation by Email submitted...',True)
            else
               LogMsg('  Request to send generated Trust Account Reconciliation by Email not submitted...',True);

            DoLine := True;
         end;

//--- Now open the Trust Account Summary if requested

         if (AutoOpen = True) then begin
            ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);

            LogMsg('  Request to open Trust Account Reconciliation for ''' + PChar(FileName + ThisFile) + ''' submitted...',True);
            DoLine := True;
         end;
      end else begin
         LogMsg('  There are no Trust records for ''' + FileArray[idx1] + ''' in the specified period...',True);
         DoLine := True;
      end;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Trust Reconcilliation');

   Close_Connection;

end;

//---------------------------------------------------------------------------
// Generate the Detail for a Simple Trust Report
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Detail_Trust(xlsSheet : IXLSWorksheet; PageRows: integer; ThisRows: integer; ThisItem: string; BalanceStr: string; ThisFile: string);
var
   idx1                     : integer;
   Pref                     : string;
   ThisAmtTrust, ThisAmt864 : double;

begin

//--- Page data

   with xlsSheet.RCRange[PageRows + 1,1,PageRows + ThisRows,lcHEC] do begin
      Borders[xlAround].LineStyle := xlContinuous;
      Borders[xlAround].Weight := xlThin;
   end;

   with xlsSheet.RCRange[PageRows + 1,lcHEC - 2, PageRows + ThisRows, lcHEC] do begin
      NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

//-- Balance Line

   with xlsSheet.RCRange[PageRows + 1,1,PageRows + 1,lcHEC] do begin
      Borders[xlEdgeBottom].LineStyle := xlContinuous;
      Borders[xlEdgeBottom].Weight := xlThin;

      Item[1,1].Value := ThisFile;
      Item[1,2].Value := LastDate;
      Item[1,2].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,3].Value := BalanceStr;
      Item[1,3].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,lcHEC - 1].Value := OpenBal864;
      Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,lcHEC].Value := OpenBalTrust;
      Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;

      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

//--- Page Detail Lines

   if (InvoiceInfo = '1') then
      Pref := 'Collect'
   else
      Pref := 'B';

   for idx1 := 2 to ThisRows - 1 do begin
      with xlsSheet.RCRange[PageRows + idx1,1,PageRows + idx1,lcHEC] do begin
         Borders[xlEdgeBottom].LineStyle := xlContinuous;
         Borders[xlEdgeBottom].Weight := xlThin;

         LastDate := Query1.FieldByName(Pref + '_Date').AsString;
         Item[1,2].Value := LastDate;
         Item[1,2].Borders[xlEdgeLeft].Weight := xlThin;
         Item[1,3].Value := ReplaceQuote(Query1.FieldByName(Pref + '_Description').AsString);
         Item[1,3].Borders[xlEdgeLeft].Weight := xlThin;

//--- Determine whether the Trust amount is positive or negative

         if (Query1.FieldByName('B_Class').AsInteger in [4,7]) then begin
            ThisAmtTrust := Query1.FieldByName('B_Amount').AsFloat;
            ThisAmt864 := 0;
            TotTotals[1] := TotTotals[1] + ThisAmtTrust;
         end else if (Query1.FieldByName('B_Class').AsInteger in [13]) then begin
            ThisAmtTrust := Query1.FieldByName('B_Amount').AsFloat;
         end else if (Query1.FieldByName('B_Class').AsInteger in [8..11,19,24]) then begin
            ThisAmtTrust := (Query1.FieldByName('B_Amount').AsFloat) * -1;
            ThisAmt864 := 0;

            if (Query1.FieldByName('B_Class').AsInteger =  8) then
               TotTotals[3] := TotTotals[3] + ThisAmtTrust;

            if (Query1.FieldByName('B_Class').AsInteger =  9) then
               TotTotals[4] := TotTotals[4] + ThisAmtTrust;

            if (Query1.FieldByName('B_Class').AsInteger = 10) then
               TotTotals[4] := TotTotals[4] + ThisAmtTrust;

            if (Query1.FieldByName('B_Class').AsInteger = 11) then
               TotTotals[5] := TotTotals[5] + ThisAmtTrust;

            if (Query1.FieldByName('B_Class').AsInteger = 19) then
               TotTotals[6] := TotTotals[6] + ThisAmtTrust;

            if (Query1.FieldByName('B_Class').AsInteger = 24) then
               TotTotals[3] := TotTotals[3] + ThisAmtTrust;

         end else if (Query1.FieldByName('B_Class').AsInteger in [12]) then begin
            ThisAmtTrust := (Query1.FieldByName('B_Amount').AsFloat) * -1;
         end;

//--- Determine whether the S86(4) amount is positive or negative

         if (Query1.FieldByName('B_Class').AsInteger in [14]) then begin
            ThisAmt864 := Query1.FieldByName('B_Amount').AsFloat;
            ThisAmtTrust := 0;
            TotTotals[2] := TotTotals[2] + ThisAmt864;
         end else if (Query1.FieldByName('B_Class').AsInteger in [12]) then begin
            ThisAmt864 := Query1.FieldByName('B_Amount').AsFloat;
         end else if (Query1.FieldByName('B_Class').AsInteger in [13]) then begin
            ThisAmt864 := (Query1.FieldByName('B_Amount').AsFloat) * -1;
         end;

         Item[1,lcHEC].Value := ThisAmtTrust;
         Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;
         OpenBalTrust := OpenBalTrust + ThisAmtTrust;
         Item[1,lcHEC - 1].Value := ThisAmt864;
         Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
         OpenBal864 := OpenBal864 + ThisAmt864;

         Query1.Next;

         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

//--- Closing Balance Line

   with xlsSheet.RCRange[PageRows + ThisRows,1,PageRows + ThisRows,lcHEC] do begin
      Borders[xlEdgeBottom].LineStyle := xlContinuous;
      Borders[xlEdgeBottom].Weight := xlThin;

      Item[1,1].Value := ThisFile;
      Item[1,2].Value := LastDate;
      Item[1,2].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,3].Value := ThisItem;
      Item[1,3].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,lcHEC - 1].Value := OpenBal864;
      Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1,lcHEC].Value := OpenBalTrust;
      Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;

      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;
end;

//---------------------------------------------------------------------------
// Generate S86(3) Summary
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_S863_Summary();
var
   Row, idx1, idx2, FilesFound  : integer;
   ThisAmount, ThisTotal        : double;
   DoLine                       : boolean;
   ThisFile, ThisStr            : string;
   xlsBook                      : IXLSWorkbook;
   xlsSheet                     : IXLSWorksheet;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Section 86(3) Summary');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Open the Excel workbook template

   FilesFound := 0;

   xlsBook       := TXLSWorkbook.Create;
   xlsSheet      := xlsBook.WorkSheets.Add;

   if (EDate < S86Date) then begin
      xlsSheet.Name := 'S78(2)(a) Summary';
      ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - S78(2)(a) Summary (' + CpyName + ').xls';
   end else begin
      xlsSheet.Name := 'S86(3) Summary';
      ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - S86(3) Summary (' + CpyName + ').xls';
   end;

   txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;

   S86Pages     := 0;
   S86FirstPage := true;
   S86PageBreak := true;

//--- Read the S78(2)(a) records from the datastore

   for idx1 := 0 to NumFiles - 1 do begin

//--- Initialise the Total Balance Fields

      OpenBal863Inv    := 0;
      OpenBal863Int    := 0;
      OpenBal863Drw    := 0;
      OpenBal863IntDrw := 0;

//--- Build the unique part of the SQL statement

      ThisStr := ' AND ((B_Class >= 15 AND B_Class <= 17) OR (B_Class = 20))';

      if ((GetBilling(FileArray[idx1],ThisStr,'Section 86(3) Summary')) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

//--- Only process files that have records

      if (Query1.RecordCount > 0) then begin
         inc(FilesFound);

//--- Perform a Page break if necessary

         if (S86PageBreak = true) then
            row := DoS86_PageBreak(xlsSheet);

//--- Write the Opening Balance

         with xlsSheet.RCRange[row,1,row,4] do begin
            Item[1,1].Value := FileArray[idx1];
            Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,1].Borders[xlEdgeBottom].Weight := xlThin;
            Item[1,2].Value := SDate;
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Borders[xlEdgeBottom].Weight := xlThin;
            Item[1,3].Value := 'Opening Balance';
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Borders[xlEdgeBottom].Weight := xlThin;
            Item[1,4].Value := OpenBal863Inv + OpenBal863Int - OpenBal863Drw - OpenBal863IntDrw;
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Borders[xlEdgeBottom].Weight := xlThin;
            Item[1,4].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';

            Font.Bold := True;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;

         ThisTotal := OpenBal863Inv + OpenBal863Int - OpenBal863Drw - OpenBal863IntDrw;
         inc(row);
         inc(S86PageRow);

         for idx2 := 0 to Query1.RecordCount - 1 do begin

//--- Perform a Page break if necessary

            if (S86PageBreak = true) then
               row := DoS86_PageBreak(xlsSheet);

            if (Query1.FieldByName('B_Class').AsInteger in [15,17]) then
               ThisAmount := Query1.FieldByName('B_Amount').AsFloat
            else
               ThisAmount := (Query1.FieldByName('B_Amount').AsFloat) * -1;

            with xlsSheet.RCRange[row,1,row,4] do begin
               Item[1,1].Value := ' ';
               Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,1].Borders[xlEdgeBottom].Weight := xlThin;
               Item[1,2].Value := Query1.FieldByName('B_Date').AsString;
               Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,2].Borders[xlEdgeBottom].Weight := xlThin;
               Item[1,3].Value := ReplaceQuote(Query1.FieldByName('B_Description').AsString);
               Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,3].Borders[xlEdgeBottom].Weight := xlThin;
               Item[1,4].Value := ThisAmount;
               Item[1,4].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
               Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,4].Borders[xlEdgeBottom].Weight := xlThin;

               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            ThisTotal := ThisTotal + ThisAmount;

            inc(row);
            inc(S86PageRow);

            if (S86PageRow > S86RowsPerPage) then
               S86PageBreak := True;

            Query1.Next;
         end;

//--- Perform a Page break if necessary

         if (S86PageBreak = true) then
            row := DoS86_PageBreak(xlsSheet);

//--- Write the Closing Balance

         with xlsSheet.RCRange[row,1,row,4] do begin
            Item[1,1].Value := FileArray[idx1];
            Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,1].Borders[xlEdgeBottom].Weight := xlThin;
            Item[1,2].Value := EDate;
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Borders[xlEdgeBottom].Weight := xlThin;
            Item[1,3].Value := 'Closing Balance';
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Borders[xlEdgeBottom].Weight := xlThin;
            Item[1,4].Value := ThisTotal;
            Item[1,4].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Borders[xlEdgeTop].Weight := xlThin;
            Item[1,4].Borders[xlEdgeBottom].Weight := xlThin;

            Font.Bold := True;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;

         inc(row);
         inc(S86PageRow);

         if (S86PageRow > S86RowsPerPage) then
            S86PageBreak := True;

//--- Perform a Page break if necessary

         if (S86PageBreak = true) then
            row := DoS86_PageBreak(xlsSheet);

      end;
   end;

//--- Write the standard copyright notice

   row := (S86Pages * lcGRows);

   with xlsSheet.RCRange[row,1,row,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet.PageSetup.Orientation    := xlLandscape;
   xlsSheet.PageSetup.FitToPagesWide := 1;
   xlsSheet.PageSetup.FitToPagesTall := S86Pages;
   xlsSheet.PageSetup.PaperSize      := xlPaperA4;
   xlsSheet.DisplayGridLines         := false;
   xlsSheet.PageSetup.CenterFooter    := 'Page &P of &N';

//--- Write the Excel file to disk

   if (FilesFound > 0) then begin
      xlsBook.SaveAs(FileName + ThisFile);
      xlsBook.Close;

      LogMsg('  Section 86(3) Summary successfully processed...',True);
      LogMsg(' ',True);

      DoLine := False;

//--- Print the generated document on the Default Printer if requested

      if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
         if (PrintDocument(ThisFile, FileName) = True) then
            LogMsg('  Document submitted for printing...',True)
         else
            LogMsg('  Printing of document failed...',True);

         DoLine := True;
      end;

//--- Create a PDF file if requested

      if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
         PDFExists := PDFDocument(ThisFile, FileName);

         if (PDFExists = True) then
            LogMsg('  PDF file creation was successfull...',True)
         else
            LogMsg('  PDF file creation failed...',True);

         DoLine := True;
      end;

//--- Send the Excel file via email if requested

      if (SendByEmail = '1') then begin
         if (SendEmail('',FileName + ThisFile,ord(PT_NORMAL)) = true) then
            LogMsg('  Request to send generated S86(3) Summary by Email submitted...',True)
         else
            LogMsg('  Request to send generated S86(3) Summary by Email not submitted...',True);

         DoLine := True;
      end;

//--- Now open the S78(2)(a) Account Summary if requested

      if (AutoOpen = True) then begin
         ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);

         LogMsg('  Request to open S86(3) Summary for ''' + PChar(FileName + ThisFile) + ''' submitted...',True);
         DoLine := True;
      end;
   end else begin
      LogMsg('  No S86(3) records found ...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Section 86(3) Summary');

   Close_Connection;

end;

//---------------------------------------------------------------------------
// Procedure to do a S78(2)(a) Recon page break
//---------------------------------------------------------------------------
function TFldExcel.DoS86_PageBreak(xlsSheet : IXLSWorksheet) : integer;
var
   row : integer;

begin

//--- Set the Page control variables

   if ((lcRepeatHeader = true) or (S86FirstPage = true)) then
      S86RowsPerPage := (lcGRows - 4)
   else
      S86RowsPerPage := (lcGRows - 1);

   if (S86FirstPage = true) then
      row := 1
   else
      row := (S86Pages * lcGRows) + 1;

   S86PageRow   := 1;
   S86PageBreak := false;
   inc(S86Pages);

//--- Check if the heading must be printed / repeated

   if ((S86FirstPage = true) or (lcRepeatHeader = true)) then begin
      S86FirstPage := false;


//--- Insert the Header (1st line) and the Heading

      with xlsSheet.Range['A' + IntToStr(row), 'D' + IntToStr(row)] do begin
         if (EDAte < S86Date) then
            Item[1,1].Value := CpyName + ': S78(2)(a) Summary for Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now())
         else
            Item[1,1].Value := CpyName + ': S86(3) Summary for Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());

         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;

      inc(row);
      inc(row);
   end;

   with xlsSheet.Range['A' + IntToStr(row), 'D' + IntToStr(row)] do begin
      Item[1,1].Value := 'File';
      Item[1,1].ColumnWidth := 10;
      Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,2].Value := 'Date';
      Item[1,2].ColumnWidth := 14;
      Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,3].Value := 'Desription';
      Item[1,3].ColumnWidth := lcGMRWidth - 38;
      Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,4].Value := 'Amount ';
      Item[1,4].ColumnWidth := 14;
      Item[1,4].HorizontalAlignment := xlHAlignRight;
      Borders[xlAround].Weight := xlThin;
      Interior.Color := integer(ColDHF);
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   inc(row);
   Result := row;
end;

//---------------------------------------------------------------------------
// Generate the Fee Earner Report
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_FeeReport();
var
   row1, row2, idx1, FilesFound, Sema1, Sema2, PageRow1 : integer;
   PageRow2, Pages1, Pages2, RowsPerPage1, RowsPerPage2 : integer;
   PageBreak1, PageBreak2, FirstPage1, FirstPage2       : boolean;
   DoLine                                               : boolean;
   ThisAmount, ThisTotal, ThisSub, FullTotal            : double;
   ThisFill, ThisText, ConsFill, ConsText               : TColor;
   ThisFile, CurrFile, DispFile, FeeType, FeeUnique     : string;
   xlsBook                                              : IXLSWorkbook;
   xlsSheet1, xlsSheet2                                 : IXLSWorksheet;

begin

   if (AccountType = 1) then
      FeeType := 'Taxation Account'
   else
      FeeType := 'Client Account';

   ThisCount := 1;
   ThisMax   := IntToStr(NumFiles);

   prbProgress.Max := NumFiles;
   LogMsg(ord(PT_PROLOG),True,True,'Fee Earner Report');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Set defaults then create a new Excel Workbook

   FilesFound := 0;
   Sema1      := 0;
   Sema2      := 0;
   FullTotal  := 0;

//--- Create the Excel Workbook

   xlsBook       := TXLSWorkbook.Create;

//--- Create the Consolidation Sheet

   xlsSheet2      := xlsBook.WorkSheets.Add;
   xlsSheet2.Name := 'Consolidation';

   Pages2     := 0;
   FirstPage2 := true;
   PageBreak2 := true;

//-- Build the file name

   ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - Fee Earner Report (' + CpyName + ').xls';
   txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;

//--- Read the Fee Earner Records from the datastore

   for idx1 := 0 to NumFiles - 1 do begin

      xlsSheet1      := xlsBook.WorkSheets.Add;
      xlsSheet1.Name := FileArray[idx1];

      ThisTotal := 0;

//--- Get the detail for the current Fee Earner

      if (GetUser(FileArray[idx1]) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      FeeEarnerName   := Query2.FieldByName('Control_Name').AsString;
      FeeEarnerEmail  := Query2.FieldByName('Control_Email').AsString;
      FeeUnique       := Query2.FieldByName('Control_Unique').AsString;

      txtError.Text := 'Processing: ' + FileArray[idx1];
      txtError.Refresh;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

//--- Get all the Fee records for the current Fee Earner

      if ((GetFeeRecs(FileArray[idx1],' AND B_FeeEarner = ' + FeeUnique + ' AND B_AccountType = ' + IntTostr(AccountType) + ' AND ((B_Class >= 0 AND B_Class <= 2) OR (B_Class = 5) OR (B_Class = 18) OR (B_CLASS = 21))')) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

//--- Only process files that have records

      if (Query1.RecordCount > 0) then begin

         Pages1     := 0;
         FirstPage1 := true;
         PageBreak1 := true;

         while Query1.Eof = false do begin

//--- Set this block's colour

            if (Sema1 = 0) then begin
               Sema1 := 1;
               ThisFill := ColAB1F;
               ThisText := ColAB1T;
            end else begin
               Sema1 := 0;
               ThisFill := ColAB2F;
               ThisText := ColAB2T;
            end;

            ThisSub := 0;
            CurrFile := Query1.FieldByName('B_Owner').AsString;
            DispFile := GetFileDesc(CurrFile);
            inc(FilesFound);

            while CurrFile = Query1.FieldByName('B_Owner').AsString do begin

               if (Query1.Eof = true) then
                  break;

               if (PageBreak1 = True) then
                  DoPageBreak(ord(PB_FEEEARNERS),xlsSheet1,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1,FileArray[idx1]);

               if (Query1.FieldByName('B_Class').AsInteger in [0..2,18]) then
                  ThisAmount := Query1.FieldByName('B_Amount').AsFloat
               else
                  ThisAmount := (Query1.FieldByName('B_Amount').AsFloat) * -1;

               with xlsSheet1.RCRange[row1,1,row1,4] do begin
                  Item[1,1].Value := DispFile;
                  Item[1,1].Borders[xlAround].Weight := xlThin;
                  Item[1,2].Value := Query1.FieldByName('B_Date').AsString;
                  Item[1,2].Borders[xlAround].Weight := xlThin;
                  Item[1,3].Value := ReplaceQuote(Query1.FieldByName('B_Description').AsString);
                  Item[1,3].Borders[xlAround].Weight := xlThin;
                  Item[1,4].Value := ThisAmount;
                  Item[1,4].Borders[xlAround].Weight := xlThin;
                  Item[1,4].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ThisFill;
                  Font.Color     := ThisText;
                  Font.Bold      := false;
                  Font.Name      := 'Arial';
                  Font.Size      := 10;
               end;

               DispFile := ' ';

               ThisSub := ThisSub + ThisAmount;
               Query1.Next;

               inc(row1);
               inc(PageRow1);

//--- If we've reached the maximum rows per page then it is PageBreak time.

               if (PageRow1 >= RowsPerPage1) then begin
                  PageBreak1 := True;
                  DoPageBreak(ord(PB_FEEEARNERS),xlsSheet1,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1,FileArray[idx1]);
               end;

            end;

//--- Write the Sub-Total

            with xlsSheet1.RCRange[row1,1,row1,4] do begin
               Borders[xlAround].Weight := xlThin;
               Item[1,1].Value := DispFile;
               Item[1,1].Borders[xlAround].Weight := xlThin;
               Item[1,2].Value := EDate;
               Item[1,2].Borders[xlAround].Weight := xlThin;
               Item[1,3].Value := 'Sub-Total';
               Item[1,3].Borders[xlAround].Weight := xlThin;
               Item[1,4].Value := ThisSub;
               Item[1,4].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
               Item[1,4].Borders[xlAround].Weight := xlThin;
               Interior.Color := ThisFill;
               Font.Color := ThisText;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
               Item[1,4].Font.Bold := true;
            end;

            ThisTotal := ThisTotal + ThisSub;

            inc(row1);
            inc(PageRow1);

//--- If we've reached the maximum rows per page then it is PageBreak time.

            if (PageRow1 >= RowsPerPage1) then begin
               PageBreak1 := True;
               DoPageBreak(ord(PB_FEEEARNERS),xlsSheet1,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1,FileArray[idx1]);
            end;

         end;
      end else begin

         Pages1     := 0;
         FirstPage1 := true;
         PageBreak1 := true;

         if (PageBreak1 = True) then
            DoPageBreak(ord(PB_FEEEARNERS),xlsSheet1,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1,FileArray[idx1]);

         with xlsSheet1.RCRange[row1,1,row1,4] do begin
            Item[1,1].Value := 'No billing records for period: ' + SDate + ' to ' + Edate;
            Item[1,1].Font.Bold := false;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := ColAB1F;
            Font.Color := ColAB1T;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;

         inc(row1);
         inc(PageRow1);

      end;

//--- Write the Total Line

      with xlsSheet1.RCRange[row1,1,row1,4] do begin
         Item[1,1].Value := 'Total Fees, Disbursements and Expenses:';
         Item[1,3].Borders[xlAround].Weight := xlThin;
         Item[1,4].Value := ThisTotal;
         Item[1,4].Borders[xlAround].Weight := xlThin;
         Item[1,4].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;

      inc(row1);
      inc(PageRow1);

//--- If we've reached the maximum rows per page then it is PageBreak time.

      if (PageRow1 >= RowsPerPage1) then begin
         PageBreak1 := True;
         DoPageBreak(ord(PB_FEEEARNERS),xlsSheet1,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1,FileArray[idx1]);
      end;

//--- Write the standard copyright notice

      row1 := (Pages1 * lcGRows);
      with xlsSheet1.RCRange[row1,1,row1,1] do begin
         Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 8;
      end;

//--- Remove the Gridlines which are added by default and set the Page orientation

      xlsSheet1.PageSetup.Orientation := xlLandscape;
      xlsSheet1.PageSetup.FitToPagesWide := 1;
      xlsSheet1.PageSetup.FitToPagesTall := Pages1;
      xlsSheet1.PageSetup.PaperSize := xlPaperA4;
      xlsSheet1.DisplayGridLines := false;
      xlsSheet1.PageSetup.CenterFooter := 'Page &P of &N';

//--- Set this Consolidation Line's colour

      if (Sema2 = 0) then begin
         Sema2 := 1;
         ConsFill := ColAB1F;
         ConsText := ColAB1T;
      end else begin
         Sema2 := 0;
         ConsFill := ColAB2F;
         ConsText := ColAB2T;
      end;

//--- Page Handling for the Consolidation Sheet

      if (PageBreak2 = True) then
         DoPageBreak(ord(PB_FEECONS),xlsSheet2,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2,'');

//--- Write the consolidation line

      with xlsSheet2.RCRange[row2,1,row2,3] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := FileArray[idx1];
         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,2].Value := FeeEarnerName;
         Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,3].Value := ThisTotal;
         Item[1,3].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
         Interior.Color := ConsFill;
         Font.Color := ConsText;
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;

      FullTotal := FullTotal + ThisTotal;

      inc(row2);
      inc(PageRow2);

//--- If we've reached the maximum rows per page then it is PageBreak time.

      if (PageRow2 >= RowsPerPage2) then begin
         PageBreak2 := True;
         DoPageBreak(ord(PB_FEECONS),xlsSheet2,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2,'');
      end;

   end;

//--- Write the Grand Total

   with xlsSheet2.RCRange[row2,1,row2,3] do begin
      Borders[xlAround].Weight := xlThin;
      Item[1,1].Value := 'Total Fees, Disbursements and Expenses for Period: ' + SDate + ' to ' + EDate;
      Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,3].Value := FullTotal;
      Item[1,3].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Interior.Color := ColDHF;
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   inc(row2);
   inc(PageRow2);

//--- If we've reached the maximum rows per page then it is PageBreak time.

   if (PageRow2 >= RowsPerPage2) then begin
      PageBreak2 := True;
      DoPageBreak(ord(PB_FEECONS),xlsSheet2,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2,'');
   end;

   row2 := (Pages2 * lcGrows);
   with xlsSheet2.RCRange[row2,1,row2,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet2.PageSetup.Orientation := xlLandscape;
   xlsSheet2.PageSetup.FitToPagesWide := 1;
   xlsSheet2.PageSetup.FitToPagesTall := Pages2;
   xlsSheet2.PageSetup.PaperSize := xlPaperA4;
   xlsSheet2.DisplayGridLines := false;
   xlsSheet2.PageSetup.CenterFooter := 'Page &P of &N';

//--- Write the Excel file to disk

   if (FilesFound > 0) then begin
      xlsBook.SaveAs(FileName + ThisFile);
      xlsBook.Close;

      LogMsg('  Fee Earner Report successfully processed...',True);
      LogMsg(' ',True);

      DoLine := False;

//--- Print the generated document on the Default Printer if requested

      if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
         if (PrintDocument(FileName + ThisFile) = True) then
            LogMsg('  Document submitted for printing...',True)
         else
            LogMsg('  Printing of document failed...',True);

         DoLine := True;
      end;

//--- Create a PDF file if requested

      if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
         PDFExists := PDFDocument(FileName + ThisFile);

         if (PDFExists = True) then
            LogMsg('  PDF file creation was successfull...',True)
         else
            LogMsg('  PDF file creation failed...',True);

         DoLine := True;
      end;

//--- Send the Excel file via email if requested

      if (SendByEmail = '1') then begin
         if (SendEmail('',FileName + ThisFile,ord(PT_NORMAL)) = true) then
            LogMsg('  Request to send generated Fee Earner Report by Email submitted...',True)
         else
            LogMsg('  Request to send generated Fee Earner Report by Email not submitted...',True);

         DoLine := True;
      end;

//--- Now open the Fee Earner Report if requested

      if (AutoOpen = True) then begin
         ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);
         LogMsg('  Request to open Fee Earner Report for ''' + PChar(FileName + ThisFile) + ''' submitted...',True);
         DoLine := True;
      end;
   end else begin
      LogMsg('  No Fee Earner Records found ...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Fee Earner Report');

   Close_Connection;

end;

//---------------------------------------------------------------------------
// Generate the Accountant Report
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_AcctReport();
var
   row1, row2, row3, row4, idx1, idx2, idx3   : integer;
   PageRow1, PageRow2, PageRow3, Pages1       : integer;
   Pages2, Pages3, RowsPerPage1, RowsPerPage2 : integer;
   RowsPerPage3, FilesFound, InvFound, Sema   : integer;
   ThisFill, ThisText                         : TColor;
   PageBreak1, PageBreak2, PageBreak3         : boolean;
   FirstPage1, FirstPage2, FirstPage3         : boolean;
   DoLine                                     : boolean;
   ThisFees, ThisTrust, ThisS864, ThisS863    : double;
   OpenFees, OpenTrust, OpenS864, OpenS863    : double;
   InFees, InTrust, InS864, InS863            : double;
   OutFees, OutTrust, OutS864, OutS863        : double;
   Paid, Amount, Fees, Disburse, Expenses     : double;
   ThisAmount, ThisDisburse, ThisExpenses     : double;
   ThisFile, ThisClient                       : string;
   xlsBook                                    : IXLSWorkbook;
   xlsSheet1, xlsSheet2, xlsSheet3, xlsSheet4 : IXLSWorksheet;

begin

   ThisCount := 1;
   ThisMax   := IntToStr(NumFiles);

   prbProgress.Max := NumFiles;
   LogMsg(ord(PT_PROLOG),True,True,'Accountant Report');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Set defaults then create a new Excel Workbook

   FilesFound := 0;

//--- Create the Excel Workbook

   xlsBook        := TXLSWorkbook.Create;
   xlsSheet1      := xlsBook.WorkSheets.Add;
   xlsSheet2      := xlsBook.WorkSheets.Add;
   xlsSheet3      := xlsBook.WorkSheets.Add;
   xlsSheet1.Name := 'Accountant Report (01)';
   xlsSheet2.Name := 'Accountant Report (02)';
   xlsSheet3.Name := 'Accountant Report (03)';

//-- Build the file name

   ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - Accountant Report (' + CpyName + ').xls';
   txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;

//--- Set up to Insert the Header (1st and 2nd line) and the Heading for each
//--- worksheet

   Pages1     := 0;    Pages2     := 0;    Pages3     := 0;
   FirstPage1 := True; FirstPage2 := True; FirstPage3 := True;
   PageBreak1 := True; PageBreak2 := True; PageBreak3 := True;

//--- Gather and process the billing records

   for idx1 := 0 to NumFiles - 1 do begin

//--- Get the opening balance for each category for this file

      if (GetBalance(FileArray[idx1]) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      OpenFees  := (This_Abs.Fees + This_Abs.Disbursements + This_Abs.Expenses + This_Abs.Business_To_Trust + This_Abs.Business_Debit) - (This_Abs.Credit + This_Abs.Write_off + This_Abs.Payment_Received + This_Abs.Business_Deposit + This_Abs.Trust_Transfer_Business_Fees + This_Abs.Trust_Transfer_Business_Other);
      OpenTrust := (This_Abs.Trust_Deposit + This_Abs.Business_To_Trust + This_Abs.Trust_Withdrawal_S86_4) - (This_Abs.Trust_Transfer_Business_Fees + This_Abs.Trust_Transfer_Business_Other + This_Abs.Trust_Transfer_Client + This_Abs.Trust_Transfer_Disbursements + This_Abs.Trust_Transfer_Trust + This_Abs.Trust_Investment_S86_4 + This_Abs.Trust_Debit);
      OpenS864 := (This_Abs.Trust_Investment_S86_4 + This_Abs.Trust_Interest_S86_4) - This_Abs.Trust_Withdrawal_S86_4;
      OpenS863 := (This_Abs.Trust_Investment_S86_3 + This_Abs.Trust_Interest_S86_3) - (This_Abs.Trust_Withdrawal_S86_3 + This_Abs.Trust_Interest_Withdrawal_S86_3);

//--- Get the current period balances for each category for this file

      if (GetCurrent(FileArray[idx1]) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

//--- Calculate the opening balance and the current period's balance for the
//--- current file

      txtError.Text := 'Processing: ' + FileArray[idx1];
      txtError.Refresh;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

      InFees  := This_Abs.Fees + This_Abs.Disbursements + This_Abs.Expenses + This_Abs.Business_To_Trust + This_Abs.Business_Debit;
      InTrust := This_Abs.Trust_Deposit + This_Abs.Business_To_Trust + This_Abs.Trust_Withdrawal_S86_4;
      InS864   := This_Abs.Trust_Investment_S86_4 + This_Abs.Trust_Interest_S86_4;
      InS863 := This_Abs.Trust_Investment_S86_3 + This_Abs.Trust_Interest_S86_3;

      OutFees  := (This_Abs.Credit + This_Abs.Write_off + This_Abs.Payment_Received + This_Abs.Business_Deposit + This_Abs.Trust_Transfer_Business_Fees + This_Abs.Trust_Transfer_Business_Other) * -1;
      OutTrust := (This_Abs.Trust_Transfer_Business_Fees + This_Abs.Trust_Transfer_Business_Other + This_Abs.Trust_Transfer_Client + This_Abs.Trust_Transfer_Disbursements + This_Abs.Trust_Transfer_Trust + This_Abs.Trust_Investment_S86_4 + This_Abs.Trust_Debit) * -1;
      OutS864  := This_Abs.Trust_Withdrawal_S86_4 * -1;
      OutS863  := (This_Abs.Trust_Withdrawal_S86_3 + This_Abs.Trust_Interest_Withdrawal_S86_3) * -1;

      ThisFees  := OpenFees  + InFees  + OutFees;
      ThisTrust := OpenTrust + InTrust + OutTrust;
      ThisS864  := OpenS864  + InS864  + OutS864;
      ThisS863  := OpenS863  + InS863 +  OutS863;

//--- Only process files that have records

      if ((Round(ThisFees) = 0) and (Round(ThisTrust) = 0) and (Round(ThisS864) = 0) and (Round(ThisS863) = 0)) then
         continue;

//--- Set this block's colour

      if (Sema = 0) then begin
         Sema := 1;
         ThisFill := ColAB1F;
         ThisText := ColAB1T;
      end else begin
         Sema := 0;
         ThisFill := ColAB2F;
         ThisText := ColAB2T;
      end;

      inc(FilesFound);

//--- Perform a Page break on Report 01 if necessary

      if (PageBreak1 = True) then
         DoPageBreak(ord(PB_ACCOUNTANT01),xlsSheet1,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1,'');

//--- Report 01 - Write the fixed data

      if (FileArray[idx1] = CpyFile) then begin

{
         if (EDate < S86Date) then begin
            AcctReport_Line1(xlsSheet1,FileArray[idx1],'Section 78(2)(a) Investments',This_Bal.Trust_Investment_S86_3,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
            AcctReport_Line1(xlsSheet1,' ','Section 78(2)(a) Withdrawals',(This_Bal.Trust_Withdrawal_S86_3 + This_Bal.Trust_Interest_Withdrawal_S86_3),'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
            AcctReport_Line1(xlsSheet1,' ','Section 78(2)(a) Interest',This_Bal.Trust_Interest_S86_3,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         end else begin
}
            AcctReport_Line1(xlsSheet1,FileArray[idx1],'Section 86(3) Investments',This_Bal.Trust_Investment_S86_3,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
            AcctReport_Line1(xlsSheet1,' ','Section 86(3) Withdrawals',(This_Bal.Trust_Withdrawal_S86_3 + This_Bal.Trust_Interest_Withdrawal_S86_3),'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
            AcctReport_Line1(xlsSheet1,' ','Section 86(3) Interest',This_Bal.Trust_Interest_S86_3,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
{
         end;
}

      end else begin

         AcctReport_Line1(xlsSheet1,FileArray[idx1],'Fees',This_Bal.Fees,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Disbursements',This_Bal.Disbursements,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Expenses',This_Bal.Expenses,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Payments Received',This_Bal.Payment_Received,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Business To Trust Transfer',This_Bal.Business_To_Trust,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Credit Allowed',(This_Bal.Credit + This_Bal.Write_off),'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Deposit to Business Account',This_Bal.Business_Deposit,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Deposits To Trust Account',This_Bal.Trust_Deposit,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Transfers from Trust for Fees',(This_Bal.Trust_Transfer_Business_Fees + This_Bal.Trust_Transfer_Business_Other),'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Transfer from Trust for Disbursements',This_Bal.Trust_Transfer_Disbursements,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Transfer from Trust to Client',This_Bal.Trust_Transfer_Client,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Transfer from Trust to Trust',This_Bal.Trust_Transfer_Trust,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
{
         if (EDate < S86Date) then begin
            AcctReport_Line1(xlsSheet1,' ','Section 78(2A) Investments',This_Bal.Trust_Investment_S86_4,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
            AcctReport_Line1(xlsSheet1,' ','Section 78(2A) Withdrawals',This_Bal.Trust_Withdrawal_S86_4,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
            AcctReport_Line1(xlsSheet1,' ','Section 78(2A) Interest',This_Bal.Trust_Interest_S86_4,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         end else begin
}
            AcctReport_Line1(xlsSheet1,' ','Section 86(4) Investments',This_Bal.Trust_Investment_S86_4,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
            AcctReport_Line1(xlsSheet1,' ','Section 86(4) Withdrawals',This_Bal.Trust_Withdrawal_S86_4,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
            AcctReport_Line1(xlsSheet1,' ','Section 86(4) Interest',This_Bal.Trust_Interest_S86_4,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
{
         end;
}
         AcctReport_Line1(xlsSheet1,' ','Business Debits',This_Bal.Business_Debit,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);
         AcctReport_Line1(xlsSheet1,' ','Sundry Trust Debits',This_Bal.Trust_Debit,'',ThisFill,ThisText,FirstPage1,RowsPerPage1,row1,Pages1,PageRow1,PageBreak1);

      end;

//--- Perform a Page break on Report 02 if necessary

      if (PageBreak2 = True) then
         DoPageBreak(ord(PB_ACCOUNTANT02),xlsSheet2,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2,'');

//--- Report 02 - Write the fixed data

      if (FileArray[idx1] = CpyFile) then begin

         AcctReport_Line2(1,xlsSheet2,FileArray[idx1],'Section 86(3) Investments',True,0,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(2,xlsSheet2,' ','Opening Balance',False,OpenS863,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(3,xlsSheet2,' ','Inflows',False,InS863,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(3,xlsSheet2,' ','Outflows',False,OutS863,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(2,xlsSheet2,' ','Closing Balance',False,ThisS863,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);

      end else begin

         AcctReport_Line2(1,xlsSheet2,FileArray[idx1],'Business Account',True,0,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(2,xlsSheet2,' ','Opening Balance',False,OpenFees,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(3,xlsSheet2,' ','Debits',False,InFees,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(3,xlsSheet2,' ','Credits',False,OutFees,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(2,xlsSheet2,' ','Closing Balance',False,ThisFees,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);

         AcctReport_Line2(1,xlsSheet2,' ','Trust Account',True,0,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(2,xlsSheet2,' ','Opening Balance',False,OpenTrust,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(3,xlsSheet2,' ','Inflows',False,InTrust,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(3,xlsSheet2,' ','Outflows',False,OutTrust,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(2,xlsSheet2,' ','Closing Balance',False,ThisTrust,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);

{
         if (EDate < S86Date) then
            AcctReport_Line2(1,xlsSheet2,' ','S78(2A) Investments',True,0,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2)
         else
}
            AcctReport_Line2(1,xlsSheet2,' ','S86(4) Investments',True,0,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);

         AcctReport_Line2(2,xlsSheet2,' ','Opening Balance',False,OpenS864,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(3,xlsSheet2,' ','Inflows',False,InS864,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(3,xlsSheet2,' ','Outflows',False,OutS864,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);
         AcctReport_Line2(2,xlsSheet2,' ','Closing Balance',False,ThisS864,'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);

         AcctReport_Line2(4,xlsSheet2,FileArray[idx1],'Nett Balance',True,(ThisTrust + ThisS864 - ThisFees),'',ThisFill,ThisText,FirstPage2,RowsPerPage2,row2,Pages2,PageRow2,PageBreak2);

      end;
   end;

//--- Do Report 03 - Start by getting all the invoices and payments for the period

   GetInvoices(ord(DT_ACCOUNTANT),'');

   InvFound  := 0;
   Sema      := 0;

   ThisCount := 1;
   ThisMax   := IntToStr(Query1.RecordCount);

   prbProgress.Max      := Query1.RecordCount;
   prbProgress.Position := 1;

   for idx3 := 0 to Query1.RecordCount - 1 do begin

//--- Process the record

      txtError.Text := 'Now processing invoice: ' + Query1.FieldByName('Inv_Invoice').AsString;
      txtError.Repaint;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

//--- Perform a Page break on Report 02 if necessary

      if (PageBreak3 = True) then
         DoPageBreak(ord(PB_ACCOUNTANT03),xlsSheet3,FirstPage3,RowsPerPage3,row3,Pages3,PageRow3,PageBreak3,'');

//--- Get the payments that were made in this period

      Paid := GetPayments(Query1.FieldByName('Inv_Invoice').AsString);

//--- Exclude this record if no payments were made in this period

      if (Paid = 0) then begin
         Query1.Next;
         continue;
      end;

//--- Allocate payments in order as follows: Disbursements, Expenses then Fees

      Amount   := StrToFloat(Query1.FieldByName('Inv_Amount').AsString);
      Fees     := StrToFloat(Query1.FieldByName('Inv_Fees').AsString);
      Disburse := StrToFloat(Query1.FieldByName('Inv_Disburse').AsString);
      Expenses := StrToFloat(Query1.FieldByName('Inv_Expenses').AsString);

      if (Paid >= Amount) then begin
         ThisFees     := Fees;
         ThisDisburse := Disburse;
         ThisExpenses := Expenses;
      end else begin
         ThisAmount := Paid;

         if (ThisAmount > Disburse) then begin
            ThisDisburse := Disburse;
            ThisAmount   := ThisAmount - Disburse;
         end else begin
            ThisDisburse := ThisAmount;
            ThisAmount   := 0;
         end;

         if (ThisAmount > Expenses) then begin
            ThisExpenses := Expenses;
            ThisAmount   := ThisAmount - Expenses;
         end else begin
            ThisExpenses := ThisAmount;
            ThisAmount   := 0;
         end;

         if (ThisAmount > Fees) then
            ThisFees := Fees
         else
            ThisFees := ThisAmount;
      end;

//--- Set this block's colour

      if (Sema = 0) then begin
         Sema := 1;
         ThisFill := ColAB1F;
         ThisText := ColAB1T;
      end else begin
         Sema := 0;
         ThisFill := ColAB2F;
         ThisText := ColAB2T;
      end;

      inc(InvFound);

//--- Write the Line

      ThisClient := GetClient(Query1.FieldByName('Inv_File').AsString);

      with xlsSheet3.RCRange[row3,1,row3 + 1,10] do begin
         Item[1, 1].Value := Query1.FieldByName('Inv_Invoice').AsString;
         Item[1, 1].Borders[xlEdgeRight].Weight := xlThin;
         Item[2, 1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1, 2].Value := Query1.FieldByName('Inv_EDate').AsString;
         Item[1, 2].Borders[xlEdgeRight].Weight := xlThin;
         Item[2, 2].Borders[xlEdgeRight].Weight := xlThin;
         Item[1, 3].Value := Query1.FieldByName('Inv_File').AsString;
         Item[1, 3].Borders[xlEdgeRight].Weight := xlThin;
         Item[2, 3].Borders[xlEdgeRight].Weight := xlThin;
         Item[1, 4].Value := ThisClient;
         Item[1, 4].Borders[xlEdgeRight].Weight := xlThin;
         Item[2, 4].Borders[xlEdgeRight].Weight := xlThin;
         Item[1, 5].Value := 'Amount:';
         Item[1, 5].Borders[xlEdgeRight].Weight := xlThin;
         Item[1, 5].Borders[xlEdgeBottom].Weight := xlThin;
         Item[2, 5].Value := 'Paid:';
         Item[2, 5].Borders[xlEdgeLeft].Weight := xlThin;
         Item[2, 5].Borders[xlEdgeRight].Weight := xlThin;
         Item[1, 6].Value := Amount;
         Item[1, 6].Borders[xlEdgeRight].Weight := xlThin;
         Item[1, 6].HorizontalAlignment := xlHAlignRight;
         Item[1, 6].Borders[xlEdgeBottom].Weight := xlThin;
         Item[2, 6].Borders[xlEdgeRight].Weight := xlThin;
         Item[2, 6].Value := Paid;
         Item[2, 6].Borders[xlEdgeRight].Weight := xlThin;
         Item[2, 6].HorizontalAlignment := xlHAlignRight;
         Item[1, 7].Value := Fees;
         Item[1, 7].Borders[xlEdgeRight].Weight := xlThin;
         Item[1, 7].Borders[xlEdgeBottom].Weight := xlThin;
         Item[1, 7].HorizontalAlignment := xlHAlignRight;
         Item[2, 7].Value := ThisFees;
         Item[2, 7].Borders[xlEdgeRight].Weight := xlThin;
         Item[2, 7].HorizontalAlignment := xlHAlignRight;
         Item[1, 8].Value := Disburse;
         Item[1, 8].Borders[xlEdgeRight].Weight := xlThin;
         Item[1, 8].Borders[xlEdgeBottom].Weight := xlThin;
         Item[1, 8].HorizontalAlignment := xlHAlignRight;
         Item[2, 8].Value := ThisDisburse;
         Item[2, 8].Borders[xlEdgeRight].Weight := xlThin;
         Item[2, 8].HorizontalAlignment := xlHAlignRight;
         Item[1, 9].Value := Expenses;
         Item[1, 9].Borders[xlEdgeRight].Weight := xlThin;
         Item[1, 9].Borders[xlEdgeBottom].Weight := xlThin;
         Item[1, 9].HorizontalAlignment := xlHAlignRight;
         Item[2, 9].Value := ThisExpenses;
         Item[2, 9].Borders[xlEdgeRight].Weight := xlThin;
         Item[2, 9].HorizontalAlignment := xlHAlignRight;
         Item[1,10].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,10].Borders[xlEdgeBottom].Weight := xlThin;
         Item[1,10].HorizontalAlignment := xlHAlignRight;

         if (ShowVAT = 1) then
            Item[2,10].Value := (ThisFees / (1 + (VATRate / 100)))
         else
            Item[2,10].Value := ThisFees;

         Item[2,10].Borders[xlEdgeRight].Weight := xlThin;
         Item[2,10].HorizontalAlignment := xlHAlignRight;
         Borders[xlAround].Weight := xlThin;
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
         Interior.Color := ThisFill;
         Font.Color := ThisText;
      end;

      with xlsSheet3.RCRange[row3,6,row3 + 1,10] do begin
         NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      end;

      inc(row3);
      inc(row3);
      inc(PageRow3);

//--- If we've reached the maximum rows per page then it is PageBreak time

      if (PageRow3 >= RowsPerPage3) then begin
         PageBreak3 := True;
         DoPageBreak(ord(PB_ACCOUNTANT03),xlsSheet3,FirstPage3,RowsPerPage3,row3,Pages3,PageRow3,PageBreak3,'');
      end;

      Query1.Next;
   end;

//--- Check for Instances where a heading was printed but no data

   if (PageRow3 = 1) then begin
      if (lcRepeatHeader = True) then begin
         row3 := row3 - 4;
         xlsSheet3.RCRange[row3,1,row3 + 2,10].Delete(xlShiftToLeft);
      end else begin
         row3 := row3 - 2;
         xlsSheet3.RCRange[row3,1,row3,10].Delete(xlShiftToLeft);
      end;
   end;

//--- Write the copyright notices for each of the sheets

   row1 := (Pages1 * lcGRows);
   with xlsSheet1.RCRange[row1,1,row1,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

   row2 := (Pages2 * lcGRows);
   with xlsSheet2.RCRange[row2,1,row2,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

   row3 := (Pages3 * lcGRows);
   with xlsSheet3.RCRange[row3,1,row3,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet1.PageSetup.Orientation    := xlLandscape;
   xlsSheet1.PageSetup.FitToPagesWide := 1;
   xlsSheet1.PageSetup.FitToPagesTall := Pages1;
   xlsSheet1.PageSetup.PaperSize      := xlPaperA4;
   xlsSheet1.DisplayGridLines         := false;
   XlsSheet1.PageSetup.CenterFooter    := 'Page &P of &N';

   xlsSheet2.PageSetup.Orientation    := xlLandscape;
   xlsSheet2.PageSetup.FitToPagesWide := 1;
   xlsSheet2.PageSetup.FitToPagesTall := Pages2;
   xlsSheet2.PageSetup.PaperSize      := xlPaperA4;
   xlsSheet2.DisplayGridLines         := false;
   XlsSheet2.PageSetup.CenterFooter    := 'Page &P of &N';

   xlsSheet3.PageSetup.Orientation    := xlLandscape;
   xlsSheet3.PageSetup.FitToPagesWide := 1;
   xlsSheet3.PageSetup.FitToPagesTall := Pages3;
   xlsSheet3.PageSetup.PaperSize      := xlPaperA4;
   xlsSheet3.DisplayGridLines         := false;
   XlsSheet3.PageSetup.CenterFooter    := 'Page &P of &N';

//--- Write the Excel file to disk. If no files were processed then delete
//--- sheets 1 and 2. If no Invoices then delete sheet 3.

   if ((FilesFound > 0) or (InvFound > 0)) then begin

      if (FilesFound = 0) then begin
         xlsSheet1.Delete;
         xlsSheet2.Delete;
      end;

      if (InvFound = 0) then
         xlsSheet3.Delete;

      xlsBook.SaveAs(FileName + ThisFile);
      xlsBook.Close;

      LogMsg('  Accountant Report successfully processed...',True);
      LogMsg(' ',True);

      DoLine := False;

//--- Print the generated document on the Default Printer if requested

      if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
         if (PrintDocument(FileName + ThisFile) = True) then
            LogMsg('  Document submitted for printing...',True)
         else
            LogMsg('  Printing of document failed...',True);

         DoLine := True;
      end;

//--- Create a PDF file if requested

      if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
         PDFExists := PDFDocument(FileName + ThisFile);

         if (PDFExists = True) then
            LogMsg('  PDF file creation was successfull...',True)
         else
            LogMsg('  PDF file creation failed...',True);

         DoLine := True;
      end;

//--- Send the Excel file via email if requested

      if (SendByEmail = '1') then begin
         if (SendEmail('',FileName + ThisFile,ord(PT_NORMAL)) = true) then
            LogMsg('  Request to send generated Accountant Report by Email submitted...',True)
         else
            LogMsg('  Request to send generated Accountant Report by Email not submitted...',True);

         DoLine := True;
      end;

//--- Now open the Accountant Report if requested

      if (AutoOpen = True) then begin
         ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);
         LogMsg('  Request to open Accountant Report for ''' + PChar(FileName + ThisFile) + ''' submitted...',True);
         DoLine := True;
      end;
   end else begin
      LogMsg('  No Billing Records found ...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Accountant Report');

   Close_Connection;

end;

//---------------------------------------------------------------------------
// Procedure to add a line to the Accountant Report 01
//---------------------------------------------------------------------------
procedure TFldExcel.AcctReport_Line1(xlsSheet: IXLSWorksheet; ThisFile: string; ThisStr: string; ThisAmt: double; ThisExtra : string; ThisFill: TColor; ThisText: TColor; var FirstPage: boolean; var RowsPerPage: integer; var row: integer; var Pages: integer; var PageRow: integer; var PageBreak: boolean);
begin

   with xlsSheet.RCRange[row,1,row,3] do begin
      Item[1,1].Value := GetFileDesc(ThisFile);
      Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,2].Value := ThisStr;
      Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,3].Value := ThisAmt;
      Item[1,3].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Interior.Color := ThisFill;
      Font.Color := ThisText;
      Font.Bold := False;
      Font.Name := 'Arial';
      Font.Size := 10;
      Borders[xlAround].Weight := xlThin;
   end;

   inc(row);
   inc(PageRow);

//--- If we've reached the maximum rows per page then it is PageBreak time

   if (PageRow >= RowsPerPage) then begin
      PageBreak := True;
      DoPageBreak(ord(PB_ACCOUNTANT01),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,ThisExtra);
   end;

end;

//---------------------------------------------------------------------------
// Procedure to add a line to the Account Report 02
//---------------------------------------------------------------------------
procedure TFldExcel.AcctReport_Line2(ThisVariant: integer; xlsSheet: IXLSWorksheet; ThisFile: string; ThisStr: string; ThisBold: boolean; ThisAmt: double; ThisExtra : string;  ThisFill: TColor; ThisText: TColor; var FirstPage: boolean; var RowsPerPage: integer; var row: integer; var Pages: integer; var PageRow: integer; var PageBreak: boolean);
begin

   with xlsSheet.RCRange[row,1,row,5] do begin
      if (ThisVariant in [1]) then
         Item[1,1].Value := GetFileDesc(ThisFile);

      if (ThisVariant in [2..4]) then
         Item[1,1].Value := ' ';

      if (ThisVariant in [1..4]) then
         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;

      if (ThisVariant in [1,4]) then
         Item[1,2].Value := ThisStr;

      if (ThisVariant in [2]) then
         Item[1,3].Value := ThisStr;

      if (ThisVariant in [3]) then
         Item[1,4].Value := ThisStr;

      if (ThisVariant in [2..4]) then begin
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].Value := ThisAmt;
            Item[1,5].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      end;

      if (ThisVariant in [1..4]) then begin
         Borders[xlAround].Weight := xlThin;
         Font.Bold := ThisBold;
         Interior.Color := ThisFill;
         Font.Color := ThisText;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;

   end;

   inc(row);
   inc(PageRow);

//--- If we've reached the maximum rows per page then it is PageBreak time

   if (PageRow >= RowsPerPage) then begin
      PageBreak := True;
      DoPageBreak(ord(PB_ACCOUNTANT02),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,ThisExtra);
   end;
end;

//---------------------------------------------------------------------------
// Generate the Alerts Report
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Alerts();
var
   idx1, row, PageRow, Pages, RowsPerPage, FilesFound   : integer;
   InkCol                                               : integer;
   PageBreak, FirstPage, DoLine                         : boolean;
   Days                                                 : double;
   ThisFile, ThisType, ThisReason, ThisFilter           : string;
   Date1, Date2                                         : TDateTime;
   xlsBook                                              : IXLSWorkbook;
   xlsSheet                                             : IXLSWorksheet;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Alerts Report');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);
      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();
   FilesFound := 0;
   ThisFilter := Parm10;

//--- Create the Excel Workbook

   xlsBook       := TXLSWorkbook.Create;

//-- Build the file name

   ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - Alerts Report (' + CpyName + ').xls';
   txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;

//--- Process each User record in turn

   for idx1 := 0 to NumFiles - 1 do begin

//--- Automatically exclude the "Backup Administrator"

      if (FileArray[idx1] = 'Backup Administrator') then
         continue;

//--- Get all the Alert records for the current User

      if (GetAlertRecs(FileArray[idx1], ThisFilter) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);
         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      xlsSheet      := xlsBook.WorkSheets.Add;
      xlsSheet.Name := FileArray[idx1];

      Pages     := 0;
      FirstPage := true;
      PageBreak := true;

//--- Only process files that have records

      if (Query1.RecordCount > 0) then begin

         inc(FilesFound);

         while Query1.Eof = false do begin

//--- Perform a Page break if necessary

            if (PageBreak = True) then
               DoPageBreak(ord(PB_ALERTS),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,FileArray[idx1]);

            InkCol     := clBlack;
            ThisType   := ReplaceQuote(Query1.FieldByName('Tracking_Type').AsString);
            ThisReason := ReplaceQuote(Query1.FieldByName('Tracking_Reason').AsString);

            if (ThisReason = '') then
               ThisReason := ' ';

            if (ThisType = 'Prescription') then begin
               ShortDateFormat := 'yyyy/MM/dd';
               DateSeparator   := '/';
               Date1 := StrToDate(Query1.FieldByName('Tracking_Order').AsString);
               Date2 := StrToDate(FormatDateTime('yyyy/mm/dd',Now()));
               Days := Date1 - Date2;

               if (Days < 0) then
                  ThisReason := 'This file Prescribed on ' + Query1.FieldByName('Tracking_Order').AsString
               else
                  ThisReason := Format('Prescription in %.0f days!!',[Days]);

               if (Days < 8) then
                  InkCol := clRed
               else if (Days < 29) then
                  InkCol := clMaroon
               else if (Days < 91) then
                  InkCol := clBlue
               else
                  InkCol := clGreen;
            end;

            with xlsSheet.RCRange[row,1,row,6] do begin
               Borders[xlAround].Weight := xlThin;
               VerticalAlignment := xlVAlignTop;
               Item[1,1].Value := Query1.FieldByName('Tracking_Order').AsString;
               Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,2].Value := ReplaceQuote(Query1.FieldByName('Tracking_Name').AsString);
               Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,3].Value := ThisType;
               Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,4].Value := ReplaceQuote(Query1.FieldByName('Tracking_Description').AsString);
               Item[1,4].WrapText := ThisWrapText;
               Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,4].Borders[xlEdgeBottom].Weight := xlThin;
               Item[1,5].Value := ThisReason;
               Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,5].WrapText := ThisWrapText;
               Item[1,5].Font.Color := InkCol;
               Item[1,6].Value := ' ';
               Item[1,6].Borders[xlEdgeRight].Weight := xlThin;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
            Query1.Next;

            inc(row);
            inc(PageRow);

//--- If we've reached the maximum rows per page then it is PageBreak time.
//--- UNLESS WarpText is true in which case we do not do Pagebreak other than
//--- on the fist page. We do the Pagebreak here so that the Copyright notice
//--- will be handled correctly

            if (ThisWrapText = False) then begin
               if (PageRow >= RowsPerPage) then begin
                  PageBreak := True;
                  DoPageBreak(ord(PB_ALERTS),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,FileArray[idx1]);
               end;
            end;
         end;

      end else begin

//--- Perform a Page break if necessary

         if (PageBreak = True) then
            DoPageBreak(ord(PB_ALERTS),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,FileArray[idx1]);

         with xlsSheet.RCRange[row,1,row,6] do begin
            Item[1,1].Value := 'No Alert records for the period ending: ' + EDate;
            Item[1,1].Font.Bold := false;
            Borders[xlAround].Weight := xlThin;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;

         inc(row);

      end;

//--- Write the standard copyright notice

      if (ThisWrapText = True) then
         inc(row)
      else
         row := (Pages * lcGRows);

      with xlsSheet.RCRange[row,1,row,1] do begin
         Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 8;
      end;

//--- Remove the Gridlines which are added by default and set the Page orientation

      xlsSheet.PageSetup.Orientation := xlLandscape;
      xlsSheet.PageSetup.PaperSize := xlPaperA4;
      xlsSheet.DisplayGridLines := false;
      xlsSheet.PageSetup.CenterFooter := 'Page &P of &N';

      if (ThisWrapText = False) then begin
         xlsSheet.PageSetup.FitToPagesWide := 1;
         xlsSheet.PageSetup.FitToPagesTall := Pages;
      end;
   end;

//--- Write the Excel file to disk

   if (FilesFound > 0) then begin
      xlsBook.SaveAs(FileName + ThisFile);
      xlsBook.Close;

      LogMsg('  Alerts Report successfully processed...',True);
      LogMsg(' ',True);

      DoLine := False;

//--- Print the generated document on the Default Printer if requested

      if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
         if (PrintDocument(ThisFile, FileName) = True) then
            LogMsg('  Document submitted for printing...',True)
         else
            LogMsg('  Printing of document failed...',True);

         DoLine := True;
      end;

//--- Create a PDF file if requested

      if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
         PDFExists := PDFDocument(ThisFile, FileName);

         if (PDFExists = True) then
            LogMsg('  PDF file creation was successfull...',True)
         else
            LogMsg('  PDF file creation failed...',True);

         DoLine := True;

      end;

//--- Send the Excel file via email if requested

      if (SendByEmail = '1') then begin
         if (SendEmail('',FileName + ThisFile,ord(PT_NORMAL)) = true) then
            LogMsg('  Request to send generated Alerts Report by Email submitted...',True)
         else
            LogMsg('  Request to send generated Alerts Report by Email not submitted...',True);

         DoLine := True;
      end;

//--- Now open the Alerts Report if requested

      if (AutoOpen = True) then begin
         ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);
         LogMsg('  Request to open Alerts Report for ''' + PChar(FileName + ThisFile) + ''' submitted...',True);
         DoLine := True;
      end;
   end else begin
      LogMsg('  No Alerts Records found ...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Alerts Report');

   Close_Connection;
end;

//---------------------------------------------------------------------------
// Export the Phonebook
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Phonebook();
var
   row, PageRow, Pages, RowsPerPage              : integer;
   PageBreak, FirstPage, DoLine                  : boolean;
   ThisFile, Filter1, Filter2, ThisType, ThisStr : string;
   xlsBook                                       : IXLSWorkbook;
   xlsSheet                                      : IXLSWorksheet;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Phonebook Export');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);
      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();
   Filter1 := Parm07;
   Filter2 := Parm10;

//--- Create the Excel Workbook

   xlsBook := TXLSWorkbook.Create;

//-- Build the file name

   ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - Phonebook Export (' + CpyName + ').xls';
   txtDocument.Caption  := 'Exporting to: ' + FileName + ThisFile;

//--- Process each User record in turn

   xlsSheet      := xlsBook.WorkSheets.Add;
   xlsSheet.Name := 'Phonebook (' + CpyName + ')';

//--- Get all the Phonebook records based on the provided filter

   if (GetPhoneRecs(Filter1, Filter2) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);
      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Only process files that have records

   if (Query1.RecordCount > 0) then begin

      Pages     := 0;
      FirstPage := true;
      PageBreak := true;

      while Query1.Eof = false do begin

//--- Perform a Page break if necessary

         if (PageBreak = True) then
            DoPageBreak(ord(PB_PHONEBOOK),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,'');

         if (ReplaceQuote(Query1.FieldByName('Cust_Customer').AsString) = '--- Not Selected ---') then begin
            Query1.Next;
            continue;
         end;

         ThisType := Query1.FieldByName('Cust_CustType').AsString;
         case StrToInt(ThisType) of
            1: ThisStr := 'Client';
            2: ThisStr := 'Opposition';
            3: ThisStr := 'Correspondent';
            4: ThisStr := 'Counsel';
            5: ThisStr := 'Opposing Attorney';
         end;

         with xlsSheet.RCRange[row,1,row,6] do begin
            Borders[xlAround].Weight := xlThin;
            Item[1,1].Value := ReplaceQuote(Query1.FieldByName('Cust_Customer').AsString);
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := ThisStr;
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := Query1.FieldByName('Cust_Telephone').AsString + ' ';
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].HorizontalAlignment := xlHAlignRight;
            Item[1,4].Value := ' ' + Query1.FieldByName('Cust_Fax').AsString + ' ';
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].HorizontalAlignment := xlHAlignRight;
            Item[1,5].Value := ' ' + Query1.FieldByName('Cust_Cellphone').AsString + ' ';
            Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].HorizontalAlignment := xlHAlignRight;
            Item[1,6].Value := ' ' + Query1.FieldByName('Cust_Worknum').AsString + ' ';
            Item[1,6].HorizontalAlignment := xlHAlignRight;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;

         Query1.Next;

         inc(row);
         inc(PageRow);

         if (PageRow > RowsPerPage) then
            PageBreak := True;

      end;

   end else begin

      row := 1;
      with xlsSheet.RCRange[row,1,row,6] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := 'No Phonebook entries found';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;

   end;

//--- Write the standard copyright notice

   row := (Pages * lcGRows);
   with xlsSheet.RCRange[row,1,row,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet.PageSetup.Orientation := xlLandscape;
   xlsSheet.PageSetup.FitToPagesWide := 1;
   xlsSheet.PageSetup.FitToPagesTall := Pages;
   xlsSheet.PageSetup.PaperSize := xlPaperA4;
   xlsSheet.DisplayGridLines := false;
   xlsSheet.PageSetup.CenterFooter := 'Page &P of &N';

//--- Write the Excel file to disk

   xlsBook.SaveAs(FileName + ThisFile);
   xlsBook.Close;

   LogMsg('  Phonebook Export successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(ThisFile, FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(ThisFile, FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName + ThisFile,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Phonebook Export by Email submitted...',True)
      else
         LogMsg('  Request to send generated Phonebook Export by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the Fee Earner Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);

      LogMsg('  Request to open Phonebook Export for ''' + PChar(FileName + ThisFile) + ''' submitted...',True);

      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Phonebook Export');

   Close_Connection;
end;

//---------------------------------------------------------------------------
// Generate Billing Preparation Report
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_BillingReport();
var
   idx1, row, PageRow, Pages, RowsPerPage, sema, ThisMonth         : integer;
   ThisYear                                                        : integer;
   PageBreak, FirstPage, AddFile, HighLight, DoLine                : boolean;
   BillC, Bill30, Bill60, Bill90, Bill120, Bill150, Bill180        : double;
   InvC, Inv30, Inv60, Inv90, Inv120, Inv150, Inv180               : double;
   TotC, Tot30, Tot60, Tot90, Tot120, Tot150, Tot180               : double;
   NotInvoiced, TotNotInvoiced, TotBill                            : double;
   PaidC, Paid30, Paid60, Paid90, Paid120, Paid150, Paid180        : double;
   ThisTrust, ThisReserved, TotPaid, Unpaid, TotUnpaid             : double;
   ThisAB1F, ThisAB1T, ThisAB2F, ThisAB2T, ThisFill, ThisText      : TColor;
   ThisYearS, ThisMonthS, ThisFile, ThisReport                     : string;
   ThisDate, SaveSDate, SaveEDate, LeadStr, Statement              : string;
   MonthC, Month1, Month2, Month3, Month4, Month5, Month6          : string;
   SDateC, SDate30, SDate60, SDate90, SDate120, SDate150, SDate180 : string;
   EDateC, EDate30, EDate60, EDate90, EDate120, EDate150, EDate180 : string;
   xlsBook                                                         : IXLSWorkbook;
   xlsSheet                                                        : IXLSWorksheet;

   LastDay : array[1..12] of integer;

begin

//--- Initialise the LastDay Array

   LastDay[ 1] := 31; LastDay[ 2] := 28; LastDay[ 3] := 31; LastDay[ 4] := 30;
   LastDay[ 5] := 31; LastDay[ 6] := 30; LastDay[ 7] := 31; LastDay[ 8] := 31;
   LastDay[ 9] := 30; LastDay[10] := 31; LastDay[11] := 30; LastDay[12] := 31;

//--- Initialise the totals

   TotC := 0; Tot30 := 0; Tot60 := 0; Tot90 := 0; Tot120 := 0; Tot150 := 0;
   Tot180 := 0; TotNotInvoiced := 0; TotUnpaid := 0;

//--- Preserve the current values of SDate and EDate

   SaveSDate := SDate;
   SaveEdate := EDate;

//--- Set the Current Period Date to the last day of the month of the end date
//--- that was passed

   ShortDateFormat := 'yyyy/MM/dd';
   DateSeparator   := '/';

   ThisYearS := FormatDateTime('yyyy',StrToDate(EDate));
   ThisYear  := StrToInt(ThisYearS);

   if (ThisYear mod 4 = 0) then
      LastDay[2] := 29;

   ThisMonthS := FormatDateTime('MM',StrToDate(EDate));
   ThisMonth  := StrToInt(ThisMonthS);
   ThisDate   := ThisYearS + '/' + ThisMonthS + '/' + IntToStr(LastDay[ThisMonth]);

   MonthC := FormatDateTime('MMM-YY',StrToDate(ThisDate));
   Month1 := FormatDateTime('MMM-YY',IncMonth(StrToDate(ThisDate),-1));
   Month2 := FormatDateTime('MMM-YY',IncMonth(StrToDate(ThisDate),-2));
   Month3 := FormatDateTime('MMM-YY',IncMonth(StrToDate(ThisDate),-3));
   Month4 := FormatDateTime('MMM-YY',IncMonth(StrToDate(ThisDate),-4));
   Month5 := FormatDateTime('MMM-YY',IncMonth(StrToDate(ThisDate),-5));
   Month6 := FormatDateTime('MMM-YY',IncMonth(StrToDate(ThisDate),-6));

//--- Now set the Start and End Dates for the various periods

   if (ThisMonth < 10) then LeadStr := '0' else LeadStr := '';
   SDateC := IntToStr(ThisYear) + '/' + LeadStr + IntToStr(ThisMonth);
   EDateC := SDateC;
   SDateC := SDateC + '/01';
   EDateC := EDateC + '/' + IntToStr(LastDay[ThisMonth]);

   Dec(ThisMonth);
   if (ThisMonth = 0) then begin
      Dec(ThisYear);
      ThisMonth := 12;
   end;

   if (ThisMonth < 10) then LeadStr := '0' else LeadStr := '';
   SDate30 := IntToStr(ThisYear) + '/' + LeadStr + IntToStr(ThisMonth);
   EDate30 := SDate30;
   SDate30 := SDate30 + '/01';
   EDate30 := EDate30 + '/' + IntToStr(LastDay[ThisMonth]);

   Dec(ThisMonth);
   if (ThisMonth = 0) then begin
      Dec(ThisYear);
      ThisMonth := 12;
   end;

   if (ThisMonth < 10) then LeadStr := '0' else LeadStr := '';
   SDate60 := IntToStr(ThisYear) + '/' + LeadStr + IntToStr(ThisMonth);
   EDate60 := SDate60;
   SDate60 := SDate60 + '/01';
   EDate60 := EDate60 + '/' + IntToStr(LastDay[ThisMonth]);

   Dec(ThisMonth);
   if (ThisMonth = 0) then begin
      Dec(ThisYear);
      ThisMonth := 12;
   end;

   if (ThisMonth < 10) then LeadStr := '0' else LeadStr := '';
   SDate90 := IntToStr(ThisYear) + '/' + LeadStr + IntToStr(ThisMonth);
   EDate90 := SDate90;
   SDate90 := SDate90 + '/01';
   EDate90 := EDate90 + '/' + IntToStr(LastDay[ThisMonth]);

   Dec(ThisMonth);
   if (ThisMonth = 0) then begin
      Dec(ThisYear);
      ThisMonth := 12;
   end;

   if (ThisMonth < 10) then LeadStr := '0' else LeadStr := '';
   SDate120 := IntToStr(ThisYear) + '/' + LeadStr + IntToStr(ThisMonth);
   EDate120 := SDate120;
   SDate120 := SDate120 + '/01';
   EDate120 := EDate120 + '/' + IntToStr(LastDay[ThisMonth]);

   Dec(ThisMonth);
   if (ThisMonth = 0) then begin
      Dec(ThisYear);
      ThisMonth := 12;
   end;

   if (ThisMonth < 10) then LeadStr := '0' else LeadStr := '';
   SDate150 := IntToStr(ThisYear) + '/' + LeadStr + IntToStr(ThisMonth);
   EDate150 := SDate150;
   SDate150 := SDate150 + '/01';
   EDate150 := EDate150 + '/' + IntToStr(LastDay[ThisMonth]);

   Dec(ThisMonth);
   if (ThisMonth = 0) then begin
      Dec(ThisYear);
      ThisMonth := 12;
   end;

   if (ThisMonth < 10) then LeadStr := '0' else LeadStr := '';
   SDate180 := '1980/01/01';
   EDate180 := IntToStr(ThisYear) + '/' + LeadStr + IntToStr(ThisMonth) + '/' + IntToStr(LastDay[ThisMonth]);

//--- Set the Page control variables

   FirstPage := True;
   PageBreak := True;
   Pages     := 0;

//--- Set up to use alternate color blocks

   sema := 1;
   ThisAB1F := ColAB1F;
   ThisAB1T := ColAB1T;
   ThisAB2F := ColAB2F;
   ThisAB2T := ColAB2T;

//--- Now process the request

   ThisCount := 1;
   ThisMax   := IntToStr(NumFiles);

   prbProgress.Max := NumFiles;
   LogMsg(ord(PT_PROLOG),True,True,'Billing Preparation Report');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);
      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Create the Excel Workbook

   xlsBook := TXLSWorkbook.Create;

//-- Build the file name

   ThisReport := FormatDateTime('yyyyMMdd',StrToDate(EDate)) + ' - Billing Preparation Report (' + CpyName + ').xls';
   txtDocument.Caption  := 'Generating to: ' + FileName + ThisReport;

//--- Process each User record in turn

   xlsSheet       := xlsBook.WorkSheets.Add;
   xlsSheet.Name := 'Billing Preparation Report (' + CpyName + ')';

   for idx1 := 0 to NumFiles - 1 do begin

      Statement := ' ';

      ThisFile := FileArray[idx1];
      txtError.Text := 'Processing: ' + ThisFile;
      txtError.Refresh;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

//--- Get the File Description and Related information

      Descrip := GetFileDesc(ThisFile);

//--- Exclude this record if Show Related is true and this is a related file

      if (ShowRelated = true) then begin
         if (SaveOwner <> SaveRelated) then
            continue;
      end;

//--- Exclude this record if it is a Collections File

      if (SaveFileType <> 0) then
         continue;

//--- Get the Billing and previous Invoices, Payments and Trust details for
//--- each File

      HighLight    := false;
      ThisTrust    := 0;
      ThisReserved := 0;

      GetBillingR(ThisFile,SDateC,  EDateC,  BillC,  InvC,  PaidC,  ThisReserved,ThisTrust);
      GetBillingR(ThisFile,SDate30, EDate30, Bill30, Inv30, Paid30, ThisReserved,ThisTrust);
      GetBillingR(ThisFile,SDate60, EDate60, Bill60, Inv60, Paid60, ThisReserved,ThisTrust);
      GetBillingR(ThisFile,SDate90, EDate90, Bill90, Inv90, Paid90, ThisReserved,ThisTrust);
      GetBillingR(ThisFile,SDate120,EDate120,Bill120,Inv120,Paid120,ThisReserved,ThisTrust);
      GetBillingR(ThisFile,SDate150,EDate150,Bill150,Inv150,Paid150,ThisReserved,ThisTrust);
      GetBillingR(ThisFile,SDate180,EDate180,Bill180,Inv180,Paid180,ThisReserved,ThisTrust);

//--- Calculate some totals

      TotBill := BillC + Bill30 + Bill60 + Bill90 + Bill120 + Bill150 + Bill180;
      TotPaid := PaidC + Paid30 + Paid60 + Paid90 + Paid120 + Paid150 + Paid180;

//--- Remove what has been invoiced from the billing for each period
//--- 180+ Days

      if (Inv180 > Bill180) then HighLight := true;
      Bill180 := Bill180 - Inv180;

//--- 150 Days

      if (Inv150 > Bill150) then HighLight := true;
      Bill150 := Bill150 - Inv150;

//--- 120 Days

      if (Inv120 > Bill120) then HighLight := true;
      Bill120 := Bill120 - Inv120;

//--- 90 Days

      if (Inv90 > Bill90) then HighLight := true;
      Bill90 := Bill90 - Inv90;

//--- 60 Days

      if (Inv60 > Bill60) then HighLight := true;
      Bill60 := Bill60 - Inv60;

//--- 30 Days

      if (Inv30 > Bill30) then HighLight := true;
      Bill30 := Bill30 - Inv30;

//--- Current

      if (InvC > BillC) then HighLight := true;
      BillC := BillC - InvC;

//--- Give effect to the "Exclude Reserve" and "Include Trust" options

      if (ExcludeReserve = true) then begin
         if ((ThisTrust > 0) AND (ThisReserved > ThisTrust)) then
            ThisReserved := ThisTrust;
      end else
         ThisReserved := 0;

      if (IncludeTrust = false) then
         ThisTrust := 0;

//--- Calculate the Unpaid and Not Billed amounts

      Unpaid      := TotBill - TotPaid - (ThisTrust - ThisReserved);
      NotInvoiced := BillC + Bill30 + Bill60 + Bill90 + Bill120 + Bill150 + Bill180;

//--- Check whether this record must be added

      if ((FloatToStrF(NotInvoiced,ffNumber,10,2) > '0.00') OR (HighLight = true) OR ((FloatToStrF(Unpaid,ffNumber,10,2) <> '0.00'))) then
         AddFile := true
      else
         AddFile := false;

//--- Process the record. Only Add records for which Billing is due or for
//--- which a Satement is due

      if (AddFile = true) then begin

//--- Determine whether 'Statement' must be checked

         if (Unpaid > 0.00) then
            Statement := 'X';

//--- Process the record

         if (PageBreak = True) then
            Billng_Prep_PageBreak(xlsSheet,PageBreak,FirstPage,SaveEDate,MonthC,Month1,Month2,Month3,Month4,Month5,Month6,RowsPerPage,row,PageRow,Pages);

         if (sema = 0) then begin;
            ThisFill := ThisAB1F;
            ThisText := ThisAB1T;
            sema := 1;
         end else begin
            ThisFill := ThisAB2F;
            ThisText := ThisAB2T;
            sema := 0;
         end;

//--- Write the Line

         with xlsSheet.RCRange[row,1,row,17] do begin
            Item[1, 1].Value := ReplaceQuote(Descrip);
            Item[1, 1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 2].Value := BillC;
            Item[1, 2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 2].HorizontalAlignment := xlHAlignRight;
            Item[1, 3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 4].Value := Bill30;
            Item[1, 4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 4].HorizontalAlignment := xlHAlignRight;
            Item[1, 5].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 6].Value := Bill60;
            Item[1, 6].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 6].HorizontalAlignment := xlHAlignRight;
            Item[1, 7].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 8].Value := Bill90;
            Item[1, 8].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 8].HorizontalAlignment := xlHAlignRight;
            Item[1, 9].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,10].Value := Bill120;
            Item[1,10].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,10].HorizontalAlignment := xlHAlignRight;
            Item[1,11].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,12].Value := Bill150;
            Item[1,12].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,12].HorizontalAlignment := xlHAlignRight;
            Item[1,13].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,14].Value := Bill180;
            Item[1,14].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,14].HorizontalAlignment := xlHAlignRight;
            Item[1,15].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,16].Value := NotInvoiced;
            Item[1,16].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,16].HorizontalAlignment := xlHAlignRight;
            Item[1,17].Value := Statement;
            Item[1,17].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,17].HorizontalAlignment := xlHAlignCenter;

            Borders[xlAround].Weight := xlThin;
            Interior.Color := ThisFill;
            Font.Color := Thistext;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;

         with xlsSheet.RCRange[row,2,row,16] do begin
            NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
         end;

         inc(row);
         inc(PageRow);

//--- Calculate the totals

         TotC           := TotC + BillC;
         Tot30          := Tot30 + Bill30;
         Tot60          := Tot60 + Bill60;
         Tot90          := Tot90 + Bill90;
         Tot120         := Tot120 + Bill120;
         Tot150         := Tot150 + Bill150;
         Tot180         := Tot180 + Bill180;
         TotNotInvoiced := TotNotInvoiced + NotInvoiced;
         TotUnpaid      := TotUnpaid + Unpaid;

         if (PageRow > RowsPerPage) then PageBreak := True;

      end;
   end;

//--- Write the Total Line

   if (PageBreak = True) then
      Billng_Prep_PageBreak(xlsSheet,PageBreak,FirstPage,SaveEDate,MonthC,Month1,Month2,Month3,Month4,Month5,Month6,RowsPerPage,row,PageRow,Pages);

//--- Write the Line

   with xlsSheet.RCRange[row,1,row,17] do begin
      Item[1, 1].Value := 'Totals per Period:';
      Item[1, 1].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 2].Value := TotC;
      Item[1, 2].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 2].HorizontalAlignment := xlHAlignRight;
      Item[1, 3].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 4].Value := Tot30;
      Item[1, 4].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 4].HorizontalAlignment := xlHAlignRight;
      Item[1, 5].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 6].Value := Tot60;
      Item[1, 6].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 6].HorizontalAlignment := xlHAlignRight;
      Item[1, 7].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 8].Value := Tot90;
      Item[1, 8].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 8].HorizontalAlignment := xlHAlignRight;
      Item[1, 9].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,10].Value := Tot120;
      Item[1,10].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,10].HorizontalAlignment := xlHAlignRight;
      Item[1,11].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,12].Value := Tot150;
      Item[1,12].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,12].HorizontalAlignment := xlHAlignRight;
      Item[1,13].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,14].Value := Tot180;
      Item[1,14].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,14].HorizontalAlignment := xlHAlignRight;
      Item[1,15].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,16].Value := TotNotInvoiced;
      Item[1,16].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,16].HorizontalAlignment := xlHAlignRight;
      Item[1,17].Value := ' ';
      Item[1,17].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,17].HorizontalAlignment := xlHAlignCenter;

      Borders[xlAround].Weight := xlThin;
      Interior.Color := ColDHF;
      Font.Color := ColDHT;
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

   with xlsSheet.RCRange[row,2,row,16] do begin
      NumberFormat := '#,##0.00_);-#,##0.00_)';
   end;

//--- Generate an appropriate message if no billing records were found

   if (row = 4) then begin
      with xlsSheet.RCRange[row,1,row,17] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := 'No Billing Records found';
         Interior.Color := ThisAB1F;
         Font.Color := ThisAB1T;
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

//--- Write the standard copyright notice

   row := (Pages * lcGRows);
   with xlsSheet.RCRange[row,1,row,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set various other Page
//--- related settings

   xlsSheet.PageSetup.Orientation := xlLandscape;
   xlsSheet.PageSetup.FitToPagesWide := 1;
   xlsSheet.PageSetup.FitToPagesTall := Pages;
   xlsSheet.PageSetup.PaperSize := xlPaperA4;
   xlsSheet.DisplayGridLines := false;
   xlsSheet.PageSetup.CenterFooter := 'Page &P of &N';

//--- Write the Excel file to disk

   xlsBook.SaveAs(FileName + ThisReport);
   xlsBook.Close;

   LogMsg('  Billing Preparation Report successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(ThisReport, FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(ThisReport, FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName + ThisReport,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Billing Preparation Report by Email submitted...',True)
      else
         LogMsg('  Request to send generated Billing Preparation Report by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the Billing Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName + ThisReport),nil,nil,SW_SHOWNORMAL);

      LogMsg('  Request to open Billing Preparation Report for ''' + PChar(FileName + ThisReport) + ''' submitted...',True);

      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Billing Preparation Report');

   EDate := SaveEDate;
   SDate := SaveSDate;

   Close_Connection;
end;

//---------------------------------------------------------------------------
// Generate Trust Management Report
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_TrustManagement();
var
   row, idx1, Pages, PageRow, RowsPerPage  : integer;
   ThisFees, ThisTrust, This864, ThisVAT  : double;
   PageBreak, FirstPage, DoLine            : boolean;
   ThisFile, ThisStr                       : string;
   xlsBook                                 : IXLSWorkbook;
   xlsSheet                                : IXLSWorksheet;

begin

//--- Set the Page control variables

   FirstPage := True;
   PageBreak := True;
   Pages     := 0;

//--- Now process the request

   ThisCount := 1;
   ThisMax   := IntToStr(NumFiles);

   prbProgress.Max := NumFiles;
   LogMsg(ord(PT_PROLOG),True,True,'Trust Management Report');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);
      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Create the Excel Workbook

   xlsBook       := TXLSWorkbook.Create;

//-- Build the file name

   ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - Trust Management Report (' + CpyName + ').xls';
   txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;

//--- Process each User record in turn

   xlsSheet      := xlsBook.WorkSheets.Add;
   xlsSheet.Name := 'Trust Management Report (' + CpyName + ')';

   for idx1 := 0 to NumFiles - 1 do begin

//--- Perform a Page break if necessary

      if (PageBreak = True) then
         DoPageBreak(ord(PB_TRUSTMAN),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,'');

//--- Initialise the Total Balance Fields

      OpenBalFees     := 0;
      OpenBalTrust    := 0;
      OpenBal864      := 0;
      OpenBalVAT      := 0;

//--- Build the unique part of the SQL statement

      ThisStr := ' AND ((B_Class >= 0 AND B_Class <= 14) OR (B_Class >= 18 AND B_Class <= 19) OR (B_Class = 21) OR (B_Class = 24)) AND B_AccountType = 0';

      if ((GetBilling(FileArray[idx1],ThisStr,'Trust Management')) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

//--- Exclude this record if Show Related is true and this is a related file

      if (ShowRelated = true) then begin
         if (SaveOwner <> SaveRelated) then
            continue;
      end;

      txtError.Text := 'Processing: ' + FileArray[idx1];
      txtError.Refresh;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

//--- Now process each file

      ThisFees  := OpenBalFees  + ((This_Abs.Fees + This_Abs.Disbursements + This_Abs.Expenses + This_Abs.Business_To_Trust + This_Abs.Business_Debit) - (This_Abs.Credit + This_Abs.Write_off + This_Abs.Payment_Received + This_Abs.Business_Deposit + This_Abs.Trust_Transfer_Business_Fees));
      ThisTrust := OpenBalTrust + ((This_Abs.Trust_Deposit + This_Abs.Business_To_Trust + This_Abs.Trust_Withdrawal_S86_4) - (This_Abs.Trust_Transfer_Business_Fees + This_Abs.Trust_Transfer_Business_Other + This_Abs.Trust_Transfer_Client + This_Abs.Trust_Transfer_Disbursements + This_Abs.Trust_Transfer_Trust + This_Abs.Trust_Debit + This_Abs.Trust_Investment_S86_4));
      This864  := OpenBal864 + (This_Abs.Trust_Investment_S86_4 + This_Abs.Trust_Interest_S86_4 - This_Abs.Trust_Withdrawal_S86_4);

//--- Add VAT if necessary

      if (VATRate > 0) then
         ThisVAT  := (ThisFees * (VATRate / 100))
      else
         ThisVAT  := 0;

//--- Exclude this record if all zeroes and 'Include Nil Balances'  is false

      ThisFees  := RoundD(ThisFees,2);
      ThisTrust := RoundD(ThisTrust,2);
      This864 := RoundD(This864,2);
      ThisVAT   := RoundD(ThisVAT,2);

      if ((ThisFees = 0) and (ThisTrust = 0) and (This864 = 0) and (NilBalance <> '1')) then
         continue;

//--- Process the record

      with xlsSheet.RCRange[row,1,row,7] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := FileArray[idx1];
         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,2].Value := ReplaceQuote(Descrip);
         Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,3].Value := This864;
         Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,3].HorizontalAlignment := xlHAlignRight;
         Item[1,4].Value := ThisTrust;
         Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,4].HorizontalAlignment := xlHAlignRight;
         Item[1,5].Value := (ThisFees + ThisVAT) * -1;
         Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,5].HorizontalAlignment := xlHAlignRight;
         Item[1,6].Value := ((ThisFees + ThisVAT) * -1) + ThisTrust + This864;
         Item[1,6].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,6].HorizontalAlignment := xlHAlignRight;

         if (ABS((This864 + ThisTrust)) = ABS((ThisFees + ThisVAT))) then
            Item[1,7].Value := ' Yes';

         Item[1,7].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,7].HorizontalAlignment := xlHAlignJustify;

         Borders[xlAround].Weight := xlThin;
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;

      with xlsSheet.RCRange[row,3,row,6] do begin
         NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      end;

      inc(row);
      inc(PageRow);

      if (PageRow > RowsPerPage) then
         PageBreak := True;

      Query1.Next;
   end;

//--- Generate an appropriate message if no billing records were found

   if (row = 4) then begin
      with xlsSheet.RCRange[row,1,row,7] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := 'No Billing Records found';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

//--- Write the standard copyright notice

   row := (Pages * lcGRows);
   with xlsSheet.RCRange[row,1,row,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet.PageSetup.Orientation := xlLandscape;
   xlsSheet.PageSetup.FitToPagesWide := 1;
   xlsSheet.PageSetup.FitToPagesTall := Pages;
   xlsSheet.PageSetup.PaperSize := xlPaperA4;
   xlsSheet.DisplayGridLines := false;
   xlsSheet.PageSetup.CenterFooter := 'Page &P of &N';

//--- Write the Excel file to disk

   xlsBook.SaveAs(FileName + ThisFile);
   xlsBook.Close;

   LogMsg('  Trust Management Report successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(ThisFile, FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(ThisFile, FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName + ThisFile,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Trust Management by Email submitted...',True)
      else
         LogMsg('  Request to send generated Trust Management by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the Billing Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);

      LogMsg('  Request to open Trust Management Report for ''' + PChar(FileName + ThisFile) + ''' submitted...',True);

      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Trust Management Report');

   Close_Connection;
end;

//---------------------------------------------------------------------------
// Generate Payments Report
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Payments();
var
   idx1, row, PageRow, Pages, RowsPerPage, Sema : integer;
   InvFound                                     : integer;
   PageBreak, FirstPage, DoLine                 : boolean;
   Paid, TotPaid, TotAmount, ThisPaid, ThisFees : double;
   ThisAmount, ThisDisburse, ThisExpenses       : double;
   TotPrevPaid                                  : double;
   ThisFill, ThisText                           : TColor;
   ThisFile, ThisClient                         : string;
   xlsBook                                      : IXLSWorkbook;
   xlsSheet                                     : IXLSWorksheet;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Payments Report');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      lbProgress.Items.Add('Unexpected Data Base error: ' + ErrMsg);
      lbProgress.TopIndex := lbProgress.Items.Count - 1;
      lbProgress.Refresh;;

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Set defaults then create a new Excel Workbook

//--- Create the Excel Workbook

   xlsBook       := TXLSWorkbook.Create;
   xlsSheet      := xlsBook.WorkSheets.Add;
   xlsSheet.Name := 'Payments Report';

//-- Build the file name

   ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - Payments Report (' + CpyName + ').xls';
   txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;

//--- Start by getting all the invoices and payments for the period

   InvFound  := 0;
   Sema      := 0;

   Pages     := 0;
   FirstPage := True;
   PageBreak := True;

   TotAmount   := 0;
   TotPaid     := 0;
   TotPrevPaid := 0;

   if (NumFiles = 1) and (CreateStatement = true) then begin

      GetInvoices(ord(DT_PAYMENT),FileArray[0]);
      InvFound := 1;

//--- Get all of the details

      PayAmount   := StrToFloat(Query1.FieldByName('Inv_Amount').AsString);
      PayFees     := StrToFloat(Query1.FieldByName('Inv_Fees').AsString);
      PayDisburse := StrToFloat(Query1.FieldByName('Inv_Disburse').AsString);
      PayExpenses := StrToFloat(Query1.FieldByName('Inv_Expenses').AsString);

//--- Process the record

      txtError.Text := 'Now processing invoice: ' + FileArray[0];
      txtError.Repaint;

//--- Get the payments that were made in this period for this invoice

      GetPayments(FileArray[0]);
      TotPaid := 0;

//--- Now process and display all the payments for this invoice

      Query2.First;

      for idx1 := 0 to Query2.RecordCount - 1 do begin

//--- Perform a Page break if necessary. We preserve the value of FirstPage
//--- in order to print once-off information on the first page only

         if (PageBreak = True) then
            DoPageBreak(ord(PB_PAYMENTS01),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,'');

//--- Set this line's colour

         if (Sema = 0) then begin
            Sema := 1;
            ThisFill := ColAB1F;
            ThisText := ColAB1T;
         end else begin
            Sema := 0;
            ThisFill := ColAB2F;
            ThisText := ColAB2T;
         end;

//--- Write the Line

         ThisPaid := Query2.FieldByName('Pay_Amount').AsFloat;
         TotPaid  := TotPaid + ThisPaid;

         if (ThisPaid >= PayAmount) then begin
            ThisFees     := PayFees;
            ThisDisburse := PayDisburse;
            ThisExpenses := PayExpenses;
         end else begin
            ThisAmount := ThisPaid;

            if (ThisAmount > PayDisburse) then begin
               ThisDisburse := PayDisburse;
               ThisAmount   := ThisAmount - PayDisburse;
               PayDisburse  := 0;
            end else begin
               ThisDisburse := ThisAmount;
               PayDisburse  := PayDisburse - ThisAmount;
               ThisAmount   := 0;
            end;

            if (ThisAmount > PayExpenses) then begin
               ThisExpenses := PayExpenses;
               ThisAmount   := ThisAmount - PayExpenses;
               PayExpenses  := 0;
            end else begin
               ThisExpenses := ThisAmount;
               PayExpenses  := PayExpenses - ThisAmount;
               ThisAmount   := 0;
            end;

            if (ThisAmount > PayFees) then
               ThisFees := PayFees
            else
               ThisFees := ThisAmount;
         end;

         with xlsSheet.RCRange[row,1,row,9] do begin
            Item[1,1].Value := Query2.FieldByName('Pay_Date').AsString;
            Item[1,1].Borders[xlAround].Weight := xlThin;
            Item[1,2].Value := RoundD(Query2.FieldByName('Pay_Amount').AsFloat,2);
            Item[1,2].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
            Item[1,2].Borders[xlAround].Weight := xlThin;
            Item[1,3].Value := ReplaceQuote(Query2.FieldByName('Pay_Note').AsString);
            Item[1,5].Value := ThisFees;
            Item[1,5].Borders[xlAround].Weight := xlThin;
            Item[1,5].HorizontalAlignment := xlHAlignRight;
            Item[1,6].Value := ThisDisburse;
            Item[1,6].Borders[xlAround].Weight := xlThin;
            Item[1,6].HorizontalAlignment := xlHAlignRight;
            Item[1,7].Value := ThisExpenses;
            Item[1,7].Borders[xlAround].Weight := xlThin;
            Item[1,7].HorizontalAlignment := xlHAlignRight;
            Item[1,8].Value := ThisPaid;
            Item[1,8].Borders[xlAround].Weight := xlThin;
            Item[1,8].HorizontalAlignment := xlHAlignRight;
            Item[1,9].Value := PayAmount - TotPaid;
            Item[1,9].Borders[xlAround].Weight := xlThin;
            Item[1,9].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := ThisFill;
            Font.Color := ThisText;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;

         with xlsSheet.RCRange[row,5,row,9] do begin
            NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
         end;

         inc(row);
         inc(PageRow);

         if (PageRow > RowsPerPage) then
            PageBreak := True;

         Query2.Next;

      end;
   end else begin

      for idx1 := 0 to NumFiles - 1 do begin

         GetInvoices(ord(DT_PAYMENT),FileArray[idx1]);

//--- Process the record

         txtError.Text := 'Now processing invoice: ' + FileArray[idx1];
         txtError.Repaint;

//--- Perform a Page break if necessary

         if (PageBreak = True) then
            DoPageBreak(ord(PB_PAYMENTS02),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,'');

//--- Get the payments that were made in this period for the current file

         Paid := GetPayments(Query1.FieldByName('Inv_Invoice').AsString);

//--- Exclude this record if no payments were made in this period

         if (Paid = 0) then begin
            Query1.Next;
            continue;
         end;

//--- Get all of the details

         PayAmount   := StrToFloat(Query1.FieldByName('Inv_Amount').AsString);
         PayFees     := StrToFloat(Query1.FieldByName('Inv_Fees').AsString);
         PayDisburse := StrToFloat(Query1.FieldByName('Inv_Disburse').AsString);
         PayExpenses := StrToFloat(Query1.FieldByName('Inv_Expenses').AsString);

         TotAmount   := TotAmount + PayAmount;
         TotPaid     := TotPaid + Paid;
         TotPrevPaid := TotPrevPaid + PrevPaid;

//--- Set this block's colour

         if (Sema = 0) then begin
            Sema := 1;
            ThisFill := ColAB1F;
            ThisText := ColAB1T;
         end else begin
            Sema := 0;
            ThisFill := ColAB2F;
            ThisText := ColAB2T;
         end;

         inc(InvFound);

//--- Write the Line

         ThisClient := GetClient(Query1.FieldByName('Inv_File').AsString);

         with xlsSheet.RCRange[row,1,row,10] do begin
            Item[1, 1].Value := Query1.FieldByName('Inv_File').AsString;
            Item[1, 1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 2].Value := Query1.FieldByName('Inv_EDate').AsString;
            Item[1, 2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 3].Value := Query1.FieldByName('Inv_Invoice').AsString;
            Item[1, 3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 4].Value := ThisClient;
            Item[1, 4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 5].Value := PayFees;
            Item[1, 5].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 5].HorizontalAlignment := xlHAlignRight;
            Item[1, 6].Value := PayDisburse;
            Item[1, 6].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 6].HorizontalAlignment := xlHAlignRight;
            Item[1, 7].Value := PayExpenses;
            Item[1, 7].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 7].HorizontalAlignment := xlHAlignRight;
            Item[1, 8].Value := PayAmount;
            Item[1, 8].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 8].HorizontalAlignment := xlHAlignRight;
            Item[1, 9].Value := Paid;
            Item[1, 9].Borders[xlEdgeRight].Weight := xlThin;
            Item[1, 9].HorizontalAlignment := xlHAlignRight;
            Item[1,10].Value := PayAmount - Paid - PrevPaid;
            Item[1,10].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,10].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlThin;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
            Interior.Color := ThisFill;
            Font.Color := ThisText;
         end;

         with xlsSheet.RCRange[row,5,row,10] do begin
            NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
         end;

         inc(row);
         inc(PageRow);

         if (PageRow > RowsPerPage) then
            PageBreak := True;

      end;

      with xlsSheet.RCRange[3,10,5,10] do begin
         Item[1, 1].Value := TotAmount;
         Item[1, 1].Borders[xlEdgeLeft].Weight := xlThin;
         Item[2, 1].Value := TotPaid;
         Item[2, 1].Borders[xlEdgeLeft].Weight := xlThin;
         Item[3, 1].Value := TotAmount - TotPaid - TotPrevPaid;
         Item[3, 1].Borders[xlEdgeLeft].Weight := xlThin;
         NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;

      inc(row);
      inc(PageRow);

      if (PageRow > RowsPerPage) then
         PageBreak := True;

//--- Perform a Page break if necessary

      if (PageBreak = True) then
         DoPageBreak(ord(PB_PAYMENTS02),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,'');

   end;

//--- Write the copyright notice

   row := (Pages * lcGRows);
   with xlsSheet.RCRange[row,1,row,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet.PageSetup.Orientation := xlLandscape;
   xlsSheet.PageSetup.FitToPagesWide := 1;
   xlsSheet.PageSetup.FitToPagesTall := Pages;
   xlsSheet.PageSetup.PaperSize := xlPaperA4;
   xlsSheet.DisplayGridLines := false;
   XlsSheet.PageSetup.CenterFooter := 'Page &P of &N';

//--- Write the Excel file to disk.

   if (InvFound > 0) then begin

      xlsBook.SaveAs(FileName + ThisFile);
      xlsBook.Close;

      LogMsg('  Payments Report successfully processed...',True);
      LogMsg(' ',True);

      DoLine := False;

//--- Print the generated document on the Default Printer if requested

      if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
         if (PrintDocument(ThisFile, FileName) = True) then
            LogMsg('  Document submitted for printing...',True)
         else
            LogMsg('  Printing of document failed...',True);

         DoLine := True;
      end;

//--- Create a PDF file if requested

      if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
         PDFExists := PDFDocument(ThisFile, FileName);

         if (PDFExists = True) then
            LogMsg('  PDF file creation was successfull...',True)
         else
            LogMsg('  PDF file creation failed...',True);

         DoLine := True;
      end;

//--- Send the Excel file via email if requested

      if (SendByEmail = '1') then begin
         if (SendEmail('',FileName + ThisFile,ord(PT_NORMAL)) = true) then
            LogMsg('  Request to send generated Payments Report by Email submitted...',True)
         else
            LogMsg('  Request to send generated Payments Report by Email not submitted...',True);

         DoLine := True;
      end;

//--- Now open the Payments Report if requested

      if (AutoOpen = True) then begin
         ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);

         LogMsg('  Request to open Payments Report for ''' + PChar(FileName + ThisFile) + ''' submitted...',True);

         DoLine := True;
      end;
   end else begin
      LogMsg('  No Invoices with Payments for the specified period ...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Payments Report');

   Close_Connection;

end;

//---------------------------------------------------------------------------
// Export Prefix Details
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_PrefixDetails();
var
   idx1, row1, RecCount             : integer;
   DoLine                           : boolean;
   MultiFiles                       : string;
   xlsBook                          : IXLSWorkbook;
   xlsSheet1                        : IXLSWorksheet;
   ExportSet, SectionSet, RecordSet : IXMLNode;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Prefix List');

   txtDocument.Caption  := 'Generating to: ' + FileName;
   txtDocument.Refresh;

//--- Get Company specific information

   GetCpyVAT();

   MultiFiles := BoolTostr(ShowRelated);

//--- Open the XML file and process the content

   XMLDoc.Active := false;
   XMLDoc.LoadFromFile(HostName);
   XMLDoc.Active := true;

   ExportSet  := XMLDoc.DocumentElement;
   SectionSet := ExportSet.ChildNodes.First;
   RecordSet  := SectionSet;

//--- Create the Excel spreadsheet

   xlsBook        := TXLSWorkbook.Create;
   xlsSheet1      := xlsBook.WorkSheets.Add;
   xlsSheet1.Name := RecordSet.ChildValues['File'];

//--- Write the Report Heading

   if (MultiFiles <> '0') then begin
      with xlsSheet1.Range['A1','K1'] do begin
         Item[1,1].Value := CpyName + ': Prefix list (' + xlsSheet1.Name + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end else begin
      with xlsSheet1.Range['A1','B1'] do begin
         Item[1,1].Value := CpyName + ': Prefix list (' + xlsSheet1.Name + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end;

//--- Write the File List Heading

   if (MultiFiles <> '0') then begin
      with xlsSheet1.Range['A3','K3'] do begin
         Item[1, 1].Value := 'Prefix';
         Item[1, 1].ColumnWidth := 10;
         Item[1, 1].Borders[xlAround].Weight := xlThin;
         Item[1, 2].Value := 'Count ';
         Item[1, 2].ColumnWidth := 10;
         Item[1, 2].HorizontalAlignment := xlHAlignRight;
         Item[1, 2].Borders[xlAround].Weight := xlThin;
         Item[1, 3].Value := 'Description';
         Item[1, 3].ColumnWidth := 100;
         Item[1, 3].Borders[xlAround].Weight := xlThin;
         Item[1, 4].Value := 'Location';
         Item[1, 4].ColumnWidth := 100;
         Item[1, 4].Borders[xlAround].Weight := xlThin;
         Item[1, 5].Value := 'Template';
         Item[1, 5].ColumnWidth := 100;
         Item[1, 5].Borders[xlAround].Weight := xlThin;
         Item[1, 6].Value := 'Created By';
         Item[1, 6].ColumnWidth := 30;
         Item[1, 6].Borders[xlAround].Weight := xlThin;
         Item[1, 7].Value := 'Create Date';
         Item[1, 7].ColumnWidth := 15;
         Item[1, 7].Borders[xlAround].Weight := xlThin;
         Item[1, 8].Value := 'Create Time';
         Item[1, 8].ColumnWidth := 15;
         Item[1, 8].Borders[xlAround].Weight := xlThin;
         Item[1, 9].Value := 'Modified By';
         Item[1, 9].ColumnWidth := 30;
         Item[1, 9].Borders[xlAround].Weight := xlThin;
         Item[1,10].Value := 'Modify Date';
         Item[1,10].ColumnWidth := 15;
         Item[1,10].Borders[xlAround].Weight := xlThin;
         Item[1,11].Value := 'Modify Time';
         Item[1,11].ColumnWidth := 15;
         Item[1,11].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end else begin
      with xlsSheet1.Range['A3','B3'] do begin
         Item[1,1].Value := 'Attribute';
         Item[1,1].ColumnWidth := 30;
         Item[1,1].Borders[xlAround].Weight := xlThin;
         Item[1,2].Value := 'Value';
         Item[1,2].ColumnWidth := 100;
         Item[1,2].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

   row1 := 4;

//--- Process the Prefix Records

   RecCount := StrToInt(RecordSet.ChildValues['Count']);

   SectionSet := SectionSet.NextSibling;
   RecordSet  := SectionSet.ChildNodes.First;

//--- Now process the records in this section if any

   if RecCount > 0 then begin
      for idx1 := 1 to RecCount do begin
         if (MultiFiles <> '0') then begin
            with xlsSheet1.RCRange[row1,1,row1,11] do begin
               Item[1, 1].Value := ReplaceXML(RecordSet.ChildValues['Prefix']);
               Item[1, 1].Borders[xlAround].Weight := xlThin;
               Item[1, 2].Value := StrToInt(ReplaceXML(RecordSet.ChildValues['Count']));
               Item[1, 2].HorizontalAlignment := xlHAlignRight;
               Item[1, 2].Borders[xlAround].Weight := xlThin;
               Item[1, 3].Value := ReplaceXML(RecordSet.ChildValues['Description']);
               Item[1, 3].Borders[xlAround].Weight := xlThin;
               Item[1, 4].Value := ReplaceXML(RecordSet.ChildValues['Location']);
               Item[1, 4].Borders[xlAround].Weight := xlThin;
               Item[1, 5].Value := ReplaceXML(RecordSet.ChildValues['Template']);
               Item[1, 5].Borders[xlAround].Weight := xlThin;
               Item[1, 6].Value := ReplaceXML(RecordSet.ChildValues['Creator']);
               Item[1, 6].Borders[xlAround].Weight := xlThin;
               Item[1, 7].Value := ReplaceXML(RecordSet.ChildValues['CreateDate']);
               Item[1, 7].Borders[xlAround].Weight := xlThin;
               Item[1, 8].Value := ReplaceXML(RecordSet.ChildValues['CreateTime']);
               Item[1, 8].Borders[xlAround].Weight := xlThin;
               Item[1, 9].Value := ReplaceXML(RecordSet.ChildValues['Modifier']);
               Item[1, 9].Borders[xlAround].Weight := xlThin;
               Item[1,10].Value := ReplaceXML(RecordSet.ChildValues['ModifyDate']);
               Item[1,10].Borders[xlAround].Weight := xlThin;
               Item[1,11].Value := ReplaceXML(RecordSet.ChildValues['ModifyTime']);
               Item[1,11].Borders[xlAround].Weight := xlThin;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
               RecordSet := RecordSet.NextSibling;
               inc(row1);
         end else begin
            with xlsSheet1.RCRange[row1,1,row1 + 11,2] do begin
               Item[ 1,1].Value := 'Prefix';
               Item[ 1,1].Borders[xlAround].Weight := xlThin;
               Item[ 1,2].Value := ReplaceXML(RecordSet.ChildValues['Prefix']);
               Item[ 1,2].Borders[xlAround].Weight := xlThin;
               Item[ 2,1].Value := 'Count';
               Item[ 2,1].Borders[xlAround].Weight := xlThin;
               Item[ 2,2].Value := StrToInt(ReplaceXML(RecordSet.ChildValues['Count']));
               Item[ 2,2].Borders[xlAround].Weight := xlThin;
               Item[ 3,1].Value := 'Description';
               Item[ 3,1].Borders[xlAround].Weight := xlThin;
               Item[ 3,2].Value := ReplaceXML(RecordSet.ChildValues['Description']);
               Item[ 3,2].Borders[xlAround].Weight := xlThin;
               Item[ 4,1].Value := 'Location';
               Item[ 4,1].Borders[xlAround].Weight := xlThin;
               Item[ 4,2].Value := ReplaceXML(RecordSet.ChildValues['Location']);
               Item[ 4,2].Borders[xlAround].Weight := xlThin;
               Item[ 5,1].Value := 'Template';
               Item[ 5,1].Borders[xlAround].Weight := xlThin;
               Item[ 5,2].Value := ReplaceXML(RecordSet.ChildValues['Template']);
               Item[ 5,2].Borders[xlAround].Weight := xlThin;
               Item[ 6,1].Value := 'Created By';
               Item[ 6,1].Borders[xlAround].Weight := xlThin;
               Item[ 6,2].Value := ReplaceXML(RecordSet.ChildValues['Creator']);
               Item[ 6,2].Borders[xlAround].Weight := xlThin;
               Item[ 7,1].Value := 'Create Date';
               Item[ 7,1].Borders[xlAround].Weight := xlThin;
               Item[ 7,2].Value := ReplaceXML(RecordSet.ChildValues['CreateDate']);
               Item[ 7,2].Borders[xlAround].Weight := xlThin;
               Item[ 8,1].Value := 'Create Time';
               Item[ 8,1].Borders[xlAround].Weight := xlThin;
               Item[ 8,2].Value := ReplaceXML(RecordSet.ChildValues['CreateTime']);
               Item[ 8,2].Borders[xlAround].Weight := xlThin;
               Item[ 9,1].Value := 'Modified By';
               Item[ 9,1].Borders[xlAround].Weight := xlThin;
               Item[ 9,2].Value := ReplaceXML(RecordSet.ChildValues['Modifier']);
               Item[ 9,2].Borders[xlAround].Weight := xlThin;
               Item[10,1].Value := 'Modify Date';
               Item[10,1].Borders[xlAround].Weight := xlThin;
               Item[10,2].Value := ReplaceXML(RecordSet.ChildValues['ModifyDate']);
               Item[10,2].Borders[xlAround].Weight := xlThin;
               Item[11,1].Value := 'Modify Time';
               Item[11,1].Borders[xlAround].Weight := xlThin;
               Item[11,2].Value := ReplaceXML(RecordSet.ChildValues['ModifyTime']);
               Item[11,2].Borders[xlAround].Weight := xlThin;
               WrapText := true;
               VerticalAlignment := xlVAlignTop;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

//            xlsSheet1.PageSetup.FitToPagesWide := 1;
            row1 := row1 + 11;
         end;
      end;
   end else begin
      with xlsSheet1.RCRange[row1,1,row1,2] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := 'No Prefix Details found';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row1);
   end;

//--- Write the standard copyright notice

   inc(row1);
   with xlsSheet1.RCRange[row1,1,row1,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      WrapText  := false;
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet1.PageSetup.Orientation    := xlLandscape;
   xlsSheet1.PageSetup.PaperSize      := xlPaperA4;
   xlsSheet1.DisplayGridLines         := false;
   xlsSheet1.PageSetup.CenterFooter   := 'Page &P of &N';

   if (MultiFiles = '0') then
      xlsSheet1.PageSetup.FitToPagesWide := 1;

//--- Write the Excel file to disk

   xlsBook.SaveAs(FileName);
   XMLDoc.Active := false;
   xlsBook.Close;

   LogMsg('  Prefix List successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Prefix List by Email submitted...',True)
      else
         LogMsg('  Request to send generated Prefix List by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the Prefix List Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName),nil,nil,SW_SHOWNORMAL);
      LogMsg('  Request to open exported Prefix List for ''' + PChar(FileName) + ''' submitted...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Prefix List');

   DeleteFile(HostName);
end;

//---------------------------------------------------------------------------
// Export Client/Opposition Details
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_ClientDetails();
var
   row1, ThisType              : integer;
   DoLine                      : boolean;
   ThisFile, Filter, ThisField : string;
   xlsBook                     : IXLSWorkbook;
   xlsSheet1                   : IXLSWorksheet;

begin

   if (RunType = 13) then begin
      LogMsg(ord(PT_PROLOG),True,True,'Client Details');
      ThisType := 1;
   end else begin
      LogMsg(ord(PT_PROLOG),True,True,'Opposition Details');
      ThisType := 2;
   end;

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      lbProgress.Items.Add('Unexpected Data Base error: ' + ErrMsg);
      lbProgress.TopIndex := lbProgress.Items.Count - 1;
      lbProgress.Refresh;;
      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   Filter    := Parm07;
   ThisField := Parm10;

//--- Create the Excel Workbook

   xlsBook       := TXLSWorkbook.Create;

//-- Build the file name

   if (RunType = 13) then
      ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - Client Details (' + ThisField + ').xls'
   else
      ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - Opposition Details (' + ThisField + ').xls';

   txtDocument.Caption  := 'Exporting to: ' + FileName + ThisFile;
   txtDocument.Refresh;

//--- Process each User record in turn

   xlsSheet1      := xlsBook.WorkSheets.Add;

   if (RunType = 13) then
      xlsSheet1.Name := 'Client Details (' + ThisField + ')'
   else
      xlsSheet1.Name := 'Opposition Details (' + ThisField + ')';

//--- Insert the Header (1st line) and the Heading

   if (Filter = '%') then begin
      with xlsSheet1.Range['A1','AB1'] do begin
         if (RunType = 13) then
            Item[1,1].Value := CpyName + ': Client Details  (' + ThisField + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now())
         else
            Item[1,1].Value := CpyName + ': Opposition Details  (' + ThisField + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());

         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end else begin
      with xlsSheet1.Range['A1','B1'] do begin
         if (RunType = 13) then
            Item[1,1].Value := CpyName + ': Client Details  (' + ThisField + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now())
         else
            Item[1,1].Value := CpyName + ': Opposition Details  (' + ThisField + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());

         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end;

   if (Filter = '%') then begin
      with xlsSheet1.Range['A3','AB3'] do begin
         if (RunType = 13) then
            Item[1,1].Value := 'Customer Name'
         else
            Item[1,1].Value := 'Opposition Name';
         Item[1, 1].ColumnWidth := 30;
         Item[1, 1].Borders[xlAround].Weight := xlThin;
         Item[1, 2].Value := 'ID/Registration Number';
         Item[1, 2].ColumnWidth := 30;
         Item[1, 2].Borders[xlAround].Weight := xlThin;
         Item[1, 3].Value := 'Contact Person';
         Item[1, 3].ColumnWidth := 30;
         Item[1, 3].Borders[xlAround].Weight := xlThin;
         Item[1, 4].Value := 'VAT Number';
         Item[1, 4].ColumnWidth := 20;
         Item[1, 4].Borders[xlAround].Weight := xlThin;
         Item[1, 5].Value := 'Employer';
         Item[1, 5].ColumnWidth := 30;
         Item[1, 5].Borders[xlAround].Weight := xlThin;
         Item[1, 6].Value := 'Telephone Number';
         Item[1, 6].ColumnWidth := 20;
         Item[1, 6].Borders[xlAround].Weight := xlThin;
         Item[1, 7].Value := 'Fax Number';
         Item[1, 7].ColumnWidth := 20;
         Item[1, 7].Borders[xlAround].Weight := xlThin;
         Item[1, 8].Value := 'Cellphone Number';
         Item[1, 8].ColumnWidth := 20;
         Item[1, 8].Borders[xlAround].Weight := xlThin;
         Item[1, 9].Value := 'Work Number';
         Item[1, 9].ColumnWidth := 20;
         Item[1, 9].Borders[xlAround].Weight := xlThin;
         Item[1,10].Value := 'Personal Email Address';
         Item[1,10].ColumnWidth := 30;
         Item[1,10].Borders[xlAround].Weight := xlThin;
         Item[1,11].Value := 'Work Email Address';
         Item[1,11].ColumnWidth := 30;
         Item[1,11].Borders[xlAround].Weight := xlThin;
         Item[1,12].Value := 'Hourly Rate';
         Item[1,12].ColumnWidth := 15;
         Item[1,12].Borders[xlAround].Weight := xlThin;
         Item[1,13].Value := 'Home Address Line 1';
         Item[1,13].ColumnWidth := 30;
         Item[1,13].Borders[xlAround].Weight := xlThin;
         Item[1,14].Value := 'Home Address Line 2';
         Item[1,14].ColumnWidth := 30;
         Item[1,14].Borders[xlAround].Weight := xlThin;
         Item[1,15].Value := 'Home Address Line 3';
         Item[1,15].ColumnWidth := 30;
         Item[1,15].Borders[xlAround].Weight := xlThin;
         Item[1,16].Value := 'Home Address Line 4';
         Item[1,16].ColumnWidth := 30;
         Item[1,16].Borders[xlAround].Weight := xlThin;
         Item[1,17].Value := 'Postal Code';
         Item[1,17].ColumnWidth := 15;
         Item[1,17].Borders[xlAround].Weight := xlThin;
         Item[1,18].Value := 'Postal Address Line 1';
         Item[1,18].ColumnWidth := 30;
         Item[1,18].Borders[xlAround].Weight := xlThin;
         Item[1,19].Value := 'Postal Address Line 2';
         Item[1,19].ColumnWidth := 30;
         Item[1,19].Borders[xlAround].Weight := xlThin;
         Item[1,20].Value := 'Postal Address Line 3';
         Item[1,20].ColumnWidth := 30;
         Item[1,20].Borders[xlAround].Weight := xlThin;
         Item[1,21].Value := 'Postal Address Line 4';
         Item[1,21].ColumnWidth := 30;
         Item[1,21].Borders[xlAround].Weight := xlThin;
         Item[1,22].Value := 'Postal Code';
         Item[1,22].ColumnWidth := 15;
         Item[1,22].Borders[xlAround].Weight := xlThin;
         Item[1,23].Value := 'Work Address Line 1';
         Item[1,23].ColumnWidth := 30;
         Item[1,23].Borders[xlAround].Weight := xlThin;
         Item[1,24].Value := 'Work Address Line 2';
         Item[1,24].ColumnWidth := 30;
         Item[1,24].Borders[xlAround].Weight := xlThin;
         Item[1,25].Value := 'Work Address Line 3';
         Item[1,25].ColumnWidth := 30;
         Item[1,25].Borders[xlAround].Weight := xlThin;
         Item[1,26].Value := 'Work Address Line 4';
         Item[1,26].ColumnWidth := 30;
         Item[1,26].Borders[xlAround].Weight := xlThin;
         Item[1,27].Value := 'Postal Code';
         Item[1,27].ColumnWidth := 16;
         Item[1,27].Borders[xlAround].Weight := xlThin;
         Item[1,28].Value := 'Free Text';
         Item[1,28].ColumnWidth := 90;
         Item[1,28].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end else begin
      with xlsSheet1.Range['A3','B3'] do begin
         Item[1,1].Value := 'Attribute';
         Item[1,1].ColumnWidth := 30;
         Item[1,1].Borders[xlAround].Weight := xlThin;
         Item[1,2].Value := 'Value';
         Item[1,2].ColumnWidth := 120;
         Item[1,2].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

   row1 := 4;

//--- Get all the Client Detail records based on the provided filter

   if (GetClientDetails(Filter,ThisType) = false) then begin
      lbProgress.Items.Add('Unexpected Data Base error: ' + ErrMsg);
      lbProgress.TopIndex := lbProgress.Items.Count - 1;
      lbProgress.Refresh;;
      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Only process files that have records

   if (Query1.RecordCount > 0) then begin

      if (Filter = '%') then begin
         while Query1.Eof = false do begin
            with xlsSheet1.RCRange[row1,1,row1,28] do begin
               Item[1, 1].Value := ReplaceQuote(Query1.FieldByName('Cust_Customer').AsString);
               Item[1, 1].Borders[xlAround].Weight := xlThin;
               Item[1, 2].Value := ReplaceQuote(Query1.FieldByName('Cust_ID').AsString);
               Item[1, 2].Borders[xlAround].Weight := xlThin;
               Item[1, 3].Value := ReplaceQuote(Query1.FieldByName('Cust_Description').AsString);
               Item[1, 3].Borders[xlAround].Weight := xlThin;
               Item[1, 4].Value := ReplaceQuote(Query1.FieldByName('Cust_VATNum').AsString);
               Item[1, 4].Borders[xlAround].Weight := xlThin;
               Item[1, 5].Value := ReplaceQuote(Query1.FieldByName('Cust_Employer').AsString);
               Item[1, 5].Borders[xlAround].Weight := xlThin;
               Item[1, 6].Value := ReplaceQuote(Query1.FieldByName('Cust_Telephone').AsString);
               Item[1, 6].Borders[xlAround].Weight := xlThin;
               Item[1, 7].Value := ReplaceQuote(Query1.FieldByName('Cust_Fax').AsString);
               Item[1, 7].Borders[xlAround].Weight := xlThin;
               Item[1, 8].Value := ReplaceQuote(Query1.FieldByName('Cust_Cellphone').AsString);
               Item[1, 8].Borders[xlAround].Weight := xlThin;
               Item[1, 9].Value := ReplaceQuote(Query1.FieldByName('Cust_Worknum').AsString);
               Item[1, 9].Borders[xlAround].Weight := xlThin;
               Item[1,10].Value := ReplaceQuote(Query1.FieldByName('Cust_Persemail').AsString);
               Item[1,10].Borders[xlAround].Weight := xlThin;
               Item[1,11].Value := ReplaceQuote(Query1.FieldByName('Cust_Workemail').AsString);
               Item[1,11].Borders[xlAround].Weight := xlThin;
               Item[1,12].Value := ReplaceQuote(Query1.FieldByName('Cust_Rate').AsString);
               Item[1,12].Borders[xlAround].Weight := xlThin;
               Item[1,13].Value := ReplaceQuote(Query1.FieldByName('Cust_Address1').AsString);
               Item[1,13].Borders[xlAround].Weight := xlThin;
               Item[1,14].Value := ReplaceQuote(Query1.FieldByName('Cust_Address2').AsString);
               Item[1,14].Borders[xlAround].Weight := xlThin;
               Item[1,15].Value := ReplaceQuote(Query1.FieldByName('Cust_Address3').AsString);
               Item[1,15].Borders[xlAround].Weight := xlThin;
               Item[1,16].Value := ReplaceQuote(Query1.FieldByName('Cust_Address4').AsString);
               Item[1,16].Borders[xlAround].Weight := xlThin;
               Item[1,17].Value := ReplaceQuote(Query1.FieldByName('Cust_PostCode').AsString);
               Item[1,17].Borders[xlAround].Weight := xlThin;
               Item[1,18].Value := ReplaceQuote(Query1.FieldByName('Cust_Postal1').AsString);
               Item[1,18].Borders[xlAround].Weight := xlThin;
               Item[1,19].Value := ReplaceQuote(Query1.FieldByName('Cust_Postal2').AsString);
               Item[1,19].Borders[xlAround].Weight := xlThin;
               Item[1,20].Value := ReplaceQuote(Query1.FieldByName('Cust_Postal3').AsString);
               Item[1,20].Borders[xlAround].Weight := xlThin;
               Item[1,21].Value := ReplaceQuote(Query1.FieldByName('Cust_Postal4').AsString);
               Item[1,21].Borders[xlAround].Weight := xlThin;
               Item[1,22].Value := ReplaceQuote(Query1.FieldByName('Cust_PostalCode').AsString);
               Item[1,22].Borders[xlAround].Weight := xlThin;
               Item[1,23].Value := ReplaceQuote(Query1.FieldByName('Cust_Work1').AsString);
               Item[1,23].Borders[xlAround].Weight := xlThin;
               Item[1,24].Value := ReplaceQuote(Query1.FieldByName('Cust_Work2').AsString);
               Item[1,24].Borders[xlAround].Weight := xlThin;
               Item[1,25].Value := ReplaceQuote(Query1.FieldByName('Cust_Work3').AsString);
               Item[1,25].Borders[xlAround].Weight := xlThin;
               Item[1,26].Value := ReplaceQuote(Query1.FieldByName('Cust_Work4').AsString);
               Item[1,26].Borders[xlAround].Weight := xlThin;
               Item[1,27].Value := ReplaceQuote(Query1.FieldByName('Cust_Postwork').AsString);
               Item[1,27].Borders[xlAround].Weight := xlThin;
               Item[1,28].Value := ReplaceQuote(Query1.FieldByName('Cust_FreeText').AsString);
               Item[1,28].Borders[xlAround].Weight := xlThin;

               WrapText := True;
               VerticalAlignment := xlVAlignTop;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
               inc(row1);
               Query1.Next;
         end;
      end else begin
         with xlsSheet1.RCRange[row1,1,row1 + 27,2] do begin
            if (RunType = 13) then
               Item[1,1].Value := 'Customer Name'
            else
               Item[1,1].Value := 'Opposition Name';
            Item[ 1,1].Borders[xlAround].Weight := xlThin;
            Item[ 1,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Customer').AsString);
            Item[ 1,2].Borders[xlAround].Weight := xlThin;
            Item[ 2,1].Value := 'ID/Registration Number';
            Item[ 2,1].Borders[xlAround].Weight := xlThin;
            Item[ 2,2].Value := ReplaceQuote(Query1.FieldByName('Cust_ID').AsString);
            Item[ 2,2].Borders[xlAround].Weight := xlThin;
            Item[ 3,1].Value := 'Contact Person';
            Item[ 3,1].Borders[xlAround].Weight := xlThin;
            Item[ 3,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Description').AsString);
            Item[ 3,2].Borders[xlAround].Weight := xlThin;
            Item[ 4,1].Value := 'VAT Number';
            Item[ 4,1].Borders[xlAround].Weight := xlThin;
            Item[ 4,2].Value := ReplaceQuote(Query1.FieldByName('Cust_VATNum').AsString);
            Item[ 4,2].Borders[xlAround].Weight := xlThin;
            Item[ 5,1].Value := 'Employer';
            Item[ 5,1].Borders[xlAround].Weight := xlThin;
            Item[ 5,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Employer').AsString);
            Item[ 5,2].Borders[xlAround].Weight := xlThin;
            Item[ 6,1].Value := 'Telephone Number';
            Item[ 6,1].Borders[xlAround].Weight := xlThin;
            Item[ 6,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Telephone').AsString);
            Item[ 6,2].Borders[xlAround].Weight := xlThin;
            Item[ 7,1].Value := 'Fax Number';
            Item[ 7,1].Borders[xlAround].Weight := xlThin;
            Item[ 7,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Fax').AsString);
            Item[ 7,2].Borders[xlAround].Weight := xlThin;
            Item[ 8,1].Value := 'Cellphone Number';
            Item[ 8,1].Borders[xlAround].Weight := xlThin;
            Item[ 8,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Cellphone').AsString);
            Item[ 8,2].Borders[xlAround].Weight := xlThin;
            Item[ 9,1].Value := 'Work Number';
            Item[ 9,1].Borders[xlAround].Weight := xlThin;
            Item[ 9,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Worknum').AsString);
            Item[ 9,2].Borders[xlAround].Weight := xlThin;
            Item[10,1].Value := 'Personal Email Address';
            Item[10,1].Borders[xlAround].Weight := xlThin;
            Item[10,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Persemail').AsString);
            Item[10,2].Borders[xlAround].Weight := xlThin;
            Item[11,1].Value := 'Work Email Address';
            Item[11,1].Borders[xlAround].Weight := xlThin;
            Item[11,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Workemail').AsString);
            Item[11,2].Borders[xlAround].Weight := xlThin;
            Item[12,1].Value := 'Hourly Rate';
            Item[12,1].Borders[xlAround].Weight := xlThin;
            Item[12,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Rate').AsString);
            Item[12,2].Borders[xlAround].Weight := xlThin;
            Item[13,1].Value := 'Home Address Line 1';
            Item[13,1].Borders[xlAround].Weight := xlThin;
            Item[13,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Address1').AsString);
            Item[13,2].Borders[xlAround].Weight := xlThin;
            Item[14,1].Value := 'Home Address Line 2';
            Item[14,1].Borders[xlAround].Weight := xlThin;
            Item[14,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Address2').AsString);
            Item[14,2].Borders[xlAround].Weight := xlThin;
            Item[15,1].Value := 'Home Address Line 3';
            Item[15,1].Borders[xlAround].Weight := xlThin;
            Item[15,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Address3').AsString);
            Item[15,2].Borders[xlAround].Weight := xlThin;
            Item[16,1].Value := 'Home Address Line 4';
            Item[16,1].Borders[xlAround].Weight := xlThin;
            Item[16,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Address4').AsString);
            Item[16,2].Borders[xlAround].Weight := xlThin;
            Item[17,1].Value := 'Postal Code';
            Item[17,1].Borders[xlAround].Weight := xlThin;
            Item[17,2].Value := ReplaceQuote(Query1.FieldByName('Cust_PostCode').AsString);
            Item[17,2].Borders[xlAround].Weight := xlThin;
            Item[18,1].Value := 'Postal Address Line 1';
            Item[18,1].Borders[xlAround].Weight := xlThin;
            Item[18,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Postal1').AsString);
            Item[18,2].Borders[xlAround].Weight := xlThin;
            Item[19,1].Value := 'Postal Address Line 2';
            Item[19,1].Borders[xlAround].Weight := xlThin;
            Item[19,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Postal2').AsString);
            Item[19,2].Borders[xlAround].Weight := xlThin;
            Item[20,1].Value := 'Postal Address Line 3';
            Item[20,1].Borders[xlAround].Weight := xlThin;
            Item[20,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Postal3').AsString);
            Item[20,2].Borders[xlAround].Weight := xlThin;
            Item[21,1].Value := 'Postal Address Line 4';
            Item[21,1].Borders[xlAround].Weight := xlThin;
            Item[21,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Postal4').AsString);
            Item[21,2].Borders[xlAround].Weight := xlThin;
            Item[22,1].Value := 'Postal Code';
            Item[22,1].Borders[xlAround].Weight := xlThin;
            Item[22,2].Value := ReplaceQuote(Query1.FieldByName('Cust_PostalCode').AsString);
            Item[22,2].Borders[xlAround].Weight := xlThin;
            Item[23,1].Value := 'Work Address Line 1';
            Item[23,1].Borders[xlAround].Weight := xlThin;
            Item[23,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Work1').AsString);
            Item[23,2].Borders[xlAround].Weight := xlThin;
            Item[24,1].Value := 'Work Address Line 2';
            Item[24,1].Borders[xlAround].Weight := xlThin;
            Item[24,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Work2').AsString);
            Item[24,2].Borders[xlAround].Weight := xlThin;
            Item[25,1].Value := 'work Address Line 3';
            Item[25,1].Borders[xlAround].Weight := xlThin;
            Item[25,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Work3').AsString);
            Item[25,2].Borders[xlAround].Weight := xlThin;
            Item[26,1].Value := 'Work Address Line 4';
            Item[26,1].Borders[xlAround].Weight := xlThin;
            Item[26,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Work4').AsString);
            Item[26,2].Borders[xlAround].Weight := xlThin;
            Item[27,1].Value := 'Postal Code';
            Item[27,1].Borders[xlAround].Weight := xlThin;
            Item[27,2].Value := ReplaceQuote(Query1.FieldByName('Cust_Postwork').AsString);
            Item[27,2].Borders[xlAround].Weight := xlThin;
            Item[28,1].Value := 'Free Text';
            Item[28,1].Borders[xlAround].Weight := xlThin;
            Item[28,2].Value := ReplaceQuote(Query1.FieldByName('Cust_FreeText').AsString);
            Item[28,2].Borders[xlAround].Weight := xlThin;

            WrapText := True;
            VerticalAlignment := xlVAlignTop;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         row1 := row1 + 28;
      end;
   end else begin
      with xlsSheet1.RCRange[row1,1,row1,2] do begin
         Borders[xlAround].Weight := xlThin;

         if (RunType = 13) then
            Item[1,1].Value := 'No Client Details found'
         else
            Item[1,1].Value := 'No Opposition Details found';

         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row1);
   end;

//--- Write the standard copyright notice

   inc(row1);
   with xlsSheet1.RCRange[row1,1,row1,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet1.PageSetup.Orientation := xlLandscape;
   xlsSheet1.PageSetup.PaperSize := xlPaperA4;
   xlsSheet1.DisplayGridLines := false;
   xlsSheet1.PageSetup.CenterFooter := 'Page &P of &N';

   if (Filter <> '%') then
      xlsSheet1.PageSetup.FitToPagesWide := 1;

//--- Write the Excel file to disk

   xlsBook.SaveAs(FileName + ThisFile);
   xlsBook.Close;

   if (RunType = 13) then
      LogMsg('  Client Details successfully processed...',True)
   else
      LogMSg('  Opposition Details successfully processed...',True);

   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName + ThisFile,ord(PT_NORMAL)) = true) then begin
         if (RunType = 13) then
            LogMsg('  Request to send generated Client Details by Email submitted...',True)
         else
            LogMsg('  Request to send generated Opposition Details by Email submitted...',True);

         DoLine := True;
      end else begin
         if (RunType = 13) then
            LogMsg('  Request to send generated Clients Details by Email not submitted...',True)
         else
            LogMsg('  Request to send generated Opposition Details by Email submitted...',True);

         DoLine := True;
      end;
   end;

//--- Now open the Client/Opposition List Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);

      if (RunType = 13) then
         LogMsg('  Request to open exported Client Details for ''' + PChar(FileName + ThisFile) + ''' submitted...',True)
      else
         LogMsg('  Request to open exported Opposition Details for ''' + PChar(FileName + ThisFile) + ''' submitted...',True);

      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ' ,True);

   if (RunType = 13) then
      LogMsg(ord(PT_EPILOG),False,False,'Client Details')
   else
      LogMsg(ord(PT_EPILOG),False,False,'Opposition Details');

   Close_Connection;
end;

//---------------------------------------------------------------------------
// Export File Details
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_FileDetails();
var
   idx1, row, RecCount, RowsPerPage, Pages, PageRow : integer;
   PageBreak, FirstPage, DoLine                     : boolean;
   MultiFiles, ThisFile                             : string;
   xlsBook                                          : IXLSWorkbook;
   xlsSheet                                         : IXLSWorksheet;
   ExportSet, SectionSet, RecordSet                 : IXMLNode;

begin

   LogMsg(ord(PT_PROLOG),True,True,'File Details');

   txtDocument.Caption  := 'Generating to: ' + FileName;
   txtDocument.Refresh;

//--- Get Company specific information

   GetCpyVAT();

   MultiFiles := BoolToStr(ShowRelated);

//--- Open the XML file and process the content

   XMLDoc.Active := false;
   XMLDoc.LoadFromFile(HostName);
   XMLDoc.Active := true;

   ExportSet  := XMLDoc.DocumentElement;
   SectionSet := ExportSet.ChildNodes.First;
   RecordSet  := SectionSet;

   ThisFile := RecordSet.ChildValues['File'];

//--- Create the Excel spreadsheet

   xlsBook  := TXLSWorkbook.Create;
   xlsSheet := xlsBook.WorkSheets.Add;

   if (MultiFiles <> '0') then
      xlsSheet.Name := 'File Export'
   else
      xlsSheet.Name := ThisFile;

//--- Process the File Records

   RecCount := StrToInt(RecordSet.ChildValues['Count']);

   ThisCount := 1;
   ThisMax   := IntToStr(RecCount);

   prbProgress.Max := RecCount;

   SectionSet := SectionSet.NextSibling;
   RecordSet  := SectionSet.ChildNodes.First;

//--- Now process the records in this section if any

   if RecCount > 0 then begin

//--- Perform a Page break if necessary


      Pages     := 0;
      FirstPage := true;
      PageBreak := true;

      for idx1 := 1 to RecCount do begin

         if (PageBreak = True) then
            DoPageBreak(ord(PB_FILEDETAILS),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,MultiFiles);

         if (MultiFiles <> '0') then begin

            txtError.Text := 'Processing: ' + ReplaceXML(RecordSet.ChildValues['File']);
            txtError.Refresh;
            prbProgress.StepIt;
            prbProgress.Refresh;

            stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
            stCount.Refresh;
            inc(ThisCount);

            with xlsSheet.RCRange[row,1,row,26] do begin
               Item[1, 1].Value := ReplaceXML(RecordSet.ChildValues['File']);
               Item[1, 1].Borders[xlAround].Weight := xlThin;
               Item[1, 2].Value := ReplaceXML(RecordSet.ChildValues['Descrip']);
               Item[1, 2].Borders[xlAround].Weight := xlThin;
               Item[1, 3].Value := ReplaceXML(RecordSet.ChildValues['DiaryDate']);
               Item[1, 3].Borders[xlAround].Weight := xlThin;
               Item[1, 4].Value := ReplaceXML(RecordSet.ChildValues['Closed']);
               Item[1, 4].Borders[xlAround].Weight := xlThin;
               Item[1, 5].Value := ReplaceXML(RecordSet.ChildValues['Settled']);
               Item[1, 5].Borders[xlAround].Weight := xlThin;
               Item[1, 6].Value := ReplaceXML(RecordSet.ChildValues['CaseNum']);
               Item[1, 6].Borders[xlAround].Weight := xlThin;
               Item[1, 7].Value := ReplaceXML(RecordSet.ChildValues['Court']);
               Item[1, 7].Borders[xlAround].Weight := xlThin;
               Item[1, 8].Value := GetFileDetails(ReplaceXML(RecordSet.ChildValues['Counsel']),2);
               Item[1, 8].Borders[xlAround].Weight := xlThin;
               Item[1, 9].Value := ReplaceXML(RecordSet.ChildValues['Owner']);
               Item[1, 9].Borders[xlAround].Weight := xlThin;
               Item[1,10].Value := ReplaceXML(RecordSet.ChildValues['Alert']);
               Item[1,10].Borders[xlAround].Weight := xlThin;
               Item[1,11].Value := ReplaceXML(RecordSet.ChildValues['AlertDate']);
               Item[1,11].Borders[xlAround].Weight := xlThin;
               Item[1,12].Value := ReplaceXML(RecordSet.ChildValues['AReason']);
               Item[1,12].Borders[xlAround].Weight := xlThin;
               Item[1,13].Value := ReplaceXML(RecordSet.ChildValues['PAlert']);
               Item[1,13].Borders[xlAround].Weight := xlThin;
               Item[1,14].Value := ReplaceXML(RecordSet.ChildValues['Prescrip']);
               Item[1,14].Borders[xlAround].Weight := xlThin;
               Item[1,15].Value := ReplaceXML(RecordSet.ChildValues['Client']);
               Item[1,15].Borders[xlAround].Weight := xlThin;
               Item[1,16].Value := ReplaceXML(RecordSet.ChildValues['Opposition']);
               Item[1,16].Borders[xlAround].Weight := xlThin;
               Item[1,17].Value := GetFileDetails(ReplaceXML(RecordSet.ChildValues['Corres']),3);
               Item[1,17].Borders[xlAround].Weight := xlThin;
               Item[1,18].Value := GetFileDetails(ReplaceXML(RecordSet.ChildValues['Oppose']),1);
               Item[1,18].Borders[xlAround].Weight := xlThin;
               Item[1,19].Value := ReplaceXML(RecordSet.ChildValues['Folder']);
               Item[1,19].Borders[xlAround].Weight := xlThin;
               Item[1,20].Value := ReplaceXML(RecordSet.ChildValues['FileType']);
               Item[1,20].Borders[xlAround].Weight := xlThin;
               Item[1,21].Value := Format('R %.2f',[StrToFloat(ReplaceXML(RecordSet.ChildValues['Rate']))]);
               Item[1,21].Borders[xlAround].Weight := xlThin;
               Item[1,21].HorizontalAlignment := xlHAlignRight;
               Item[1,22].Value := ReplaceXML(RecordSet.ChildValues['Related']);
               Item[1,22].Borders[xlAround].Weight := xlThin;
               Item[1,23].Value := GetFileDetails(ReplaceXML(RecordSet.ChildValues['Sheriff']),4);
               Item[1,23].Borders[xlAround].Weight := xlThin;
               Item[1,24].Value := ReplaceXML(RecordSet.ChildValues['Free1']);
               Item[1,24].Borders[xlAround].Weight := xlThin;
               Item[1,25].Value := ReplaceXML(RecordSet.ChildValues['Free2']);
               Item[1,25].Borders[xlAround].Weight := xlThin;
               Item[1,26].Value := ReplaceXML(RecordSet.ChildValues['Free3']);
               Item[1,26].Borders[xlAround].Weight := xlThin;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            RecordSet := RecordSet.NextSibling;
            inc(row);
            inc(PageRow);

//--- If we've reached the maximum rows per page then it is PageBreak time.

            if (PageRow >= RowsPerPage) then begin
               PageBreak := True;
               DoPageBreak(ord(PB_FILEDETAILS),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,MultiFiles);
            end;

         end else begin

            txtError.Text := 'Processing: ' + ReplaceXML(RecordSet.ChildValues['File']);
            txtError.Refresh;
            prbProgress.StepIt;
            prbProgress.Refresh;

            stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
            stCount.Refresh;
            inc(ThisCount);

            with xlsSheet.RCRange[row,1,row + 26,2] do begin
               Item[ 1,1].Value := 'File';
               Item[ 1,1].Borders[xlAround].Weight := xlThin;
               Item[ 1,2].Value := ReplaceXML(RecordSet.ChildValues['File']);
               Item[ 1,2].Borders[xlAround].Weight := xlThin;
               Item[ 2,1].Value := 'Diary Date';
               Item[ 2,1].Borders[xlAround].Weight := xlThin;
               Item[ 2,2].Value := ReplaceXML(RecordSet.ChildValues['DiaryDate']);
               Item[ 2,2].Borders[xlAround].Weight := xlThin;
               Item[ 3,1].Value := 'Closed';
               Item[ 3,1].Borders[xlAround].Weight := xlThin;
               Item[ 3,2].Value := ReplaceXML(RecordSet.ChildValues['Closed']);
               Item[ 3,2].Borders[xlAround].Weight := xlThin;
               Item[ 4,1].Value := 'Settled';
               Item[ 4,1].Borders[xlAround].Weight := xlThin;
               Item[ 4,2].Value := ReplaceXML(RecordSet.ChildValues['Settled']);
               Item[ 4,2].Borders[xlAround].Weight := xlThin;
               Item[ 5,1].Value := 'Client Description';
               Item[ 5,1].Borders[xlAround].Weight := xlThin;
               Item[ 5,2].Value := ReplaceXML(RecordSet.ChildValues['Client']);
               Item[ 5,2].Borders[xlAround].Weight := xlThin;
               Item[ 6,1].Value := 'Opposition Description';
               Item[ 6,1].Borders[xlAround].Weight := xlThin;
               Item[ 6,2].Value := ReplaceXML(RecordSet.ChildValues['Opposition']);
               Item[ 6,2].Borders[xlAround].Weight := xlThin;
               Item[ 7,1].Value := 'File Owner';
               Item[ 7,1].Borders[xlAround].Weight := xlThin;
               Item[ 7,2].Value := ReplaceXML(RecordSet.ChildValues['Owner']);
               Item[ 7,2].Borders[xlAround].Weight := xlThin;
               Item[ 8,1].Value := 'Case Number';
               Item[ 8,1].Borders[xlAround].Weight := xlThin;
               Item[ 8,2].Value := ReplaceXML(RecordSet.ChildValues['CaseNum']);
               Item[ 8,2].Borders[xlAround].Weight := xlThin;
               Item[ 9,1].Value := 'Court';
               Item[ 9,1].Borders[xlAround].Weight := xlThin;
               Item[ 9,2].Value := ReplaceXML(RecordSet.ChildValues['Court']);
               Item[ 9,2].Borders[xlAround].Weight := xlThin;
               Item[10,1].Value := 'Alert Date';
               Item[10,1].Borders[xlAround].Weight := xlThin;
               Item[10,2].Value := ReplaceXML(RecordSet.ChildValues['AlertDate']);
               Item[10,2].Borders[xlAround].Weight := xlThin;
               Item[11,1].Value := 'Alert Set';
               Item[11,1].Borders[xlAround].Weight := xlThin;
               Item[11,2].Value := ReplaceXML(RecordSet.ChildValues['Alert']);
               Item[11,2].Borders[xlAround].Weight := xlThin;
               Item[12,1].Value := 'Alert Reason';
               Item[12,1].Borders[xlAround].Weight := xlThin;
               Item[12,2].Value := ReplaceXML(RecordSet.ChildValues['AReason']);
               Item[12,2].Borders[xlAround].Weight := xlThin;
               Item[13,1].Value := 'Prescription Date';
               Item[13,1].Borders[xlAround].Weight := xlThin;
               Item[13,2].Value := ReplaceXML(RecordSet.ChildValues['Prescrip']);
               Item[13,2].Borders[xlAround].Weight := xlThin;
               Item[14,1].Value := 'Prescription Alert';
               Item[14,1].Borders[xlAround].Weight := xlThin;
               Item[14,2].Value := ReplaceXML(RecordSet.ChildValues['PAlert']);
               Item[14,2].Borders[xlAround].Weight := xlThin;
               Item[15,1].Value := 'File Description';
               Item[15,1].Borders[xlAround].Weight := xlThin;
               Item[15,2].Value := ReplaceXML(RecordSet.ChildValues['Descrip']);
               Item[15,2].Borders[xlAround].Weight := xlThin;
               Item[16,1].Value := 'File Folder';
               Item[16,1].Borders[xlAround].Weight := xlThin;
               Item[16,2].Value := ReplaceXML(RecordSet.ChildValues['Folder']);
               Item[16,2].Borders[xlAround].Weight := xlThin;
               Item[17,1].Value := 'File Type';
               Item[17,1].Borders[xlAround].Weight := xlThin;
               Item[17,2].Value := ReplaceXML(RecordSet.ChildValues['FileType']);
               Item[17,2].Borders[xlAround].Weight := xlThin;
               Item[18,1].Value := 'Opposing Attorney';
               Item[18,1].Borders[xlAround].Weight := xlThin;
               Item[18,2].Value := GetFileDetails(ReplaceXML(RecordSet.ChildValues['Oppose']),1);
               Item[18,2].Borders[xlAround].Weight := xlThin;
               Item[19,1].Value := 'Counsel';
               Item[19,1].Borders[xlAround].Weight := xlThin;
               Item[19,2].Value := GetFileDetails(ReplaceXML(RecordSet.ChildValues['Counsel']),2);
               Item[19,2].Borders[xlAround].Weight := xlThin;
               Item[20,1].Value := 'Correspondent';
               Item[20,1].Borders[xlAround].Weight := xlThin;
               Item[20,2].Value := GetFileDetails(ReplaceXML(RecordSet.ChildValues['Corres']),3);
               Item[20,2].Borders[xlAround].Weight := xlThin;
               Item[21,1].Value := 'File Rate';
               Item[21,1].Borders[xlAround].Weight := xlThin;
               Item[21,2].Value := Format('R %.2f',[StrToFloat(ReplaceXML(RecordSet.ChildValues['Rate']))]);
               Item[21,2].Borders[xlAround].Weight := xlThin;
               Item[22,1].Value := 'Related File';
               Item[22,1].Borders[xlAround].Weight := xlThin;
               Item[22,2].Value := ReplaceXML(RecordSet.ChildValues['Related']);
               Item[22,2].Borders[xlAround].Weight := xlThin;
               Item[23,1].Value := 'Sheriff Details';
               Item[23,1].Borders[xlAround].Weight := xlThin;
               Item[23,2].Value := GetFileDetails(ReplaceXML(RecordSet.ChildValues['Sheriff']),4);
               Item[23,2].Borders[xlAround].Weight := xlThin;
               Item[24,1].Value := 'Free Text 1';
               Item[24,1].Borders[xlAround].Weight := xlThin;
               Item[24,2].Value := ReplaceXML(RecordSet.ChildValues['Free1']);
               Item[24,2].Borders[xlAround].Weight := xlThin;
               Item[25,1].Value := 'Free Text 2';
               Item[25,1].Borders[xlAround].Weight := xlThin;
               Item[25,2].Value := ReplaceXML(RecordSet.ChildValues['Free2']);
               Item[25,2].Borders[xlAround].Weight := xlThin;
               Item[26,1].Value := 'Free Text 3';
               Item[26,1].Borders[xlAround].Weight := xlThin;
               Item[26,2].Value := ReplaceXML(RecordSet.ChildValues['Free3']);
               Item[26,2].Borders[xlAround].Weight := xlThin;
               WrapText := True;
               VerticalAlignment := xlVAlignTop;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;
      end;
   end else begin
      with xlsSheet.RCRange[row,1,row,1] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := 'No File Details found';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row);
   end;

//--- Write the standard copyright notice

   row := (Pages * lcGRows);
//   inc(row);
   with xlsSheet.RCRange[row,1,row,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet.PageSetup.Orientation := xlLandscape;
   xlsSheet.PageSetup.FitToPagesTall := Pages;
   xlsSheet.PageSetup.PaperSize := xlPaperA4;
   xlsSheet.DisplayGridLines := false;
   xlsSheet.PageSetup.CenterFooter := 'Page &P of &N';

   if (Multifiles = '0') then
      xlsSheet.PageSetup.FitToPagesWide := 1;

//--- Write the Excel file to disk

   xlsBook.SaveAs(FileName);
   XMLDoc.Active := false;
   xlsBook.Close;

   LogMsg('  File Details successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated File Details by Email submitted...',True)
      else
         LogMsg('  Request to send generated File Details by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the File List Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName),nil,nil,SW_SHOWNORMAL);
      LogMsg('  Request to open exported File Details for ''' + PChar(FileName) + ''' submitted...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'File Details');

   DeleteFile(HostName);
end;

//---------------------------------------------------------------------------
// Generate Invoice List
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_InvoiceList();
var
   idx1, row, RecCount, sema, PageRow, Pages, RowsPerPage     : integer;
   PageBreak, FirstPage, DoLine                               : boolean;
   ThisAB1F, ThisAB1T, ThisAB2F, ThisAB2T, ThisFill, ThisText : TColor;
   xlsBook                                                    : IXLSWorkbook;
   xlsSheet                                                   : IXLSWorksheet;
   ExportSet, SectionSet, RecordSet                           : IXMLNode;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Invoice List');

   txtDocument.Caption  := 'Generating to: ' + FileName;
   txtDocument.Refresh;

//--- Get Company specific information

   GetCpyVAT();

//--- Open the XML file and process the content

   XMLDoc.Active := false;
   XMLDoc.LoadFromFile(HostName);
   XMLDoc.Active := true;

   DeleteFile(HostName);

   ExportSet  := XMLDoc.DocumentElement;
   SectionSet := ExportSet.ChildNodes.First;
   RecordSet  := SectionSet;

//--- Create the Excel spreadsheet

   xlsBook       := TXLSWorkbook.Create;
   xlsSheet      := xlsBook.WorkSheets.Add;
   xlsSheet.Name := 'Invoice List';

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet.PageSetup.Orientation := xlLandscape;
   xlsSheet.PageSetup.FitToPagesWide := 1;
   xlsSheet.PageSetup.PaperSize := xlPaperA4;
   xlsSheet.DisplayGridLines := false;
   xlsSheet.PageSetup.CenterFooter := 'Page &P of &N';

//--- Set up to use alternate color blocks

   sema := 1;
   ThisAB1F := ColAB1F;
   ThisAB1T := ColAB1T;
   ThisAB2F := ColAB2F;
   ThisAB2T := ColAB2T;

//--- Process the Invoice Records

   RecCount := StrToInt(RecordSet.ChildValues['Count']);

   SectionSet := SectionSet.NextSibling;
   RecordSet  := SectionSet.ChildNodes.First;

   Pages     := 0;
   FirstPage := true;
   PageBreak := true;

//--- Now process the records in this section if any

   if RecCount > 0 then begin
      for idx1 := 1 to RecCount do begin

//--- Perform a Page break

         if (PageBreak = true) then begin

//--- Set the Page control variables

            if ((lcRepeatHeader = true) or (FirstPage = true)) then
               RowsPerPage := (lcGRows - 5)
            else
               RowsPerPage := (lcGRows - 2);

            if (FirstPage = True) then
               row := 1
            else
               row := (Pages * lcGRows) + 1;

            PageRow   := 1;
            PageBreak := False;
            inc(Pages);

//--- Check if the heading must be printed / repeated

            if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
               FirstPage := False;

//--- Write the Report Heading

               with xlsSheet.Range['A' + IntToStr(row), 'G' + IntToStr(row + 1)] do begin
                  Item[1,1].Value := CpyName;
                  Item[2,1].Value := 'Invoice List, Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ColDHF;
                  Font.Color := ColDHT;
                  Font.Bold := true;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;
               row := row + 3;
            end;

//--- Write the Heading

            with xlsSheet.Range['A' + IntToStr(row), 'G' + IntToStr(row)] do begin
               Item[1,1].Value := 'File';
               Item[1,1].ColumnWidth := 12;
               Item[1,1].Borders[xlAround].Weight := xlThin;
               Item[1,2].Value := 'Invoice';
               Item[1,2].ColumnWidth := 16;
               Item[1,2].Borders[xlAround].Weight := xlThin;
               Item[1,3].Value := 'Amount';
               Item[1,3].HorizontalAlignment := xlHAlignRight;
               Item[1,3].ColumnWidth := 12;
               Item[1,3].Borders[xlAround].Weight := xlThin;
               Item[1,4].Value := 'Start Date';
               Item[1,4].ColumnWidth := 12;
               Item[1,4].Borders[xlAround].Weight := xlThin;
               Item[1,5].Value := 'End Date';
               Item[1,5].ColumnWidth := 12;
               Item[1,5].Borders[xlAround].Weight := xlThin;
               Item[1,6].Value := 'Description';
               Item[1,6].ColumnWidth := lcGMRWidth - 76;
               Item[1,6].Borders[xlAround].Weight := xlThin;
               Item[1,7].Value := 'Send Out';
               Item[1,7].ColumnWidth := 12;
               Item[1,7].Borders[xlAround].Weight := xlThin;
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row);
         end;

         if (sema = 0) then begin;
            ThisFill := ThisAB1F;
            ThisText := ThisAB1T;
            sema := 1;
         end else begin
            ThisFill := ThisAB2F;
            ThisText := ThisAB2T;
            sema := 0;
         end;

         with xlsSheet.RCRange[row,1,row,7] do begin
            Item[1,1].Value := ReplaceXML(RecordSet.ChildValues['File']);
            Item[1,1].Borders[xlAround].Weight := xlThin;
            Item[1,2].Value := ReplaceXML(RecordSet.ChildValues['Invoice']);
            Item[1,2].Borders[xlAround].Weight := xlThin;
            Item[1,3].Value := StrToFloat(RecordSet.ChildValues['Amount']);
            Item[1,3].Borders[xlAround].Weight := xlThin;
            Item[1,3].HorizontalAlignment := xlHAlignRight;
            Item[1,3].NumberFormat := '#,##0.00_)';
            Item[1,4].Value := RecordSet.ChildValues['SDate'];
            Item[1,4].Borders[xlAround].Weight := xlThin;
            Item[1,5].Value := RecordSet.ChildValues['EDate'];
            Item[1,5].Borders[xlAround].Weight := xlThin;
            Item[1,6].Value := ReplaceXML(RecordSet.ChildValues['Descrip']);
            Item[1,6].Borders[xlAround].Weight := xlThin;
            Item[1,7].Value := ' ';
            Item[1,7].Borders[xlAround].Weight := xlThin;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := ThisFill;
            Font.Color := Thistext;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         RecordSet := RecordSet.NextSibling;

         inc(row);
         inc(PageRow);

         if (PageRow > RowsPerPage) then
            PageBreak := True;

      end;
   end else begin
      with xlsSheet.RCRange[row,1,row,9] do begin
         Item[1,1].Value := 'No records found';
         Borders[xlAround].Weight := xlThin;
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row);
   end;

   inc(row);

   with xlsSheet.RCRange[row,1,row,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('yyyy',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Save the workbook

   xlsBook.SaveAs(FileName);
   XMLDoc.Active := false;
   xlsBook.Close;

   LogMsg('  Invoice List successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Invoice List Export by Email submitted...',True)
      else
         LogMsg('  Request to send generated Invoice List Export by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the Invoice List Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName),nil,nil,SW_SHOWNORMAL);
      LogMsg('  Request to open exported File list for ''' + PChar(FileName) + ''' submitted...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Invoice List');

   DeleteFile(HostName);

end;

//---------------------------------------------------------------------------
// Export Safe Keeping Details
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_SafeDetails();
var
   idx1, row1, RecCount             : integer;
   DoLine                           : boolean;
   MultiFiles                       : string;
   xlsBook                          : IXLSWorkbook;
   xlsSheet1                        : IXLSWorksheet;
   ExportSet, SectionSet, RecordSet : IXMLNode;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Safe Keeping List');

   txtDocument.Caption  := 'Generating to: ' + FileName;
   txtDocument.Refresh;

//--- Get Company specific information

   GetCpyVAT();

   MultiFiles := BoolToStr(ShowRelated);

//--- Open the XML file and process the content

   XMLDoc.Active := false;
   XMLDoc.LoadFromFile(HostName);
   XMLDoc.Active := true;

   ExportSet  := XMLDoc.DocumentElement;
   SectionSet := ExportSet.ChildNodes.First;
   RecordSet  := SectionSet;

//--- Create the Excel spreadsheet

   xlsBook        := TXLSWorkbook.Create;
   xlsSheet1      := xlsBook.WorkSheets.Add;
   xlsSheet1.Name := RecordSet.ChildValues['File'];

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet1.PageSetup.Orientation := xlLandscape;
   xlsSheet1.PageSetup.PaperSize := xlPaperA4;
   xlsSheet1.DisplayGridLines := false;
   xlsSheet1.PageSetup.CenterFooter := 'Page &P of &N';

//--- Write the Report Heading

   if (MultiFiles = '1') then begin
      with xlsSheet1.Range['A1','M1'] do begin
         Item[1,1].Value := CpyName + ': Safe Keeping list (' + xlsSheet1.Name + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end else begin
      with xlsSheet1.Range['A1','B1'] do begin
         Item[1,1].Value := CpyName + ': Safe Keeping list (' + xlsSheet1.Name + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end;

//--- Write the File List Heading

   if (MultiFiles = '1') then begin
      with xlsSheet1.Range['A3','M3'] do begin
         Item[1, 1].Value := 'File';
         Item[1, 1].ColumnWidth := 15;
         Item[1, 1].Borders[xlAround].Weight := xlThin;
         Item[1, 2].Value := 'Customer';
         Item[1, 2].ColumnWidth := 80;
         Item[1, 2].Borders[xlAround].Weight := xlThin;
         Item[1, 3].Value := 'From Date';
         Item[1, 3].ColumnWidth := 15;
         Item[1, 3].Borders[xlAround].Weight := xlThin;
         Item[1, 4].Value := 'To Date';
         Item[1, 4].ColumnWidth := 15;
         Item[1, 4].Borders[xlAround].Weight := xlThin;
         Item[1, 5].Value := 'Description';
         Item[1, 5].ColumnWidth := 100;
         Item[1, 5].Borders[xlAround].Weight := xlThin;
         Item[1, 6].Value := 'Billing Reminder';
         Item[1, 6].ColumnWidth := 18;
         Item[1, 6].Borders[xlAround].Weight := xlThin;
         Item[1, 7].Value := 'Customer ID';
         Item[1, 7].ColumnWidth := 30;
         Item[1, 7].Borders[xlAround].Weight := xlThin;
         Item[1, 8].Value := 'Created By';
         Item[1, 8].ColumnWidth := 30;
         Item[1, 8].Borders[xlAround].Weight := xlThin;
         Item[1, 9].Value := 'Create Date';
         Item[1, 9].ColumnWidth := 15;
         Item[1, 9].Borders[xlAround].Weight := xlThin;
         Item[1,10].Value := 'Create Time';
         Item[1,10].ColumnWidth := 15;
         Item[1,10].Borders[xlAround].Weight := xlThin;
         Item[1,11].Value := 'Modified By';
         Item[1,11].ColumnWidth := 30;
         Item[1,11].Borders[xlAround].Weight := xlThin;
         Item[1,12].Value := 'Modify Date';
         Item[1,12].ColumnWidth := 15;
         Item[1,12].Borders[xlAround].Weight := xlThin;
         Item[1,13].Value := 'Modify Time';
         Item[1,13].ColumnWidth := 15;
         Item[1,13].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end else begin
      with xlsSheet1.Range['A3','B3'] do begin
         Item[1,1].Value := 'Attribute';
         Item[1,1].ColumnWidth := 30;
         Item[1,1].Borders[xlAround].Weight := xlThin;
         Item[1,2].Value := 'Value';
         Item[1,2].ColumnWidth := 100;
         Item[1,2].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

   row1 := 4;

//--- Process the Safe Keeping Records

   RecCount := StrToInt(RecordSet.ChildValues['Count']);

   SectionSet := SectionSet.NextSibling;
   RecordSet  := SectionSet.ChildNodes.First;

//--- Now process the records in this section if any

   if RecCount > 0 then begin
      for idx1 := 1 to RecCount do begin
         if (MultiFiles = '1') then begin
            with xlsSheet1.RCRange[row1,1,row1,13] do begin
               Item[1, 1].Value := ReplaceXML(RecordSet.ChildValues['File']);
               Item[1, 1].Borders[xlAround].Weight := xlThin;
               Item[1, 2].Value := ReplaceXML(RecordSet.ChildValues['Customer']);
               Item[1, 2].Borders[xlAround].Weight := xlThin;
               Item[1, 3].Value := ReplaceXML(RecordSet.ChildValues['FromDate']);
               Item[1, 3].Borders[xlAround].Weight := xlThin;
               Item[1, 4].Value := ReplaceXML(RecordSet.ChildValues['ToDate']);
               Item[1, 4].Borders[xlAround].Weight := xlThin;
               Item[1, 5].Value := ReplaceXML(RecordSet.ChildValues['Description']);
               Item[1, 5].Borders[xlAround].Weight := xlThin;
               Item[1, 6].Value := ReplaceXML(RecordSet.ChildValues['BillingReminder']);
               Item[1, 6].Borders[xlAround].Weight := xlThin;
               Item[1, 7].Value := ReplaceXML(RecordSet.ChildValues['CustID']);
               Item[1, 7].Borders[xlAround].Weight := xlThin;
               Item[1, 8].Value := ReplaceXML(RecordSet.ChildValues['Creator']);
               Item[1, 8].Borders[xlAround].Weight := xlThin;
               Item[1, 9].Value := ReplaceXML(RecordSet.ChildValues['CreateDate']);
               Item[1, 9].Borders[xlAround].Weight := xlThin;
               Item[1,10].Value := ReplaceXML(RecordSet.ChildValues['CreateTime']);
               Item[1,10].Borders[xlAround].Weight := xlThin;
               Item[1,11].Value := ReplaceXML(RecordSet.ChildValues['Modifier']);
               Item[1,11].Borders[xlAround].Weight := xlThin;
               Item[1,12].Value := ReplaceXML(RecordSet.ChildValues['ModifyDate']);
               Item[1,12].Borders[xlAround].Weight := xlThin;
               Item[1,13].Value := ReplaceXML(RecordSet.ChildValues['ModifyTime']);
               Item[1,13].Borders[xlAround].Weight := xlThin;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
               RecordSet := RecordSet.NextSibling;
               inc(row1);
         end else begin
            with xlsSheet1.RCRange[row1,1,row1 + 13,2] do begin
               Item[ 1,1].Value := 'File';
               Item[ 1,1].Borders[xlAround].Weight := xlThin;
               Item[ 1,2].Value := ReplaceXML(RecordSet.ChildValues['File']);
               Item[ 1,2].Borders[xlAround].Weight := xlThin;
               Item[ 2,1].Value := 'Customer';
               Item[ 2,1].Borders[xlAround].Weight := xlThin;
               Item[ 2,2].Value := ReplaceXML(RecordSet.ChildValues['Customer']);
               Item[ 2,2].HorizontalAlignment := xlHAlignLeft;
               Item[ 2,2].Borders[xlAround].Weight := xlThin;
               Item[ 3,1].Value := 'From Date';
               Item[ 3,1].Borders[xlAround].Weight := xlThin;
               Item[ 3,2].Value := ReplaceXML(RecordSet.ChildValues['FromDate']);
               Item[ 3,2].Borders[xlAround].Weight := xlThin;
               Item[ 4,1].Value := 'To Date';
               Item[ 4,1].Borders[xlAround].Weight := xlThin;
               Item[ 4,2].Value := ReplaceXML(RecordSet.ChildValues['ToDate']);
               Item[ 4,2].Borders[xlAround].Weight := xlThin;
               Item[ 5,1].Value := 'Description';
               Item[ 5,1].Borders[xlAround].Weight := xlThin;
               Item[ 5,2].Value := ReplaceXML(RecordSet.ChildValues['Description']);
               Item[ 5,2].Borders[xlAround].Weight := xlThin;
               Item[ 6,1].Value := 'Billing Reminder';
               Item[ 6,1].Borders[xlAround].Weight := xlThin;
               Item[ 6,2].Value := ReplaceXML(RecordSet.ChildValues['BillingReminder']);
               Item[ 6,2].Borders[xlAround].Weight := xlThin;
               Item[ 7,1].Value := 'Customer ID';
               Item[ 7,1].Borders[xlAround].Weight := xlThin;
               Item[ 7,2].Value := ReplaceXML(RecordSet.ChildValues['CustID']);
               Item[ 7,2].Borders[xlAround].Weight := xlThin;
               Item[ 8,1].Value := 'Created By';
               Item[ 8,1].Borders[xlAround].Weight := xlThin;
               Item[ 8,2].Value := ReplaceXML(RecordSet.ChildValues['Creator']);
               Item[ 8,2].Borders[xlAround].Weight := xlThin;
               Item[ 9,1].Value := 'Create Date';
               Item[ 9,1].Borders[xlAround].Weight := xlThin;
               Item[ 9,2].Value := ReplaceXML(RecordSet.ChildValues['CreateDate']);
               Item[ 9,2].Borders[xlAround].Weight := xlThin;
               Item[10,1].Value := 'Create Time';
               Item[10,1].Borders[xlAround].Weight := xlThin;
               Item[10,2].Value := ReplaceXML(RecordSet.ChildValues['CreateTime']);
               Item[10,2].Borders[xlAround].Weight := xlThin;
               Item[11,1].Value := 'Modified By';
               Item[11,1].Borders[xlAround].Weight := xlThin;
               Item[11,2].Value := ReplaceXML(RecordSet.ChildValues['Modifier']);
               Item[11,2].Borders[xlAround].Weight := xlThin;
               Item[12,1].Value := 'Modify Date';
               Item[12,1].Borders[xlAround].Weight := xlThin;
               Item[12,2].Value := ReplaceXML(RecordSet.ChildValues['ModifyDate']);
               Item[12,2].Borders[xlAround].Weight := xlThin;
               Item[13,1].Value := 'Modify Time';
               Item[13,1].Borders[xlAround].Weight := xlThin;
               Item[13,2].Value := ReplaceXML(RecordSet.ChildValues['ModifyTime']);
               Item[13,2].Borders[xlAround].Weight := xlThin;
               WrapText := true;
               VerticalAlignment := xlVAlignTop;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            xlsSheet1.PageSetup.FitToPagesWide := 1;
            row1 := row1 + 13;
         end;
      end;
   end else begin
      with xlsSheet1.RCRange[row1,1,row1,2] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := 'No Safe Keeping Details found';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row1);
   end;

//--- Write the standard copyright notice

   inc(row1);
   with xlsSheet1.RCRange[row1,1,row1,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      WrapText  := false;
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Write the Excel file to disk

   xlsBook.SaveAs(FileName);
   XMLDoc.Active := false;
   xlsBook.Close;

   LogMsg('  Safe Keeping List successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Safe Keeping List by Email submitted...',True)
      else
         LogMsg('  Request to send generated Safe Keeping List by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the Safe Keeping List Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName),nil,nil,SW_SHOWNORMAL);
      LogMsg('  Request to open exported Safe Keeping List for ''' + PChar(FileName) + ''' submitted...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Safe Keeping List');

   DeleteFile(HostName);
end;

//---------------------------------------------------------------------------
// Export Document Control Details
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_DocCntlDetails();
var
   idx1, row1, RecCount             : integer;
   DoLine                           : boolean;
   MultiFiles, ThisStatus           : string;
   xlsBook                          : IXLSWorkbook;
   xlsSheet1                        : IXLSWorksheet;
   ExportSet, SectionSet, RecordSet : IXMLNode;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Dcoument Control List');

   txtDocument.Caption  := 'Generating to: ' + FileName;
   txtDocument.Refresh;

//--- Get Company specific information

   GetCpyVAT();

   MultiFiles := BoolToStr(ShowRelated);

//--- Open the XML file and process the content

   XMLDoc.Active := false;
   XMLDoc.LoadFromFile(HostName);
   XMLDoc.Active := true;

   ExportSet  := XMLDoc.DocumentElement;
   SectionSet := ExportSet.ChildNodes.First;
   RecordSet  := SectionSet;

//--- Create the Excel spreadsheet

   xlsBook        := TXLSWorkbook.Create;
   xlsSheet1      := xlsBook.WorkSheets.Add;
   xlsSheet1.Name := RecordSet.ChildValues['File'];

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet1.PageSetup.Orientation := xlLandscape;
   xlsSheet1.PageSetup.PaperSize := xlPaperA4;
   xlsSheet1.DisplayGridLines := false;
   xlsSheet1.PageSetup.CenterFooter := 'Page &P of &N';

//--- Write the Report Heading

   if (MultiFiles = '1') then begin
      with xlsSheet1.Range['A1','T1'] do begin
         Item[1,1].Value := CpyName + ': Document Control list (' + xlsSheet1.Name + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end else begin
      with xlsSheet1.Range['A1','B1'] do begin
         Item[1,1].Value := CpyName + ': Document Control list (' + xlsSheet1.Name + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end;

//--- Write the File List Heading

   if (MultiFiles = '1') then begin
      with xlsSheet1.Range['A3','T3'] do begin
         Item[1, 1].Value := 'File';
         Item[1, 1].ColumnWidth := 15;
         Item[1, 1].Borders[xlAround].Weight := xlThin;
         Item[1, 2].Value := 'Description';
         Item[1, 2].ColumnWidth := 100;
         Item[1, 2].Borders[xlAround].Weight := xlThin;
         Item[1, 3].Value := 'Reason';
         Item[1, 3].ColumnWidth := 100;
         Item[1, 3].Borders[xlAround].Weight := xlThin;
         Item[1, 4].Value := 'Date Out';
         Item[1, 4].ColumnWidth := 15;
         Item[1, 4].Borders[xlAround].Weight := xlThin;
         Item[1, 5].Value := 'Time Out';
         Item[1, 5].ColumnWidth := 15;
         Item[1, 5].Borders[xlAround].Weight := xlThin;
         Item[1, 6].Value := 'Date In';
         Item[1, 6].ColumnWidth := 15;
         Item[1, 6].Borders[xlAround].Weight := xlThin;
         Item[1, 7].Value := 'Time In';
         Item[1, 7].ColumnWidth := 15;
         Item[1, 7].Borders[xlAround].Weight := xlThin;
         Item[1, 8].Value := 'Person';
         Item[1, 8].ColumnWidth := 40;
         Item[1, 8].Borders[xlAround].Weight := xlThin;
         Item[1, 9].Value := 'Status';
         Item[1, 9].ColumnWidth := 30;
         Item[1, 9].Borders[xlAround].Weight := xlThin;
         Item[1,10].Value := 'Facility';
         Item[1,10].ColumnWidth := 50;
         Item[1,10].Borders[xlAround].Weight := xlThin;
         Item[1,11].Value := 'Container';
         Item[1,11].ColumnWidth := 30;
         Item[1,11].Borders[xlAround].Weight := xlThin;
         Item[1,12].Value := 'Box';
         Item[1,12].ColumnWidth := 10;
         Item[1,12].Borders[xlAround].Weight := xlThin;
         Item[1,13].Value := 'Created By';
         Item[1,13].ColumnWidth := 40;
         Item[1,13].Borders[xlAround].Weight := xlThin;
         Item[1,14].Value := 'Create Date';
         Item[1,14].ColumnWidth := 20;
         Item[1,14].Borders[xlAround].Weight := xlThin;
         Item[1,15].Value := 'Create Time';
         Item[1,15].ColumnWidth := 20;
         Item[1,15].Borders[xlAround].Weight := xlThin;
         Item[1,16].Value := 'Modified By';
         Item[1,16].ColumnWidth := 40;
         Item[1,16].Borders[xlAround].Weight := xlThin;
         Item[1,17].Value := 'Modify Date';
         Item[1,17].ColumnWidth := 20;
         Item[1,17].Borders[xlAround].Weight := xlThin;
         Item[1,18].Value := 'Modify Time';
         Item[1,18].ColumnWidth := 20;
         Item[1,18].Borders[xlAround].Weight := xlThin;
         Item[1,19].Value := 'File Key';
         Item[1,19].ColumnWidth := 10;
         Item[1,19].Borders[xlAround].Weight := xlThin;
         Item[1,20].Value := 'Time Stamp';
         Item[1,20].ColumnWidth := 40;
         Item[1,20].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end else begin
      with xlsSheet1.Range['A3','B3'] do begin
         Item[1,1].Value := 'Attribute';
         Item[1,1].ColumnWidth := 30;
         Item[1,1].Borders[xlAround].Weight := xlThin;
         Item[1,2].Value := 'Value';
         Item[1,2].ColumnWidth := 100;
         Item[1,2].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

   row1 := 4;

//--- Process the Document Control Records

   RecCount := StrToInt(RecordSet.ChildValues['Count']);

   SectionSet := SectionSet.NextSibling;
   RecordSet  := SectionSet.ChildNodes.First;

//--- Now process the records in this section if any

   if RecCount > 0 then begin
      for idx1 := 1 to RecCount do begin
//--- Set the Status

         if (RecordSet.ChildValues['Status'] = '0') then
            ThisStatus := 'Checked Out'
         else if (RecordSet.ChildValues['Status'] = '1') then
            ThisStatus := 'Checked In'
         else if (RecordSet.ChildValues['Status'] = '2') then
            ThisStatus := 'Service'
         else if (RecordSet.ChildValues['Status'] = '3') then
            ThisStatus := 'Storage'
         else
            ThisStatus := 'Unknown';

         if (MultiFiles = '1') then begin
            with xlsSheet1.RCRange[row1,1,row1,20] do begin
               Item[1, 1].Value := ReplaceXML(RecordSet.ChildValues['File']);
               Item[1, 1].Borders[xlAround].Weight := xlThin;
               Item[1, 2].Value := ReplaceXML(RecordSet.ChildValues['Description']);
               Item[1, 2].Borders[xlAround].Weight := xlThin;
               Item[1, 3].Value := ReplaceXML(RecordSet.ChildValues['Reason']);
               Item[1, 3].Borders[xlAround].Weight := xlThin;
               Item[1, 4].Value := ReplaceXML(RecordSet.ChildValues['DateOut']);
               Item[1, 4].Borders[xlAround].Weight := xlThin;
               Item[1, 5].Value := ReplaceXML(RecordSet.ChildValues['TimeOut']);
               Item[1, 5].Borders[xlAround].Weight := xlThin;
               Item[1, 6].Value := ReplaceXML(RecordSet.ChildValues['DateIn']);
               Item[1, 6].Borders[xlAround].Weight := xlThin;
               Item[1, 7].Value := ReplaceXML(RecordSet.ChildValues['TimeIn']);
               Item[1, 7].Borders[xlAround].Weight := xlThin;
               Item[1, 8].Value := ReplaceXML(RecordSet.ChildValues['Person']);
               Item[1, 8].Borders[xlAround].Weight := xlThin;
               Item[1, 9].Value := ThisStatus;
               Item[1, 9].Borders[xlAround].Weight := xlThin;
               Item[1,10].Value := ReplaceXML(RecordSet.ChildValues['Facility']);
               Item[1,10].Borders[xlAround].Weight := xlThin;
               Item[1,11].Value := ReplaceXML(RecordSet.ChildValues['Container']);
               Item[1,11].Borders[xlAround].Weight := xlThin;
               Item[1,12].Value := ReplaceXML(RecordSet.ChildValues['Box']);
               Item[1,12].Borders[xlAround].Weight := xlThin;
               Item[1,13].Value := ReplaceXML(RecordSet.ChildValues['Creator']);
               Item[1,13].Borders[xlAround].Weight := xlThin;
               Item[1,14].Value := ReplaceXML(RecordSet.ChildValues['CreateDate']);
               Item[1,14].Borders[xlAround].Weight := xlThin;
               Item[1,15].Value := ReplaceXML(RecordSet.ChildValues['CreateTime']);
               Item[1,15].Borders[xlAround].Weight := xlThin;
               Item[1,16].Value := ReplaceXML(RecordSet.ChildValues['Modifier']);
               Item[1,16].Borders[xlAround].Weight := xlThin;
               Item[1,17].Value := ReplaceXML(RecordSet.ChildValues['ModifyDate']);
               Item[1,17].Borders[xlAround].Weight := xlThin;
               Item[1,18].Value := ReplaceXML(RecordSet.ChildValues['ModifyTime']);
               Item[1,18].Borders[xlAround].Weight := xlThin;
               Item[1,19].Value := ReplaceXML(RecordSet.ChildValues['FileKey']);
               Item[1,19].Borders[xlAround].Weight := xlThin;
               Item[1,20].Value := ReplaceXML(RecordSet.ChildValues['TimeStamp']);
               Item[1,20].Borders[xlAround].Weight := xlThin;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
               RecordSet := RecordSet.NextSibling;
               inc(row1);
         end else begin
            with xlsSheet1.RCRange[row1,1,row1 + 20,2] do begin
               Item[ 1,1].Value := 'Prefix';
               Item[ 1,1].Borders[xlAround].Weight := xlThin;
               Item[ 1,2].Value := ReplaceXML(RecordSet.ChildValues['File']);
               Item[ 1,2].Borders[xlAround].Weight := xlThin;
               Item[ 2,1].Value := 'Description';
               Item[ 2,1].Borders[xlAround].Weight := xlThin;
               Item[ 2,2].Value := ReplaceXML(RecordSet.ChildValues['Description']);
               Item[ 2,2].Borders[xlAround].Weight := xlThin;
               Item[ 3,1].Value := 'Reason';
               Item[ 3,1].Borders[xlAround].Weight := xlThin;
               Item[ 3,2].Value := ReplaceXML(RecordSet.ChildValues['Reason']);
               Item[ 3,2].Borders[xlAround].Weight := xlThin;
               Item[ 4,1].Value := 'Date Out';
               Item[ 4,1].Borders[xlAround].Weight := xlThin;
               Item[ 4,2].Value := ReplaceXML(RecordSet.ChildValues['DateOut']);
               Item[ 4,2].Borders[xlAround].Weight := xlThin;
               Item[ 5,1].Value := 'Time Out';
               Item[ 5,1].Borders[xlAround].Weight := xlThin;
               Item[ 5,2].Value := ReplaceXML(RecordSet.ChildValues['TimeOut']);
               Item[ 5,2].Borders[xlAround].Weight := xlThin;
               Item[ 6,1].Value := 'Date In';
               Item[ 6,1].Borders[xlAround].Weight := xlThin;
               Item[ 6,2].Value := ReplaceXML(RecordSet.ChildValues['DateIn']);
               Item[ 6,2].Borders[xlAround].Weight := xlThin;
               Item[ 7,1].Value := 'Time In';
               Item[ 7,1].Borders[xlAround].Weight := xlThin;
               Item[ 7,2].Value := ReplaceXML(RecordSet.ChildValues['TimeIn']);
               Item[ 7,2].Borders[xlAround].Weight := xlThin;
               Item[ 8,1].Value := 'Person';
               Item[ 8,1].Borders[xlAround].Weight := xlThin;
               Item[ 8,2].Value := ReplaceXML(RecordSet.ChildValues['Person']);
               Item[ 8,2].Borders[xlAround].Weight := xlThin;
               Item[ 9,1].Value := 'Status';
               Item[ 9,1].Borders[xlAround].Weight := xlThin;
               Item[ 9,2].Value := ThisStatus;
               Item[ 9,2].Borders[xlAround].Weight := xlThin;
               Item[10,1].Value := 'Storage Facility';
               Item[10,1].Borders[xlAround].Weight := xlThin;
               Item[10,2].Value := ReplaceXML(RecordSet.ChildValues['Facility']);
               Item[10,2].Borders[xlAround].Weight := xlThin;
               Item[11,1].Value := 'Container';
               Item[11,1].Borders[xlAround].Weight := xlThin;
               Item[11,2].Value := ReplaceXML(RecordSet.ChildValues['Container']);
               Item[11,2].Borders[xlAround].Weight := xlThin;
               Item[12,1].Value := 'Box';
               Item[12,1].Borders[xlAround].Weight := xlThin;
               Item[12,2].Value := ReplaceXML(RecordSet.ChildValues['Box']);
               Item[12,2].Borders[xlAround].Weight := xlThin;
               Item[13,1].Value := 'Created By';
               Item[13,1].Borders[xlAround].Weight := xlThin;
               Item[13,2].Value := ReplaceXML(RecordSet.ChildValues['Creator']);
               Item[13,2].Borders[xlAround].Weight := xlThin;
               Item[14,1].Value := 'Create Date';
               Item[14,1].Borders[xlAround].Weight := xlThin;
               Item[14,2].Value := ReplaceXML(RecordSet.ChildValues['CreateDate']);
               Item[14,2].Borders[xlAround].Weight := xlThin;
               Item[15,1].Value := 'Create Time';
               Item[15,1].Borders[xlAround].Weight := xlThin;
               Item[15,2].Value := ReplaceXML(RecordSet.ChildValues['CreateTime']);
               Item[15,2].Borders[xlAround].Weight := xlThin;
               Item[16,1].Value := 'Modified By';
               Item[16,1].Borders[xlAround].Weight := xlThin;
               Item[16,2].Value := ReplaceXML(RecordSet.ChildValues['Modifier']);
               Item[16,2].Borders[xlAround].Weight := xlThin;
               Item[17,1].Value := 'Modify Date';
               Item[17,1].Borders[xlAround].Weight := xlThin;
               Item[17,2].Value := ReplaceXML(RecordSet.ChildValues['ModifyDate']);
               Item[17,2].Borders[xlAround].Weight := xlThin;
               Item[18,1].Value := 'Modify Time';
               Item[18,1].Borders[xlAround].Weight := xlThin;
               Item[18,2].Value := ReplaceXML(RecordSet.ChildValues['ModifyTime']);
               Item[18,2].Borders[xlAround].Weight := xlThin;
               Item[19,1].Value := 'File Key';
               Item[19,1].Borders[xlAround].Weight := xlThin;
               Item[19,2].Value := ReplaceXML(RecordSet.ChildValues['FileKey']);
               Item[19,2].Borders[xlAround].Weight := xlThin;
               Item[20,1].Value := 'Time Stamp';
               Item[20,1].Borders[xlAround].Weight := xlThin;
               Item[20,2].Value := ReplaceXML(RecordSet.ChildValues['TimeStamp']);
               Item[20,2].Borders[xlAround].Weight := xlThin;
               WrapText := true;
               VerticalAlignment := xlVAlignTop;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            xlsSheet1.PageSetup.FitToPagesWide := 1;
            row1 := row1 + 20;
         end;
      end;
   end else begin
      with xlsSheet1.RCRange[row1,1,row1,2] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := 'No Document Control Details found';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row1);
   end;

//--- Write the standard copyright notice

   inc(row1);
   with xlsSheet1.RCRange[row1,1,row1,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      WrapText  := false;
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Write the Excel file to disk

   xlsBook.SaveAs(FileName);
   XMLDoc.Active := false;
   xlsBook.Close;

   LogMsg('  Document Control List successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := true;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Document Control List by Email submitted...',True)
      else
         LogMsg('  Request to send generated Document Control List by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the Document Control List Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName),nil,nil,SW_SHOWNORMAL);
      LogMsg('  Request to open exported Document Control List ist for ''' + PChar(FileName) + ''' submitted...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Document Control List');

   DeleteFile(HostName);
end;

//---------------------------------------------------------------------------
// Export Log Details
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_LogDetails();
var
   idx1, row1, RecCount             : integer;
   DoLine                           : boolean;
   MultiFiles                       : string;
   xlsBook                          : IXLSWorkbook;
   xlsSheet1                        : IXLSWorksheet;
   ExportSet, SectionSet, RecordSet : IXMLNode;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Log Details Export');

   txtDocument.Caption  := 'Generating to: ' + FileName;
   txtDocument.Refresh;

//--- Get Company specific information

   GetCpyVAT();

   MultiFiles := BoolToStr(ShowRelated);

//--- Open the XML file and process the content

   XMLDoc.Active := false;
   XMLDoc.LoadFromFile(HostName);
   XMLDoc.Active := true;

   ExportSet  := XMLDoc.DocumentElement;
   SectionSet := ExportSet.ChildNodes.First;
   RecordSet  := SectionSet;

//--- Create the Excel spreadsheet

   xlsBook        := TXLSWorkbook.Create;
   xlsSheet1      := xlsBook.WorkSheets.Add;
   xlsSheet1.Name := RecordSet.ChildValues['File'];

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet1.PageSetup.Orientation := xlLandscape;
   xlsSheet1.PageSetup.PaperSize := xlPaperA4;
   xlsSheet1.DisplayGridLines := false;
   xlsSheet1.PageSetup.CenterFooter := 'Page &P of &N';

//--- Write the Report Heading

   if (MultiFiles = '1') then begin
      with xlsSheet1.Range['A1','I1'] do begin
         Item[1,1].Value := CpyName + ': Log details (' + xlsSheet1.Name + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end else begin
      with xlsSheet1.Range['A1','B1'] do begin
         Item[1,1].Value := CpyName + ': Log details (' + xlsSheet1.Name + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end;

//--- Write the File List Heading

   if (MultiFiles = '1') then begin
      with xlsSheet1.Range['A3','I3'] do begin
         Item[1,1].Value := 'Date';
         Item[1,1].ColumnWidth := 12;
         Item[1,1].Borders[xlAround].Weight := xlThin;
         Item[1,2].Value := 'Time';
         Item[1,2].ColumnWidth := 12;
         Item[1,2].Borders[xlAround].Weight := xlThin;
         Item[1,3].Value := 'User';
         Item[1,3].ColumnWidth := 12;
         Item[1,3].Borders[xlAround].Weight := xlThin;
         Item[1,4].Value := 'Description';
         Item[1,4].ColumnWidth := 75;
         Item[1,4].Borders[xlAround].Weight := xlThin;
         Item[1,5].Value := 'Created By';
         Item[1,5].ColumnWidth := 12;
         Item[1,5].Borders[xlAround].Weight := xlThin;
         Item[1,6].Value := 'Create Date';
         Item[1,6].ColumnWidth := 12;
         Item[1,6].Borders[xlAround].Weight := xlThin;
         Item[1,7].Value := 'Create Time';
         Item[1,7].ColumnWidth := 12;
         Item[1,7].Borders[xlAround].Weight := xlThin;
         Item[1,8].Value := 'TimeStamp';
         Item[1,8].ColumnWidth := 36;
         Item[1,8].Borders[xlAround].Weight := xlThin;
         Item[1,9].Value := 'Key';
         Item[1,9].ColumnWidth := 8;
         Item[1,9].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end else begin
      with xlsSheet1.Range['A3','B3'] do begin
         Item[1,1].Value := 'Attribute';
         Item[1,1].ColumnWidth := 30;
         Item[1,1].Borders[xlAround].Weight := xlThin;
         Item[1,2].Value := 'Value';
         Item[1,2].ColumnWidth := 100;
         Item[1,2].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

   row1 := 4;

//--- Process the User Records

   RecCount := StrToInt(RecordSet.ChildValues['Count']);

   SectionSet := SectionSet.NextSibling;
   RecordSet  := SectionSet.ChildNodes.First;

//--- Now process the records in this section if any

   if RecCount > 0 then begin
      for idx1 := 1 to RecCount do begin
         if (MultiFiles = '1') then begin
            with xlsSheet1.RCRange[row1,1,row1,9] do begin
               Item[1,1].Value := ReplaceXML(RecordSet.ChildValues['Date']);
               Item[1,1].Borders[xlAround].Weight := xlThin;
               Item[1,2].Value := ReplaceXML(RecordSet.ChildValues['Time']);
               Item[1,2].Borders[xlAround].Weight := xlThin;
               Item[1,3].Value := ReplaceXML(RecordSet.ChildValues['User']);
               Item[1,3].Borders[xlAround].Weight := xlThin;
               Item[1,4].Value := ReplaceXML(RecordSet.ChildValues['Activity']);
               Item[1,4].Borders[xlAround].Weight := xlThin;
               Item[1,5].Value := ReplaceXML(RecordSet.ChildValues['Create_By']);
               Item[1,5].Borders[xlAround].Weight := xlThin;
               Item[1,6].Value := ReplaceXML(RecordSet.ChildValues['Create_Date']);
               Item[1,6].Borders[xlAround].Weight := xlThin;
               Item[1,7].Value := ReplaceXML(RecordSet.ChildValues['Create_Time']);
               Item[1,7].Borders[xlAround].Weight := xlThin;
               Item[1,8].Value := ReplaceXML(RecordSet.ChildValues['TimeStamp']);
               Item[1,8].Borders[xlAround].Weight := xlThin;
               Item[1,9].Value := ReplaceXML(RecordSet.ChildValues['Key']);
               Item[1,9].Borders[xlAround].Weight := xlThin;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
               RecordSet := RecordSet.NextSibling;
               inc(row1);
         end else begin
            with xlsSheet1.RCRange[row1,1,row1 + 9,2] do begin
               Item[1,1].Value := 'Date';
               Item[1,1].Borders[xlAround].Weight := xlThin;
               Item[1,2].Value := ReplaceXML(RecordSet.ChildValues['Date']);
               Item[1,2].Borders[xlAround].Weight := xlThin;
               Item[2,1].Value := 'Time';
               Item[2,1].Borders[xlAround].Weight := xlThin;
               Item[2,2].Value := ReplaceXML(RecordSet.ChildValues['Time']);
               Item[2,2].Borders[xlAround].Weight := xlThin;
               Item[3,1].Value := 'User';
               Item[3,1].Borders[xlAround].Weight := xlThin;
               Item[3,2].Value := ReplaceXML(RecordSet.ChildValues['User']);
               Item[3,2].Borders[xlAround].Weight := xlThin;
               Item[4,1].Value := 'Description';
               Item[4,1].Borders[xlAround].Weight := xlThin;
               Item[4,2].Value := ReplaceXML(RecordSet.ChildValues['Activity']);
               Item[4,2].Borders[xlAround].Weight := xlThin;
               Item[5,1].Value := 'Created By';
               Item[5,1].Borders[xlAround].Weight := xlThin;
               Item[5,2].Value := ReplaceXML(RecordSet.ChildValues['Create_By']);
               Item[5,2].Borders[xlAround].Weight := xlThin;
               Item[6,1].Value := 'Create Date';
               Item[6,1].Borders[xlAround].Weight := xlThin;
               Item[6,2].Value := ReplaceXML(RecordSet.ChildValues['Create_Date']);
               Item[6,2].Borders[xlAround].Weight := xlThin;
               Item[7,1].Value := 'Create Time';
               Item[7,1].Borders[xlAround].Weight := xlThin;
               Item[7,2].Value := ReplaceXML(RecordSet.ChildValues['Create_Time']);
               Item[7,2].Borders[xlAround].Weight := xlThin;
               Item[8,1].Value := 'Time Stamp';
               Item[8,1].Borders[xlAround].Weight := xlThin;
               Item[8,2].Value := ReplaceXML(RecordSet.ChildValues['TimeStamp']);
               Item[8,2].Borders[xlAround].Weight := xlThin;
               Item[9,1].Value := 'Key';
               Item[9,1].Borders[xlAround].Weight := xlThin;
               Item[9,2].Value := ReplaceXML(RecordSet.ChildValues['Key']);
               Item[9,2].Borders[xlAround].Weight := xlThin;
               WrapText := true;
               VerticalAlignment := xlVAlignTop;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

//            xlsSheet1.PageSetup.FitToPagesWide := 1;
            row1 := row1 + 17;
         end;
      end;
   end else begin
      with xlsSheet1.RCRange[row1,1,row1,2] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := 'No Log records found';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row1);
   end;

//--- Write the standard copyright notice

   inc(row1);
   with xlsSheet1.RCRange[row1,1,row1,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Write the Excel file to disk

   xlsSheet1.PageSetup.FitToPagesWide := 1;
   xlsBook.SaveAs(FileName);
   XMLDoc.Active := false;
   xlsBook.Close;


   LogMsg('  Log Details Export successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Log Details Export by Email submitted...',True)
      else
         LogMsg('  Request to send generated Log Details Export by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the Log Details Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName),nil,nil,SW_SHOWNORMAL);
      LogMsg('  Request to open exported Log Details for ''' + PChar(FileName) + ''' submitted...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Log Details Export');

   DeleteFile(HostName);
end;

//---------------------------------------------------------------------------
// Export User Details
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_UserDetails();
var
   idx1, row1, RecCount             : integer;
   DoLine                           : boolean;
   MultiFiles                       : string;
   xlsBook                          : IXLSWorkbook;
   xlsSheet1                        : IXLSWorksheet;
   ExportSet, SectionSet, RecordSet : IXMLNode;

begin

   LogMsg(ord(PT_PROLOG),True,True,'User Details');

   txtDocument.Caption  := 'Generating to: ' + FileName;
   txtDocument.Refresh;

//--- Get Company specific information

   GetCpyVAT();

   MultiFiles := BoolToStr(ShowRelated);

//--- Open the XML file and process the content

   XMLDoc.Active := false;
   XMLDoc.LoadFromFile(HostName);
   XMLDoc.Active := true;

   ExportSet  := XMLDoc.DocumentElement;
   SectionSet := ExportSet.ChildNodes.First;
   RecordSet  := SectionSet;

//--- Create the Excel spreadsheet

   xlsBook        := TXLSWorkbook.Create;
   xlsSheet1      := xlsBook.WorkSheets.Add;
   xlsSheet1.Name := RecordSet.ChildValues['File'];

//--- Write the Report Heading

   if (MultiFiles = '1') then begin
      with xlsSheet1.Range['A1','R1'] do begin
         Item[1,1].Value := CpyName + ': User Details (' + xlsSheet1.Name + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end else begin
      with xlsSheet1.Range['A1','B1'] do begin
         Item[1,1].Value := CpyName + ': User Details (' + xlsSheet1.Name + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 11;
      end;
   end;

//--- Write the File List Heading

   if (MultiFiles = '1') then begin
      with xlsSheet1.Range['A3','R3'] do begin
         Item[1, 1].Value := 'User Name';
         Item[1, 1].ColumnWidth := 30;
         Item[1, 1].Borders[xlAround].Weight := xlThin;
         Item[1, 2].Value := 'User Type';
         Item[1, 2].ColumnWidth := 15;
         Item[1, 2].Borders[xlAround].Weight := xlThin;
         Item[1, 3].Value := 'Description';
         Item[1, 3].ColumnWidth := 100;
         Item[1, 3].Borders[xlAround].Weight := xlThin;
         Item[1, 4].Value := 'Email Address';
         Item[1, 4].ColumnWidth := 50;
         Item[1, 4].Borders[xlAround].Weight := xlThin;
         Item[1, 5].Value := 'Fee Earner';
         Item[1, 5].ColumnWidth := 15;
         Item[1, 5].Borders[xlAround].Weight := xlThin;
         Item[1, 6].Value := 'Do Billing';
         Item[1, 6].ColumnWidth := 15;
         Item[1, 6].Borders[xlAround].Weight := xlThin;
         Item[1, 7].Value := 'Update Billing';
         Item[1, 7].ColumnWidth := 15;
         Item[1, 7].Borders[xlAround].Weight := xlThin;
         Item[1, 8].Value := 'Can Delete';
         Item[1, 8].ColumnWidth := 15;
         Item[1, 8].Borders[xlAround].Weight := xlThin;
         Item[1, 9].Value := 'Do Invoices';
         Item[1, 9].ColumnWidth := 15;
         Item[1, 9].Borders[xlAround].Weight := xlThin;
         Item[1,10].Value := 'Do Payments';
         Item[1,10].ColumnWidth := 15;
         Item[1,10].Borders[xlAround].Weight := xlThin;
         Item[1,11].Value := 'Blocked';
         Item[1,11].ColumnWidth := 15;
         Item[1,11].Borders[xlAround].Weight := xlThin;
         Item[1,12].Value := 'Unique';
         Item[1,12].ColumnWidth := 15;
         Item[1,12].Borders[xlAround].Weight := xlThin;
         Item[1,13].Value := 'Created By';
         Item[1,13].ColumnWidth := 40;
         Item[1,13].Borders[xlAround].Weight := xlThin;
         Item[1,14].Value := 'Create Date';
         Item[1,14].ColumnWidth := 20;
         Item[1,14].Borders[xlAround].Weight := xlThin;
         Item[1,15].Value := 'Create Time';
         Item[1,15].ColumnWidth := 20;
         Item[1,15].Borders[xlAround].Weight := xlThin;
         Item[1,16].Value := 'Modified By';
         Item[1,16].ColumnWidth := 40;
         Item[1,16].Borders[xlAround].Weight := xlThin;
         Item[1,17].Value := 'Modify Date';
         Item[1,17].ColumnWidth := 20;
         Item[1,17].Borders[xlAround].Weight := xlThin;
         Item[1,18].Value := 'Modify Time';
         Item[1,18].ColumnWidth := 20;
         Item[1,18].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end else begin
      with xlsSheet1.Range['A3','B3'] do begin
         Item[1,1].Value := 'Attribute';
         Item[1,1].ColumnWidth := 30;
         Item[1,1].Borders[xlAround].Weight := xlThin;
         Item[1,2].Value := 'Value';
         Item[1,2].ColumnWidth := 100;
         Item[1,2].Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

   row1 := 4;

//--- Process the User Records

   RecCount := StrToInt(RecordSet.ChildValues['Count']);

   SectionSet := SectionSet.NextSibling;
   RecordSet  := SectionSet.ChildNodes.First;

//--- Now process the records in this section if any

   if RecCount > 0 then begin
      for idx1 := 1 to RecCount do begin
         if (MultiFiles = '1') then begin
            with xlsSheet1.RCRange[row1,1,row1,18] do begin
               Item[1, 1].Value := ReplaceXML(RecordSet.ChildValues['Username']);
               Item[1, 1].Borders[xlAround].Weight := xlThin;
               Item[1, 2].Value := ReplaceXML(RecordSet.ChildValues['UserType']);
               Item[1, 2].Borders[xlAround].Weight := xlThin;
               Item[1, 3].Value := ReplaceXML(RecordSet.ChildValues['Description']);
               Item[1, 3].Borders[xlAround].Weight := xlThin;
               Item[1, 4].Value := ReplaceXML(RecordSet.ChildValues['EmailAddress']);
               Item[1, 4].Borders[xlAround].Weight := xlThin;
               Item[1, 5].Value := ReplaceXML(RecordSet.ChildValues['FeeEarner']);
               Item[1, 5].Borders[xlAround].Weight := xlThin;
               Item[1, 6].Value := ReplaceXML(RecordSet.ChildValues['DoBilling']);
               Item[1, 6].Borders[xlAround].Weight := xlThin;
               Item[1, 7].Value := ReplaceXML(RecordSet.ChildValues['UpdateBilling']);
               Item[1, 7].Borders[xlAround].Weight := xlThin;
               Item[1, 8].Value := ReplaceXML(RecordSet.ChildValues['AllowDelete']);
               Item[1, 8].Borders[xlAround].Weight := xlThin;
               Item[1, 9].Value := ReplaceXML(RecordSet.ChildValues['CreateInvoices']);
               Item[1, 9].Borders[xlAround].Weight := xlThin;
               Item[1,10].Value := ReplaceXML(RecordSet.ChildValues['ProcessPayments']);
               Item[1,10].Borders[xlAround].Weight := xlThin;
               Item[1,11].Value := ReplaceXML(RecordSet.ChildValues['Blocked']);
               Item[1,11].Borders[xlAround].Weight := xlThin;
               Item[1,12].Value := ReplaceXML(RecordSet.ChildValues['UniqueNumber']);
               Item[1,12].Borders[xlAround].Weight := xlThin;
               Item[1,13].Value := ReplaceXML(RecordSet.ChildValues['Creator']);
               Item[1,13].Borders[xlAround].Weight := xlThin;
               Item[1,14].Value := ReplaceXML(RecordSet.ChildValues['CreateDate']);
               Item[1,14].Borders[xlAround].Weight := xlThin;
               Item[1,15].Value := ReplaceXML(RecordSet.ChildValues['CreateTime']);
               Item[1,15].Borders[xlAround].Weight := xlThin;
               Item[1,16].Value := ReplaceXML(RecordSet.ChildValues['Modifier']);
               Item[1,16].Borders[xlAround].Weight := xlThin;
               Item[1,17].Value := ReplaceXML(RecordSet.ChildValues['ModifyDate']);
               Item[1,17].Borders[xlAround].Weight := xlThin;
               Item[1,18].Value := ReplaceXML(RecordSet.ChildValues['ModifyTime']);
               Item[1,18].Borders[xlAround].Weight := xlThin;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
               RecordSet := RecordSet.NextSibling;
               inc(row1);
         end else begin
            with xlsSheet1.RCRange[row1,1,row1 + 18,2] do begin
               Item[ 1,1].Value := 'User Name';
               Item[ 1,1].Borders[xlAround].Weight := xlThin;
               Item[ 1,2].Value := ReplaceXML(RecordSet.ChildValues['Username']);
               Item[ 1,2].Borders[xlAround].Weight := xlThin;
               Item[ 2,1].Value := 'User Type';
               Item[ 2,1].Borders[xlAround].Weight := xlThin;
               Item[ 2,2].Value := ReplaceXML(RecordSet.ChildValues['UserType']);
               Item[ 2,2].Borders[xlAround].Weight := xlThin;
               Item[ 3,1].Value := 'Description';
               Item[ 3,1].Borders[xlAround].Weight := xlThin;
               Item[ 3,2].Value := ReplaceXML(RecordSet.ChildValues['Description']);
               Item[ 3,2].Borders[xlAround].Weight := xlThin;
               Item[ 4,1].Value := 'Email Address';
               Item[ 4,1].Borders[xlAround].Weight := xlThin;
               Item[ 4,2].Value := ReplaceXML(RecordSet.ChildValues['EmailAddress']);
               Item[ 4,2].Borders[xlAround].Weight := xlThin;
               Item[ 5,1].Value := 'Fee Earner';
               Item[ 5,1].Borders[xlAround].Weight := xlThin;
               Item[ 5,2].Value := ReplaceXML(RecordSet.ChildValues['FeeEarner']);
               Item[ 5,2].Borders[xlAround].Weight := xlThin;
               Item[ 6,1].Value := 'Can Do Billing';
               Item[ 6,1].Borders[xlAround].Weight := xlThin;
               Item[ 6,2].Value := ReplaceXML(RecordSet.ChildValues['DoBilling']);
               Item[ 6,2].Borders[xlAround].Weight := xlThin;
               Item[ 7,1].Value := 'Can Update Billing';
               Item[ 7,1].Borders[xlAround].Weight := xlThin;
               Item[ 7,2].Value := ReplaceXML(RecordSet.ChildValues['UpdateBilling']);
               Item[ 7,2].Borders[xlAround].Weight := xlThin;
               Item[ 8,1].Value := 'Can Delete';
               Item[ 8,1].Borders[xlAround].Weight := xlThin;
               Item[ 8,2].Value := ReplaceXML(RecordSet.ChildValues['AllowDelete']);
               Item[ 8,2].Borders[xlAround].Weight := xlThin;
               Item[ 9,1].Value := 'Can Create Invoices';
               Item[ 9,1].Borders[xlAround].Weight := xlThin;
               Item[ 9,2].Value := ReplaceXML(RecordSet.ChildValues['CreateInvoices']);
               Item[ 9,2].Borders[xlAround].Weight := xlThin;
               Item[10,1].Value := 'Can Process Payments';
               Item[10,1].Borders[xlAround].Weight := xlThin;
               Item[10,2].Value := ReplaceXML(RecordSet.ChildValues['ProcessPayments']);
               Item[10,2].Borders[xlAround].Weight := xlThin;
               Item[11,1].Value := 'Blocked';
               Item[11,1].Borders[xlAround].Weight := xlThin;
               Item[11,2].Value := ReplaceXML(RecordSet.ChildValues['Blocked']);
               Item[11,2].Borders[xlAround].Weight := xlThin;
               Item[12,1].Value := 'Unique Number';
               Item[12,1].Borders[xlAround].Weight := xlThin;
               Item[12,2].Value := ReplaceXML(RecordSet.ChildValues['UniqueNumber']);
               Item[12,2].Borders[xlAround].Weight := xlThin;
               Item[13,1].Value := 'Created By';
               Item[13,1].Borders[xlAround].Weight := xlThin;
               Item[13,2].Value := ReplaceXML(RecordSet.ChildValues['Creator']);
               Item[13,2].Borders[xlAround].Weight := xlThin;
               Item[14,1].Value := 'Create Date';
               Item[14,1].Borders[xlAround].Weight := xlThin;
               Item[14,2].Value := ReplaceXML(RecordSet.ChildValues['CreateDate']);
               Item[14,2].Borders[xlAround].Weight := xlThin;
               Item[15,1].Value := 'Create Time';
               Item[15,1].Borders[xlAround].Weight := xlThin;
               Item[15,2].Value := ReplaceXML(RecordSet.ChildValues['CreateTime']);
               Item[15,2].Borders[xlAround].Weight := xlThin;
               Item[16,1].Value := 'Modified By';
               Item[16,1].Borders[xlAround].Weight := xlThin;
               Item[16,2].Value := ReplaceXML(RecordSet.ChildValues['Modifier']);
               Item[16,2].Borders[xlAround].Weight := xlThin;
               Item[17,1].Value := 'Modify Date';
               Item[17,1].Borders[xlAround].Weight := xlThin;
               Item[17,2].Value := ReplaceXML(RecordSet.ChildValues['ModifyDate']);
               Item[17,2].Borders[xlAround].Weight := xlThin;
               Item[18,1].Value := 'Modify Time';
               Item[18,1].Borders[xlAround].Weight := xlThin;
               Item[18,2].Value := ReplaceXML(RecordSet.ChildValues['ModifyTime']);
               Item[18,2].Borders[xlAround].Weight := xlThin;
               WrapText := true;
               VerticalAlignment := xlVAlignTop;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
            row1 := row1 + 17;
         end;
      end;
   end else begin
      with xlsSheet1.RCRange[row1,1,row1,2] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := 'No User Details found';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row1);
   end;

//--- Write the standard copyright notice

   inc(row1);
   inc(row1);

   with xlsSheet1.RCRange[row1,1,row1,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      WrapText  := false;
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet1.PageSetup.Orientation := xlLandscape;
   xlsSheet1.PageSetup.PaperSize := xlPaperA4;
   xlsSheet1.DisplayGridLines := false;
   xlsSheet1.PageSetup.CenterFooter := 'Page &P of &N';
   xlsSheet1.PageSetup.FitToPagesWide := 1;

//--- Write the Excel file to disk

   xlsBook.SaveAs(FileName);
   XMLDoc.Active := false;
   xlsBook.Close;

   LogMsg('  User Details successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated User Details by Email submitted...',True)
      else
         LogMsg('  Request to send generated User Details by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the User List Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName),nil,nil,SW_SHOWNORMAL);
      LogMsg('  Request to open exported User Details for ''' + PChar(FileName) + ''' submitted...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'User Details');

   DeleteFile(HostName);
end;

//---------------------------------------------------------------------------
// Generate Quotation
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Quotation();
var
   PageNum, ThisRows, ThisCol, ThisClass, Row, NumPages   : integer;
   idx1, idx2, idx3, RemainRows                           : integer;
   ThisAmount                                             : double;
   DoLine                                                 : boolean;
   ThisMsg, ThisVal, ThisItem, ThisFile, ThisStr, ThisVat : string;
   ThisInvoice, CurrentFile, S1, ThisQuote                : string;
   ThisDate                                               : TDateTime;
   xlsBook                                                : IXLSWorkbook;
   xlsSheet                                               : IXLSWorksheet;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Quotation');

   txtDocument.Caption  := 'Generating to: ' + FileName;
   txtDocument.Refresh;

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Get the Heading and Layout information

   GetLayout('Quote');

//--- Read the Quotation records from the datastore

   for idx1 := 0 to NumFiles - 1 do begin

      ThisQuote   := FileArray[idx1];
      CurrentFile := GetQuoteFile(ThisQuote);

      if (CurrentFile = '') then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      ThisVat := GetVATNum(CurrentFile);
      PageNum := 1;
      ThisDate := StrToDate(EDate);
      ThisFile := FormatDateTime('yyyyMMdd',ThisDate) + ' - Quote (' + CurrentFile + ').xls';
      txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;
      txtDocument.Refresh;

//--- Initialise the Summary Fields

      SummaryFees        := 0;
      SummaryDisburse    := 0;
      SummaryExpenses    := 0;
      SummaryVAT         := 0;

//--- Now process this Quote

      if ((GetQuote(CurrentFile,ThisQuote)) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      txtError.Text := 'Processing: ' + ThisQuote;
      txtError.Refresh;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

//--- Only process files that have records

      if (Query1.RecordCount > 0) then begin

//--- Open the Excel workbook template

         xlsBook := TXLSWorkbook.Create;
         xlsBook.Open(Template_Q);
         xlsSheet := xlsBook.ActiveSheet;
         xlsSheet.Name := 'Quotation (' + CurrentFile + ')';

//--- Clear everything but keep the Header information if lcShowHeader is set

         if (lcShowHeader = true) then
            xlsSheet.RCRange[lcHER + 1, 1, 999, lcHEC].Clear
         else
            xlsSheet.RCRange[1, 1, 999, lcSMaxCols].Clear;

//--- Insert the Page 1 Heading

         Generate_DocHeading(xlsSheet,PageNum,idx1,lcHER + 1,lcHEC,'Quote');

//--- Insert the Customer Information

         if (lcShowAddress = true) then begin

            GetAddress(CurrentFile);

            with xlsSheet.RCRange[lcASR,1,lcAER,lcHEC] do begin
               Item[1,lcASC].Value := Customer;
               Item[2,lcASC].Value := Address1;
               Item[3,lcASC].Value := Address2;
               Item[4,lcASC].Value := Address3;
               Item[5,lcASC].Value := Address4;
               Item[6,lcASC].Value := Address5;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;

//--- Insert the instruction information

         if (lcShowInstruct = true) then begin
            with xlsSheet.RCRange[lcISR,1,lcISR + 2,lcHEC] do begin
               Item[1,1].Value := 'Client:';
               Item[2,1].Value := 'Instruction:';
               Item[3,1].Value := 'VAT Num:';
               Item[1,lcISCD].Value := Customer;
               Item[2,lcISCD].Value := Descrip;
               Item[3,lcISCD].Value := ThisVat;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;

//--- Insert the Summary Information

         if (lcShowSummary = true) then begin
            with xlsSheet.RCRange[lcXSR,lcXSCL,lcXSR + 3,lcXSCD] do begin
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;

               Item[1,1].Value := 'Fees:';
               Item[1,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[1,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[1,2].Borders[xlEdgeTop].Weight := xlThin;
               Item[2,1].Value := 'Disbursements:';
               Item[2,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[2,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[2,2].Borders[xlEdgeTop].Weight := xlThin;
               Item[3,1].Value := 'Expenses:';
               Item[3,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[3,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[3,2].Borders[xlEdgeTop].Weight := xlThin;
               Item[4,1].Value := 'Total for this Quote:';
               Item[4,1].Font.Bold := True;
               Item[4,1].Borders[xlEdgeTop].Weight := xlThin;
               Item[4,1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[4,1].Borders[xlEdgeBottom].Weight := xlThin;
               Item[4,2].Borders[xlEdgeTop].Weight := xlThin;
               Item[4,2].Borders[xlEdgeBottom].Weight := xlThin;

               Item[1,(lcXSCD - lcXSCL) + 1].Value := SummaryFees * -1;
               Item[1,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[2,(lcXSCD - lcXSCL) + 1].Value := SummaryDisburse * -1;
               Item[2,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[3,(lcXSCD - lcXSCL) + 1].Value := SummaryExpenses * -1;
               Item[3,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[4,(lcXSCD - lcXSCL) + 1].Value := (SummaryFees + SummaryDisburse + SummaryExpenses) * -1;
               Item[4,(lcXSCD - lcXSCL) + 1].Borders[xlAround].Weight := xlThin;
               Item[4,(lcXSCD - lcXSCL) + 1].Font.Bold := True;
            end;

            with xlsSheet.RCRange[lcXSR,lcXSCD,lcXSR + 3,lcXSCD] do begin
               NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
            end;
         end;

//--- Insert the Banking details

         if (lcShowBanking = true) then begin
            for idx3 := 0 to QDisclaimerStr.Count - 1 do begin
               with xlsSheet.RCRange[lcBSR + idx3,1,lcBSR + idx3,lcHEC] do begin
                  Item[1,1].Value := DoSymVars(QDisclaimerStr.Strings[idx3],CurrentFile);
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;
            end;
         end;

//--- Write the Data Heading Information

         Generate_DataHeading(xlsSheet,lcPSR,lcHEC,'Quote');

//--- Page 1 data

         if (Query1.RecordCount <= (lcPRows - 2)) then begin
            ThisRows := Query1.RecordCount + 1;
            ThisItem := 'Total for this quote';
         end else begin
            ThisRows := lcPRows - 1;
            ThisItem := 'Carried Over';
         end;

         OpenBalFees   := 0;
         OpenBalVAT    := 0;
         OpenBalAmount := 0;

         LastDate := Sdate;
         Generate_Detail_Quote(xlsSheet, lcPSR, ThisRows, ThisItem, '1');
      end;

//--- Data on subsequent pages - compensate for Document Heading (3 rows)

      if (Query1.RecordCount > (lcPRows - 2)) then begin
         RemainRows := (Query1.RecordCount - (lcProws - 2));
         NumPages := ((Query1.RecordCount - (lcPRows - 3)) div (lcSRows - 3)) + 1;

//--- Compensate for cases where we have an exact page size

         if ((Query1.RecordCount - (lcPRows - 3)) mod (lcSRows - 3) = 0) then
            NumPages := NumPages - 1;

         Row := lcSSR;
         PageNum := PageNum + 1;

         for idx3 := 0 to NumPages -1 do begin
            if (lcHeaderPageOne = false) then begin
               xlsSheet.RCRange[lcHSR,lcHSC,lcHER,lcHEC].Copy(xlsSheet.RCRange[lcSSR,lcHSC,lcSSR + lcHER,lcHEC]);

               Row := Row + lcHER + 1;
            end;

            Generate_DocHeading(xlsSheet,PageNum,idx1,Row,lcHEC,'Quote');
            Row := Row + 4;
            Generate_DataHeading(xlsSheet,Row,lcHEC,'Quote');

            if (RemainRows <= (lcSRows - 3)) then begin
               ThisRows := RemainRows + 2;
               ThisItem := 'Total for this quote';
            end else begin
               ThisRows := lcSRows - 1;
               ThisItem := 'Carried Over';
            end;

            Generate_Detail_Quote(xlsSheet, Row, ThisRows, ThisItem, '2');

            PageNum := PageNum + 1;
            RemainRows := RemainRows - lcSRows + 3;
            Row := Row + lcSMaxRows - 4;
         end;
      end else begin
         Row := lcSSR;
      end;

//--- Write the standard copyright notice - Note we do not write this if no
//---    records were found

      if (Query1.RecordCount > 0) then begin
         dec(Row);

         with xlsSheet.RCRange[Row,1,Row,1] do begin
            Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 8;
         end;

//--- Put the Quote number on the Quote

            if (lcShowAge = true) then begin
               with xlsSheet.RCRange[lcAASR,lcAASC,lcAASR,lcAASC + 2] do begin
                  Item[1,1].Value := 'QUOTE NUMBER: ';
                  Item[1,3].Value := ThisQuote;
                  Item[1,3].HorizontalAlignment := xlHAlignRight;
                  Font.Bold := true;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;
            end;
//         end;

//--- Remove the Gridlines which are added by default and set the Page orientation

         xlsSheet.PageSetup.Orientation    := xlLandscape;
         xlsSheet.PageSetup.FitToPagesWide := 1;
         xlsSheet.PageSetup.FitToPagesTall := PageNum - 1;
         xlsSheet.PageSetup.PaperSize      := xlPaperA4;
         xlsSheet.DisplayGridLines         := false;
         xlsSheet.PageSetup.CenterFooter    := 'Page &P of &N';

//--- Write the Excel file to disk

         xlsBook.SaveAs(FileName + ThisFile);
         xlsBook.Close;

         LogMsg('  Quote ''' + ThisQuote + ''' successfully processed...',True);
         LogMsg(' ',True);

         DoLine := True;

//--- Print the generated document on the Default Printer if requested

         if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
            if (PrintDocument(ThisFile, FileName) = True) then
               LogMsg('  Document submitted for printing...',True)
            else
               LogMsg('  Printing of document failed...',True);

            DoLine := True;
         end;

//--- Create a PDF file if requested

         if ((CreatePDF = true) and (PDFPrinter <> 'Not Found')) then begin
            PDFExists := PDFDocument(ThisFile, FileName);

            if (PDFExists = True) then
               LogMsg('  PDF file creation was successfull...',True)
            else
               LogMsg('  PDF file creation failed...',True);

            DoLine := True;
         end;

//--- Send the Excel file via email if requested.

         if (SendByEmail = '1') then begin
            if (GroupAttach = true) then
               AttachList.Add(FileName + ThisFile)
            else begin
               if (SendEmail(CurrentFile,FileName + ThisFile,ord(PT_NORMAL)) = true) then
                  LogMsg('  Request to send generated Quote by Email submitted...',True)
               else
                  LogMsg('  Request to send generated Quote by Email not submitted...',True);

               DoLine := True;
            end;
         end;

//--- Now open the Quote if requested

         if (AutoOpen = True) then begin
            ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);
            LogMsg('  Request to open Quote for ''' + CurrentFile + ''' submitted...',True);
            DoLine := True;
         end;

      end else begin
         LogMsg('  No Quotation data found for ''' + CurrentFile + '''',True);
         DoLine := True;
      end;
      FldExcel.Refresh;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

//--- Finish up

   LogMsg(ord(PT_EPILOG),False,False,'Quotation');

   Close_Connection;

end;

//---------------------------------------------------------------------------
// Generate the Detail for Quotes
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_Detail_Quote(xlsSheet : IXLSWorksheet; PageRows: integer; ThisRows: integer; ThisItem: string; BalanceStr: string);
var
   idx1, ThisClass   : integer;
   ThisAmount        : double;

begin

//--- Page data

   with xlsSheet.RCRange[PageRows + 1,1,PageRows + ThisRows,lcHEC] do begin
      Borders[xlAround].LineStyle := xlContinuous;
      Borders[xlAround].Weight := xlThin;
   end;

   with xlsSheet.RCRange[PageRows + 1,lcHEC - 2, PageRows + ThisRows, lcHEC] do begin
      NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;

//-- Opening Balance Line - only used on 2nd Page and onwards

   if (BalanceStr = '2') then begin
      with xlsSheet.RCRange[PageRows + 1,1,PageRows + 1,lcHEC] do begin
         Borders[xlEdgeBottom].LineStyle := xlContinuous;
         Borders[xlEdgeBottom].Weight := xlThin;

         Item[1,1].Value := LastDate;
         Item[1,2].Value := 'Carried Down';

         if (VATRate = 0) then begin
            Item[1,lcHEC].Value := OpenBalFees;
         end else begin
            Item[1,lcHEC - 2].Value := OpenBalFees;
            Item[1,lcHEC - 1].Value := OpenBalVAT;
            Item[1,lcHEC].Value := OpenBalAmount;
            Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
         end;

         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;

         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

//--- Page Detail Lines

   for idx1 := StrToInt(BalanceStr) to ThisRows - 1 do begin
      with xlsSheet.RCRange[PageRows + idx1,1,PageRows + idx1,lcHEC] do begin
         Borders[xlEdgeBottom].LineStyle := xlContinuous;
         Borders[xlEdgeBottom].Weight := xlThin;
         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;

         if (VATRate = 0) then begin
            Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;
         end else begin
            Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;
         end;

         LastDate := Query1.FieldByName('Q_Date').AsString;
         Item[1,1].Value := LastDate;
         ThisAmount := Query1.FieldByName('Q_Amount').AsFloat;
         ThisClass := Query1.FieldByName('Q_Class').AsInteger;

         if (ThisClass in [0..2]) then begin

            if (Query1.FieldByName('Q_DrCr').AsInteger = 2) then begin
               ThisAmount := ThisAmount * -1;
            end else begin
               ThisAmount := ThisAmount;
            end;

            if (ThisClass = 0) then begin
               Item[1,2].Value := ReplaceQuote(Query1.FieldByName('Q_Description').AsString);
            end else if (ThisClass = 1) then begin
               Item[1,2].Value := '[**Disbursement] ' + ReplaceQuote(Query1.FieldByName('Q_Description').AsString);
            end else if (ThisClass = 2) then begin
               Item[1,2].Value := '[**Expense] ' + ReplaceQuote(Query1.FieldByName('Q_Description').AsString);
            end;

            if (VATRate > 0) then begin
               Item[1,lcHEC - 2].Value := ThisAmount;
               OpenBalFees := OpenBalFees + ThisAmount;

               if (ThisClass in [0]) then begin
                  Item[1,lcHec - 1].Value := (ThisAmount * VATRate) / 100;
                  Item[1,lcHec].Value := (ThisAmount * (100 + VATRate)) / 100;
                  OpenBalVAT := OpenBalVAT + ((ThisAmount * VATRate) / 100);
                  OpenBalAmount := OpenBalAmount + ((ThisAmount * (100 + VATRate)) / 100);
               end else begin
                  Item[1,lcHEC - 1].Value := 0;
                  Item[1,lcHEC].Value := ThisAmount;
                  OpenBalAmount := OpenBalAmount + ThisAmount;
               end;
            end else begin
               Item[1,lcHEC].Value := ThisAmount;
               OpenBalFees := OpenBalFees + ThisAmount;
            end;
         end;

         Query1.Next;

         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
   end;

//--- Closing Balance Line

   with xlsSheet.RCRange[PageRows + ThisRows,1,PageRows + ThisRows,lcHEC] do begin
      Borders[xlEdgeBottom].LineStyle := xlContinuous;
      Borders[xlEdgeBottom].Weight := xlThin;

      if (ThisItem = 'Total for this quote') then
         Item[1,1].Value := EDate
      else
         Item[1,1].Formula := LastDate;

      Item[1,2].Value := ThisItem;

      if (VATRate = 0) then begin
         Item[1,lcHEC].Formula := OpenBalFees;

         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;

         SaveAmount := OpenBalFees * -1;
      end else begin
         Item[1,lcHEC - 2].Value := OpenBalFees;
         Item[1,lcHEC - 1].Value := OpenBalVAT;
         Item[1,lcHEC].Formula := OpenBalAmount;

         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
         Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
         Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
         Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;

         SaveAmount := OpenBalAmount * -1;
      end;

      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;
end;

//---------------------------------------------------------------------------
// Generate Task List for accepted Quote
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_TaskList();
var
   idx1, row, PageRow, Pages, RowsPerPage  : integer;
   FileCount                               : integer;
   PageBreak, FirstPage, DoLine            : boolean;
   ThisFile, ThisName, ThisQuote           : string;
   CurrentFile                             : string;
   xlsBook                                 : IXLSWorkbook;
   xlsSheet                                : IXLSWorksheet;

begin

   ThisCount := 1;
   ThisMax   := IntToStr(NumFiles);

   prbProgress.Max := NumFiles;
   LogMsg(ord(PT_PROLOG),True,True,'Task List for Accepted Quote');

//--- Open a connection to the datastore named in HostName

   if ((Open_Connection(HostName)) = false) then begin
      LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

      CloseOnComplete  := False;
      AutoOpen         := False;
      txtError.Text    := 'There are errors...';
      Exit;
   end;

//--- Get Company specific information

   GetCpyVAT();

//--- Process the Quote in FileArray

   for idx1 := 0 to NumFiles - 1 do begin

      ThisQuote   := FileArray[idx1];
      CurrentFile := GetQuoteFile(ThisQuote);

      ThisFile := FormatDateTime('yyyyMMdd',Now()) + ' - Task List for ' + ThisQuote + ' (' + CurrentFile + ').xls';
      txtDocument.Caption  := 'Generating to: ' + FileName + ThisFile;
      txtDocument.Refresh;

//--- Read the Quote records records from the datastore

      if ((GetQuote(CurrentFile,ThisQuote)) = false) then begin
         LogMsg('Unexpected Data Base error: ' + ErrMsg,True);

         CloseOnComplete  := False;
         AutoOpen         := False;
         txtError.Text    := 'There are errors...';
         Exit;
      end;

      txtError.Text := 'Processing: ' + ThisQuote;
      txtError.Refresh;
      prbProgress.StepIt;
      prbProgress.Refresh;

      stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
      stCount.Refresh;
      inc(ThisCount);

//--- Only process if there are Quote records

      if (Query1.RecordCount > 0) then begin

         Pages     := 0;
         FirstPage := true;
         PageBreak := true;

         ThisName := CurrentFile;

//--- Open the Excel workbook template that will contain the Specified Account

         xlsBook       := TXLSWorkbook.Create;
         xlsSheet      := xlsBook.WorkSheets.Add;
         xlsSheet.Name := 'Task List for ' + ThisQuote + ' (' + CurrentFile + ')';

//--- Now step through each Quotes record

         Query1.First;

         while Query1.Eof = False do begin

//--- Perform a Page break if necessary

            if (PageBreak = True) then
               DoPageBreak(ord(PB_TASKLIST),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,'Quote ' + ThisQuote + ' on File ' + CurrentFile);

            with xlsSheet.RCRange[row,1,row,4] do begin
               Item[1,1].Value := Query1.FieldByName('Q_Description').AsString;
               Item[1,1].VerticalAlignment := xlVAlignTop;
               Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,2].Value := ' ';
               Item[1,2].VerticalAlignment := xlVAlignTop;
               Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,3].Value := ' ';
               Item[1,3].WrapText := ThisWrapText;
               Item[1,3].VerticalAlignment := xlVAlignTop;
               Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,4].Value := ' ';
               Item[1,4].VerticalAlignment := xlVAlignTop;
               Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
               Font.Bold := false;
               Font.Name := 'Arial';
               Font.Size := 10;
               Borders[xlAround].Weight := xlThin;
            end;

            Query1.Next;

            inc(row);
            inc(PageRow);

//--- If we've reached the maximum rows per page then it is PageBreak time.
//--- UNLESS WrapText is true in which case we do not do Pagebreak other than
//--- on the fist page. We do the Pagebreak here so that the Copyright notice
//--- will be handled correctly

            if (ThisWrapText = False) then begin
               if (PageRow >= RowsPerPage) then begin
                  PageBreak := True;
                  DoPageBreak(ord(PB_TASKLIST),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,'Quote ' + ThisQuote + ' on File ' + CurrentFile);
               end;
            end;
         end;

         inc(row);

         with xlsSheet.RCRange[row,1,row,1] do begin
            Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 8;
         end;

//--- Remove the Gridlines which are added by default and set the Page orientation

         xlsSheet.PageSetup.Orientation := xlLandscape;
         xlsSheet.PageSetup.PaperSize := xlPaperA4;
         xlsSheet.DisplayGridLines := false;
         xlsSheet.PageSetup.CenterFooter := 'Page &P of &N';
         xlsSheet.PageSetup.FitToPagesWide := 1;

         if (ThisWrapText = False) then
            xlsSheet.PageSetup.FitToPagesTall := Pages;

//--- Write the Excel file to disk

         xlsBook.SaveAs(FileName + ThisFile);
         LogMsg('  Task List for Quote ''' + ThisQuote + ''' successfully processed...',True);
         LogMsg(' ',True);

         DoLine := False;

//--- Print the generated document on the Default Printer if requested

         if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
            if (PrintDocument(ThisFile, FileName) = True) then
               LogMsg('  Document submitted for printing...',True)
            else
               LogMsg('  Printing of document failed...',True);

            DoLine := True;
         end;

//--- Create a PDF file if requested

         if ((CreatePDF = true) and (PDFPrinter <> 'Not Found')) then begin
            PDFExists := PDFDocument(ThisFile, FileName);

            if (PDFExists = True) then
               LogMsg('  PDF file creation was successfull...',True)
            else
               LogMsg('  PDF file creation failed...',True);

            DoLine := True;
         end;

//--- Add the File to the list of files to be attached

         if (SendByEmail = '1') then
            AttachList.Add(FileName + ThisFile);

//--- Now open the Exported Notes if requested

         if (AutoOpen = True) then begin
            ShellExecute(Handle,'open',PChar(FileName + ThisFile),nil,nil,SW_SHOWNORMAL);

            LogMsg('  Request to open Task List for Quote ' + ThisQuote + ' submitted...',True);

            DoLine := True;
         end;
      end else begin
         LogMsg('  No quotation records found for Quote ' + ThisQuote,True);
         DoLine := True;
      end;
   end;

//--- Send the Excel files via email if requested and if there are any

   if (FileCount > 1) then ThisName := '';

   if (GroupAttach = true) and (SendByEmail = '1') then begin
      if (SendEmail(ThisName,'',ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Task List by Email submitted...',True)
      else
         LogMsg('  Request to send generated Task List by Email failed...',True);

      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Task List for Accepted Quote');

end;

//---------------------------------------------------------------------------
// Generate Collection Account
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_CollectAcct();
var
   idx1, row, RecCount, sema, PageRow, Pages, RowsPerPage     : integer;
   PageBreak, FirstPage, DoLine                               : boolean;
   ThisBalance, ThisCapital, ThisInterest                     : double;
   ThisAB1F, ThisAB1T, ThisAB2F, ThisAB2T, ThisFill, ThisText : TColor;
   ThisFile, ThisDescrip                                      : string;
   xlsBook                                                    : IXLSWorkbook;
   xlsSheet                                                   : IXLSWorksheet;
   ExportSet, SectionSet, RecordSet                           : IXMLNode;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Collection Account');

   txtDocument.Caption  := 'Generating to: ' + FileName;
   txtDocument.Refresh;

//--- Get Company specific information

   GetCpyVAT();

//--- Open the XML file and process the content

   XMLDoc.Active := false;
   XMLDoc.LoadFromFile(HostName);
   XMLDoc.Active := true;

   DeleteFile(HostName);

   ExportSet  := XMLDoc.DocumentElement;
   SectionSet := ExportSet.ChildNodes.First;
   RecordSet  := SectionSet;

//--- Extract values that will be used in the headers

   ThisFile     := ReplaceXML(RecordSet.ChildValues['File']);
   ThisDescrip  := ReplaceXML(RecordSet.ChildValues['Descrip']);
   ThisBalance  := StrToFloat(RecordSet.ChildValues['Balance']);
   ThisCapital  := StrToFloat(RecordSet.ChildValues['Capital']);
   ThisInterest := StrToFloat(RecordSet.ChildValues['Interest']);

//--- Create the Excel spreadsheet

   xlsBook       := TXLSWorkbook.Create;
   xlsSheet      := xlsBook.WorkSheets.Add;
   xlsSheet.Name := RecordSet.ChildValues['File'];

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet.PageSetup.Orientation := xlLandscape;
   xlsSheet.PageSetup.FitToPagesWide := 1;
   xlsSheet.PageSetup.PaperSize := xlPaperA4;
   xlsSheet.DisplayGridLines := false;
   xlsSheet.PageSetup.CenterFooter := 'Page &P of &N';

//--- Set up to use alternate color blocks

   sema := 1;
   ThisAB1F := ColAB1F;
   ThisAB1T := ColAB1T;
   ThisAB2F := ColAB2F;
   ThisAB2T := ColAB2T;

//--- Process the Account Records

   RecCount := StrToInt(RecordSet.ChildValues['Count']);

   SectionSet := SectionSet.NextSibling;
   RecordSet  := SectionSet.ChildNodes.First;

   Pages     := 0;
   FirstPage := true;
   PageBreak := true;

//--- Now process the records if any

   if RecCount > 0 then begin

      for idx1 := 1 to RecCount do begin

         if (PageBreak = true) then begin

//--- Perform a Page break
//--- Set the Page control variables

            if ((lcRepeatHeader = true) or (FirstPage = true)) then
               RowsPerPage := (lcGRows - 9)
            else
               RowsPerPage := (lcGRows - 2);

            if (FirstPage = True) then
               row := 1
            else
               row := (Pages * lcGRows) + 1;

            PageRow   := 1;
            PageBreak := False;
            inc(Pages);

//--- Check if the heading must be printed / repeated

            if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
               FirstPage := False;

//--- Write the Report Heading

               with xlsSheet.Range['A' + IntToStr(row), 'E' + IntToStr(row + 1)] do begin
                  Item[1,1].Value := CpyName;
                  Item[2,1].Value := 'Collections: Statement of Account for ' + ThisFile + ' (' + ThisDescrip + '), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ColDHF;
                  Font.Color := ColDHT;
                  Font.Bold := true;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;

               row := row + 3;

//--- Write the Account Information

               with xlsSheet.Range['A' + IntToStr(row),'E' + IntToStr(row)] do begin
                  Item[1,1].Value := 'Current Balance:';
                  Item[1,5].Value := RoundD(ThisBalance,2);
                  Item[1,5].NumberFormat := '#,##0.00;-#,##0.00';
                  Item[1,5].Borders[xlAround].Weight := xlThin;
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ColAB1F;
                  Font.Color := ColAB1T;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;

               inc(row);

               with xlsSheet.Range['A' + IntToStr(row),'E' + IntToStr(row)] do begin
                  Item[1,1].Value := 'Capital:';
                  Item[1,5].Value := RoundD(ThisCapital,2);
                  Item[1,5].NumberFormat := '#,##0.00;-#,##0.00';
                  Item[1,5].Borders[xlAround].Weight := xlThin;
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ColAB2F;
                  Font.Color := ColAB2T;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;

               inc(row);

               with xlsSheet.Range['A' + IntToStr(row),'E' + IntToStr(row)] do begin
                  Item[1,1].Value := 'Arrears Interest:';
                  Item[1,5].Value := RoundD(ThisInterest,2);
                  Item[1,5].NumberFormat := '#,##0.00;-#,##0.00';
                  Item[1,5].Borders[xlAround].Weight := xlThin;
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ColAB1F;
                  Font.Color := ColAB1T;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;

               inc(row);
               inc(row);
            end;

//--- Write the Account Heading

            with xlsSheet.Range['A' + IntToStr(row), 'E' + IntToStr(row)] do begin
               Item[1,1].Value := 'Date';
               Item[1,1].ColumnWidth := 12;
               Item[1,1].Borders[xlAround].Weight := xlThin;
               Item[1,2].Value := 'Description';
               Item[1,2].ColumnWidth := lcGMRWidth - 48;
               Item[1,2].Borders[xlAround].Weight := xlThin;
               Item[1,3].Value := 'Trust';
               Item[1,3].HorizontalAlignment := xlHAlignRight;
               Item[1,3].ColumnWidth := 12;
               Item[1,3].Borders[xlAround].Weight := xlThin;
               Item[1,4].Value := 'Business';
               Item[1,4].HorizontalAlignment := xlHAlignRight;
               Item[1,4].ColumnWidth := 12;
               Item[1,4].Borders[xlAround].Weight := xlThin;
               Item[1,5].Value := 'Balance';
               Item[1,5].HorizontalAlignment := xlHAlignRight;
               Item[1,5].ColumnWidth := 12;
               Item[1,5].Borders[xlAround].Weight := xlThin;
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row);

         end;

         if (sema = 0) then begin;
            ThisFill := ThisAB1F;
            ThisText := ThisAB1T;
            sema := 1;
         end else begin
            ThisFill := ThisAB2F;
            ThisText := ThisAB2T;
            sema := 0;
         end;

         with xlsSheet.RCRange[row,1,row,5] do begin
            Item[1,1].Value := RecordSet.ChildValues['Date'];
            Item[1,1].Borders[xlAround].Weight := xlThin;
            Item[1,1].Borders[xlBelow].Weight := xlThin;
            Item[1,2].Value := ReplaceXML(RecordSet.ChildValues['Descrip']);
            Item[1,2].Borders[xlAround].Weight := xlThin;
            Item[1,3].Value := RoundD(StrToFloat(RecordSet.ChildValues['Trust']),2);
            Item[1,3].NumberFormat := '#,##0.00;-#,##0.00';
            Item[1,3].Borders[xlAround].Weight := xlThin;
            Item[1,4].Value := RoundD(StrToFloat(RecordSet.ChildValues['Business']),2);
            Item[1,4].NumberFormat := '#,##0.00;-#,##0.00';
            Item[1,4].Borders[xlAround].Weight := xlThin;
            Item[1,5].Value := RoundD(StrToFloat(RecordSet.ChildValues['Balance']),2);
            Item[1,5].NumberFormat := '#,##0.00;-#,##0.00';
            Item[1,5].Borders[xlAround].Weight := xlThin;
            Interior.Color := ThisFill;
            Font.Color := Thistext;
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;

         RecordSet := RecordSet.NextSibling;
         inc(row);
         inc(PageRow);

         if (PageRow > RowsPerPage) then
            PageBreak := True;

      end;
   end else begin
      with xlsSheet.RCRange[row,1,row,5] do begin
         Item[1,1].Value := 'No records found';
         Borders[xlAround].Weight := xlThin;
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row);
   end;

   inc(row);

   with xlsSheet.RCRange[row,1,row,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Save the workbook

   xlsBook.SaveAs(FileName);
   XMLDoc.Active := false;
   xlsBook.Close;

   LogMsg('  Collections Account successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Workbook via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send Collections Account by Email submitted...',True)
      else
         LogMsg('  Request to send Collections Account by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the Collections Account if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName),nil,nil,SW_SHOWNORMAL);
      LogMsg('  Request to open Collections Account submitted...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Collections Account');

   DeleteFile(HostName);

end;

//---------------------------------------------------------------------------
// Export Quote Details
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_QuoteDetails();
var
   idx1, row, RecCount, RowsPerPage, Pages, PageRow : integer;
   PageBreak, FirstPage, DoLine                     : boolean;
   ThisHost                                         : string;
   xlsBook                                          : IXLSWorkbook;
   xlsSheet                                         : IXLSWorksheet;
   ExportSet, SectionSet, RecordSet                 : IXMLNode;

begin

   LogMsg(ord(PT_PROLOG),True,True,'Quotes List');

   txtDocument.Caption  := 'Generating to: ' + FileName;
   txtDocument.Refresh;

//--- Get Company specific information

   GetCpyVAT();

//--- Open the XML file and process the content

   XMLDoc.Active := false;
   XMLDoc.LoadFromFile(HostName);
   XMLDoc.Active := true;

   ExportSet  := XMLDoc.DocumentElement;
   SectionSet := ExportSet.ChildNodes.First;
   RecordSet  := SectionSet;

//--- Create the Excel spreadsheet

   xlsBook       := TXLSWorkbook.Create;
   xlsSheet      := xlsBook.WorkSheets.Add;
   xlsSheet.Name := 'Quote Export';

//--- Process the File Records

   RecCount := StrToInt(RecordSet.ChildValues['Count']);

   ThisCount := 1;
   ThisMax   := IntToStr(RecCount);

   prbProgress.Max := RecCount;

   SectionSet := SectionSet.NextSibling;
   RecordSet  := SectionSet.ChildNodes.First;

//--- Now process the records in this section if any

   if RecCount > 0 then begin

//--- Perform a Page break if necessary


      Pages     := 0;
      FirstPage := true;
      PageBreak := true;

      for idx1 := 1 to RecCount do begin

         if (PageBreak = True) then
            DoPageBreak(ord(PB_QUOTELIST),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,'');

         txtError.Text := 'Processing: ' + ReplaceXML(RecordSet.ChildValues['Quote']);
         txtError.Refresh;
         prbProgress.StepIt;
         prbProgress.Refresh;

         stCount.Caption := IntToStr(ThisCount) + ' of ' + ThisMax;
         stCount.Refresh;
         inc(ThisCount);

         if (Parm07 = '_') then
            ThisHost := ReplaceXML(RecordSet.ChildValues['HostName'])
         else
            ThisHost := Parm07;

         with xlsSheet.RCRange[row,1,row,7] do begin
            Item[1,1].Value := ReplaceXML(RecordSet.ChildValues['File']);
            Item[1,1].Borders[xlAround].Weight := xlThin;
            Item[1,2].Value := ReplaceXML(RecordSet.ChildValues['Quote']);
            Item[1,2].Borders[xlAround].Weight := xlThin;
            Item[1,3].Value := ReplaceXML(RecordSet.ChildValues['Date']);
            Item[1,3].Borders[xlAround].Weight := xlThin;
            Item[1,4].Value := ReplaceXML(RecordSet.ChildValues['Description']);
            Item[1,4].Borders[xlAround].Weight := xlThin;
            Item[1,5].Value := ReplaceXML(RecordSet.ChildValues['Accepted']);
            Item[1,5].Borders[xlAround].Weight := xlThin;
            Item[1,6].Value := ThisHost;
            Item[1,6].Borders[xlAround].Weight := xlThin;
            Item[1,7].Value := GetQuoteVal(ReplaceXML(RecordSet.ChildValues['Quote']),ThisHost);
            Item[1,7].Borders[xlAround].Weight := xlThin;
            Item[1,7].HorizontalAlignment := xlHAlignRight;
            Item[1,7].NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
            Font.Bold := false;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;

         RecordSet := RecordSet.NextSibling;
         inc(row);
         inc(PageRow);

//--- If we've reached the maximum rows per page then it is PageBreak time.

         if (PageRow >= RowsPerPage) then begin
            PageBreak := True;
            DoPageBreak(ord(PB_FILEDETAILS),xlsSheet,FirstPage,RowsPerPage,row,Pages,PageRow,PageBreak,'');
         end;
      end;
   end else begin
      with xlsSheet.RCRange[row,1,row,1] do begin
         Borders[xlAround].Weight := xlThin;
         Item[1,1].Value := 'No Quote Details found';
         Font.Bold := false;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row);
   end;

//--- Write the standard copyright notice

   row := (Pages * lcGRows);
   with xlsSheet.RCRange[row,1,row,1] do begin
      Value := 'Generated by ' + LPMSHeading + ' (LPMS) © 2008-' + FormatDateTime('YYYY',Now()) + ' BlueCrane Software Development CC (www.bluecrane.cc)';
      Font.Bold := false;
      Font.Name := 'Arial';
      Font.Size := 8;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet.PageSetup.Orientation := xlLandscape;
   xlsSheet.PageSetup.FitToPagesTall := Pages;
   xlsSheet.PageSetup.PaperSize := xlPaperA4;
   xlsSheet.DisplayGridLines := false;
   xlsSheet.PageSetup.CenterFooter := 'Page &P of &N';
   xlsSheet.PageSetup.FitToPagesWide := 1;

//--- Write the Excel file to disk

   xlsBook.SaveAs(FileName);
   XMLDoc.Active := false;
   xlsBook.Close;

   LogMsg(' Quote List successfully processed...',True);
   LogMsg(' ',True);

   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(FileName) = True) then
         LogMsg('  Document submitted for printing...',True)
      else
         LogMsg('  Printing of document failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(FileName);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Send the Excel file via email if requested

   if (SendByEmail = '1') then begin
      if (SendEmail('',FileName,ord(PT_NORMAL)) = true) then
         LogMsg('  Request to send generated Quote by Email submitted...',True)
      else
         LogMsg('  Request to send generated Quote by Email not submitted...',True);

      DoLine := True;
   end;

//--- Now open the Quote List Report if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(FileName),nil,nil,SW_SHOWNORMAL);
      LogMsg('  Request to open exported Quote List for ''' + PChar(FileName) + ''' submitted...',True);
      DoLine := True;
   end;

   if (DoLine = True) then
      LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'Quote Details');

   DeleteFile(HostName);
end;

//---------------------------------------------------------------------------
// Generate the File Cover, File Notes and Billing Sheet for a selected File
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_PrintSheets();
var
   SheetsType  : integer;

begin

   LogMsg(ord(PT_PROLOG),True,True,'File Documents');

   txtDocument.Caption  := 'Generating to Folder: ' + FileName;
   txtDocument.Refresh;

//--- Get Company specific information

   GetCpyVAT();

//--- Detrmine what type of documens must be produced
//--- (0=Normal & 1=Conveyancing)

   SheetsType := StrToInt(DocStrings[0]);

//--- Set up the SymVars specific to PrintSheets

   SymVars_LPMS.SV[ord(SV_CLIENT)].Value     := DocStrings[ 2];
   SymVars_LPMS.SV[ord(SV_OPPOSE)].Value     := DocStrings[ 3];
   SymVars_LPMS.SV[ord(SV_TYPEFILE)].Value   := DocStrings[ 4];
   SymVars_LPMS.SV[ord(SV_FILE)].Value       := DocStrings[ 5];
   SymVars_LPMS.SV[ord(SV_DATE)].Value       := DocStrings[ 6];
   SymVars_LPMS.SV[ord(SV_PRESCRIPTD)].Value := DocStrings[ 7];
   SymVars_LPMS.SV[ord(SV_DIALCODE)].Value   := DocStrings[ 8];
   SymVars_LPMS.SV[ord(SV_EARNER)].Value     := DocStrings[ 9];
   SymVars_LPMS.SV[ord(SV_RATE)].Value       := DocStrings[10];
   SymVars_LPMS.SV[ord(SV_SUFFIX)].Value     := DocStrings[11];
   SymVars_LPMS.SV[ord(SV_BUYER)].Value      := DocStrings[12];
   SymVars_LPMS.SV[ord(SV_SELLER)].Value     := DocStrings[13];

//--- Now process the PrintSheets

   if (SheetsType = 0) then
      Create_PrintSheet('File Cover',Template_FC,lcRowsFC,lcColsFC)
   else
      Create_PrintSheet('Conveyancing File Cover',Template_CV,lcRowsCV,lcColsCV);

   Create_PrintSheet('File Notes',Template_FN,lcRowsFN,lcColsFN);
   Create_PrintSheet('Billing Sheet',Template_BS,lcRowsBS,lcColsBS);

//--- Wrap it up

   LogMsg(' File Documents successfully processed...',True);
   LogMsg(' ',True);

   LogMsg(ord(PT_EPILOG),False,False,'File Documents');
end;

//---------------------------------------------------------------------------
// Utility procedure to generate the content of a PrintSheet
//---------------------------------------------------------------------------
procedure TFldExcel.Create_PrintSheet(SheetName: string; ThisTemplate: string; Rows: integer; Cols: integer);
var
   Pages, row, col                   : integer;
   DoLine                            : boolean;
   OutFile, OutName, ThisStr, NewStr : string;
   xlsBook                           : IXLSWorkbook;
   xlsSheet                          : IXLSWorksheet;


begin

//--- Create the Excel spreadsheet

   xlsBook       := TXLSWorkbook.Create;

   LogMsg('--- Creating ' + SheetName + ' Document',True);
   LogMsg(' ',True);

   xlsBook.Open(ThisTemplate);
   xlsSheet      := xlsBook.ActiveSheet;
   xlsSheet.Name := SheetName;

   OutName := '\' + SheetName + ' (' + DocStrings[5] + ').xls';
   OutFile := FileName + OutName;

   for row := 1 to Rows do begin
      for col := 1 to Cols do begin
         if (xlsSheet.Cells.Item[row,col].Value <> Null) then begin
            ThisStr := xlsSheet.Cells.Item[row,col].Value;

            if (ThisStr = '$[Col]') then begin
               xlsSheet.Cells.Item[row,col].Interior.Color := StrToInt(DocStrings[14]);
               xlsSheet.Cells.Item[row,col].Value := '';
            end;

            NewStr  := DoSymVars(ThisStr,DocStrings[5]);

            if (ThisStr <> NewStr) then
               xlsSheet.Cells.Item[row,col].Value := NewStr;
         end;
      end;
   end;

//--- Remove the Gridlines which are added by default and set the Page orientation

   xlsSheet.PageSetup.Orientation := xlPortrait;
   xlsSheet.PageSetup.FitToPagesTall := Pages;
   xlsSheet.PageSetup.PaperSize := xlPaperA4;
   xlsSheet.DisplayGridLines := false;
   xlsSheet.PageSetup.FitToPagesWide := 1;
   xlsSheet.PageSetup.FitToPagesTall := 1;

//--- Write the Excel file to disk

   xlsBook.SaveAs(OutFile);
   xlsBook.Close;
   DoLine := False;

//--- Print the generated document on the Default Printer if requested

   if ((DoPrint = true) and (DefPrinter <> 'Not Found')) then begin
      if (PrintDocument(OutFile) = True) then
         LogMsg('  ' + SheetName + ' submitted for printing...',True)
      else
         LogMsg('  Printing of ' + SheetName + ' failed...',True);

      DoLine := True;
   end;

//--- Create a PDF file if requested

   PDFExists := False;

   if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then begin
      PDFExists := PDFDocument(OutFile);

      if (PDFExists = True) then
         LogMsg('  PDF file creation was successfull...',True)
      else
         LogMsg('  PDF file creation failed...',True);

      DoLine := True;
   end;

//--- Now open the File Cover if requested

   if (AutoOpen = True) then begin
      ShellExecute(Handle,'open',PChar(OutFile),nil,nil,SW_SHOWNORMAL);
      LogMsg('  Request to open ' + SheetName + ' ''' + PChar(OutFile) + ''' submitted...',True);
      DoLine := True;
   end;

//--- Check whether the XLS file must be deleted - this will only be done if
//--- a PDF was sucessfull generated

   if((PDFExists = True) and (StrToBool(DocStrings[1]) = True)) then
      DeleteFile(OutFile);

//--- Print a blank line if necessary

   if (DoLine = True) then
      LogMsg(' ',True);
end;

//===========================================================================
//===========================================================================
//===                                                                     ===
//=== Report support functions and procedures                             ===
//===                                                                     ===
//===========================================================================
//===========================================================================

//---------------------------------------------------------------------------
// Procedure to generate the Document Type heading
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_DocHeading(xlsSheet: IXLSWorksheet; PageNum: integer; idx1: integer; Row: integer; Col: integer; ThisType: string);
begin

//--- Insert the Page 1 Heading

   with xlsSheet.RCRange[Row,1,Row + 2,Col] do begin
      Borders[xlAround].LineStyle := xlContinuous;
      Borders[xlAround].Weight := xlMedium;
      Font.Name := 'Arial';
   end;

   with xlsSheet.RCRange[Row + 1,1,Row + 1,Col] do begin
      Merge(true);

      if (ThisType = 'Specified Account') then
         Value := DoSymVars(Header_A,FileArray[idx1])
      else if (ThisType = 'Invoice') then
         Value := DoSymVars(Header_X,FileArray[idx1])
      else if (ThisType = 'Statement') then
         Value := DoSymVars(Header_S,FileArray[idx1])
      else if (ThisType = 'Trust') then
         Value := DoSymVars(Header_T,FileArray[idx1])
      else if (ThisType = 'Quote') then
         Value := DoSymVars(Header_X,FileArray[idx1]);

      HorizontalAlignment := xlHAlignCenter;
      VerticalAlignment := xlVAlignCenter;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 12;
   end;
end;

//---------------------------------------------------------------------------
// Procedure to generate the Data heading
//---------------------------------------------------------------------------
procedure TFldExcel.Generate_DataHeading(xlsSheet : IXLSWorksheet; Row: integer; Col: integer; ThisType: string);
begin

   if (ThisType = 'Statement') then begin
      with xlsSheet.RCRange[Row,lcSCB,Row,lcSCBD + 1] do begin
         Borders[xlAround].Weight := xlThin;

         Interior.Color := ColDHF;
         Font.Color     := ColDHT;
         Font.Bold      := true;
         Font.Name      := 'Arial';
         Font.Size      := 10;

         Item[1,1].Value := 'Business Account';
         Item[1,lcSCBD - lcSCB + 1].Value := 'Amount ';
         Item[1,lcSCBD - lcSCB + 1].Borders[xlAround].Weight := xlThin;
         Item[1,lcSCBD - lcSCB + 1].HorizontalAlignment := xlHAlignRight;
      end;

      with xlsSheet.RCRange[Row,lcSCT,Row,lcSCTD] do begin
         Borders[xlAround].Weight := xlThin;

         Interior.Color := ColDHF;
         Font.Color     := ColDHT;
         Font.Bold      := true;
         Font.Name      := 'Arial';
         Font.Size      := 10;

         Item[1,1].Value := 'Trust Account';
         Item[1,lcSCTD - lcSCT + 1].Value := 'Amount ';
         Item[1,lcSCTD - lcSCT + 1].Borders[xlAround].Weight := xlThin;
         Item[1,lcSCTD - lcSCT + 1].HorizontalAlignment := xlHAlignRight;
      end;
   end else begin
      with xlsSheet.RCRange[Row,1,Row,Col] do begin
         Borders[xlAround].Weight := xlThin;
         Borders[xlAround].LineStyle := xlContinuous;
         Interior.Color := ColDHF;
         Item[1,1].Borders[xlEdgeRight].Weight := xlThin;

         if (ThisType = 'Specified Account') then begin
            Item[1,1].Value := 'Date';
            Item[1,2].Value := 'Item';

            if (EDate < S86Date) then
               Item[1,lcHEC - 2].Value := 'S78(2A) '
            else
               Item[1,lcHEC - 2].Value := 'S86(4) ';

            Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC - 2].HorizontalAlignment := xlHAlignRight;

            Item[1,lcHEC - 1].Value := 'Trust ';
            Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC - 1].HorizontalAlignment := xlHAlignRight;

            Item[1,lcHEC].Value := 'Client ';
            Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC].HorizontalAlignment := xlHAlignRight;
         end;

         if (ThisType = 'Invoice') then begin
            Item[1,1].Value := 'Date';
            Item[1,2].Value := 'Item';

            if (VATRate = 0) then begin
               Item[1,lcHEC].Value := 'Amount ';
            end else begin
               Item[1,lcHEC - 2].Value := 'Amount (excl) ';
               Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
               Item[1,lcHEC - 2].HorizontalAlignment := xlHAlignRight;

               Item[1,lcHEC - 1].Value := 'VAT ';
               Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[1,lcHEC - 1].HorizontalAlignment := xlHAlignRight;

               Item[1,lcHEC].Value := 'Amount (incl) ';
            end;

            Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC].HorizontalAlignment := xlHAlignRight;
         end;

         if (ThisType = 'Trust') then begin
            Item[1,1].Value := 'File';
            Item[1,2].Value := 'Date';
            Item[1,3].Value := 'Description';
            Item[1,3].Borders[xlEdgeLeft].Weight := xlThin;

            if (EDate < S86Date) then
               Item[1,lcHEC - 1].Value := 'S78(2A) '
            else
               Item[1,lcHEC - 1].Value := 'S86(4) ';

            Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC - 1].HorizontalAlignment := xlHAlignRight;

            Item[1,lcHEC].Value := 'Trust ';
            Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC].HorizontalAlignment := xlHAlignRight;
         end;

         if (ThisType = 'Quote') then begin
            Item[1,1].Value := 'Date';
            Item[1,2].Value := 'Item';

            if (VATRate = 0) then begin
               Item[1,lcHEC].Value := 'Amount ';
            end else begin
               Item[1,lcHEC - 2].Value := 'Amount (excl) ';
               Item[1,lcHEC - 2].Borders[xlEdgeLeft].Weight := xlThin;
               Item[1,lcHEC - 2].HorizontalAlignment := xlHAlignRight;

               Item[1,lcHEC - 1].Value := 'VAT ';
               Item[1,lcHEC - 1].Borders[xlEdgeLeft].Weight := xlThin;
               Item[1,lcHEC - 1].HorizontalAlignment := xlHAlignRight;

               Item[1,lcHEC].Value := 'Amount (incl) ';
            end;

            Item[1,lcHEC].Borders[xlEdgeLeft].Weight := xlThin;
            Item[1,lcHEC].HorizontalAlignment := xlHAlignRight;
         end;

         Font.Color := ColDHT;
         Font.Bold  := true;
         Font.Name  := 'Arial';
         Font.Size  := 10;

         Interior.Color := ColDHF;
      end;
   end;
end;

//---------------------------------------------------------------------------
// Procedure to do a page break for various reports
//---------------------------------------------------------------------------
procedure TFldExcel.DoPageBreak(ThisType: integer; xlsSheet: IXLSWorksheet; var FirstPage: boolean; var RowsPerPage: integer; var row: integer; var Pages: integer; var PageRow: integer; var PageBreak: boolean; ThisFile: string);
var
   Compensate          : integer;
   UserName, UserEmail : string;

begin

   case ThisType of

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Pagebreak for Statement details page
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_STATEMENT): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 4)
         else
            RowsPerPage := (lcGRows - 2);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Insert the Header (1st line) and the Heading

            with xlsSheet.Range['A' + IntToStr(row), 'E' + IntToStr(row)] do begin
               Item[1,1].Value := CpyName + ': Statement detail for File ' + ThisFile + ' for the Period ending: ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());         Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
            inc(row);
            inc(row);
         end;

         with xlsSheet.Range['A' + IntToStr(row), 'E' + IntToStr(row)] do begin
            Item[1,1].Value := 'Date';
            Item[1,1].ColumnWidth := 12;
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := 'Item';
            Item[1,2].ColumnWidth := lcGMRWidth - (12+14+14+14);
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;

            if (EDate < S86Date) then
               Item[1,3].Value := 'S78(2A) '
            else
               Item[1,3].Value := 'S86(4) ';

            Item[1,3].ColumnWidth := 14;
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].HorizontalAlignment := xlHAlignRight;
            Item[1,4].Value := 'Trust ';
            Item[1,4].ColumnWidth := 14;
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].HorizontalAlignment := xlHAlignRight;
            Item[1,5].Value := 'Business ';
            Item[1,5].ColumnWidth := 14;
            Item[1,5].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := integer(ColDHF);
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Export Notes for a File
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_FILENOTES): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 3)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Insert the Header (1st line) and the Heading

            with xlsSheet.RCRange[row,1,row,4] do begin
               Item[1,1].Value := CpyName + ': Notes for File ' + ThisFile + ' for the Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());         Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
            inc(row);
            inc(row);
         end;

         with xlsSheet.RCRange[row,1,row,4] do begin
            Item[1,1].Value := 'Date';
            Item[1,1].ColumnWidth := 12;
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := 'Time';
            Item[1,2].ColumnWidth := 12;
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := 'Note';
            Item[1,3].ColumnWidth := lcGMRWidth - (12+12+15);
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Value := 'User';
            Item[1,4].ColumnWidth := 15;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := integer(ColDHF);
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Detail page for Trust Reconciliation
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_TRUSTDETAIL): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 9)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Insert the Header (1st line) and the Heading

            with xlsSheet.RCRange[row,1,row,5] do begin
               Item[1,1].Value := CpyName;
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 11;
            end;

            inc(row);

            with xlsSheet.RCRange[row,1,row,5] do begin
               Item[1,1].Value := 'Trust Account Detail for the Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 11;
            end;
            row     := row + 2;

//--- Trust Account and S86(4) details

            with xlsSheet.RCRange[row,1,row,5] do begin
               Item[1,1].Value := 'File';
               Item[1,1].ColumnWidth := 8;
               Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,2].Value := 'Description';
               Item[1,2].ColumnWidth := 12;
               Item[1,3].Value := '';
               Item[1,3].ColumnWidth := lcGMRWidth - (8+12+16+16);
               Item[1,4].Value := '';
               Item[1,4].ColumnWidth := 16;
               Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
               Item[1,4].HorizontalAlignment := xlHAlignRight;

               if (EDate < S86Date) then
                  Item[1,5].Value := 'Section 78(2)(a) '
               else
                  Item[1,5].Value := 'Section 86(3) ';

               Item[1,5].ColumnWidth := 16;
               Item[1,5].HorizontalAlignment := xlHAlignRight;
               Borders[xlAround].Weight := xlThin;
               Interior.Color := integer(ColDHF);
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
            row     := row + 5;
         end;

         with xlsSheet.RCRange[row,1,row,5] do begin
            Item[1,1].Value := 'File';
            Item[1,1].ColumnWidth := 8;
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := 'Date';
            Item[1,2].ColumnWidth := 12;
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := 'Desription';
            Item[1,3].ColumnWidth := lcGMRWidth - (8+12+16+16);
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;

            if (EDate < S86Date) then
               Item[1,4].Value := 'S78(2A) '
            else
               Item[1,4].Value := 'S86(4) ';

            Item[1,4].ColumnWidth := 16;
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].HorizontalAlignment := xlHAlignRight;
            Item[1,5].Value := 'Trust ';
            Item[1,5].ColumnWidth := 16;
            Item[1,5].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := integer(ColDHF);
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Summary page for Trust Reconciliation
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_TRUSTSUMMARY): begin

         Compensate := 0;

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 5)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            Compensate := 1;
            FirstPage  := False;

//--- Insert the Header (1st line) and the Heading

            with xlsSheet.RCRange[row,1,row,6] do begin
               Item[1,1].Value := CpyName;
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 11;
            end;

            inc(row);

            with xlsSheet.RCRange[row,1,row,6] do begin
               Item[1,1].Value := 'Trust Account Summary for Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 11;
            end;
            row := row + 2;
         end;

         with xlsSheet.RCRange[row,1,row,6] do begin
            Item[1,1].Value := 'File';
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,1].ColumnWidth := 8;
            Item[1,2].Value := 'Description';
            Item[1,2].ColumnWidth := 14;
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].ColumnWidth := lcGMRWidth - (8+14+16+16+16);

            if (EDate < S86Date) then
               Item[1,4].Value := 'S78(2)(a) '
            else
               Item[1,4].Value := 'S86(3) ';

            Item[1,4].HorizontalAlignment := xlHAlignRight;
            Item[1,4].ColumnWidth := 16;

            if (EDate < S86Date) then
               Item[1,5].Value := 'S78(2A) '
            else
               Item[1,5].Value := 'S86(4) ';

            Item[1,5].HorizontalAlignment := xlHAlignRight;
            Item[1,5].ColumnWidth := 16;
            Item[1,6].Value := 'Trust Balance ';
            Item[1,6].HorizontalAlignment := xlHAlignRight;
            Item[1,6].ColumnWidth := 16;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := integer(ColDHF);
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
         row := row + Compensate;
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Details of Sec 78(2)(a) investments
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_TRUSTS864): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 4)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Insert the Header (1st line) and the Heading

            with xlsSheet.RCRange[row,1,row,3] do begin
               Item[1,1].Value := CpyName;
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 11;
            end;

            inc(row);

            with xlsSheet.RCRange[row,1,row,3] do begin
               Item[1,1].Value := 'Section 86(3) Interest for current Financial Year Year-To-Date (' + FYYear + '/' + FYMonth + '/01 to ' + EDate +'), Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 11;
            end;
            row := row + 2;
         end;

//--- Section 78(2)(a) Interest details

         with xlsSheet.RCRange[row,1,row,3] do begin
            Item[1,1].Value := 'Date';
            Item[1,1].ColumnWidth := 14;
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := 'Description ';
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].ColumnWidth := lcGMRWidth - (14+14);
            Item[1,3].Value := 'Amount ';
            Item[1,3].HorizontalAlignment := xlHAlignRight;
            Item[1,3].ColumnWidth := 14;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := integer(ColDHF);
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Export Fee Earner details
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_FEEEARNERS): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 6)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Insert the Header (1st line) and the Heading

            with xlsSheet.RCRange[row,1,row,4] do begin
               Item[1,1].Value := CpyName + ': Fee Earner Report (VAT Excl) for Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row);
            inc(row);

            with xlsSheet.RCRange[row,1,row,4] do begin
               Item[1,1].Value := 'Fee Earner:';
               Item[1,1].Font.Bold := true;
               Item[1,2].Value := ThisFile + ' (' + FeeEarnerName + ')';
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row);

            with xlsSheet.RCRange[row,1,row,4] do begin
               Item[1,1].Value := 'Email:';
               Item[1,1].Font.Bold := true;
               Item[1,2].Value := FeeEarnerEmail;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row);
            inc(row);
         end;

///--- Write the Column Heading

         with xlsSheet.RCRange[row,1,row,4] do begin
            Item[1,1].Value := 'File';
            Item[1,1].ColumnWidth := 42;
            Item[1,1].Borders[xlAround].Weight := xlThin;
            Item[1,2].Value := 'Date';
            Item[1,2].ColumnWidth := 12;
            Item[1,2].Borders[xlAround].Weight := xlThin;
            Item[1,3].Value := 'Desription';
            Item[1,3].ColumnWidth := lcGMRWidth - (42+12+14);
            Item[1,3].Borders[xlAround].Weight := xlThin;
            Item[1,4].Value := 'Amount ';
            Item[1,4].ColumnWidth := 14;
            Item[1,4].Borders[xlAround].Weight := xlThin;
            Item[1,4].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := integer(ColDHF);
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Export Fee Earner Consolidation
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_FEECONS): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 3)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

            with xlsSheet.RCRange[row,1,row,3] do begin
               Item[1,1].Value := CpyName + ': Fee Earner Report for the Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row);
            inc(row);

         end;

//--- Write the Column Heading

         with xlsSheet.RCRange[row,1,row,3] do begin
            Item[1,1].Value := 'Fee Earner';
            Item[1,1].ColumnWidth := 32;
            Item[1,1].Borders[xlAround].Weight := xlThin;
            Item[1,2].Value := 'Description';
            Item[1,2].ColumnWidth := lcGMRWidth - (32+14);
            Item[1,2].Borders[xlAround].Weight := xlThin;
            Item[1,3].Value := 'Amount';
            Item[1,3].ColumnWidth := 14;
            Item[1,3].Borders[xlAround].Weight := xlThin;
            Item[1,3].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := integer(ColDHF);
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Accountant Report 01
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_ACCOUNTANT01): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 3)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Insert the Header (1st line)

            with xlsSheet.RCRange[row,1,row,3] do begin
               Item[1,1].Value := CpyName + ': Accountant Report 1 for the Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlMedium;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 11;
            end;

            inc(row);
            inc(row);
         end;

//--- Write the section heading

         with xlsSheet.RCRange[row,1,row,3] do begin
            Item[1,1].Value := 'File';
            Item[1,1].ColumnWidth := 65;
            Item[1,1].Borders[xlAround].Weight := xlThin;
            Item[1,2].Value := 'Desription';
            Item[1,2].ColumnWidth := lcGMRWidth - (65+14);
            Item[1,2].Borders[xlAround].Weight := xlThin;
            Item[1,3].Value := 'Amount ';
            Item[1,3].ColumnWidth := 14;
            Item[1,3].Borders[xlAround].Weight := xlThin;
            Item[1,3].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlMedium;
            Interior.Color := ColDHF;
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Accountant Report 02
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_ACCOUNTANT02): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 3)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Insert the Header

            with xlsSheet.RCRange[row,1,row,5] do begin
               Item[1,1].Value := CpyName + ': Accountant Report 2 for the Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlMedium;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 11;
            end;

            inc(row);
            inc(row);
         end;

//--- Write the section heading

         with xlsSheet.RCRange[row,1,row,5] do begin
            Item[1,1].Value := 'File';
            Item[1,1].ColumnWidth := 72;
            Item[1,1].Borders[xlAround].Weight := xlThin;
            Item[1,2].Value := 'Description';
            Item[1,2].ColumnWidth := 1;
            Item[1,3].ColumnWidth := 1;
            Item[1,4].ColumnWidth := lcGMRWidth - (72+1+1+14);
            Item[1,4].Borders[xlAround].Weight := xlThin;
            Item[1,5].Value := 'Amount ';
            Item[1,5].ColumnWidth := 14;
            Item[1,5].Borders[xlAround].Weight := xlThin;
            Item[1,5].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlMedium;
            Interior.Color := ColDHF;
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Accountant Report 03
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_ACCOUNTANT03): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = True) or (FirstPage = True)) then
            RowsPerPage := (lcGRows - 4) div 2
         else
            RowsPerPage := (lcGRows - 2) div 2;

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

            with xlsSheet.RCRange[row,1,row,10] do begin
               Item[1,1].Value := CpyName + ': Accountant Report 3 for the Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
            inc(row);
            inc(row);
         end;

         with xlsSheet.RCRange[row,1,row,10] do begin
            Item[1, 1].Value := 'Invoice';
            Item[1, 1].ColumnWidth := 18;
            Item[1, 1].Borders[xlAround].Weight := xlThin;
            Item[1, 2].Value := 'Date';
            Item[1, 2].ColumnWidth := 10;
            Item[1, 2].Borders[xlAround].Weight := xlThin;
            Item[1, 3].Value := 'File';
            Item[1, 3].ColumnWidth := 14;
            Item[1, 3].Borders[xlAround].Weight := xlThin;
            Item[1, 4].Value := 'Client';
            Item[1, 4].ColumnWidth := lcGMRWidth - (18+10+14+7+12+12+12+12+12+1);
            Item[1, 4].Borders[xlAround].Weight := xlThin;
            Item[1, 5].Value := 'Type';
            Item[1, 5].ColumnWidth := 7;
            Item[1, 5].Borders[xlAround].Weight := xlThin;
            Item[1, 6].Value := 'Totals ';
            Item[1, 6].ColumnWidth := 12;
            Item[1, 6].Borders[xlAround].Weight := xlThin;
            Item[1, 6].HorizontalAlignment := xlHAlignRight;
            Item[1, 7].Value := 'Fees ';
            Item[1, 7].ColumnWidth := 12;
            Item[1, 7].Borders[xlAround].Weight := xlThin;
            Item[1, 7].HorizontalAlignment := xlHAlignRight;
            Item[1, 8].Value := 'Disburse ';
            Item[1, 8].ColumnWidth := 12;
            Item[1, 8].Borders[xlAround].Weight := xlThin;
            Item[1, 8].HorizontalAlignment := xlHAlignRight;
            Item[1, 9].Value := 'Expenses ';
            Item[1, 9].ColumnWidth := 12;
            Item[1, 9].Borders[xlAround].Weight := xlThin;
            Item[1, 9].HorizontalAlignment := xlHAlignRight;
            Item[1,10].Value := 'Control ';
            Item[1,10].ColumnWidth := 12;
            Item[1,10].Borders[xlAround].Weight := xlThin;
            Item[1,10].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := ColDHF;
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Alerts Report
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_ALERTS): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 7)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Insert the Header (1st line) and the Heading

            with xlsSheet.RCRange[row,1,row,6] do begin
               Item[1,1].Value := CpyName + ': Alerts Report for the period ending: ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/MM/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row);
            inc(row);

//--- Get the detail for the current File Owner

            if (GetUser(ThisFile) = false) then begin
               LogMsg('Unexpected Data Base error: ' + ErrMsg,True);
               CloseOnComplete  := False;
               AutoOpen         := False;
               txtError.Text    := 'There are errors...';
               Exit;
            end;

            UserName   := Query2.FieldByName('Control_Name').AsString;
            UserEmail  := Query2.FieldByName('Control_Email').AsString;

            with xlsSheet.RCRange[row,1,row,6] do begin
               Item[1,1].Value := 'User:';
               Item[1,1].Font.Bold := true;
               Item[1,3].Value := ThisFile + ' (' + UserName + ')';
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row);

            with xlsSheet.RCRange[row,1,row,6] do begin
               Item[1,1].Value := 'Email:';
               Item[1,1].Font.Bold := true;
               Item[1,3].Value := UserEmail;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row);
            inc(row);
         end;

//--- Write the Column Heading

         with xlsSheet.RCRange[row,1,row,6] do begin
            Item[1,1].Value := 'Date';
            Item[1,1].ColumnWidth := 10;
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := 'File';
            Item[1,2].ColumnWidth := 8;
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := 'Type';
            Item[1,3].ColumnWidth := 12;
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Value := 'Desription';
            Item[1,4].ColumnWidth := 56;
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].Value := 'Alert ';
            Item[1,5].ColumnWidth := lcGMRWidth - (10+8+12+56+8);
            Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,6].Value := 'Action ';
            Item[1,6].ColumnWidth := 8;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := ColDHF;
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Phonebook Export
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_PHONEBOOK): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 4)
         else
            RowsPerPage := (lcGRows - 2);

         if (FirstPage = true) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := false;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = true) or (lcRepeatHeader = true)) then begin
            FirstPage := false;

//--- Insert the Header (1st line)

            with xlsSheet.RCRange[row,1,row,6] do begin
               Item[1,1].Value := CpyName + ': Phonebook Export, Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
            inc(row);
            inc(row);
         end;

//--- Write the Column Heading

         with xlsSheet.RCRange[row,1,row,6] do begin
            Item[1,1].Value := 'Name';
            Item[1,1].ColumnWidth := lcGMRWidth - (16+16+16+16+16);
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := 'Type';
            Item[1,2].ColumnWidth := 16;
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := 'Telephone ';
            Item[1,3].ColumnWidth := 16;
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].HorizontalAlignment := xlHAlignRight;
            Item[1,4].Value := 'Fax Number ';
            Item[1,4].ColumnWidth := 16;
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].HorizontalAlignment := xlHAlignRight;
            Item[1,5].Value := 'Cellphone ';
            Item[1,5].ColumnWidth := 16;
            Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].HorizontalAlignment := xlHAlignRight;
            Item[1,6].Value := 'Work Number ';
            Item[1,6].ColumnWidth := 16;
            Item[1,6].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := ColDHF;
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Trust Management Report
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_TRUSTMAN): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = True) or (FirstPage = True)) then
            RowsPerPage := (lcGRows - 4)
         else
            RowsPerPage := (lcGRows - 2);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         Inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Insert the Header (1st line) and the Heading

            with xlsSheet.RCRange[row,1,row,7] do begin
               Item[1,1].Value := CpyName + ': Trust Management Report for the period ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
            inc(row);
            inc(row);
         end;

         with xlsSheet.RCRange[row,1,row,7] do begin
            Item[1,1].Value := 'File';
            Item[1,1].ColumnWidth := 10;
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := 'Description';
            Item[1,2].ColumnWidth := lcGMRWidth - (10+15+15+15+15+10);
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := 'Sec 78(2A) ';
            Item[1,3].ColumnWidth := 15;
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].HorizontalAlignment := xlHAlignRight;
            Item[1,4].Value := 'Trust Account ';
            Item[1,4].ColumnWidth := 15;
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].HorizontalAlignment := xlHAlignRight;
            Item[1,5].Value := 'Business Acct ';
            Item[1,5].ColumnWidth := 15;
            Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].HorizontalAlignment := xlHAlignRight;
            Item[1,6].Value := 'Balance ';
            Item[1,6].ColumnWidth := 15;
            Item[1,6].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,6].HorizontalAlignment := xlHAlignRight;
            Item[1,7].Value := ' Manage';
            Item[1,7].ColumnWidth := 10;
            Item[1,7].Borders[xlEdgeRight].Weight := xlThin;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := ColDHF;
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Payments Report - Variant 01
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_PAYMENTS01): begin

         if ((lcRepeatHeader = true) or (FirstPage = true)) then begin
            if (FirstPage = true) then
               RowsPerPage := lcGRows - 7
            else
               RowsPerPage := lcGRows - 4
         end else
            RowsPerPage := lcGRows - 2;

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin

//--- Insert the Header (1st line) and the Heading

            with xlsSheet.RCRange[row,1,row,9] do begin
               Item[1,1].Value := CpyName + ': Payments Report for the Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row);
            inc(row);

            if (FirstPage = true) then begin

               with xlsSheet.RCRange[row,1,row,9] do begin
                  Item[1,1].Value := 'Start Date';
                  Item[1,1].ColumnWidth := 12;
                  Item[1,1].Borders[xlAround].Weight := xlThin;
                  Item[1,2].Value := 'End Date';
                  Item[1,2].ColumnWidth := 12;
                  Item[1,2].Borders[xlAround].Weight := xlThin;
                  Item[1,3].Value := 'Client';
                  Item[1,3].ColumnWidth := lcGMRWidth - (12+12+18+12+12+12+12+12);
                  Item[1,3].Borders[xlAround].Weight := xlThin;
                  Item[1,4].Value := 'Invoice';
                  Item[1,4].ColumnWidth := 18;
                  Item[1,4].Borders[xlAround].Weight := xlThin;
                  Item[1,5].Value := 'File';
                  Item[1,5].ColumnWidth := 12;
                  Item[1,5].Borders[xlAround].Weight := xlThin;
                  Item[1,6].Value := 'Fees ';
                  Item[1,6].ColumnWidth := 12;
                  Item[1,6].Borders[xlAround].Weight := xlThin;
                  Item[1,6].HorizontalAlignment := xlHAlignRight;
                  Item[1,7].Value := 'Disburse ';
                  Item[1,7].ColumnWidth := 12;
                  Item[1,7].Borders[xlAround].Weight := xlThin;
                  Item[1,7].HorizontalAlignment := xlHAlignRight;
                  Item[1,8].Value := 'Expenses ';
                  Item[1,8].ColumnWidth := 12;
                  Item[1,8].Borders[xlAround].Weight := xlThin;
                  Item[1,8].HorizontalAlignment := xlHAlignRight;
                  Item[1,9].Value := 'Total ';
                  Item[1,9].ColumnWidth := 12;
                  Item[1,9].Borders[xlAround].Weight := xlThin;
                  Item[1,9].HorizontalAlignment := xlHAlignRight;
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ColDHF;
                  Font.Color := ColDHT;
                  Font.Bold := true;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;

               inc(row);

               with xlsSheet.RCRange[row,1,row,9] do begin
                  Item[1,1].Value := Query1.FieldByName('Inv_SDate').AsString;
                  Item[1,1].Borders[xlAround].Weight := xlThin;
                  Item[1,2].Value := Query1.FieldByName('Inv_EDate').AsString;
                  Item[1,2].Borders[xlAround].Weight := xlThin;
                  Item[1,3].Value := ReplaceQuote(Query1.FieldByName('Inv_Description').AsString);
                  Item[1,3].Borders[xlAround].Weight := xlThin;
                  Item[1,4].Value := Query1.FieldByName('Inv_Invoice').AsString;
                  Item[1,4].Borders[xlAround].Weight := xlThin;
                  Item[1,5].Value := Query1.FieldByName('Inv_File').AsString;
                  Item[1,5].Borders[xlAround].Weight := xlThin;
                  Item[1,6].Value := PayFees;
                  Item[1,6].Borders[xlAround].Weight := xlThin;
                  Item[1,6].HorizontalAlignment := xlHAlignRight;
                  Item[1,7].Value := PayDisburse;
                  Item[1,7].Borders[xlAround].Weight := xlThin;
                  Item[1,7].HorizontalAlignment := xlHAlignRight;
                  Item[1,8].Value := PayExpenses;
                  Item[1,8].Borders[xlAround].Weight := xlThin;
                  Item[1,8].HorizontalAlignment := xlHAlignRight;
                  Item[1,9].Value := PayAmount;
                  Item[1,9].Borders[xlAround].Weight := xlThin;
                  Item[1,9].HorizontalAlignment := xlHAlignRight;
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ColAB1F;
                  Font.Color := ColAB1T;
                  Font.Bold := false;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;

               with xlsSheet.RCRange[row,6,row,9] do begin
                  NumberFormat := '#,##0.00_);' + NegativeStr + '-#,##0.00_)';
               end;

               inc(row);
               inc(row);

               FirstPage := False;

            end;

         end;

         with xlsSheet.RCRange[row,1,row,9] do begin
            Item[1,1].Value := 'Date';
            Item[1,1].Borders[xlAround].Weight := xlThin;
            Item[1,2].Value := 'Amount ';
            Item[1,2].Borders[xlAround].Weight := xlThin;
            Item[1,3].Value := 'Description';
            Item[1,5].Value := 'Fees ';
            Item[1,5].Borders[xlAround].Weight := xlThin;
            Item[1,5].HorizontalAlignment := xlHAlignRight;
            Item[1,6].Value := 'Disburse ';
            Item[1,6].Borders[xlAround].Weight := xlThin;
            Item[1,6].HorizontalAlignment := xlHAlignRight;
            Item[1,7].Value := 'Expenses ';
            Item[1,7].Borders[xlAround].Weight := xlThin;
            Item[1,7].HorizontalAlignment := xlHAlignRight;
            Item[1,8].Value := 'Paid ';
            Item[1,8].Borders[xlAround].Weight := xlThin;
            Item[1,8].HorizontalAlignment := xlHAlignRight;
            Item[1,9].Value := 'Balance ';
            Item[1,9].Borders[xlAround].Weight := xlThin;
            Item[1,9].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := ColDHF;
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Payments Report - Variant 02
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_PAYMENTS02): begin

         if ((lcRepeatHeader = true) or (FirstPage = true)) then begin
            if (FirstPage = true) then
               RowsPerPage := lcGRows - 8
            else
               RowsPerPage := lcGRows - 4
         end else
            RowsPerPage := lcGRows - 2;

         if (FirstPage = true) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = true) or (lcRepeatHeader = true)) then begin

            with xlsSheet.RCRange[row,1,row,10] do begin
               Item[1,1].Value := CpyName + ': Payments Report for the Period: ' + SDate + ' to ' + EDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;

            inc(row);
            inc(row);

            if (FirstPage = true) then begin
               with xlsSheet.RCRange[row,1,row,10] do begin
                  Item[1,1].Value := 'Total for all Invoices';
                  Interior.Color := ColAB1F;
                  Borders[xlAround].Weight := xlThin;
                  Font.Color := ColAB1T;
                  Font.Bold := false;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;

               inc(row);

               with xlsSheet.RCRange[row,1,row,10] do begin
                  Item[1,1].Value := 'Total Paid for all Invoices';
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ColAB2F;
                  Font.Color := ColAB2T;
                  Font.Bold := false;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;

               inc(row);

               with xlsSheet.RCRange[row,1,row,10] do begin
                  Item[1,1].Value := 'Total Unpaid for all Invoices';
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ColAB1F;
                  Font.Color := ColAB1T;
                  Font.Bold := false;
                  Font.Name := 'Arial';
                  Font.Size := 10;
               end;

               inc(row);
               inc(row);

            end;

            FirstPage := False;
         end;

         with xlsSheet.RCRange[row,1,row,10] do begin
            Item[1, 1].Value := 'File';
            Item[1, 1].ColumnWidth := 12;
            Item[1, 1].Borders[xlAround].Weight := xlThin;
            Item[1, 2].Value := 'Date';
            Item[1, 2].ColumnWidth := 12;
            Item[1, 2].Borders[xlAround].Weight := xlThin;
            Item[1, 3].Value := 'Invoice';
            Item[1, 3].ColumnWidth := 18;
            Item[1, 3].Borders[xlAround].Weight := xlThin;
            Item[1, 4].Value := 'Client';
            Item[1, 4].ColumnWidth := lcGMRWidth - (12+12+18+12+12+12+12+12+14+3);
            Item[1, 4].Borders[xlAround].Weight := xlThin;
            Item[1, 5].Value := 'Fees ';
            Item[1, 5].ColumnWidth := 12;
            Item[1, 5].Borders[xlAround].Weight := xlThin;
            Item[1, 5].HorizontalAlignment := xlHAlignRight;
            Item[1, 6].Value := 'Disburse ';
            Item[1, 6].ColumnWidth := 12;
            Item[1, 6].Borders[xlAround].Weight := xlThin;
            Item[1, 6].HorizontalAlignment := xlHAlignRight;
            Item[1, 7].Value := 'Expenses ';
            Item[1, 7].ColumnWidth := 12;
            Item[1, 7].Borders[xlAround].Weight := xlThin;
            Item[1, 7].HorizontalAlignment := xlHAlignRight;
            Item[1, 8].Value := 'Total ';
            Item[1, 8].ColumnWidth := 12;
            Item[1, 8].Borders[xlAround].Weight := xlThin;
            Item[1, 8].HorizontalAlignment := xlHAlignRight;
            Item[1, 9].Value := 'Paid ';
            Item[1, 9].ColumnWidth := 12;
            Item[1, 9].Borders[xlAround].Weight := xlThin;
            Item[1, 9].HorizontalAlignment := xlHAlignRight;
            Item[1,10].Value := 'Unpaid ';
            Item[1,10].ColumnWidth := 14;
            Item[1,10].Borders[xlAround].Weight := xlThin;
            Item[1,10].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := ColDHF;
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Export File details
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_FILEDETAILS): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 3)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Write the Report Heading

            if (ThisFile <> '0') then begin
               with xlsSheet.RCRange[row,1,row,26] do begin
                  Item[1,1].Value := CpyName + ': File Details Export, Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ColDHF;
                  Font.Color := ColDHT;
                  Font.Bold := true;
                  Font.Name := 'Arial';
                  Font.Size := 11;
               end;
            end else begin
               with xlsSheet.RCRange[row,1,row,2] do begin
                  Item[1,1].Value := CpyName + ': File Details Export, Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
                  Borders[xlAround].Weight := xlThin;
                  Interior.Color := ColDHF;
                  Font.Color := ColDHT;
                  Font.Bold := true;
                  Font.Name := 'Arial';
                  Font.Size := 11;
               end;
            end;

            inc(row);
            inc(row);
         end;

//--- Write the File List Heading

         if (ThisFile <> '0') then begin
            with xlsSheet.RCRange[row,1,row,26] do begin
               Item[1, 1].Value := 'File';
               Item[1, 1].ColumnWidth := 15;
               Item[1, 1].Borders[xlAround].Weight := xlThin;
               Item[1, 2].Value := 'Description';
               Item[1, 2].ColumnWidth := 100;
               Item[1, 2].Borders[xlAround].Weight := xlThin;
               Item[1, 3].Value := 'Diary Date';
               Item[1, 3].ColumnWidth := 15;
               Item[1, 3].Borders[xlAround].Weight := xlThin;
               Item[1, 4].Value := 'Inactive';
               Item[1, 4].ColumnWidth := 10;
               Item[1, 4].Borders[xlAround].Weight := xlThin;
               Item[1, 5].Value := 'Settled';
               Item[1, 5].ColumnWidth := 10;
               Item[1, 5].Borders[xlAround].Weight := xlThin;
               Item[1, 6].Value := 'Case Number';
               Item[1, 6].ColumnWidth := 20;
               Item[1, 6].Borders[xlAround].Weight := xlThin;
               Item[1, 7].Value := 'Court';
               Item[1, 7].ColumnWidth := 90;
               Item[1, 7].Borders[xlAround].Weight := xlThin;
               Item[1, 8].Value := 'Counsel';
               Item[1, 8].ColumnWidth := 40;
               Item[1, 8].Borders[xlAround].Weight := xlThin;
               Item[1, 9].Value := 'File Owner';
               Item[1, 9].ColumnWidth := 20;
               Item[1, 9].Borders[xlAround].Weight := xlThin;
               Item[1,10].Value := 'Alert';
               Item[1,10].ColumnWidth := 10;
               Item[1,10].Borders[xlAround].Weight := xlThin;
               Item[1,11].Value := 'Alert Date';
               Item[1,11].ColumnWidth := 15;
               Item[1,11].Borders[xlAround].Weight := xlThin;
               Item[1,12].Value := 'Alert Reason';
               Item[1,12].ColumnWidth := 120;
               Item[1,12].Borders[xlAround].Weight := xlThin;
               Item[1,13].Value := 'Prescription';
               Item[1,13].ColumnWidth := 15;
               Item[1,13].Borders[xlAround].Weight := xlThin;
               Item[1,14].Value := 'Prescription Date';
               Item[1,14].ColumnWidth := 20;
               Item[1,14].Borders[xlAround].Weight := xlThin;
               Item[1,15].Value := 'Client';
               Item[1,15].ColumnWidth := 90;
               Item[1,15].Borders[xlAround].Weight := xlThin;
               Item[1,16].Value := 'Opposition';
               Item[1,16].ColumnWidth := 90;
               Item[1,16].Borders[xlAround].Weight := xlThin;
               Item[1,17].Value := 'Correspondent';
               Item[1,17].ColumnWidth := 90;
               Item[1,17].Borders[xlAround].Weight := xlThin;
               Item[1,18].Value := 'Opposing Attorney';
               Item[1,18].ColumnWidth := 90;
               Item[1,18].Borders[xlAround].Weight := xlThin;
               Item[1,19].Value := 'Folder';
               Item[1,19].ColumnWidth := 200;
               Item[1,19].Borders[xlAround].Weight := xlThin;
               Item[1,20].Value := 'File Type';
               Item[1,20].ColumnWidth := 15;
               Item[1,20].Borders[xlAround].Weight := xlThin;
               Item[1,21].Value := 'Rate';
               Item[1,21].ColumnWidth := 10;
               Item[1,21].Borders[xlAround].Weight := xlThin;
               Item[1,22].Value := 'Related File';
               Item[1,22].ColumnWidth := 15;
               Item[1,22].Borders[xlAround].Weight := xlThin;
               Item[1,23].Value := 'Sheriff';
               Item[1,23].ColumnWidth := 90;
               Item[1,23].Borders[xlAround].Weight := xlThin;
               Item[1,24].Value := 'Free Text 1';
               Item[1,24].ColumnWidth := 100;
               Item[1,24].Borders[xlAround].Weight := xlThin;
               Item[1,25].Value := 'Free Text 2';
               Item[1,25].ColumnWidth := 100;
               Item[1,25].Borders[xlAround].Weight := xlThin;
               Item[1,26].Value := 'Free Text 3';
               Item[1,26].ColumnWidth := 100;
               Item[1,26].Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end else begin
            with xlsSheet.RCRange[row,1,row,2] do begin
               Item[1,1].Value := 'Attribute';
               Item[1,1].ColumnWidth := 30;
               Item[1,1].Borders[xlAround].Weight := xlThin;
               Item[1,2].Value := 'Value';
               Item[1,2].ColumnWidth := lcGMRWidth - 30;
               Item[1,2].Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Export Quote items for a Task List
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_TASKLIST): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 3)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Insert the Header (1st line) and the Heading

            with xlsSheet.RCRange[row,1,row,4] do begin
               Item[1,1].Value := CpyName + ': Task List for ' + ThisFile + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
            inc(row);
            inc(row);
         end;

         with xlsSheet.RCRange[row,1,row,4] do begin
            Item[1,1].Value := 'Description/Activity/Task';
            Item[1,1].ColumnWidth := lcGMRWidth - (12+12+15);
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := 'Completed';
            Item[1,2].ColumnWidth := 12;
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := 'By';
            Item[1,3].ColumnWidth := 12;
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Value := 'Date';
            Item[1,4].ColumnWidth := 15;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := integer(ColDHF);
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Export Quote List
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

      ord(PB_QUOTELIST): begin

//--- Set the Page control variables

         if ((lcRepeatHeader = true) or (FirstPage = true)) then
            RowsPerPage := (lcGRows - 3)
         else
            RowsPerPage := (lcGRows - 1);

         if (FirstPage = True) then
            row := 1
         else
            row := (Pages * lcGRows) + 1;

         PageRow   := 1;
         PageBreak := False;
         inc(Pages);

//--- Check if the heading must be printed / repeated

         if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
            FirstPage := False;

//--- Insert the Header (1st line) and the Heading

            with xlsSheet.RCRange[row,1,row,7] do begin
               Item[1,1].Value := CpyName + ': Quote List generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
               Borders[xlAround].Weight := xlThin;
               Interior.Color := ColDHF;
               Font.Color := ColDHT;
               Font.Bold := true;
               Font.Name := 'Arial';
               Font.Size := 10;
            end;
            inc(row);
            inc(row);
         end;

         with xlsSheet.RCRange[row,1,row,7] do begin
            Item[1,1].Value := 'File';
            Item[1,1].ColumnWidth := 12;
            Item[1,1].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,2].Value := 'Quote';
            Item[1,2].ColumnWidth := 18;
            Item[1,2].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,3].Value := 'Date';
            Item[1,3].ColumnWidth := 12;
            Item[1,3].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,4].Value := 'Description';
            Item[1,4].ColumnWidth := lcGMRWidth - (12+18+12+10+20+18);
            Item[1,4].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,5].Value := 'Accepted';
            Item[1,5].ColumnWidth := 10;
            Item[1,5].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,6].Value := 'Hostname';
            Item[1,6].ColumnWidth := 20;
            Item[1,6].Borders[xlEdgeRight].Weight := xlThin;
            Item[1,7].Value := 'Value ';
            Item[1,7].ColumnWidth := 18;
            Item[1,7].HorizontalAlignment := xlHAlignRight;
            Borders[xlAround].Weight := xlThin;
            Interior.Color := integer(ColDHF);
            Font.Color := ColDHT;
            Font.Bold := true;
            Font.Name := 'Arial';
            Font.Size := 10;
         end;
         inc(row);
      end;
   end;
end;

//---------------------------------------------------------------------------
// Procedure to perform a Page Break when generating the Billing Prep Report
//---------------------------------------------------------------------------
procedure TFldExcel.Billng_Prep_PageBreak(xlsSheet: IXLSWorksheet; var PageBreak: boolean; var FirstPage: boolean; SaveEDate: string; MonthC: string; Month1: string; Month2: string; Month3: string; Month4: string; Month5: string; Month6: string; var RowsPerPage: integer; var row: integer; var PageRow: integer; var Pages: integer);
begin

   if ((lcRepeatHeader = True) or (FirstPage = True)) then
      RowsPerPage := (lcGRows - 4)
   else
      RowsPerPage := (lcGRows - 2);

   if (FirstPage = True) then
      row := 1
   else
      row := (Pages * lcGRows) + 1;

   PageRow   := 1;
   PageBreak := False;
   Inc(Pages);

//--- Check if the heading must be printed / repeated

   if ((FirstPage = True) or (lcRepeatHeader = True)) then begin
      FirstPage := False;

      with xlsSheet.RCRange[row,1,row,17] do begin
         Item[1,1].Value := CpyName + ': Billing Preparation Report for the period ending ' +  SaveEDate + ', Generated on: ' + FormatDateTime('yyyy/mm/dd',Now());
         Borders[xlAround].Weight := xlThin;
         Interior.Color := ColDHF;
         Font.Color := ColDHT;
         Font.Bold := true;
         Font.Name := 'Arial';
         Font.Size := 10;
      end;
      inc(row);
      inc(row);
   end;

   with xlsSheet.RCRange[row,1,row,17] do begin
      Item[1, 1].Value := 'File';
      Item[1, 1].ColumnWidth := lcGMRWidth - (12+2+12+2+12+2+12+2+12+2+12+2+12+2+12+2);
      Item[1, 1].Borders[xlEdgeLeft].Weight := xlThin;
      Item[1, 1].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 2].Value := MonthC;
      Item[1, 2].ColumnWidth := 12;
      Item[1, 2].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 2].HorizontalAlignment := xlHAlignRight;
      Item[1, 3].Value := 'I';
      Item[1, 3].ColumnWidth := 2;
      Item[1, 3].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 3].HorizontalAlignment := xlHAlignCenter;
      Item[1, 4].Value := Month1;
      Item[1, 4].ColumnWidth := 12;
      Item[1, 4].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 4].HorizontalAlignment := xlHAlignRight;
      Item[1, 5].Value := 'I';
      Item[1, 5].ColumnWidth := 2;
      Item[1, 5].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 5].HorizontalAlignment := xlHAlignCenter;
      Item[1, 6].Value := Month2;
      Item[1, 6].ColumnWidth := 12;
      Item[1, 6].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 6].HorizontalAlignment := xlHAlignRight;
      Item[1, 7].Value := 'I';
      Item[1, 7].ColumnWidth := 2;
      Item[1, 7].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 7].HorizontalAlignment := xlHAlignCenter;
      Item[1, 8].Value := Month3;
      Item[1, 8].ColumnWidth := 12;
      Item[1, 8].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 8].HorizontalAlignment := xlHAlignRight;
      Item[1, 9].Value := 'I';
      Item[1, 9].ColumnWidth := 2;
      Item[1, 9].Borders[xlEdgeRight].Weight := xlThin;
      Item[1, 9].HorizontalAlignment := xlHAlignCenter;
      Item[1,10].Value := Month4;
      Item[1,10].ColumnWidth := 12;
      Item[1,10].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,10].HorizontalAlignment := xlHAlignRight;
      Item[1,11].Value := 'I';
      Item[1,11].ColumnWidth := 2;
      Item[1,11].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,11].HorizontalAlignment := xlHAlignCenter;
      Item[1,12].Value := Month5;
      Item[1,12].ColumnWidth := 12;
      Item[1,12].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,12].HorizontalAlignment := xlHAlignRight;
      Item[1,13].Value := 'I';
      Item[1,13].ColumnWidth := 2;
      Item[1,13].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,13].HorizontalAlignment := xlHAlignCenter;
      Item[1,14].Value := Month6 + '+ < ';
      Item[1,14].ColumnWidth := 12;
      Item[1,14].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,14].HorizontalAlignment := xlHAlignRight;
      Item[1,15].Value := 'I';
      Item[1,15].ColumnWidth := 2;
      Item[1,15].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,15].HorizontalAlignment := xlHAlignCenter;
      Item[1,16].Value := 'Not Invoiced';
      Item[1,16].ColumnWidth := 12;
      Item[1,16].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,16].HorizontalAlignment := xlHAlignRight;
      Item[1,17].Value := 'S';
      Item[1,17].ColumnWidth := 2;
      Item[1,17].Borders[xlEdgeRight].Weight := xlThin;
      Item[1,17].HorizontalAlignment := xlHAlignCenter;
      Interior.Color := ColDHF;
      Font.Color := ColDHT;
      Font.Bold := true;
      Font.Name := 'Arial';
      Font.Size := 10;
   end;
   Inc(row);
end;

//---------------------------------------------------------------------------
// Function to extract the heading and layout information
//---------------------------------------------------------------------------
procedure TFldExcel.GetLayout(RunTimeTemplate: string);
var
   StrLength, Delimiter : integer;
   ThisLayout           : string;

begin

//--- Ensure that the Header is handled correctly for Invoices

   if (RunTimeTemplate = 'Invoice') then begin
      ThisLayout := Layout_I;

      StrLength := Length(Header_I);
      Delimiter := Pos('|',Header_I);

      if (ShowVat = 0) then begin
         Header_X := Copy(Header_I,1,Delimiter - 1);
      end else begin
         Header_X := Copy(Header_I,Delimiter + 1, (StrLength - Delimiter));
      end;

   end else if (RunTimeTemplate = 'Specified Account') then begin
      ThisLayout := Layout_A;
      Header_X   := Header_A;
   end else if (RunTimeTemplate = 'Statement') then begin
      ThisLayout := Layout_S;
      Header_X   := Header_S;
   end else if (RunTimeTemplate = 'Trust') then begin
      ThisLayout := Layout_T;
      Header_X   := Header_T;
   end else if (RunTimeTemplate = 'Quote') then begin
      ThisLayout := Layout_Q;
      Header_X   := Header_Q;
   end;

//--- Now extract the layout variables from the Layout Code

   lcShowHeader    := StrToBool(Copy(ThisLayout, 2,1));
   lcHeaderPageOne := StrToBool(Copy(ThisLayout, 3,1));
   lcHSR           := StrToInt(Copy(ThisLayout,  4,2));
   lcHER           := StrToInt(Copy(ThisLayout,  6,2));
   lcHSC           := StrToInt(Copy(ThisLayout,  8,2));
   lcHEC           := StrToInt(Copy(ThisLayout, 10,2));
   lcShowAddress   := StrToBool(Copy(ThisLayout,12,1));
   lcASR           := StrToInt(Copy(ThisLayout, 13,2));
   lcAER           := StrToInt(Copy(ThisLayout, 15,2));
   lcASC           := StrToInt(Copy(ThisLayout, 17,2));
   lcShowInstruct  := StrToBool(Copy(ThisLayout,19,1));
   lcISR           := StrToInt(Copy(ThisLayout, 20,2));
   lcISCL          := StrToInt(Copy(ThisLayout, 22,2));
   lcISCD          := StrToInt(Copy(ThisLayout, 24,2));
   lcShowSummary   := StrToBool(Copy(ThisLayout,26,1));
   lcXSR           := StrToInt(Copy(ThisLayout, 27,2));
   lcXSCL          := StrToInt(Copy(ThisLayout, 29,2));
   lcXSCD          := StrToInt(Copy(ThisLayout, 31,2));
   lcPSR           := StrToInt(Copy(ThisLayout, 33,2));
   lcPRows         := StrToInt(Copy(ThisLayout, 35,2));
   lcShowBanking   := StrToBool(Copy(ThisLayout,37,1));
   lcBSR           := StrToInt(Copy(ThisLayout, 38,2));
   lcBSC           := StrToInt(Copy(ThisLayout, 40,2));

//--- Following could be either subsequent page or Statement details data

   lcSSR           := StrToInt(Copy(ThisLayout, 42,2));
   lcSRows         := StrToInt(Copy(ThisLayout, 44,2));
   lcSMaxRows      := StrToInt(Copy(ThisLayout, 46,2));
   lcSMaxCols      := StrToInt(Copy(ThisLayout, 48,2));

   lcSCB           := StrToInt(Copy(ThisLayout, 42,2));
   lcSCBD          := StrToInt(Copy(ThisLayout, 46,2));
   lcSCT           := StrToInt(Copy(ThisLayout, 44,2));
   lcSCTD          := StrToInt(Copy(ThisLayout, 48,2));

//--- Get the rest

   lcShowAge       := StrToBool(Copy(ThisLayout,50,1));
   lcAASR          := StrToInt(Copy(ThisLayout, 51,2));
   lcAASC          := StrToInt(Copy(ThisLayout, 53,2));

end;

//===========================================================================
//===========================================================================
//===                                                                     ===
//=== General support functions and procedures                            ===
//===                                                                     ===
//===========================================================================
//===========================================================================

//---------------------------------------------------------------------------
// Function to retrieve the Company's VAT number
//---------------------------------------------------------------------------
procedure TFldExcel.GetCpyVAT();
var
   S1  : string;

begin

   S1 := 'SELECT VATRegistered, VATNumber, VATRate, CpyName, CpyFile FROM lpms WHERE Signature = ''Legal Practice Management System - LPMS''';

//   Screen.Cursor := crHourGlass;

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ShowVAT   := 0;
      VATNumber := '';
      Exit;
   end;

   ShowVAT := Query2.FieldByName('VATRegistered').AsInteger;

   if (ShowVAT = 1) then begin
      VATNumber := Query2.FieldByName('VATNumber').AsString;
      VATRate   := Query2.FieldByName('VATRate').AsFloat;
   end else begin
      VATNumber := '';
      VATRate   := 0.00;
   end;

   CpyName := Query2.FieldByName('CpyName').AsString;
   CpyFile := Query2.FieldByName('CpyFile').AsString;

end;

//---------------------------------------------------------------------------
// Function to retrieve the VAT number for a customer
//---------------------------------------------------------------------------
function TFldExcel.GetVATNum(FileName: string): string;
var
   S1, S2  : string;

begin

   S1 := 'SELECT Tracking_ClientKey FROM tracking WHERE Tracking_Name = ''' +
         FileName + '''';

//--- Get unique key for this customer from the tracking table

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (tracking)';
      Result := '';
      Exit;
   end;

   S2 := 'SELECT Cust_VATNum FROM customers WHERE Cust_TimeStamp = ''' +
         Query2.FieldByName('Tracking_ClientKey').AsString + '''';

   try
      Query2.Close;
      Query2.SQL.Text := S2;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (customers)';
      Result := '';
      Exit;
   end;

   Result := Query2.FieldByName('Cust_VATNum').AsString;
end;

//---------------------------------------------------------------------------
// Procedure to retrieve the Address for a customer
//---------------------------------------------------------------------------
procedure TFldExcel.GetAddress(FileName: string);
var
   A1, A2, A3 : string;
   S1, S2     : string;

begin

   S1 := 'SELECT Tracking_ClientKey FROM tracking WHERE Tracking_Name = ''' +
         FileName + '''';

//--- Get unique key for this customer from the tracking table

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (tracking)';
      Exit;
   end;

   S2 := 'SELECT * FROM customers WHERE Cust_TimeStamp = ''' +
         Query2.FieldByName('Tracking_ClientKey').AsString + '''';

   try
      Query2.Close;
      Query2.SQL.Text := S2;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (customers)';
      Exit;
   end;

//--- Check which addresses are available

   A1 := Query2.FieldByName('Cust_Address1').AsString;
   A2 := Query2.FieldByName('Cust_Postal1').AsString;
   A3 := Query2.FieldByName('Cust_Work1').AsString;

//--- Get the Customer's address

   if (A1 <> '') then begin
      Address1 := Query2.FieldByName('Cust_Address1').AsString;
      Address2 := Query2.FieldByName('Cust_Address2').AsString;
      Address3 := Query2.FieldByName('Cust_Address3').AsString;
      Address4 := Query2.FieldByName('Cust_Address4').AsString;
      Address5 := Query2.FieldByName('Cust_PostCode').AsString;
   end else if (A2 <> '') then begin
      Address1 := Query2.FieldByName('Cust_Postal1').AsString;
      Address2 := Query2.FieldByName('Cust_Postal2').AsString;
      Address3 := Query2.FieldByName('Cust_Postal3').AsString;
      Address4 := Query2.FieldByName('Cust_Postal4').AsString;
      Address5 := Query2.FieldByName('Cust_PostalCode').AsString;
   end else begin
      Address1 := Query2.FieldByName('Cust_Work1').AsString;
      Address2 := Query2.FieldByName('Cust_Work2').AsString;
      Address3 := Query2.FieldByName('Cust_Work3').AsString;
      Address4 := Query2.FieldByName('Cust_Work4').AsString;
      Address5 := Query2.FieldByName('Cust_Postwork').AsString;
   end;

//--- Optimise the addresses by removing blank fields << CHANGE >>

   if ((Address5 <> '') and (Address4 = '')) then begin
      Address4 := Address5;
      Address5 := '';
   end;

   if ((Address4 <> '') and (Address3 = '')) then begin
      Address3 := Address4;
      Address4 := '';
   end;

   if ((Address3 <> '') and (Address2 = '')) then begin
      Address2 := Address3;
      Address3 := '';
   end;

   if ((Address2 <> '') and (Address1 = '')) then begin
      Address1 := Address2;
      Address2 := '';
   end;
end;

//---------------------------------------------------------------------------
// Procedure to retrieve the Description for a File
//---------------------------------------------------------------------------
function TFldExcel.GetDescription(FileName: string): string;
var
   S1 : string;

begin

   S1 := 'SELECT Tracking_Description FROM tracking WHERE Tracking_Name = ''' +
         FileName + '''';

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (tracking)';
      Result := '** Not Found';
      Exit;
   end;

   Result := Query2.FieldByName('Tracking_Description').AsString;

end;

//---------------------------------------------------------------------------
// Procedure to retrieve the Fee Earner's details
//---------------------------------------------------------------------------
function TFldExcel.GetUser(UserID: string): boolean;
var
   S1 : string;

begin

   S1 := 'SELECT Control_Name, Control_Email, Control_Unique FROM control WHERE Control_UserID = ''' +
         UserID + '''';

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (tracking)';
      Result := false;
      Exit;
   end;

   Result := true;

end;

//---------------------------------------------------------------------------
// Procedure to retrieve all the User Records
//---------------------------------------------------------------------------
function TFldExcel.GetAllUsers(ThisStr: string): boolean;
var
   S1 : string;

begin

   S1 := 'SELECT Control_UserID FROM control' + ThisStr;

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (tracking)';
      Result := false;
      Exit;
   end;

   Result := true;

end;

//---------------------------------------------------------------------------
// Procedure to retrieve all the FileNames
//---------------------------------------------------------------------------
function TFldExcel.GetFileNames(Filter: string): boolean;
var
   S1 : string;

begin

   S1 := 'SELECT Tracking_Name FROM tracking WHERE Tracking_Name LIKE ''' + Filter + ''' ORDER BY Tracking_Name';

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (tracking)';
      Result := false;
      Exit;
   end;

   Result := true;

end;

//---------------------------------------------------------------------------
// Function to retrieve the Billing records
//---------------------------------------------------------------------------
function TFldExcel.GetBilling(FileName: string; ThisStr: string; ThisType: string): boolean;
var
   S1, S2  : string;

begin

//--- Get customer details

   S1 := 'SELECT Tracking_Description, Tracking_Client, Tracking_Opposition FROM tracking WHERE Tracking_Name = ''' + FileName + '''';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (tracking)';
      Result := false;
      Exit;
   end;

   Customer   := Query1.FieldByName('Tracking_Client').AsString;
   Descrip    := Query1.FieldByName('Tracking_Description').AsString;

//--- Calculate the Opening Balances

   if (GetBalance(FileName) = false) then begin
      Result := false;
      Exit;
   end;

   if (VATRate > 0) then
      OpenBalVAT := (This_Bal.Fees * (VATRate /100)) - ((This_Bal.Credit * (VATRate / 100)) + (This_Bal.Write_off * (VATRate / 100)))
   else
      OpenBalVAT := 0;

   if (ThisType = 'Specified Account') then begin
      OpenBalFees     := ((This_Bal.Fees + This_Bal.Disbursements + This_Bal.Expenses) + (This_Bal.Credit + This_Bal.Write_off));
      OpenBalTrust    := (This_Bal.Trust_Deposit + This_Bal.Business_To_Trust + This_Bal.Trust_Withdrawal_S86_4) - (This_Bal.Trust_Transfer_Business_Fees + This_Bal.Trust_Transfer_Business_Other + This_Bal.Trust_Transfer_Client + This_Bal.Trust_Transfer_Disbursements + This_Bal.Trust_Transfer_Trust + This_Bal.Trust_Debit + This_Bal.Trust_Investment_S86_4);
      OpenBalReserve  := This_Bal.Reserved_Trust;
      OpenBal864Int := This_Bal.Trust_Interest_S86_4;
      OpenBal864Inv := This_Bal.Trust_Investment_S86_4;
      OpenBal864Drw := This_Bal.Trust_Withdrawal_S86_4;
      OpenBal864    := OpenBal864Int + OpenBal864Inv + OpenBal864Drw;
   end;

   if (ThisType = 'Invoice') then begin
      //
   end;

   if (ThisType = 'Statement') then begin
      StatementFees     := This_Bal.Fees + This_Bal.Credit + This_Bal.Write_off + OpenBalVAT;
      StatementDisburse := This_Bal.Disbursements;
      StatementExpenses := This_Bal.Expenses;
      StatementDeposits := This_Bal.Business_Deposit + This_Bal.Payment_Received;
      StatementBusPay   := This_Bal.Business_Debit + This_Bal.Business_To_Trust;
      StatementTrustDep := This_Bal.Trust_Deposit + (This_Bal.Business_To_Trust * -1);
      StatementReserve  := This_Bal.Reserved_Trust;
      StatementTrustInt := This_Bal.Trust_Interest_S86_4;
      StatementTrustPay := This_Bal.Trust_Transfer_Business_Fees + This_Bal.Trust_Transfer_Business_Other;
      StatementTrustDis := This_Bal.Trust_Transfer_Disbursements + This_Bal.Trust_Transfer_Client + This_Bal.Trust_Transfer_Trust + This_Bal.Trust_Debit;
   end;

   if (ThisType = 'Trust Simple') then begin
      OpenBalTrust   := (This_Abs.Trust_Deposit + This_Abs.Trust_Withdrawal_S86_4 + This_Abs.Business_To_Trust) - (This_Abs.Trust_Transfer_Business_Fees + This_Abs.Trust_Transfer_Client + This_Abs.Trust_Transfer_Disbursements + This_Abs.Trust_Transfer_Trust + This_Abs.Trust_Investment_S86_4 + This_Abs.Trust_Debit);
      OpenBal864     := (This_Abs.Trust_Investment_S86_4 + This_Abs.Trust_Interest_S86_4) - This_Abs.Trust_Withdrawal_S86_4;
      OpenBalReserve := This_Abs.Reserved_Trust;

      TotTotals[1] := This_Bal.Trust_Deposit + This_Bal.Business_To_Trust;
      TotTotals[2] := This_Bal.Trust_Interest_S86_4;
      TotTotals[3] := This_Bal.Trust_Transfer_Business_Fees + This_Bal.Trust_Transfer_Business_Other;
      TotTotals[4] := This_Bal.Trust_Transfer_Disbursements + This_Bal.Trust_Transfer_Client;
      TotTotals[5] := This_Bal.Trust_Transfer_Trust;
      TotTotals[6] := This_Bal.Trust_Debit;
   end;

   if (ThisType = 'Trust Recon') then begin
      OpenBalTrust   := (This_Abs.Trust_Deposit + This_Abs.Trust_Withdrawal_S86_4 + This_Abs.Business_To_Trust) - (This_Abs.Trust_Transfer_Business_Fees + This_Abs.Trust_Transfer_Client + This_Abs.Trust_Transfer_Disbursements + This_Abs.Trust_Transfer_Trust + This_Abs.Trust_Investment_S86_4 + This_Abs.Trust_Debit);
      OpenBal864     := (This_Abs.Trust_Investment_S86_4 + This_Abs.Trust_Interest_S86_4) - This_Abs.Trust_Withdrawal_S86_4;
      OpenBalReserve := This_Abs.Reserved_Trust;
   end;

   if (ThisType = 'Section 86(3) Summary') then begin
      OpenBal863Inv    := This_Abs.Trust_Investment_S86_3;
      OpenBal863Int    := This_Abs.Trust_Interest_S86_3;
      OpenBal863Drw    := This_Abs.Trust_Withdrawal_S86_3;
      OpenBal863IntDrw := This_Abs.Trust_Interest_Withdrawal_S86_3;
   end;

   if (ThisType = 'Trust Management') then begin
      OpenBalFees  := ((This_Abs.Fees + This_Abs.Disbursements + This_Abs.Expenses + This_Abs.Business_To_Trust + This_Abs.Business_Debit) - (This_Abs.Credit + This_Abs.Write_off + This_Abs.Payment_Received + This_Abs.Business_Deposit + This_Abs.Trust_Transfer_Business_Fees));
      OpenBalTrust := (This_Abs.Trust_Deposit + This_Abs.Business_To_Trust + This_Abs.Trust_Withdrawal_S86_4) - (This_Abs.Trust_Transfer_Business_Fees + This_Abs.Trust_Transfer_Business_Other + This_Abs.Trust_Transfer_Client + This_Abs.Trust_Transfer_Disbursements + This_Abs.Trust_Transfer_Trust + This_Abs.Trust_Debit + This_Abs.Trust_Investment_S86_4);
      OpenBal864   := This_Abs.Trust_Investment_S86_4 + This_Abs.Trust_Interest_S86_4 - This_Abs.Trust_Withdrawal_S86_4;
   end;

//--- Get the current amounts for each class

   if (GetCurrent(FileName) = false) then begin
      Result := false;
      Exit;
   end;

//--- Now calculate the Summary amounts

   if (ThisType = 'Specified Account') then begin
      if (VATRate > 0) then
         SummaryVAT := (This_Bal.Fees + This_Bal.Credit + This_Bal.Write_off) * (VATRate /100);
   end;

   if (ThisType = 'Invoice') then begin
      if (VATRate > 0) then begin
         SummaryFees := (This_Bal.Fees + This_Bal.Credit + This_Bal.Write_off) * ((100 + VATRate) / 100);
         SummaryVAT  := (This_Bal.Fees + This_Bal.Credit + This_Bal.Write_off) * (VATRate / 100);
      end else begin
         SummaryFees := This_Bal.Fees + This_Bal.Credit + This_Bal.Write_off;
         SummaryVAT  := 0;
      end;

      SummaryDisburse := This_Bal.Disbursements;
      SummaryExpenses := This_Bal.Expenses;
   end;

   if (ThisType = 'Statement') then begin
      StatementFees     := StatementFees     + This_Bal.Fees + This_Bal.Credit + This_Bal.Write_off + SummaryVAT;
      StatementDisburse := StatementDisburse + This_Bal.Disbursements;
      StatementExpenses := StatementExpenses + This_Bal.Expenses;
      StatementBusPay   := StatementBusPay   + This_Bal.Business_Debit + This_Bal.Business_To_Trust;
      StatementDeposits := StatementDeposits + This_Bal.Business_Deposit + This_Bal.Payment_Received;
      StatementTrustDep := StatementTrustDep + This_Bal.Trust_Deposit + (This_Bal.Business_To_Trust * -1);
      StatementReserve  := StatementReserve  + This_Bal.Reserved_Trust;
      StatementTrustInt := StatementTrustInt + This_Bal.Trust_Interest_S86_4;
      StatementTrustPay := StatementTrustPay + This_Bal.Trust_Transfer_Business_Fees + This_Bal.Trust_Transfer_Business_Other;
      StatementTrustDis := StatementTrustDis + This_Bal.Trust_Transfer_Disbursements + This_Bal.Trust_Transfer_Client + This_Bal.Trust_Transfer_Trust + This_Bal.Trust_Debit;

      SDate := '1980/01/01';
   end;
{
   SummaryFees := SummaryFees * -1;

   SummaryDisburse := This_Bal.Disbursements * -1;
   SummaryExpenses := This_Bal.Expenses * -1;
}
{
   StatementFees     := StatementFees + This_Bal.Fees + This_Bal.Credit + This_Bal.Write_off + SummaryVAT;
   StatementDisburse := StatementDisburse + This_Bal.Disbursements;
   StatementExpenses := StatementExpenses + This_Bal.Expenses;
   StatementDeposits := StatementDeposits + This_Bal.Business_Deposit + This_Bal.Payment_Received;
   StatementTrustDep := StatementTrustDep + This_Bal.Trust_Deposit + This_Bal.Business_To_Trust;
   StatementReserve  := StatementReserve + This_Bal.Reserved_Trust;
   StatementTrustInt := StatementTrustInt + This_Bal.Trust_Interest_S86_4;
   StatementTrustPay := StatementTrustPay + This_Bal.Trust_Transfer_Business_Fees + This_Bal.Trust_Transfer_Business_Other;
   StatementTrustDis := StatementTrustDis + This_Bal.Trust_Transfer_Disbursements + This_Bal.Trust_Transfer_Client + This_Bal.Trust_Transfer_Trust + This_Bal.Trust_Debit;
}
{
   if (VATRate > 0) then begin
      if (AccountType = 0) then
         SummaryFees := ((AmtFees + AmtBusDebit) * (100 + VATRate) / 100) + AmtBusToTrust - (AmtCredit * (100 + VATRate) / 100)
      else
         SummaryFees := ((AmtFees + AmtBusDebit) * (100 + VATRate) / 100) - (AmtCredit * (100 + VATRate) / 100);

      SummaryVAT  := (AmtFees * (VATRate / 100)) - (AmtCredit * (VATRate / 100));
   end else begin
      if (AccountType = 0) then
         SummaryFees := AmtFees + AmtBusToTrust - AmtCredit + AmtBusDebit
      else
         SummaryFees := AmtFees - AmtCredit;

      SummaryVAT  := 0;
   end;

   SummaryFees := SummaryFees * -1;

   SummaryDisburse := AmtDisburse * -1;
   SummaryExpenses := AmtExpense * -1;

   StatementFees     := StatementFees + (AmtFees + AmtBusToTrust + AmtBusDebit) - AmtCredit + SummaryVAT;
   StatementDisburse := StatementDisburse + AmtDisburse;
   StatementExpenses := StatementExpenses + AmtExpense;
   StatementDeposits := StatementDeposits + AmtBusDeposit + AmtPayment;
   StatementTrustDep := StatementTrustDep + AmtTrustDeposit + AmtBusToTrust;
   StatementReserve  := StatementReserve + AmtTrustReserve;
   StatementTrustInt := StatementTrustInt + AmtInt782BA;
   StatementTrustPay := StatementTrustPay + AmtXfrFees;
   StatementTrustDis := StatementTrustDis + (AmtXfrDisburse + AmtXfrClient + AmtXfrTrust + AmtTrustDebit);

}
//--- Now retrieve the detail information for this request

   if (InvoiceInfo = '1') then begin
      if (ShowRelated = true) then
         S2 := 'Collect_Related'
      else
         S2 := 'Collect_Owner';
   end else begin
      if (ShowRelated = true) then
         S2 := 'B_Related'
      else
         S2 := 'B_Owner';
   end;

   if (InvoiceInfo = '1') then
      S1 := 'SELECT Collect_Owner, Collect_Related, Collect_Date, Collect_Description, Collect_DrCr, Collect_Amount, Collect_Class FROM collect WHERE ' + S2 + ' = ''' + FileName +
            ''' AND Collect_Date >= ''' + SDate + ''' AND Collect_Date <= ''' + EDate +  '''' +
            ThisStr + ' ORDER BY Collect_Date, Create_Date, Create_Time'
   else
      S1 := 'SELECT B_Owner, B_Related, B_Date, B_Description, B_DrCr, B_Amount, B_Class, B_ReserveAmt FROM billing WHERE ' + S2 + ' = ''' + FileName +
            ''' AND B_Date >= ''' + SDate + ''' AND B_Date <= ''' + EDate +  '''' +
            ThisStr + ' ORDER BY B_Date, Create_Date, Create_Time';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      if (InvoiceInfo = '1') then
         ErrMsg := '''Unable to read from ' + HostName + ''' (collect)'
      else
         ErrMsg := '''Unable to read from ' + HostName + ''' (billing)';

      Result := false;
      Exit;
   end;

   Result := true;
end;

//---------------------------------------------------------------------------
// Function to retrieve the Quote records
//---------------------------------------------------------------------------
function TFldExcel.GetQuote(FileName: string; QuoteName: string): boolean;
var
   S1  : string;

begin

//--- Get customer details

   S1 := 'SELECT Tracking_Description, Tracking_Client, Tracking_Opposition FROM tracking WHERE Tracking_Name = ''' + FileName + '''';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (tracking)';
      Result := false;
      Exit;
   end;

   Customer   := Query1.FieldByName('Tracking_Client').AsString;
   Descrip    := Query1.FieldByName('Tracking_Description').AsString;

//--- Get the current amounts for each class

   if (GetQuoteDetails(QuoteName) = false) then begin
      Result := false;
      Exit;
   end;

//--- Now calculate the Summary amounts

   if (VATRate > 0) then begin
      SummaryFees := (This_Bal.Fees + This_Bal.Credit + This_Bal.Write_off) * ((100 + VATRate) / 100);
      SummaryVAT  := (This_Bal.Fees + This_Bal.Credit + This_Bal.Write_off) * (VATRate / 100);
   end else begin
      SummaryFees := This_Bal.Fees + This_Bal.Credit + This_Bal.Write_off;
      SummaryVAT  := 0;
   end;

   SummaryDisburse := This_Bal.Disbursements;
   SummaryExpenses := This_Bal.Expenses;

//--- Now retrieve the detail information for this request

   S1 := 'SELECT Q_Owner, Q_Date, Q_Description, Q_DrCr, Q_Amount, Q_Class FROM quotes WHERE Q_Quote = ''' +
         QuoteName + ''' AND (Q_Date >= ''' + SDate + ''' AND Q_Date <= ''' +
         EDate + ''') ORDER BY Q_Date, Create_Date, Create_Time ASC';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (quotes)';

      Result := false;
      Exit;
   end;

   Result := true;
end;

//---------------------------------------------------------------------------
// Function to retrieve the value of a Quote
//---------------------------------------------------------------------------
function TFldExcel.GetQuoteVal(Quote: string; ThisHost: string): double;
var
   idx1, ThisClass                    : integer;
   Amount, Fees, Disburse, Expenses   : double;
   S1                                 : string;

begin

   S1 := 'SELECT Q_Class, Q_Amount FROM quotes WHERE Q_Quote = ''' + Quote +
         ''' ORDER BY Q_Date ASC';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + ThisHost + ''' (quotes)';
      Result := 0;
      Exit;
   end;

   Fees     := 0;
   Disburse := 0;
   Expenses := 0;

   Query1.First;

   for idx1 := 0 to Query1.RecordCount - 1 do begin
      ThisClass := Query1.FieldByName('Q_Class').AsInteger;

      case ThisClass of
         0: Fees     := Fees     + Query1.FieldByName('Q_Amount').AsFloat;
         1: Disburse := Disburse + Query1.FieldByName('Q_Amount').AsFloat;
         2: Expenses := Expenses + Query1.FieldByName('Q_Amount').AsFloat;
      end;

      Query1.Next;
   end;

   Amount := Fees + Disburse + Expenses;
   Result := Amount;
end;

//---------------------------------------------------------------------------
// Function to retrieve Billing records for the Billing Preparation report
//---------------------------------------------------------------------------
function TFldExcel.GetBillingR(ThisFile: string; StartDate: string; EndDate: string; var Billing: double; var Invoiced: double; var Paid: double; var ThisReserved: double; var ThisTrust: double): boolean;
var
   idx                    : integer;
   ThisVAT, InvAmount     : double;
   S1                     : string;

begin

//--- Get the Billing and Paid Amounts

   S1 := 'SELECT B_Amount, B_DrCr, B_Class, B_ReserveDep, B_ReserveAmt FROM billing WHERE B_Owner = ''' +
         ThisFile + ''' AND B_Date >= ''' + StartDate + ''' AND B_Date <= ''' +
         EndDate + ''' AND B_AccountType = 0';

   if (GetRecord(S1,2) = false) then begin
      MessageDlg(ErrMsg + ' - Report Generation aborted', mtError, [mbOK], 0);
      Result := false;
      Exit;
   end;

   GetAmounts(0,Query2);

//--- Inclue VAT if registered for VAT

   ThisVAT := 0;
   if VATRate > 0 then
      ThisVAT := (This_Abs.Fees * (VATRate / 100)) - ((This_Abs.Credit + This_Abs.Write_off) * (VATRate / 100));

//--- Calculate the Billing and Payments for the Period

   Billing := (This_Abs.Fees + This_Abs.Disbursements + This_Abs.Expenses + ThisVAT) - (This_Abs.Credit + This_Abs.Write_off);
   Paid    := This_Abs.Payment_Received + This_Abs.Business_Deposit + This_Abs.Trust_Transfer_Business_Fees - This_Abs.Business_To_Trust;

//--- Calculate what is on Trust and what is reserved

   ThisTrust    := ThisTrust + (This_Abs.Trust_Deposit + This_Abs.Business_To_Trust + This_Abs.Trust_Interest_S86_4) - (This_Abs.Trust_Transfer_Business_Fees + This_Abs.Trust_Transfer_Client + This_Abs.Trust_Transfer_Disbursements + This_Abs.Trust_Debit);
   ThisReserved := ThisReserved + This_Abs.Reserved_Trust;

//--- Calculate the Invoiced Amount

   S1 := 'SELECT Inv_Amount FROM invoices WHERE Inv_File = ''' + ThisFile +
         ''' AND Inv_SDate >= ''' + StartDate + ''' AND Inv_SDate <= ''' +
         EndDate + '''';

   if (GetRecord(S1,2) = false) then begin
      MessageDlg(ErrMsg + ' - Report Generation aborted', mtError, [mbOK], 0);
      Result := false;
      Exit;
   end;

   InvAmount := 0;

   Query2.First;
   for idx := 0 to Query2.RecordCount -1 do begin
      InvAmount := InvAmount + Query2.FieldByName('Inv_Amount').AsFloat;
      Query2.Next;
   end;

   Invoiced := InvAmount;
   Result := true;
end;

//---------------------------------------------------------------------------
// Function to get the opening balance for each type of billing class
//---------------------------------------------------------------------------
function  TFldExcel.GetBalance(FileName: string): boolean;
var
//   idx, ThisClass : integer;
   S1, S2, S3, S4 : string;
//   This_Bal, This_Abs : LPMS_Amounts;

begin

   if (InvoiceInfo = '1') then begin
      S3 := 'Collect';
      S4 := 'collect';
      if (ShowRelated = true) then
         S2 := 'Collect_Related'
      else
         S2 := 'Collect_Owner';
   end else begin
      S3 := 'B';
      S4 := 'billing';
      if (ShowRelated = true) then
         S2 := 'B_Related'
      else
         S2 := 'B_Owner';
   end;

//--- Set up and read billing information from the Database

   S1 := 'SELECT * FROM ' + S4 + ' WHERE ' + S2 + ' = ''' + FileName + ''' AND ' + S3 + '_Date < ''' + SDate + ''' AND ' + S3 + '_AccountType = ' + IntToStr(AccountType);
   if (GetTotals(S1) = -1) then begin
      Result := false;
      Exit;
   end;

//--- Calculate the totals for each Billing Class

   GetAmounts(StrToInt(InvoiceInfo),Query1{,This_Bal,This_Abs});

{
   if (ParamStr(15) = '1') then
      S1 := 'SELECT Collect_Amount, Collect_Class FROM collect WHERE ((Collect_Class >= 0 AND Collect_Class <= 3) OR (Collect_Class >= 5 AND Collect_Class <= 6) OR (Collect_Class = 18)) AND ' + S2 + ' = ''' + FileName + ''' AND Collect_Date < ''' + SDate + ''' AND Collect_AccountType = ' + IntToStr(AccountType)
   else
      S1 := 'SELECT B_Amount, B_Class FROM billing WHERE ((B_Class >= 0 AND B_Class <= 3) OR (B_Class >= 5 AND B_Class <= 6) OR (B_Class = 18)) AND ' + S2 + ' = ''' + FileName + ''' AND B_Date < ''' + SDate + ''' AND B_AccountType = ' + IntToStr(AccountType);

   if (GetTotals(S1) = -1) then begin
      Result := false;
      Exit;
   end else begin
      BalFees       := 0;
      BalDisburse   := 0;
      BalExpense    := 0;
      BalPayment    := 0;
      BalCredit     := 0;
      BalBusDeposit := 0;
      BalBusDebit   := 0;

      for idx := 0 to Query1.RecordCount - 1 do begin

         if (ParamStr(15) = '1') then begin
            ThisClass := Query1.FieldByName('Collect_Class').AsInteger;
            case ThisClass of
                0: BalFees       := BalFees       + Query1.FieldByName('Collect_Amount').AsFloat;
                1: BalDisburse   := BalDisburse   + Query1.FieldByName('Collect_Amount').AsFloat;
                2: BalExpense    := BalExpense    + Query1.FieldByName('Collect_Amount').AsFloat;
                3: BalPayment    := BalPayment    + Query1.FieldByName('Collect_Amount').AsFloat;
                5: BalCredit     := BalCredit     + Query1.FieldByName('Collect_Amount').AsFloat;
                6: BalBusDeposit := BalBusDeposit + Query1.FieldByName('Collect_Amount').AsFloat;
               18: BalBusDebit   := BalBusDebit   + Query1.FieldByName('Collect_Amount').AsFloat;
            end;
         end else begin
            ThisClass := Query1.FieldByName('B_Class').AsInteger;
            case ThisClass of
                0: BalFees       := BalFees       + Query1.FieldByName('B_Amount').AsFloat;
                1: BalDisburse   := BalDisburse   + Query1.FieldByName('B_Amount').AsFloat;
                2: BalExpense    := BalExpense    + Query1.FieldByName('B_Amount').AsFloat;
                3: BalPayment    := BalPayment    + Query1.FieldByName('B_Amount').AsFloat;
                5: BalCredit     := BalCredit     + Query1.FieldByName('B_Amount').AsFloat;
                6: BalBusDeposit := BalBusDeposit + Query1.FieldByName('B_Amount').AsFloat;
               18: BalBusDebit   := BalBusDebit   + Query1.FieldByName('B_Amount').AsFloat;
            end;
         end;

         Query1.Next;
      end;
   end;

   if (ParamStr(15) = '1') then
      S1 := 'SELECT Collect_Amount FROM collect WHERE Collect_Class = 4 AND ' + S2 + ' = ''' + FileName + ''' AND Collect_Date < ''' + SDate + ''''
   else
      S1 := 'SELECT B_Amount FROM billing WHERE B_Class = 4 AND ' + S2 + ' = ''' + FileName + ''' AND B_Date < ''' + SDate + '''';

   if (GetTotals(S1) = -1) then begin
      Result := false;
      Exit;
   end else begin
      BalBusToTrust := 0;
      for idx := 0 to Query1.RecordCount - 1 do begin
         BalBusToTrust := BalBusToTrust + Query1.FieldByName('B_Amount').AsFloat;
         Query1.Next;
      end;
   end;

//--- Get the information to calculate the Opening Balances for Trust

   if (ParamStr(15) = '1') then
      S1 := 'SELECT Collect_Amount, Collect_Class FROM collect WHERE ((Collect_Class >= 7 AND Collect_Class <= 17) OR (Collect_Class = 19)) AND ' + S2 + ' = ''' + FileName + ''' AND Collect_Date < ''' + SDate + ''''
   else
      S1 := 'SELECT B_Amount, B_ReserveAmt, B_Class FROM billing WHERE ((B_Class >= 7 AND B_Class <= 17) OR (B_Class = 19)) AND ' + S2 + ' = ''' + FileName + ''' AND B_Date < ''' + SDate + '''';

   if (GetTotals(S1) = -1) then begin
      Result := false;
      Exit;
   end else begin
      BalTrustDeposit := 0;
      BalTrustReserve := 0;
      BalXfrFees      := 0;
      BalXfrDisburse  := 0;
      BalXfrClient    := 0;
      BalXfrTrust     := 0;
      BalInv782BA     := 0;
      BalDrw782BA     := 0;
      BalInt782BA     := 0;
      BalInv782Sa     := 0;
      BalDrw782Sa     := 0;
      BalInt782Sa     := 0;
      BalTrustDebit   := 0;

      for idx := 0 to Query1.RecordCount - 1 do begin

         if (ParamStr(15) = '1') then begin
            ThisClass := Query1.FieldByName('Collect_Class').AsInteger;
            case ThisClass of
                7: BalTrustDeposit := BalTrustDeposit + Query1.FieldByName('Collect_Amount').AsFloat;
                8: BalXfrFees      := BalXfrFees      + Query1.FieldByName('Collect_Amount').AsFloat;
                9: BalXfrDisburse  := BalXfrDisburse  + Query1.FieldByName('Collect_Amount').AsFloat;
               10: BalXfrClient    := BalXfrClient    + Query1.FieldByName('Collect_Amount').AsFloat;
               11: BalXfrTrust     := BalXfrTrust     + Query1.FieldByName('Collect_Amount').AsFloat;
               12: BalInv782BA     := BalInv782BA     + Query1.FieldByName('Collect_Amount').AsFloat;
               13: BalDrw782BA     := BalDrw782BA     + Query1.FieldByName('Collect_Amount').AsFloat;
               14: BalInt782BA     := BalInt782BA     + Query1.FieldByName('Collect_Amount').AsFloat;
               15: BalInv782Sa     := BalInv782Sa     + Query1.FieldByName('Collect_Amount').AsFloat;
               16: BalDrw782Sa     := BalDrw782Sa     + Query1.FieldByName('Collect_Amount').AsFloat;
               17: BalInt782Sa     := BalInt782Sa     + Query1.FieldByName('Collect_Amount').AsFloat;
               19: BalTrustDebit   := BalTrustDebit   + Query1.FieldByName('Collect_Amount').AsFloat;
            end;
         end else begin
            ThisClass := Query1.FieldByName('B_Class').AsInteger;
            case ThisClass of
                7: begin
                      BalTrustDeposit := BalTrustDeposit + Query1.FieldByName('B_Amount').AsFloat;
                      BalTrustReserve := BalTrustReserve + Query1.FieldByName('B_ReserveAmt').AsFloat;
                   end;
                8: BalXfrFees      := BalXfrFees      + Query1.FieldByName('B_Amount').AsFloat;
                9: BalXfrDisburse  := BalXfrDisburse  + Query1.FieldByName('B_Amount').AsFloat;
               10: BalXfrClient    := BalXfrClient    + Query1.FieldByName('B_Amount').AsFloat;
               11: BalXfrTrust     := BalXfrTrust     + Query1.FieldByName('B_Amount').AsFloat;
               12: BalInv782BA     := BalInv782BA     + Query1.FieldByName('B_Amount').AsFloat;
               13: BalDrw782BA     := BalDrw782BA     + Query1.FieldByName('B_Amount').AsFloat;
               14: BalInt782BA     := BalInt782BA     + Query1.FieldByName('B_Amount').AsFloat;
               15: BalInv782Sa     := BalInv782Sa     + Query1.FieldByName('B_Amount').AsFloat;
               16: BalDrw782Sa     := BalDrw782Sa     + Query1.FieldByName('B_Amount').AsFloat;
               17: BalInt782Sa     := BalInt782Sa     + Query1.FieldByName('B_Amount').AsFloat;
               19: BalTrustDebit   := BalTrustDebit   + Query1.FieldByName('B_Amount').AsFloat;
            end;
         end;

         Query1.Next;
      end;
   end;
}

   Result := true;
end;

//---------------------------------------------------------------------------
// Function to get the amounts for the current period for each billing class
//---------------------------------------------------------------------------
function  TFldExcel.GetCurrent(FileName: string): boolean;
var
//   idx, ThisClass : integer;
   S1, S2, S3, S4 : string;

begin

   if (InvoiceInfo = '1') then begin
      S3 := 'Collect';
      S4 := 'collect';
      if (ShowRelated = true) then
         S2 := 'Collect_Related'
      else
         S2 := 'Collect_Owner';
   end else begin
      S3 := 'B';
      S4 := 'billing';
      if (ShowRelated = true) then
         S2 := 'B_Related'
      else
         S2 := 'B_Owner';
   end;

//--- Get the information to calculate Fees, Disbursements and Expenses

   S1 := 'SELECT * FROM ' + S4 + ' WHERE ' + S2 + ' = ''' + FileName + ''' AND ' + S3 + '_Date >= ''' + SDate + ''' AND ' + S3 + '_Date <= ''' + EDate + ''' AND ' + S3 + '_AccountType = ' + IntToStr(AccountType);
   if (GetTotals(S1) = -1) then begin
      Result := false;
      Exit;
   end;

//--- Calculate the totals for each Billing Class

   GetAmounts(StrToInt(InvoiceInfo),Query1{,This_Bal,This_Abs});

{
   if ((RunType = 1) or (RunType = 3)) then begin

      if (ParamStr(15) = '1') then
         S1 := 'SELECT Collect_Amount, Collect_Class, Collect_Owner, Collect_Related FROM collect WHERE ((Collect_Class >= 0 AND Collect_Class <= 3) OR (Collect_Class >= 5 AND Collect_Class <= 6) OR (Collect_Class = 18)) AND ' + S2 + ' = ''' + FileName + ''' AND Collect_Date >= ''' + SDate + ''' AND Collect_Date <= ''' + EDate + ''' AND Collect_AccountType = ' + IntToStr(AccountType)
      else
         S1 := 'SELECT B_Amount, B_Class, B_Owner, B_Related FROM billing WHERE ((B_Class >= 0 AND B_Class <= 3) OR (B_Class >= 5 AND B_Class <= 6) OR (B_Class = 18)) AND ' + S2 + ' = ''' + FileName + ''' AND B_Date >= ''' + SDate + ''' AND B_Date <= ''' + EDate + ''' AND B_AccountType = ' + IntToStr(AccountType);

      if (GetTotals(S1) = -1) then begin
         Result := false;
         Exit;
      end else begin
         if (ParamStr(15) = '1') then begin
            SaveRelated := Query1.FieldByName('Collect_Related').AsString;
            SaveOwner   := Query1.FieldByName('Collect_Owner').AsString;
         end else begin
            SaveRelated := Query1.FieldByName('B_Related').AsString;
            SaveOwner   := Query1.FieldByName('B_Owner').AsString;
         end;

         AmtFees       := 0;
         AmtDisburse   := 0;
         AmtExpense    := 0;
         AmtPayment    := 0;
         AmtCredit     := 0;
         AmtBusDeposit := 0;
         AmtBusDebit   := 0;

         for idx := 0 to Query1.RecordCount - 1 do begin
            if (ParamStr(15) = '1') then begin
               ThisClass := Query1.FieldByName('Collect_Class').AsInteger;
               case ThisClass of
                   0: AmtFees       := AmtFees       + Query1.FieldByName('Collect_Amount').AsFloat;
                   1: AmtDisburse   := AmtDisburse   + Query1.FieldByName('Collect_Amount').AsFloat;
                   2: AmtExpense    := AmtExpense    + Query1.FieldByName('Collect_Amount').AsFloat;
                   3: AmtPayment    := AmtPayment    + Query1.FieldByName('Collect_Amount').AsFloat;
                   5: AmtCredit     := AmtCredit     + Query1.FieldByName('Collect_Amount').AsFloat;
                   6: AmtBusDeposit := AmtBusDeposit + Query1.FieldByName('Collect_Amount').AsFloat;
                  18: AmtBusDebit   := AmtBusDebit   + Query1.FieldByName('Collect_Amount').AsFloat;
               end;
            end else begin
               ThisClass := Query1.FieldByName('B_Class').AsInteger;
               case ThisClass of
                   0: AmtFees       := AmtFees       + Query1.FieldByName('B_Amount').AsFloat;
                   1: AmtDisburse   := AmtDisburse   + Query1.FieldByName('B_Amount').AsFloat;
                   2: AmtExpense    := AmtExpense    + Query1.FieldByName('B_Amount').AsFloat;
                   3: AmtPayment    := AmtPayment    + Query1.FieldByName('B_Amount').AsFloat;
                   5: AmtCredit     := AmtCredit     + Query1.FieldByName('B_Amount').AsFloat;
                   6: AmtBusDeposit := AmtBusDeposit + Query1.FieldByName('B_Amount').AsFloat;
                  18: AmtBusDebit   := AmtBusDebit   + Query1.FieldByName('B_Amount').AsFloat;
               end;
            end;

            Query1.Next;
         end;
      end;

      if (ParamStr(15) = '1') then
         S1 := 'SELECT Collect_Amount, Collect_Owner, Collect_Related FROM collect WHERE Collect_Class = 4 AND ' + S2 + ' = ''' + FileName + ''' AND Collect_Date >= ''' + SDate + ''' AND Collect_Date <= ''' + EDate + ''''
      else
         S1 := 'SELECT B_Amount, B_Owner, B_Related FROM billing WHERE B_Class = 4 AND ' + S2 + ' = ''' + FileName + ''' AND B_Date >= ''' + SDate + ''' AND B_Date <= ''' + EDate + '''';

      if (GetTotals(S1) = -1) then begin
         Result := false;
         Exit;
      end else begin
         if (ParamStr(15) = '1') then begin
            SaveRelated := Query1.FieldByName('Collect_Related').AsString;
            SaveOwner   := Query1.FieldByName('Collect_Owner').AsString;
         end else begin
            SaveRelated := Query1.FieldByName('B_Related').AsString;
            SaveOwner   := Query1.FieldByName('B_Owner').AsString;
         end;

         AmtBusToTrust := 0;
         for idx := 0 to Query1.RecordCount - 1 do begin
            if (ParamStr(15) = '1') then
               AmtBusToTrust := AmtBusToTrust + Query1.FieldByName('Collect_Amount').AsFloat
            else
               AmtBusToTrust := AmtBusToTrust + Query1.FieldByName('B_Amount').AsFloat;

            Query1.Next;
         end;
      end;
   end;

//--- Get the information to calculate Trust Amounts

   if ((RunType = 2) or (RunType = 3)) then begin
      if (ParamStr(15) = '1') then
         S1 := 'SELECT Collect_Amount, Collect_Class, Collect_Owner, Collect_Related FROM collect WHERE ((Collect_Class >= 7 AND Collect_Class <= 17) OR (Collect_Class = 19)) AND ' + S2 + ' = ''' + FileName + ''' AND Collect_Date >= ''' + SDate + ''' AND Collect_Date <= ''' + EDate + ''' AND Collect_AccountType = ' + IntToStr(AccountType)
      else
         S1 := 'SELECT B_Amount, B_Class, B_Owner, B_Related, B_ReserveAmt FROM billing WHERE ((B_Class >= 7 AND B_Class <= 17) OR (B_Class = 19)) AND ' + S2 + ' = ''' + FileName + ''' AND B_Date >= ''' + SDate + ''' AND B_Date <= ''' + EDate + ''' AND B_AccountType = ' + IntToStr(AccountType);

      if (GetTotals(S1) = -1) then begin
         Result := false;
         Exit;
      end else begin
         if (ParamStr(15) = '1') then begin
            SaveRelated := Query1.FieldByName('Collect_Related').AsString;
            SaveOwner   := Query1.FieldByName('Collect_Owner').AsString;
         end else begin
            SaveRelated := Query1.FieldByName('B_Related').AsString;
            SaveOwner   := Query1.FieldByName('B_Owner').AsString;
         end;

         AmtTrustDeposit := 0;
         AmtTrustReserve := 0;
         AmtXfrFees      := 0;
         AmtXfrDisburse  := 0;
         AmtXfrClient    := 0;
         AmtXfrTrust     := 0;
         AmtInv782BA     := 0;
         AmtDrw782BA     := 0;
         AmtInt782BA     := 0;
         AmtInv782Sa     := 0;
         AmtDrw782Sa     := 0;
         AmtInt782Sa     := 0;
         AmtTrustDebit   := 0;

         for idx := 0 to Query1.RecordCount - 1 do begin

            if (ParamStr(15) = '1') then begin
               ThisClass := Query1.FieldByName('Collect_Class').AsInteger;
               case ThisClass of
                   7: AmtTrustDeposit := AmtTrustDeposit + Query1.FieldByName('Collect_Amount').AsFloat;
                   8: AmtXfrFees      := AmtXfrFees      + Query1.FieldByName('Collect_Amount').AsFloat;
                   9: AmtXfrDisburse  := AmtXfrDisburse  + Query1.FieldByName('Collect_Amount').AsFloat;
                  10: AmtXfrClient    := AmtXfrClient    + Query1.FieldByName('Collect_Amount').AsFloat;
                  11: AmtXfrTrust     := AmtXfrTrust     + Query1.FieldByName('Collect_Amount').AsFloat;
                  12: AmtInv782BA     := AmtInv782BA     + Query1.FieldByName('Collect_Amount').AsFloat;
                  13: AmtDrw782BA     := AmtDrw782BA     + Query1.FieldByName('Collect_Amount').AsFloat;
                  14: AmtInt782BA     := AmtInt782BA     + Query1.FieldByName('Collect_Amount').AsFloat;
                  15: AmtInv782Sa     := AmtInv782Sa     + Query1.FieldByName('Collect_Amount').AsFloat;
                  16: AmtDrw782Sa     := AmtDrw782Sa     + Query1.FieldByName('Collect_Amount').AsFloat;
                  17: AmtInt782Sa     := AmtInt782Sa     + Query1.FieldByName('Collect_Amount').AsFloat;
                  19: AmtTrustDebit   := AmtTrustDebit   + Query1.FieldByName('Collect_Amount').AsFloat;
               end;
            end else begin
               ThisClass := Query1.FieldByName('B_Class').AsInteger;
               case ThisClass of
                   7: begin
                         AmtTrustDeposit := AmtTrustDeposit + Query1.FieldByName('B_Amount').AsFloat;
                         AmtTrustReserve := AmtTrustReserve + Query1.FieldByName('B_ReserveAmt').AsFloat;
                      end;
                   8: AmtXfrFees      := AmtXfrFees      + Query1.FieldByName('B_Amount').AsFloat;
                   9: AmtXfrDisburse  := AmtXfrDisburse  + Query1.FieldByName('B_Amount').AsFloat;
                  10: AmtXfrClient    := AmtXfrClient    + Query1.FieldByName('B_Amount').AsFloat;
                  11: AmtXfrTrust     := AmtXfrTrust     + Query1.FieldByName('B_Amount').AsFloat;
                  12: AmtInv782BA     := AmtInv782BA     + Query1.FieldByName('B_Amount').AsFloat;
                  13: AmtDrw782BA     := AmtDrw782BA     + Query1.FieldByName('B_Amount').AsFloat;
                  14: AmtInt782BA     := AmtInt782BA     + Query1.FieldByName('B_Amount').AsFloat;
                  15: AmtInv782Sa     := AmtInv782Sa     + Query1.FieldByName('B_Amount').AsFloat;
                  16: AmtDrw782Sa     := AmtDrw782Sa     + Query1.FieldByName('B_Amount').AsFloat;
                  17: AmtInt782Sa     := AmtInt782Sa     + Query1.FieldByName('B_Amount').AsFloat;
                  19: AmtTrustDebit   := AmtTrustDebit   + Query1.FieldByName('B_Amount').AsFloat;
               end;
            end;

            Query1.Next;
         end;
      end;
   end;

}
   Result := true;
end;

//---------------------------------------------------------------------------
// Function to Extract the FileName for a Quote
//---------------------------------------------------------------------------
function TfldExcel.GetQuoteFile(QuoteName: string): string;
var
   S1    : string;

begin
   S1 := 'SELECT DISTINCT Q_Owner FROM quotes WHERE Q_Quote = ''' + QuoteName + '''';

   if (GetTotals(S1) = -1) then begin
      Result := '';
      Exit;
   end;

   Result := Query1.FieldByName('Q_Owner').AsString;
end;

//---------------------------------------------------------------------------
// Function to get the amounts for the current period for each billing class
//---------------------------------------------------------------------------
function  TFldExcel.GetQuoteDetails(QuoteName: string): boolean;
var
   idx1, DrCr  : integer;
   S1          : string;
   ThisAmount  : double;

begin

//--- Get the information to calculate Fees, Disbursements and Expenses

   S1 := 'SELECT * FROM quotes WHERE Q_Quote = ''' + QuoteName +
         ''' AND (Q_Date >= ''' + SDate + ''' AND Q_Date <= ''' + EDate +
         ''') ORDER BY Q_Date, Create_Date, Create_Time ASC';

   if (GetTotals(S1) = -1) then begin
      Result := false;
      Exit;
   end;

//--- Clear the structures that will hold the amounts for each kind of Billig record

   Clear_Amounts();

//--- Calculate the totals for Fees, Expenses and Disbursement

   Query1.First;
   for idx1 := 0 to Query1.RecordCount - 1 do begin
      DrCr        := Query1.FieldByName('Q_DrCr').AsInteger;
      ThisAmount  := Query1.FieldByName('Q_Amount').AsFloat;

      case Query1.FieldByName('Q_Class').AsInteger of
         0: begin                  // Fees
            if (DrCr = 1) then
               This_Bal.Fees := This_Bal.Fees + (ThisAmount * -1.00)
            else
               This_Bal.Fees := This_Bal.Fees + ThisAmount;
            This_Abs.Fees := This_Abs.Fees + ThisAmount;
         end;

         1: begin                  // Disbursements
            if (DrCr = 1) then
               This_Bal.Disbursements := This_Bal.Disbursements + (ThisAmount * -1.00)
            else
               This_Bal.Disbursements := This_Bal.Disbursements + ThisAmount;
            This_Abs.Disbursements := This_Abs.Disbursements + ThisAmount;
         end;

         2: begin                  // Expenses
            if (DrCr = 1) then
               This_Bal.Expenses := This_Bal.Expenses + (ThisAmount * -1.00)
            else
               This_Bal.Expenses := This_Bal.Expenses + ThisAmount;
            This_Abs.Expenses := This_Abs.Expenses + ThisAmount;
         end;
      end;
      Query1.Next;
   end;

   Result := true;
end;

{
//---------------------------------------------------------------------------
// Function to retrieve the Trust records
//---------------------------------------------------------------------------
function  TFldExcel.GetTrust(FileName: string): boolean;
var
   idx, ThisClass : integer;
   S1             : string;

begin

//--- Get the information to calculate the Opening Balances for Trust

{
   S1 := 'SELECT B_Amount FROM billing WHERE B_Class = 4 AND B_Related = ''' + FileName + ''' AND B_Date < ''' + SDate + '''';
   if (GetTotals(S1) = -1) then begin
      Result := false;
      Exit;
   end else begin
      BalBusToTrust := 0;
      for idx := 0 to Query1.RecordCount - 1 do begin
         BalBusToTrust := BalBusToTrust + Query1.FieldByName('B_Amount').AsFloat;
         Query1.Next;
      end;
   end;
}

{
   S1 := 'SELECT B_Amount, B_Class, B_ReserveAmt FROM billing WHERE ((B_Class = 4) OR (B_Class >= 7 AND B_Class <= 14) OR (B_Class = 19) OR (B_Class = 24)) AND B_Related = ''' + FileName + ''' AND B_Date < ''' + SDate + '''';
   if (GetTotals(S1) = -1) then begin
      Result := false;
      Exit;
   end else begin
      BalBusToTrust   := 0;
      BalTrustDeposit := 0;
      BalTrustReserve := 0;
      BalXfrFees      := 0;
      BalXfrDisburse  := 0;
      BalXfrClient    := 0;
      BalXfrTrust     := 0;
      BalInv782BA     := 0;
      BalDrw782BA     := 0;
      BalInt782BA     := 0;
      BalTrustDebit   := 0;

      for idx := 0 to Query1.RecordCount - 1 do begin
         ThisClass := Query1.FieldByName('B_Class').AsInteger;
         case ThisClass of
             4: BalBusToTrust   := BalBusToTrust   + Query1.FieldByName('B_Amount').AsFloat;
             7: begin
                   BalTrustDeposit := BalTrustDeposit + Query1.FieldByName('B_Amount').AsFloat;
                   BalTrustReserve := BalTrustReserve + Query1.FieldByName('B_ReserveAmt').AsFloat;
                end;
             8: BalXfrFees      := BalXfrFees      + Query1.FieldByName('B_Amount').AsFloat;
             9: BalXfrDisburse  := BalXfrDisburse  + Query1.FieldByName('B_Amount').AsFloat;
            10: BalXfrClient    := BalXfrClient    + Query1.FieldByName('B_Amount').AsFloat;
            11: BalXfrTrust     := BalXfrTrust     + Query1.FieldByName('B_Amount').AsFloat;
            12: BalInv782BA     := BalInv782BA     + Query1.FieldByName('B_Amount').AsFloat;
            13: BalDrw782BA     := BalDrw782BA     + Query1.FieldByName('B_Amount').AsFloat;
            14: BalInt782BA     := BalInt782BA     + Query1.FieldByName('B_Amount').AsFloat;
            19: BalTrustDebit   := BalTrustDebit   + Query1.FieldByName('B_Amount').AsFloat;
         end;
         Query1.Next;
      end;
   end;

//--- Calcluate the Opening Balances

{
//=== TO DO:
//===    Add S78(2)(a) records as these are effected in the Trust Account
//===
}

{
   OpenBalTrust   := (BalTrustDeposit + BalDrw782BA + BalBusToTrust) - (BalXfrFees + BalXfrClient + BalXfrDisburse + BalXfrTrust + BalInv782BA + BalTrustDebit);
   OpenBal782BA   := (BalInv782BA + BalInt782BA) - BalDrw782BA;
   OpenBalReserve := BalTrustReserve;

//--- Get all the applicable Trust Records

   S1 := 'SELECT B_Date, B_Description, B_DrCr, B_Amount, B_Class, B_ReserveAmt FROM billing WHERE B_Related = ''' +
         FileName + ''' AND B_Date >= ''' + SDate + ''' AND B_Date <= ''' +
         EDate +  ''' AND ((B_Class >= 7 AND B_Class <= 14) OR (B_Class = 4) OR (B_Class = 19))' +
         ' ORDER BY B_Date, Create_Date, Create_Time';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (billing)';
      Result := false;
      Exit;
   end;

   Result := true;
end;
}

//---------------------------------------------------------------------------
// Function to retrieve the S86(3)Trust records
//---------------------------------------------------------------------------
function  TFldExcel.GetS863(FileName: string; StartDate: string; EndDate: string): boolean;
var
   idx, ThisClass : integer;
   S1             : string;

begin

//--- Get the information to calculate the Opening Balances for S78(2)(a)/S86(3)

   S1 := 'SELECT B_Amount, B_Class FROM billing WHERE ((B_Class >= 15 AND B_Class <= 17) OR B_Class = 20) AND B_Owner = ''' + FileName + ''' AND B_Date < ''' + StartDate + '''';
   if (GetTotals(S1) = -1) then begin
      Result := false;
      Exit;
   end else begin
      BalInv863     := 0;
      BalDrw863     := 0;
      BalInt863     := 0;
      BalIntDrw863  := 0;

      for idx := 0 to Query1.RecordCount - 1 do begin
         ThisClass := Query1.FieldByName('B_Class').AsInteger;
         case ThisClass of
            15: BalInv863    := BalInv863    + Query1.FieldByName('B_Amount').AsFloat;
            16: BalDrw863    := BalDrw863    + Query1.FieldByName('B_Amount').AsFloat;
            17: BalInt863    := BalInt863    + Query1.FieldByName('B_Amount').AsFloat;
            20: BalIntDrw863 := BalIntDrw863 + Query1.FieldByName('B_Amount').AsFloat;
         end;
         Query1.Next;
      end;
   end;

//--- Calcluate the Opening Balances

   OpenBal863Inv    := BalInv863;
   OpenBal863Int    := BalInt863;
   OpenBal863Drw    := BalDrw863;
   OpenBal863IntDrw := BalIntDrw863;

//--- Get all the applicable Trust Records

   S1 := 'SELECT B_Date, B_Description, B_DrCr, B_Amount, B_Class FROM billing WHERE B_Owner = ''' +
         FileName + ''' AND B_Date >= ''' + StartDate + ''' AND B_Date <= ''' +
         EndDate +  ''' AND ((B_Class >= 15 AND B_Class <= 17) OR B_Class = 20)' +
         ' ORDER BY B_Date, Create_Date, Create_Time';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (billing)';
      Result := false;
      Exit;
   end;

   Result := true;
end;

//---------------------------------------------------------------------------
// Function to retrieve an individual record from the DB
//---------------------------------------------------------------------------
function TFldExcel.GetRecord(S1: string; QryType: integer): boolean;

begin

   if (QryType = 1) then begin
      try
         Query1.Close;
         Query1.SQL.Text := S1;
         Query1.Open;
      except
         ErrMsg := '''Unable to read from ' + HostName + ''' (GetRecord)';
         Result := false;
         Exit;
      end;
   end else begin
      try
         Query2.Close;
         Query2.SQL.Text := S1;
         Query2.Open;
      except
         ErrMsg := '''Unable to read from ' + HostName + ''' (GetRecord)';
         Result := false;
         Exit;
      end;
   end;

   Result := true;

end;

//---------------------------------------------------------------------------
// Function to retrieve an individual record from the DB
//---------------------------------------------------------------------------
function TFldExcel.GetTotals(S1: String): double;

begin

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (tracking)';
      Result := -1;
      Exit;
   end;

   Result := 0;
end;

{
//---------------------------------------------------------------------------
// Function to retrieve the Total invoiced amount for a period for a File
//---------------------------------------------------------------------------
function TFldExcel.GetInvoiced(FileName: string): double;
var
   idx   : integer;
   S1    : string;
   Total : double;

begin

   S1 := 'SELECT Inv_Amount FROM invoices WHERE Inv_File = ''' + FileName + ''' AND Inv_EDate >= ''' + SDate + ''' AND Inv_EDate <= ''' + EDate + '''';

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (invoices)';
      Result := -1;
      Exit;
   end;

   Total := 0;
   Query2.First;

   for idx := 0 to Query2.RecordCount - 1 do begin
      Total := Total + Query2.FieldByName('Inv_Amount').AsFloat;
      Query2.Next;
   end;

   Result := Total;

end;
}

//---------------------------------------------------------------------------
// Function to retrieve specific information about an Invoice
//---------------------------------------------------------------------------
function TFldExcel.GetInvoiceData(Invoice: string): string;
var
   S1 : string;

begin

   S1 := 'SELECT * FROM invoices WHERE Inv_Invoice = ''' + Invoice + '''';

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (invoices)';
      Result := 'Not Found';
      Exit;
   end;

   if (Query2.RecordCount = 0) then begin
      Result := 'Not Found';
      Exit;
   end;

   AccountType := Query2.FieldByName('Inv_AcctType').AsInteger;
   ShowRelated := boolean(Query2.FieldByName('Inv_ShowRelated').AsInteger);
   SDate       := Query2.FieldByName('Inv_SDate').AsString;
   EDate       := Query2.FieldByName('Inv_EDate').AsString;

   Result := Query2.FieldByName('Inv_File').AsString;
end;

//---------------------------------------------------------------------------
// Function to retrieve the Total Paid amount for a period for a File
//---------------------------------------------------------------------------
{
function TFldExcel.GetPaid(FileName: string): double;
var
   idx   : integer;
   Total : double;
   S1    : string;

begin

   S1 := 'SELECT Pay_Amount FROM payments WHERE Pay_File = ''' + FileName + ''' AND Pay_Date >= ''' + SDate + ''' AND Pay_Date <= ''' + EDate + '''';

//--- Get the amount for all the payments in this period

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (payments)';
      Result := -1;
      Exit;
   end;

   Total := 0;
   Query2.First;

   for idx := 0 to Query2.RecordCount - 1 do begin
      Total := Total + Query2.FieldByName('Pay_Amount').AsFloat;
      Query2.Next;
   end;

   Result := Total;

end;
}

{
//---------------------------------------------------------------------------
// Function to retrieve a Trust balance for a spcecified period
//---------------------------------------------------------------------------
function  TFldExcel.GetTrust_Period(FileName: string): double;
var
   idx, ThisClass     : integer;
   BalTrust, BalS782A : double;
   S1                 : string;

begin

//--- Get the information to calculate the Opening Balances for Trust

   S1 := 'SELECT B_Amount, B_Class, B_ReserveAmt FROM billing WHERE ((B_Class = 4) OR (B_Class >= 7 AND B_Class <= 14) OR (B_Class = 19)) AND B_Related = ''' + FileName + ''' AND B_Date >= ''' + SDate + ''' AND B_DATE <= ''' + EDate + '''';

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (billing preparation report)';
      Result := 0;
      Exit;
   end;

   BalBusToTrust   := 0;
   BalTrustDeposit := 0;
   BalTrustReserve := 0;
   BalXfrFees      := 0;
   BalXfrDisburse  := 0;
   BalXfrClient    := 0;
   BalXfrTrust     := 0;
   BalInv782BA     := 0;
   BalDrw782BA     := 0;
   BalInt782BA     := 0;
   BalTrustDebit   := 0;

   for idx := 0 to Query2.RecordCount - 1 do begin
      ThisClass := Query2.FieldByName('B_Class').AsInteger;
      case ThisClass of
          4: BalBusToTrust   := BalBusToTrust   + Query2.FieldByName('B_Amount').AsFloat;
          7: begin
                BalTrustDeposit := BalTrustDeposit + Query2.FieldByName('B_Amount').AsFloat;
                BalTrustReserve := BalTrustReserve + Query2.FieldByName('B_ReserveAmt').AsFloat;
             end;
          8: BalXfrFees      := BalXfrFees      + Query2.FieldByName('B_Amount').AsFloat;
          9: BalXfrDisburse  := BalXfrDisburse  + Query2.FieldByName('B_Amount').AsFloat;
         10: BalXfrClient    := BalXfrClient    + Query2.FieldByName('B_Amount').AsFloat;
         11: BalXfrTrust     := BalXfrTrust     + Query2.FieldByName('B_Amount').AsFloat;
         12: BalInv782BA     := BalInv782BA     + Query2.FieldByName('B_Amount').AsFloat;
         13: BalDrw782BA     := BalDrw782BA     + Query2.FieldByName('B_Amount').AsFloat;
         14: BalInt782BA     := BalInt782BA     + Query2.FieldByName('B_Amount').AsFloat;
         19: BalTrustDebit   := BalTrustDebit   + Query2.FieldByName('B_Amount').AsFloat;
      end;
      Query2.Next;
   end;

//--- Calcluate and return the amount on Trust for the specified period

   BalTrust := (BalTrustDeposit + BalDrw782BA + BalBusToTrust) - (BalXfrFees + BalXfrClient + BalXfrDisburse + BalXfrTrust + BalInv782BA + BalTrustDebit);
   BalS782A := (BalInv782BA + BalInt782BA) - BalDrw782BA;

   Result := BalTrust + BalS782A;
end;
}

//---------------------------------------------------------------------------
// Function to retrieve the Notes records
//---------------------------------------------------------------------------
function TFldExcel.GetNotes(FileName: String): boolean;
var
   S1, S2 : string;

begin

   if (StrToInt(Parm07) in [1..2]) then
      S2 := ' Notes_Type = ' + Parm07 + ' AND '
   else
      S2 := ' ';

   S1 := 'SELECT Notes_Date, Notes_Time, Notes_Note, Notes_User FROM notes WHERE' +
         S2 + 'Notes_Tracking = ''' + FileName + ''' AND Notes_Date >= ''' +
         SDate + ''' AND Notes_Date <= ''' + EDate +
         ''' ORDER BY Notes_Date, Notes_Time';

//--- Get the notes for this File

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (Notes)';
      Result := false;
      Exit;
   end;

   Result := true;
end;

//---------------------------------------------------------------------------
// Function to retrieve the Fee Earner Records
//---------------------------------------------------------------------------
function TFldExcel.GetFeeRecs(FileName: String; ThisStr : string): boolean;
var
   S1 : string;

begin

   S1 := 'SELECT B_Date, B_Owner, B_Description, B_Amount, B_Class FROM billing WHERE B_Date >= ''' +
         SDate + ''' AND B_Date <= ''' + EDate +  '''' + ThisStr +
         ' ORDER BY B_Owner, B_Date, Create_Date, Create_Time';

//--- Get the notes for this File

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (Fee Earner)';
      Result := false;
      Exit;
   end;

   Result := true;
end;

//---------------------------------------------------------------------------
// Function to retrieve the Description for the File
//---------------------------------------------------------------------------
function TFldExcel.GetFileDesc(FileName: String): String;
var
   S1 : string;

begin

   S1 := 'SELECT Tracking_Description, Tracking_Related, Tracking_FileType FROM tracking WHERE Tracking_Name = ''' +
         FileName + '''';

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (Fee Earner- File Description)';
      Result := 'Error';
      Exit;
   end;

   SaveOwner    := FileName;
   SaveRelated  := Query2.FieldByName('Tracking_Related').AsString;
   SaveFileType := Query2.FieldByName('Tracking_FileType').AsInteger;

   Result := Query2.FieldByName('Tracking_Description').AsString;
end;

//---------------------------------------------------------------------------
// Function to retrieve the Alert Records
//---------------------------------------------------------------------------
function TFldExcel.GetAlertRecs(ThisUser: string; ThisFilter: string): boolean;
var
   S0, S1, S2, S3, S4           : string;
   StartDate, EndDate : string;

begin

   ShortDateFormat := 'yyyy/MM/dd';
   DateSeparator   := '/';

   StartDate := SDate;
   EndDate   := EDate;

//--- Load the data

   S1 := 'SELECT Tracking_Date AS Tracking_Order, Tracking_Name, ''Diary Alert'' AS Tracking_Type, Tracking_Owner, Tracking_Description, Tracking_Inactive AS Tracking_Active, '''' AS Tracking_Reason FROM tracking WHERE (Tracking_Date >= ''' +
         StartDate + ''' AND Tracking_Date <= ''' + EndDate +
         ''') AND Tracking_Inactive = 0 AND Tracking_Owner = ''' +
         ThisUser + ''' AND Tracking_Name LIKE ''' + ThisFilter + '''';

   S2 := 'SELECT Tracking_AlertDate AS Tracking_Order, Tracking_Name, ''Alert'' AS Tracking_Type, Tracking_Owner, Tracking_Description, Tracking_Alert AS Tracking_Active, Tracking_AlertReason AS Tracking_Reason FROM tracking WHERE Tracking_AlertDate <= ''' +
         EndDate + ''' AND Tracking_Alert = 1 AND Tracking_Owner = ''' +
         ThisUser + ''' AND Tracking_Name LIKE ''' + ThisFilter + '''';

   S3 := 'SELECT Tracking_PrescriptionDate AS Tracking_Order, Tracking_Name, ''Prescription'' AS Tracking_Type, Tracking_Owner, Tracking_Description, Tracking_Prescription AS Tracking_Active, ''Prescription in %d days!'' AS Tracking_Reason FROM tracking ' +
         'WHERE Tracking_PrescriptionDate <= ''' +
         FormatDateTime('yyyy/mm/dd',(Now() + 180)) +
         ''' AND Tracking_Prescription = 1 AND Tracking_Owner = ''' +
         ThisUser + ''' AND Tracking_Name LIKE ''' + ThisFilter + '''';

   S4 := ' ORDER BY Tracking_Order';

   case AccountType of
      0: S0 := S1 + ' UNION ' + S2 + ' UNION ' + S3 + S4;
      1: S0 := S1 + S4;
      2: S0 := S2 + S4;
      3: S0 := S3 + S4;
   end;

   try
      Query1.Close;
      Query1.SQL.Text := S0;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (Alerts Report)';
      Result := false;
      Exit;
   end;

   Result := true;
end;

//---------------------------------------------------------------------------
// Function to retrieve the Phonebook Records
//---------------------------------------------------------------------------
function TFldExcel.GetPhoneRecs(Filter1: string; Filter2: string): boolean;
var
   S1, ThisDate : string;

begin

   ShortDateFormat := 'yyyy/MM/dd';
   DateSeparator   := '/';

   ThisDate := FormatDateTime('yyyy/mm/dd',Now());

//--- Load the data

   S1 := 'SELECT Cust_Customer, Cust_Telephone, Cust_Cellphone, Cust_Fax, Cust_Worknum, Cust_CustType FROM customers WHERE Cust_Customer LIKE ''' +
         Filter1 + ''' AND (Cust_Telephone LIKE ''' + Filter2 +
         ''' OR Cust_Cellphone LIKE ''' + Filter2 +
         ''' OR Cust_Fax LIKE ''' + Filter2 + ''' OR Cust_Worknum LIKE ''' +
         Filter2 + ''')' +
         ' UNION ' +
         'SELECT Cor_Attorney AS Cust_Customer, Cor_Telephone AS Cust_Telephone, Cor_Cellphone AS Cust_Cellphone, Cor_Fax AS Cust_Fax, '''' AS Cust_Worknum, 3 AS Cust_CustType FROM corres WHERE Cor_Attorney LIKE ''' +
         Filter1 + ''' AND (Cor_Telephone LIKE ''' + Filter2 +
         ''' OR Cor_Cellphone LIKE ''' + Filter2 +
         ''' OR Cor_Fax LIKE ''' + Filter2 + ''')' +
         ' UNION ' +
         'SELECT Counsel_Counsel AS Cust_Customer, Counsel_Telephone AS Cust_Telephone, Counsel_Cellphone AS Cust_Cellphone, Counsel_Fax AS Cust_Fax, '''' AS Cust_Worknum, 4 AS Cust_CustType FROM counsel WHERE Counsel_Counsel LIKE ''' +
         Filter1 + ''' AND (Counsel_Telephone LIKE ''' + Filter2 +
         ''' OR Counsel_Cellphone LIKE ''' + Filter2 +
         ''' OR Counsel_Fax LIKE ''' + Filter2 + ''')' +
         ' UNION ' +
         'SELECT Opp_Attorney AS Cust_Customer, Opp_Telephone AS Cust_Telephone, Opp_Cellphone AS Cust_Cellphone, Opp_Fax AS Cust_Fax, '''' AS Cust_Worknum, 5 AS Cust_CustType FROM opposing WHERE Opp_Attorney LIKE ''' +
         Filter1 + ''' AND (Opp_Telephone LIKE ''' + Filter2 +
         ''' OR Opp_Cellphone LIKE ''' + Filter2 +
         ''' OR Opp_Fax LIKE ''' + Filter2 + ''') ORDER BY Cust_Customer';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (Phonebook Export)';
      Result := false;
      Exit;
   end;

   Result := true;
end;

//---------------------------------------------------------------------------
// Function to retrieve the Client/Opposition Records for Export
//---------------------------------------------------------------------------
function TFldExcel.GetClientDetails(Filter: string; ThisType: integer): boolean;
var
   S1 : string;

begin

//--- Load the data

   S1 := 'SELECT * FROM customers WHERE Cust_Customer LIKE ''' + Filter +
         ''' AND Cust_CustType = ' + IntToStr(ThisType) + ' ORDER BY Cust_Customer';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (Client/Opposition Export)';
      Result := false;
      Exit;
   end;

   Result := true;
end;

//---------------------------------------------------------------------------
// Function to retrieve the File Records for Export
//---------------------------------------------------------------------------
function TFldExcel.GetFileDetails(TimeStamp: string; ThisType: integer): string;
var
   S1 : string;

begin

   if (TimeStamp = '') then begin
      Result := '';
      Exit;
   end;

//--- 1 = Opposing Attorney, 2 = Counsel, 3 = Correspondent, 4 = Sheriff

   case ThisType of
      1: S1 := 'SELECT Opp_Attorney FROM opposing WHERE Opp_TimeStamp = ''' + TimeStamp + '''';
      2: S1 := 'SELECT Counsel_Counsel FROM counsel WHERE Counsel_TimeStamp = ''' + TimeStamp + '''';
      3: S1 := 'SELECT Cor_Attorney FROM corres WHERE Cor_TimeStamp = ''' + TimeStamp + '''';
      4: S1 := 'SELECT Sheriff_Sheriff FROM sheriff WHERE Sheriff_TimeStamp = ''' + TimeStamp + '''';
   end;

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (File Export)';
      Result := '';
      Exit;
   end;

   case ThisType of
      1: Result := ReplaceQuote(Query1.FieldByName('Opp_Attorney').AsString);
      2: Result := ReplaceQuote(Query1.FieldByName('Counsel_Counsel').AsString);
      3: Result := ReplaceQuote(Query1.FieldByName('Cor_Attorney').AsString);
      4: Result := ReplaceQuote(Query1.FieldByName('Sheriff_Sheriff').AsString);
   end;
end;

//---------------------------------------------------------------------------
// Function to retrieve the known client email addresses for a File
//---------------------------------------------------------------------------
function TFldExcel.GetAddresses(FileName: string): string;
var
   S1, S2, WorkEmail, CustEmail, CustName : string;

begin

   S1 := 'SELECT Tracking_Client, Tracking_Opposition, Tracking_ClientKey FROM tracking WHERE Tracking_Name = ''' + FileName + ''' ORDER BY Tracking_Name';

//--- Get availble client information

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (Send Email)';
      Result := ThisAddrList;
      Exit;
   end;

   ThisSubject   := 'IN RE: ' + Query2.FieldByName('Tracking_Client').AsString;

   if (Query2.FieldByName('Tracking_Opposition').AsString = 'State') then
      ThisSubject := 'IN RE: State vs ' + Query2.FieldByName('Tracking_Client').AsString
   else if (Query2.FieldByName('Tracking_Opposition').AsString <> '--- None ---') then
      ThisSubject := ThisSubject + ' vs ' + Query2.FieldByName('Tracking_Opposition').AsString;

   ThisAddrList := ThisAddrList + ';R' + FileName;

//--- Get the client's email addresses

   S2 := 'SELECT Cust_Description, Cust_Persemail, Cust_Workemail FROM customers WHERE Cust_CustType = 1 AND Cust_TimeStamp = ''' + Query2.FieldByName('Tracking_ClientKey').AsString + '''';

   try
      Query2.Close;
      Query2.SQL.Text := S2;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (Send Email)';
      Result := ThisAddrList;
      Exit;
   end;

   CustEmail := Query2.FieldByName('Cust_Persemail').AsString;
   WorkEmail := Query2.FieldByName('Cust_Workemail').AsString;
   CustName  := Query2.FieldByName('Cust_Description').AsString;

   if (CustEmail <> '') then
      ThisAddrList := ThisAddrList + ';P' + CustEmail;

   if (WorkEmail <> '') then
      ThisAddrList := ThisAddrList + ';W' + WorkEmail;

   if (CustName <> '') then
      ThisAddrList := ThisAddrList + ';N' + CustName;

   Result := ThisAddrList;
end;

//---------------------------------------------------------------------------
// Procedure to send the workbook via Email
//---------------------------------------------------------------------------
function TFldExcel.SendEmail(FileRef: string; FileName: string; ProcessType: integer): boolean;
var
   idx, Len                                : integer;
   ThisList, ThisStr, ToStr, CcStr, BCcStr : string;
   ToDelim, CcDelim, BccDelim, AddrType    : string;
   ThisAttach, AttachFile, ThisDate        : string;
   ThisTime                                : string;
   ThisBody                                : TStringList;
   IntAttach                               : TIdAttachmentFile;

begin

//--- If AttachList is empty then no email will be sent

   if (GroupAttach = true) and (AttachList.Count < 1) then begin
      Result := false;
      Exit;
   end;

//--- Set the default Subject

   ThisSubject := CpyName + ': ' + ThisLabels[RunType + 1];

   ThisBody  := TStringList.Create;

//--- Get additional information if a valid File Reference was passed

   if (FileRef <> '') then
      ThisAddrList := GetAddresses(FileRef);

   if (ProcessType = ord(PT_BILLING)) then begin

      if (IntEmail = False) then begin

         emlSend.SendTo    := Billing_To;
         emlSend.Cc        := Billing_CC;
         emlSend.Bcc       := Billing_BCC;
         emlSend.Subject   := Billing_Subject;
         emlSend.Body.Text := Billing_EMail;

//--- Set the email flags

         if ((EditEmail = true) and (ReadReceipt = true)) then
            emlSend.Flags := [sfDialog,sfReceiptRequested]
         else if (ReadReceipt = true) then
            emlSend.Flags := [sfReceiptRequested]
         else if (EditEmail = true) then
            emlSend.Flags := [sfDialog];
      end else begin
         idMessage.ClearBody;
         idMessage.ContentType               := 'multipart/mixed';
         idMessage.From.Address              := UserEmail;
         idMessage.Recipients.EMailAddresses := Billing_To;
         idMessage.CCList.EMailAddresses     := Billing_CC;
         idMessage.BccList.EMailAddresses    := Billing_BCC;
         idMessage.Subject                   := Billing_Subject;
         idMessage.Body.Text                 := Billing_EMail;

         if (ReadReceipt = true) then
            idMessage.ReceiptRecipient.Address := UserEmail;
      end;

   end else begin

//--- Call ldGetEmailP to get the list of email addresses to use. We begin by
//--- initialising the Symvars Lists so that these are ready in case the user
//--- click on the button to show the available Symvars

      SetUpSymVars;

//--- We also set all the Symvars for which we have known values

      ThisDate := FormatDateTime('yyyy/MM/dd',Now());
      ThisTime := FormatDateTime('HH:NN:SS',Now());

      SymVars_LPMS.SV[ord(SV_COMBINED)].Value   := ThisLabels[RunType + 1];
      SymVars_LPMS.SV[ord(SV_CURRDATE)].Value   := ThisDate;
      SymVars_LPMS.SV[ord(SV_DATE)].Value       := ThisDate;
      SymVars_LPMS.SV[ord(SV_ENDDATE)].Value    := EDate;
      SymVars_LPMS.SV[ord(SV_FILE)].Value       := FileRef;
      SymVars_LPMS.SV[ord(SV_FROMDATE)].Value   := SDate;
      SymVars_LPMS.SV[ord(SV_MAILNOTICE)].Value := EmailStrings.Text;
      SymVars_LPMS.SV[ord(SV_SHORTTIME)].Value  := Copy(ThisTime,1,5);
      SymVars_LPMS.SV[ord(SV_SHORTYEAR)].Value  := Copy(ThisDate,1,4);
      SymVars_LPMS.SV[ord(SV_TIME)].Value       := ThisTime;
      SymVars_LPMS.SV[ord(SV_USER)].Value       := UserName;
      SymVars_LPMS.SV[ord(SV_VATNUM)].Value     := VATNumber;

//--- Now display FldGetEmailP

      FldGetEmailP.UserEmail          := ThisAddrList;
      FldGetEmailP.edtSubject.Text    := ThisSubject;
      FldGetEmailP.rteEmailBody1.Text := EmailBody;

      if (IntEmail = True) then
         FldGetEmailP.chkEdit.Enabled := False;

      FldGetEmailP.ShowModal;

//--- Abort the email sending if ldGetEmailP returned false

      if (GetEmailResult = false) then begin
         AttachList.Clear;
         Result := false;
         Exit;
      end;

//--- Get the list of information needed to send the email

      ThisList  := FldGetEmailP.EmailStr;

//--- Extract the email addresses and labels from the string constructed by
//--- ldGetEmailP

      EmailLst.Clear;
      ExtractStrings(['|'],[' '],PChar(ThisList),EmailLst);

//--- Determine whether the Compose dialog must be displayed and or whether
//--- a ReadReceipt must be requested

      if ((EmailLst[EmailLst.Count - 2] = '1') and (EmailLst[EmailLst.Count - 1] = '1')) then
         emlSend.Flags := [sfDialog,sfReceiptRequested]
      else if (EmailLst[EmailLst.Count - 1] = '1') then
         emlSend.Flags := [sfReceiptRequested]
      else if (EmailLst[EmailLst.Count - 2] = '1') then
         emlSend.Flags := [sfDialog];

//--- Initialise the variables that must have a value

      ToStr      := '';
      CcStr      := '';
      BccStr     := '';
      ToDelim    := '';
      CcDelim    := '';
      BccDelim   := '';
      ThisAttach := '';

//--- Get the list of recipients

      for idx := 0 to EmailLst.Count -5 do begin
         if (EmailLst[idx] <> '') then begin
            ThisStr := EmailLst[idx];
            Len := Length(ThisStr);
            AddrType := Copy(ThisStr,1,1);

            if (AddrType = 'T') then begin
               ToStr := ToStr + ToDelim + Copy(ThisStr,2,Len -1);
               ToDelim := ';';
            end;

            if (AddrType = 'C') then begin
               CcStr := CcStr + CcDelim + Copy(ThisStr,2,Len - 1);
               CcDelim := ';';
            end;

            if (AddrType = 'B') then begin
               BccStr := BccStr + BccDelim + Copy(ThisStr,2,Len - 1);
               BccDelim := ';';
            end;
         end;
      end;

      if (IntEmail = False) then begin
         emlSend.SendTo  := ToStr;
         emlSend.Cc      := CcStr;
         emlSend.Bcc     := BccStr;
         emlSend.Subject := EmailLst[EmailLst.Count - 5];
      end else begin
         idMessage.ClearBody;
         idMessage.ContentType               := 'multipart/mixed';
         idMessage.From.Address              := UserEmail;
         idMessage.Recipients.EMailAddresses := ToStr;
         idMessage.CCList.EMailAddresses     := CcStr;
         idMessage.BccList.EMailAddresses    := BccStr;
         idMessage.Subject                   := EmailLst[EmailLst.Count -5];

         if (EmailLst[EmailLst.Count - 1] = '1') then
            idMessage.ReceiptRecipient.Address := UserEmail;
      end;

//--- Set some further Symvar Values

      SymVars_LPMS.SV[ord(SV_CLIENT)].Value := EmailLst[EmailLst.Count - 4];

//--- Create the body of the email

      for idx := 0 to FldGetEmailP.rteEmailBody1.Lines.Count - 1 do begin
         ThisBody.Add(DoSymVars(FldGetEmailP.rteEmailBody1.Lines.Strings[idx],EmailLst[EmailLst.Count - 3]));
      end;

      if (IntEmail = False) then
         emlSend.Body := ThisBody
      else
         idMessage.Body := ThisBody;

   end;

//--- Check whether the original file or the PDF file must be attached. If
//--- GroupAttach is true then all of the files in AttachList are attached
//--- in a single email. If PDFPrefer is true then an attempt is made to
//--- attach the PDF version of each file in AttachList

   if (PDFPrefer = True) then begin
      if (GroupAttach = False) then begin
         if (PDFExists = True) then begin
            if (IntEmail = False) then
               emlSend.Attach := ChangeFileExt(FileName,'.pdf')
            else begin
               AttachFile := ChangeFileExt(FileName,'.pdf');
               IntAttach := TIdAttachmentFile.Create(idMessage.MessageParts,AttachFile);
               IntAttach.FileName := AttachFile;
               IntAttach.DisplayName := ExtractFileName(AttachFile);
               IntAttach.ContentType := 'application/pdf';
            end;
         end else begin
            if (IntEmail = False) then
               emlSend.Attach := FileName
            else begin
               AttachFile := FileName;
               IntAttach := TIdAttachmentFile.Create(idMessage.MessageParts,AttachFile);
               IntAttach.FileName := AttachFile;
               IntAttach.DisplayName := ExtractFileName(AttachFile);
               IntAttach.ContentType := 'application/vnd.ms-excel';
            end;
         end;
      end else begin
         for idx := 0 to AttachList.Count - 1 do begin
            if (FileExists(ChangeFileExt(AttachList.Strings[idx],'.pdf'))) then begin
               if (IntEmail = False) then
                  ThisAttach := ThisAttach + ChangeFileExt(AttachList.Strings[idx],'.pdf') + ';'
               else begin
                  AttachFile := ChangeFileExt(AttachList.Strings[idx],'.pdf');
                  IntAttach := TIdAttachmentFile.Create(idMessage.MessageParts,AttachFile);
                  IntAttach.FileName := AttachFile;
                  IntAttach.DisplayName := ExtractFileName(AttachFile);
                  IntAttach.ContentType := 'application/pdf';
               end;
            end else begin
               if (IntEmail = False) then
                  ThisAttach := ThisAttach + AttachList.Strings[idx] + ';'
               else begin
                  AttachFile := AttachList.Strings[idx];
                  IntAttach := TIdAttachmentFile.Create(idMessage.MessageParts,AttachFile);
                  IntAttach.FileName := AttachFile;
                  IntAttach.DisplayName := ExtractFileName(AttachFile);
                  IntAttach.ContentType := 'application/vnd.ms-excel';
               end;
            end;
         end;

         if (IntEmail = False) then
            emlSend.Attach := ThisAttach
      end;
   end else begin
      if (GroupAttach = false) then begin
         if (IntEmail = False) then
            emlSend.Attach := FileName
         else begin
            AttachFile := FileName;
            IntAttach := TIdAttachmentFile.Create(idMessage.MessageParts,AttachFile);
            IntAttach.FileName := AttachFile;
            IntAttach.DisplayName := ExtractFileName(AttachFile);
            IntAttach.ContentType := 'application/vnd.ms-excel';
         end;
      end else begin
         for idx := 0 to AttachList.Count - 1 do begin
            if (IntEmail = False) then
               ThisAttach := ThisAttach + AttachList.Strings[idx] + ';'
            else begin
               AttachFile := AttachList.Strings[idx];
               IntAttach := TIdAttachmentFile.Create(idMessage.MessageParts,AttachFile);
               IntAttach.FileName := AttachFile;
               IntAttach.DisplayName := ExtractFileName(AttachFile);
               IntAttach.ContentType := 'application/vnd.ms-excel';
            end;
         end;

         if (IntEmail = False) then
            emlSend.Attach := ThisAttach
      end;
   end;

//--- Send the Email

   if (IntEmail = True) then begin

      try
         idSMTP.Host     := SMTPServer;
         idSMTP.Port     := StrToInt(SMTPPort);

         if (SMTPAuth = True) then begin
            case SMTPAuthType of
               0: idSMTP.AuthType := satDefault;
               1: idSMTP.AuthType := satNone;
               2: idSMTP.AuthType := satSASL;
            end;
         end else
            idSMTP.AuthType := SatNone;

         idSMTP.Username := SMTPUser;
         idSMTP.Password := SMTPPass;

         idSMTP.Connect;

         try
            idSMTP.Send(idMessage);
         finally
            idSMTP.Disconnect();
         end;
      finally
//--- Reset Attachlist as SendMail may be called repeatedly

         AttachList.Clear;
      end;
   end else
      emlSend.Send;

   ThisBody.Destroy;
   Result := true;
end;

//---------------------------------------------------------------------------
// Function to retrieve/set the next sequential Invoice Number
//---------------------------------------------------------------------------
function TFldExcel.GetNextInvoice(QryType: integer): boolean;
var
   S1 : string;

begin

//--- Get or Set the next sequential Invoice Number

   if (QryType = 1) then begin
      S1 := 'SELECT InvoiceIdx, InvoicePref FROM lpms';

      try
         Query1.Close;
         Query1.SQL.Text := S1;
         Query1.Open;
      except
         ErrMsg := '''Unable to read from ' + HostName + ''' (Invoice Idx)';
         InvoicePref := '0';
         InvoiceStr  := '000000';
         Result := false;
         Exit;
      end;

      InvoicePref := Query1.FieldByName('InvoicePref').AsString;
      InvoiceNum  := Query1.FieldByName('InvoiceIdx').AsInteger;

      inc(InvoiceNum);
      InvoiceStr := Format('%.6d',[InvoiceNum]);

   end else begin
      S1 := 'UPDATE lpms SET InvoiceIdx = ' + IntToStr(InvoiceNum);

      try
         Query1.Close;
         Query1.SQL.Text := S1;
         Query1.ExecSQL;
      except
         ErrMsg := '''Unable to update to ' + HostName + ''' (Invoice Idx)';
         Result := false;
         Exit;
      end;

   end;

   Result := true;
end;

//---------------------------------------------------------------------------
// Function to get the Invoices to look for payments
//---------------------------------------------------------------------------
function TFldExcel.GetInvoices(ThisType: integer; ThisFile: string): boolean;
var
   S1, EndDate : string;

begin

   case ThisType of
      ord(DT_ACCOUNTANT):
         S1 := 'SELECT Inv_Invoice, Inv_Amount, Inv_File, Inv_Description, ' +
               'Inv_Fees, Inv_Disburse, Inv_Expenses, Inv_EDate FROM invoices ' +
               'ORDER BY Inv_Invoice';

      ord(DT_PAYMENT):
         S1 := 'SELECT Inv_Invoice, Inv_Amount, Inv_File, Inv_Description, ' +
               'Inv_Fees, Inv_Disburse, Inv_Expenses, Inv_SDate, Inv_EDate ' +
               'FROM invoices WHERE Inv_Invoice = ''' + ThisFile + '''';

      ord(DT_INVOICES):
         S1 := 'SELECT Inv_Invoice FROM invoices';

      ord(DT_STATEMENT): begin
         EndDate := Copy(EDate,1,8) + '31';

         S1 := 'SELECT Inv_Amount, Inv_Invoice, Inv_EDate FROM invoices ' +
               'WHERE Inv_File = ''' + ThisFile + ''' AND Inv_EDate <= ''' +
               EndDate + ''' ORDER BY Inv_EDate ASC';
      end;
   end;
{
   if (ThisType = ord(IT_ACCOUNTANT)) then
      S1 := 'SELECT Inv_Invoice, Inv_Amount, Inv_File, Inv_Description, ' +
            'Inv_Fees, Inv_Disburse, Inv_Expenses, Inv_EDate FROM invoices ' +
            'ORDER BY Inv_Invoice'
   else if (ThisType = ord(IT_PAYMENT)) then
      S1 := 'SELECT Inv_Invoice, Inv_Amount, Inv_File, Inv_Description, ' +
            'Inv_Fees, Inv_Disburse, Inv_Expenses, Inv_SDate, Inv_EDate ' +
            'FROM invoices WHERE Inv_Invoice = ''' + ThisFile + ''''
   else if (ThisType = ord(IT_INVOICES)) then
      S1 := 'SELECT Inv_Invoice FROM invoices'
   else begin
      EndDate := Copy(EDate,1,8) + '31';
      S1 := 'SELECT Inv_Amount, Inv_Invoice, Inv_EDate FROM invoices WHERE Inv_File = ''' +
            ThisFile + ''' AND Inv_EDate <= ''' + EndDate +
            ''' ORDER BY Inv_EDate ASC';
   end;
}

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read invoices from ' + HostName + '''';
      Result := false;
      Exit;
   end;

   Result := true;
end;

//---------------------------------------------------------------------------
// Function to get the payments between SDate and EDate for a specific invoice
//---------------------------------------------------------------------------
function TFldExcel.GetPayments(ThisInvoice: string): double;
var
   idx    : integer;
   Amount : double;
   S1, S2 : string;

begin

   S1 := 'SELECT Pay_Amount FROM payments WHERE Pay_Date < ''' + SDate +
         ''' AND Pay_Invoice = ''' + ThisInvoice + ''' ORDER BY Pay_Date ASC';

   S2 := 'SELECT Pay_Date, Pay_Amount, Pay_Note FROM payments WHERE Pay_Date >= ''' +
         SDate + ''' AND Pay_Date <= ''' + EDate + ''' AND Pay_Invoice = ''' +
         ThisInvoice + ''' ORDER BY Pay_Date ASC';

//--- Get the amount for all the payments before this period

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (payments)';
      Result := -1;
      Exit;
   end;

   PrevPaid := 0;
   Query2.First;

   for idx := 0 to Query2.RecordCount - 1 do begin
      PrevPaid := PrevPaid + Query2.FieldByName('Pay_Amount').AsFloat;
      Query2.Next;
   end;

//--- Get all the payments for the period in which the Invoice falls

   try
      Query2.Close;
      Query2.SQL.Text := S2;
      Query2.Open;
   except
      ErrMsg := '''Unable to read payments for invoice ' + ThisInvoice + ' from ' + HostName + '''';
      Result := Amount;
      Exit;
   end;

   Amount := 0;
   Query2.First;

   for idx := 0 to Query2.RecordCount - 1 do begin
      Amount := Amount + StrToFloat(Query2.FieldByName('Pay_Amount').AsString);
      Query2.Next;
   end;

   Result := Amount;
end;

//---------------------------------------------------------------------------
// Function to get the Client Name for a File
//---------------------------------------------------------------------------
function TFldExcel.GetClient(FileName: string): string;
var
   S1 : string;

begin

   S1 := 'SELECT Tracking_Client FROM tracking WHERE Tracking_Name = ''' +
         FileName + '''';

   try
      Query2.Close;
      Query2.SQL.Text := S1;
      Query2.Open;
   except
      ErrMsg := '''Unable to read Client Name for File ' + FileName + ' from ' + HostName + '''';
      Result := '';
      Exit;
   end;

   Result := Query2.FieldByName('Tracking_Client').AsString;
end;

//---------------------------------------------------------------------------
// Function to store an Invoice
//---------------------------------------------------------------------------
function TFldExcel.StoreInvoice(ThisFile: string; Amount: double; Fees: double; Disburse: double; Expenses: double): boolean;
var
   S1, ThisDescrip, TimeStamp : string;

begin

   ShortDateFormat := 'yyyy/MM/dd';
   DateSeparator   := '/';

   TimeStamp := FormatDateTime('yyyy/mm/dd',Now()) + '+' + FormatDateTime('hh:nn:ss:zzz',Now()) + '+System';

//--- First we check if the Invoice already exists

   S1 := 'SELECT Inv_Invoice FROM invoices WHERE Inv_Invoice = ''' + InvoicePref + InvoiceStr + '''';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (Invoice Idx)';
      Result := false;
      Exit;
   end;

   if (Query1.RecordCount > 0) then begin
      ErrMsg := '''Invoice "' + InvoicePref + InvoiceStr + '" already exist - unable to store invoice for File "' + ThisFile + '" ...''';
      Result := false;
      Exit;
   end;

//--- If we get here then the Invoice does not exist - Get the Description

   S1 := 'SELECT Tracking_Description FROM tracking WHERE Tracking_Name = ''' + ThisFile + '''';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      ErrMsg := '''Unable to read from ' + HostName + ''' (Invoice Description)';
      Result := false;
      Exit;
   end;

   ThisDescrip := Query1.FieldByName('Tracking_Description').AsString;

//--- Store the Invoice

   S1 := 'INSERT INTO invoices (Inv_Invoice, Inv_File, Inv_Description, Inv_Hostname, Inv_SDate, Inv_EDate, Inv_AcctType, Inv_ShowRelated, Create_By, Create_Date, Create_Time, Inv_Amount, Inv_Fees, Inv_Disburse, Inv_Expenses, Inv_TimeStamp) Values(''' +
         InvoicePref + InvoiceStr + ''', ''' + ThisFile + ''', ''' +
         ThisDescrip + ''', ''' + HostName + ''', ''' + SDate +
         ''', ''' + Edate + ''', ' + IntToStr(AccountType) + ', ' +
         BoolToStr(ShowRelated) + ', ''System'', ''' +
         FormatDateTime('yyy/MM/dd',Now()) + ''', ''' +
         FormatDateTime('hh:nn:ss', Now()) + ''', ''' +
         Format('%.2f',[Amount]) + ''', ''' +  Format('%.2f',[Fees]) +
         ''', ''' + Format('%.2f',[Disburse]) + ''', ''' +
         Format('%.2f',[Expenses]) + ''', ''' + TimeStamp + ''')';
   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.ExecSQL;
   except
      ErrMsg := '''Unable to store invoice to ' + HostName + ''' (Invoice Idx)';
      Result := false;
      Exit;
   end;

//--- Now update the Invoice index

   Result := GetNextInvoice(2);
end;

{//---------------------------------------------------------------------------
// Temporary function to fix an Invoice that was generated before the
// introduction of Fees, Disbursments and Expenses amounts in the Invoice
// record
//---------------------------------------------------------------------------
function TFldExcel.FixInvoice(ThisInvoice: string; Amount: double; Fees: double; Disburse: double; Expenses: double): boolean;
var
   S1          : string;

begin

//--- Fix the Invoice

   S1 := 'UPDATE invoices SET Inv_Amount = ''' + Format('%.2f',[Amount]) +
         ''', Inv_Fees = ''' + Format('%.2f',[Fees]) +
         ''', Inv_Disburse = ''' + Format('%.2f',[Disburse]) +
         ''', Inv_Expenses = ''' + Format('%.2f',[Expenses]) +
         ''' WHERE Inv_Invoice = ''' +  ThisInvoice + '''';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.ExecSQL;
   except
      ErrMsg := '''Unable to fix invoice: ''' + ThisInvoice + '''';
      Result := false;
      Exit;
   end;

   Result := true;
end;
}

//---------------------------------------------------------------------------
// Function to get the Actual and Absolute totals for a billing record set
//---------------------------------------------------------------------------
procedure TFldExcel.GetAmounts(ThisType: integer; adoQry: TADOQuery{; This_Bal: LPMS_Amounts; This_Abs: LPMS_Amounts});
var
   idx1, DrCr, IsReserved  : integer;
   Pref                    : string;
   ThisAmount, ReservedAmt : double;

begin

// Ensure that the correct type of Billing records are read

   if ((ThisType = 0) or (ThisType = 2)) then
      Pref := 'B'
   else
      Pref := 'Collect';

// Clear the structures that will hold the amounts for each kind of Billig record

   Clear_Amounts;

// Collect the Billing amounts for each type of Billing record

   adoQry.First;
   for idx1 := 0 to adoQry.RecordCount - 1 do begin
      DrCr        := adoQry.FieldByName(Pref + '_DrCr').AsInteger;
      ThisAmount  := adoQry.FieldByName(Pref + '_Amount').AsFloat;
      IsReserved  := adoQry.FieldByName(Pref + '_ReserveDep').AsInteger;
      ReservedAmt := adoQry.FieldByName(Pref + '_ReserveAmt').AsFloat;

      case adoQry.FieldByName(Pref + '_Class').AsInteger of
         0: begin                  // Fees
            if (DrCr = 1) then
               This_Bal.Fees := This_Bal.Fees + (ThisAmount * -1.00)
            else
               This_Bal.Fees := This_Bal.Fees + ThisAmount;
            This_Abs.Fees := This_Abs.Fees + ThisAmount;
         end;

         1: begin                  // Disbursements
            if (DrCr = 1) then
               This_Bal.Disbursements := This_Bal.Disbursements + (ThisAmount * -1.00)
            else
               This_Bal.Disbursements := This_Bal.Disbursements + ThisAmount;
            This_Abs.Disbursements := This_Abs.Disbursements + ThisAmount;
         end;

         2: begin                  // Expenses
            if (DrCr = 1) then
               This_Bal.Expenses := This_Bal.Expenses + (ThisAmount * -1.00)
            else
               This_Bal.Expenses := This_Bal.Expenses + ThisAmount;
            This_Abs.Expenses := This_Abs.Expenses + ThisAmount;
         end;

         3: begin                  // Payment Received
            if (DrCr = 1) then
               This_Bal.Payment_Received := This_Bal.Payment_Received + (ThisAmount * -1.00)
            else
               This_Bal.Payment_Received := This_Bal.Payment_Received + ThisAmount;
            This_Abs.Payment_Received := This_Abs.Payment_Received + ThisAmount;
         end;

         4: begin                  // Business to Trust
            if (DrCr = 1) then
               This_Bal.Business_To_Trust := This_Bal.Business_To_Trust + (ThisAmount * -1.00)
            else
               This_Bal.Business_To_Trust := This_Bal.Business_To_Trust + ThisAmount;
            This_Abs.Business_To_Trust := This_Abs.Business_To_Trust + ThisAmount;
         end;

         5: begin                  // Credit
            if (DrCr = 1) then
               This_Bal.Credit := This_Bal.Credit + (ThisAmount * -1.00)
            else
               This_Bal.Credit := This_Bal.Credit + ThisAmount;
            This_Abs.Credit := This_Abs.Credit + ThisAmount;
         end;

         6: begin                  // Business Deposit
            if (DrCr = 1) then
               This_Bal.Business_Deposit := This_Bal.Business_Deposit + (ThisAmount * -1.00)
            else
               This_Bal.Business_Deposit := This_Bal.Business_Deposit + ThisAmount;
            This_Abs.Business_Deposit := This_Abs.Business_Deposit + ThisAmount;
         end;

         7: begin                  // Trust Deposit
            if (DrCr = 1) then
               This_Bal.Trust_Deposit := This_Bal.Trust_Deposit + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Deposit := This_Bal.Trust_Deposit + ThisAmount;
            This_Abs.Trust_Deposit := This_Abs.Trust_Deposit + ThisAmount;

            if (IsReserved = 1) then begin
               if (DrCr = 1) then
                  This_Bal.Reserved_Trust := This_Bal.Reserved_Trust + (ReservedAmt * -1.00)
               else
                  This_Bal.Reserved_Trust := This_Bal.Reserved_Trust + ReservedAmt;
               This_Abs.Reserved_Trust := This_Abs.Reserved_Trust + ReservedAmt;
            end;
         end;

         8: begin                  // Trust Transfer (Business)
            if (DrCr = 1) then
               This_Bal.Trust_Transfer_Business_Fees := This_Bal.Trust_Transfer_Business_Fees + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Transfer_Business_Fees := This_Bal.Trust_Transfer_Business_Fees + ThisAmount;
            This_Abs.Trust_Transfer_Business_Fees := This_Abs.Trust_Transfer_Business_Fees + ThisAmount;
         end;

         9: begin                  // Trust Transfer (Disbursements)
            if (DrCr = 1) then
               This_Bal.Trust_Transfer_Disbursements := This_Bal.Trust_Transfer_Disbursements + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Transfer_Disbursements := This_Bal.Trust_Transfer_Disbursements + ThisAmount;
            This_Abs.Trust_Transfer_Disbursements := This_Abs.Trust_Transfer_Disbursements + ThisAmount;
         end;

         10: begin                  // Trust Transfer (Client)
            if (DrCr = 1) then
               This_Bal.Trust_Transfer_Client := This_Bal.Trust_Transfer_Client + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Transfer_Client := This_Bal.Trust_Transfer_Client + ThisAmount;
            This_Abs.Trust_Transfer_Client := This_Abs.Trust_Transfer_Client + ThisAmount;
         end;

         11: begin                  // Trust Transfer (Trust)
            if (DrCr = 1) then
               This_Bal.Trust_Transfer_Trust := This_Bal.Trust_Transfer_Trust + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Transfer_Trust := This_Bal.Trust_Transfer_Trust + ThisAmount;
            This_Abs.Trust_Transfer_Trust := This_Abs.Trust_Transfer_Trust + ThisAmount;
         end;

         12: begin                  // Trust Investment S86(4)
            if (DrCr = 1) then
               This_Bal.Trust_Investment_S86_4 := This_Bal.Trust_Investment_S86_4 + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Investment_S86_4 := This_Bal.Trust_Investment_S86_4 + ThisAmount;
            This_Abs.Trust_Investment_S86_4 := This_Abs.Trust_Investment_S86_4 + ThisAmount;
         end;

         13: begin                  // Trust Withdrawal S86(4)
            if (DrCr = 1) then
               This_Bal.Trust_Withdrawal_S86_4 := This_Bal.Trust_Withdrawal_S86_4 + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Withdrawal_S86_4 := This_Bal.Trust_Withdrawal_S86_4 + ThisAmount;
            This_Abs.Trust_Withdrawal_S86_4 := This_Abs.Trust_Withdrawal_S86_4 + ThisAmount;
         end;

         14: begin                  // Trust Interest S86(4)
            if (DrCr = 1) then
               This_Bal.Trust_Interest_S86_4 := This_Bal.Trust_Interest_S86_4 + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Interest_S86_4 := This_Bal.Trust_Interest_S86_4 + ThisAmount;
            This_Abs.Trust_Interest_S86_4 := This_Abs.Trust_Interest_S86_4 + ThisAmount;
         end;

         15: begin                  // Trust Investment S86(3)
            if (DrCr = 1) then
               This_Bal.Trust_Investment_S86_3 := This_Bal.Trust_Investment_S86_3 + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Investment_S86_3 := This_Bal.Trust_Investment_S86_3 + ThisAmount;
            This_Abs.Trust_Investment_S86_3 := This_Abs.Trust_Investment_S86_3 + ThisAmount;
         end;

         16: begin                  // Trust Withdrawal S86(3)
            if (DrCr = 1) then
               This_Bal.Trust_Withdrawal_S86_3 := This_Bal.Trust_Withdrawal_S86_3 + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Withdrawal_S86_3 := This_Bal.Trust_Withdrawal_S86_3 + ThisAmount;
            This_Abs.Trust_Withdrawal_S86_3 := This_Abs.Trust_Withdrawal_S86_3 + ThisAmount;
         end;

         17: begin                  // Trust Interest S86(3)
            if (DrCr = 1) then
               This_Bal.Trust_Interest_S86_3 := This_Bal.Trust_Interest_S86_3 + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Interest_S86_3 := This_Bal.Trust_Interest_S86_3 + ThisAmount;
            This_Abs.Trust_Interest_S86_3 := This_Abs.Trust_Interest_S86_3 + ThisAmount;
         end;

         18: begin                  // Business Debit
            if (DrCr = 1) then
               This_Bal.Business_Debit := This_Bal.Business_Debit + (ThisAmount * -1.00)
            else
               This_Bal.Business_Debit := This_Bal.Business_Debit + ThisAmount;
            This_Abs.Business_Debit := This_Abs.Business_Debit + ThisAmount;
         end;

         19: begin                  // Trust Debit
            if (DrCr = 1) then
               This_Bal.Trust_Debit := This_Bal.Trust_Debit + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Debit := This_Bal.Trust_Debit + ThisAmount;
            This_Abs.Trust_Debit := This_Abs.Trust_Debit + ThisAmount;
         end;

         20: begin                  // Trust Interest Withdrawal S86(3)
            if (DrCr = 1) then
               This_Bal.Trust_Interest_Withdrawal_S86_3 := This_Bal.Trust_Interest_Withdrawal_S86_3 + (ThisAmount * -1.00)
            else
               This_Bal.Trust_Interest_Withdrawal_S86_3 := This_Bal.Trust_Interest_Withdrawal_S86_3 + ThisAmount;
            This_Abs.Trust_Interest_Withdrawal_S86_3 := This_Abs.Trust_Interest_Withdrawal_S86_3 + ThisAmount;
         end;

         21: begin                  // Write-off
            if (DrCr = 1) then
               This_Bal.Write_off := This_Bal.Write_off + (ThisAmount * -1.00)
            else
               This_Bal.Write_off := This_Bal.Write_off + ThisAmount;
            This_Abs.Write_off := This_Abs.Write_off + ThisAmount;
         end;

         22: begin                  // Collection Debit
            if (DrCr = 1) then
               This_Bal.Collection_Debit := This_Bal.Collection_Debit + (ThisAmount * -1.00)
            else
               This_Bal.Collection_Debit := This_Bal.Collection_Debit + ThisAmount;
            This_Abs.Collection_Debit := This_Abs.Collection_Debit + ThisAmount;
         end;

         23: begin                  // Collection Credit
            if (DrCr = 1) then
               This_Bal.Collection_Credit := This_Bal.Collection_Credit + (ThisAmount * -1.00)
            else
               This_Bal.Collection_Credit := This_Bal.Collection_Credit + ThisAmount;
            This_Abs.Collection_Credit := This_Abs.Collection_Credit + ThisAmount;
         end;

         24: begin                  // Trust Transfer (Business - Other)
            if (DrCr = 1) then
               This_Bal.Trust_Transfer_Business_Other := This_Bal.Trust_Transfer_Business_Other + (ThisAmount * -1.00)
            else
               This_Abs.Trust_Transfer_Business_Other := This_Bal.Trust_Transfer_Business_Other + ThisAmount;
         end;
      end;
      adoQry.Next;
   end;
end;

//---------------------------------------------------------------------------
// Procedure to clear the passed amounts record
//---------------------------------------------------------------------------
procedure TFldExcel.Clear_Amounts();
begin
   This_Bal.Fees                              := 0.00;
   This_Bal.Disbursements                     := 0.00;
   This_Bal.Expenses                          := 0.00;
   This_Bal.Payment_Received                  := 0.00;
   This_Bal.Business_To_Trust                 := 0.00;
   This_Bal.Credit                            := 0.00;
   This_Bal.Business_Deposit                  := 0.00;
   This_Bal.Trust_Deposit                     := 0.00;
   This_Bal.Trust_Transfer_Business_Fees      := 0.00;
   This_Bal.Trust_Transfer_Disbursements      := 0.00;
   This_Bal.Trust_Transfer_Client             := 0.00;
   This_Bal.Trust_Transfer_Trust              := 0.00;
   This_Bal.Trust_Investment_S86_4            := 0.00;
   This_Bal.Trust_Withdrawal_S86_4            := 0.00;
   This_Bal.Trust_Interest_S86_4              := 0.00;
   This_Bal.Trust_Investment_S86_3            := 0.00;
   This_Bal.Trust_Withdrawal_S86_3            := 0.00;
   This_Bal.Trust_Interest_S86_3              := 0.00;
   This_Bal.Business_Debit                    := 0.00;
   This_Bal.Trust_Debit                       := 0.00;
   This_Bal.Trust_Interest_Withdrawal_S86_3   := 0.00;
   This_Bal.Write_off                         := 0.00;
   This_Bal.Collection_Debit                  := 0.00;
   This_Bal.Collection_Credit                 := 0.00;
   This_Bal.Reserved_Trust                    := 0.00;
   This_Bal.Trust_Transfer_Business_Other     := 0.00;
   This_Bal.Trust_FF_Interest_S86_4           := 0.00;

   This_Abs.Fees                              := 0.00;
   This_Abs.Disbursements                     := 0.00;
   This_Abs.Expenses                          := 0.00;
   This_Abs.Payment_Received                  := 0.00;
   This_Abs.Business_To_Trust                 := 0.00;
   This_Abs.Credit                            := 0.00;
   This_Abs.Business_Deposit                  := 0.00;
   This_Abs.Trust_Deposit                     := 0.00;
   This_Abs.Trust_Transfer_Business_Fees      := 0.00;
   This_Abs.Trust_Transfer_Disbursements      := 0.00;
   This_Abs.Trust_Transfer_Client             := 0.00;
   This_Abs.Trust_Transfer_Trust              := 0.00;
   This_Abs.Trust_Investment_S86_4            := 0.00;
   This_Abs.Trust_Withdrawal_S86_4            := 0.00;
   This_Abs.Trust_Interest_S86_4              := 0.00;
   This_Abs.Trust_Investment_S86_3            := 0.00;
   This_Abs.Trust_Withdrawal_S86_3            := 0.00;
   This_Abs.Trust_Interest_S86_3              := 0.00;
   This_Abs.Business_Debit                    := 0.00;
   This_Abs.Trust_Debit                       := 0.00;
   This_Abs.Trust_Interest_Withdrawal_S86_3   := 0.00;
   This_Abs.Write_off                         := 0.00;
   This_Abs.Collection_Debit                  := 0.00;
   This_Abs.Collection_Credit                 := 0.00;
   This_Abs.Reserved_Trust                    := 0.00;
   This_Abs.Trust_Transfer_Business_Other     := 0.00;
   This_Abs.Trust_FF_Interest_S86_4           := 0.00;
end;

//---------------------------------------------------------------------------
// Function to Get the current Default Printer
//---------------------------------------------------------------------------
function TFldExcel.GetDefaultPrinter: string;
var
   ResStr: array[0..255] of Char;
begin
   GetProfileString('Windows', 'device', '', ResStr, 255);
   Result := StrPas(ResStr);
end;

//---------------------------------------------------------------------------
// Function to Set the current Default Printer
//---------------------------------------------------------------------------
procedure TFldExcel.SetDefaultPrinter(NewDefPrinter: string);
var
   ResStr: array[0..255] of Char;
begin
   StrPCopy(ResStr, NewdefPrinter);
   WriteProfileString('windows', 'device', ResStr);
   StrCopy(ResStr, 'windows');
   SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, Longint(@ResStr));
end;

//---------------------------------------------------------------------------
// Intermediate Function to handle instances where a fuly qualified filename
// is passed to Print the Excel file to PDF format. {Overloaded}
//---------------------------------------------------------------------------
function TFldExcel.PDFDocument(FileName: string): boolean;
var
   FolderPart, FilePart : string;

begin

//--- Split the Filename into a Folder part and a File part

   FolderPart := ExtractFilePath(FileName);
   FilePart   := ExtractFileName(FileName);

//--- Now call the Print function

   Result := PDFDocument(FilePart, FolderPart);
end;

//---------------------------------------------------------------------------
// Function to Print the Excel file using the default printer which should
// be set to a PDF printer. {Overloaded}
//---------------------------------------------------------------------------
function TFldExcel.PDFDocument(FileName: string; FolderName: string): boolean;
var
   ReturnValue, idx1, idx2, SheetCount, FileCount  : Integer;
   CurrDefault, FullName, PDFFile, PDFMerge        : string;
   Delim, ThisCmd, ThisPDF, OutFile                : string;
   xlsBook                                         : IXLSWorkbook;

begin

//--- Return if the Printer name is invalid or not set

   if (PDFPrinter = 'Not Found') then begin
      LogMsg('  Unable to create PDF file. Hint: Check PDF options to select a valid PDF Printer',True);
      Result := False;
      Exit;
   end;

   FullName := FolderName + FileName;
   ThisPDF  := PDFFolder + ChangeFileExt(FileName,'.pdf');

//--- Get the Worksheet count to determine whether more than one PDF file must
//--- be produced and then merged

   xlsBook := TXLSWorkbook.Create;
   xlsBook.Open(FullName);
   SheetCount := xlsBook.Sheets.Count;
   xlsBook.Close;

//--- If there is more than 1 Sheet then the PDF Merge utility must be defined
//--- and must exist

   if (SheetCount > 1) then begin
      if (PDFMergeSel = 1) then begin
         if (FileExists(PDFMergeBullzip) = False) then begin
            LogMsg('  Unable to create PDF file. Hint: Check PDF options to select a valid PDF Merge Utility',True);
            Result := False;
            Exit;
         end;
         PDFMerge := PDFMergeBullzip;
      end else if (PDFMergeSel = 2) then begin
         if (FileExists(PDFMergePDFtk) = False) then begin
            LogMsg('  Unable to create PDF file. Hint: Check PDF options to select a valid PDF Merge Utility',True);
            Result := False;
            Exit;
         end;
         PDFMerge := PDFMergePDFtk;
      end else begin
         LogMsg('  Unable to create PDF file. Hint: Check PDF options to select a valid PDF Merge Utility',True);
         Result := False;
         Exit;
      end;
   end;

//--- Save the current default printer then set the PDFPrinter as the default

   CurrDefault := GetDefaultPrinter();
   SetDefaultPrinter(PDFPrinter);

//--- Create the PDF file. If there is only one Sheet then it is just a
//--- straight forward print. If there are more than one sheet then we need to
//--- create each individual sheet as a file and then at the end merge all the
//--- individual sheets into a single PDF file

   if (SheetCount = 1) then begin
      ReturnValue := ShellExecute(0, 'print', PChar(FullName), nil, nil, SW_HIDE);
   end else begin
      FileCount := 0;

//--- Step through each sheet, make it the default then print as PDF and
//--- finally store as sequentially numbered files

     for idx2 := 1 to SheetCount do begin

         xlsBook.Open(FullName);
         xlsBook.Sheets[idx2].Activate;
         xlsBook.Save;
         xlsBook.Close;
         ReturnValue := ShellExecute(0, 'print', PChar(FullName), nil, nil, SW_HIDE);

         if (ReturnValue < 33) then
            break;

         Sleep(3000);

         OutFile := PDFFolder + 'A' + IntToStr(idx2) + '.pdf';
         DeleteFile(OutFile);
         RenameFile(ThisPDF,OutFile);

         inc(FileCount);
      end;

//--- Merge the sequentially numbered files into the correctly named output file

      if (FileCount = SheetCount) then begin

         case PDFMergeSel of
            1: begin                   // Bullzip Merge Utility
               Delim := '';
               ThisCmd := 'command=merge input="';

               for idx2 := 1 to SheetCount do begin
                  OutFile := PDFFolder + 'A' + IntToStr(idx2) + '.pdf';

                  ThisCmd := ThisCmd + Delim + OutFile;
                  Delim := '|';
               end;

               ThisCmd := ThisCmd + '" output="' + ThisPDF + '"';
            end;

            2: begin                   // PDFtk Merge Utility
               Delim := '';
               ThisCmd := '"';

               for idx2 := 1 to SheetCount do begin
                  OutFile := PDFFolder + 'A' + IntToStr(idx2) + '.pdf';

                  ThisCmd := ThisCmd + Delim + OutFile;
                  Delim := '" "';
               end;

               ThisCmd := ThisCmd + '" cat output "' + ThisPDF + '"';
            end;
         end;

         ReturnValue := ShellExecute(Handle,'open',PChar(PDFMerge), PChar(ThisCmd), nil,SW_HIDE);

//--- Allow enough time for the PDF Merge operation to complete

         sleep(3000);

//--- Delete the temporary files

         for idx2 := 1 to SheetCount do
            DeleteFile(PDFFolder + 'A' + IntToStr(idx2) + '.pdf');

//--- Make the First sheet in the Workbook the default again

         xlsBook.Open(FullName);
         xlsBook.Sheets[1].Activate;
         xlsBook.Save;
         xlsBook.Close;

      end;
   end;

//--- Restore the previous default printer

   SetDefaultPrinter(CurrDefault);

//--- Indicate the result of the operation

   if (ReturnValue < 33) then begin
      LogMsg('  Unable to create PDF file - return code = ''' + IntToStr(ReturnValue) + '''',True);
      Result := False;
      Exit;
   end else begin

//--- Move the generated PDF file from the default PDF Folder to the Folder
//--- where the Excel file was saved.

      PDFFile := ChangeFileExt(FileName,'.pdf');
      DeleteFile(PChar(FolderName + PDFFile));
      if (PDFRetry = 0) then PDFRetry := 1;

//--- We will try at most PDFRetry times

      for idx1 := 0 to PDFRetry do begin

         Sleep(PDFInterval * 1000);

//--- Check whether the PDF utility has finished creating the file and if so
//--- then move the file from the default PDF folder to the gesignated folder

         if (FileExists(PChar(PDFFolder + PDFFile)) = true) then begin

            MoveFile(PChar(PDFFolder + PDFFile),PChar(FolderName + PDFFile));
            break;

         end;
      end;

//--- Make sure that the move was successful, this is required when a large
//--- number of files are processed

      for idx1 := 0 to PDFRetry do begin

         Sleep(PDFInterval * 1000);

         if (FileExists(PChar(FolderName + PDFFile)) = true) then
            break
         else
            MoveFile(PChar(PDFFolder + PDFFile),PChar(FolderName + PDFFile));

      end;

//--- If after the maximum retries the file has still not appeared in the
//--- target folder then give up

      if (FileExists(PChar(FolderName + PDFFile)) = false) then begin
         LogMsg('  Unable to move ''' + PChar(PDFFile) + ''' to ''' + PChar(FolderName) + '''', True);
         Result := False;
         Exit;
      end;

      Result := True;
   end;
end;

//---------------------------------------------------------------------------
// Intermediate Function to handle instances where a fuly qualified filename
// is passed to Print the Excel file on the LPMS Default Printer
//---------------------------------------------------------------------------
function TFldExcel.PrintDocument(FileName: string): boolean;
var
   FolderPart, FilePart : string;

begin

//--- Split the Filename into a Folder part and a File part

   FolderPart := ExtractFilePath(FileName);
   FilePart   := ExtractFileName(FileName);

//--- Now call the Print function

   Result := PrintDocument(FilePart, FolderPart);
end;

//---------------------------------------------------------------------------
// Function to Print the Excel file on the LPMS Default Printer
//---------------------------------------------------------------------------
function TFldExcel.PrintDocument(FileName: string; FolderName: string): boolean;
var
   ReturnValue           : Integer;
   CurrDefault, FullName : string;

begin

//--- Return if the Printer name is invalid or not set

   if (DefPrinter = 'Not Found') then begin
      LogMsg('Unable to print file - check print options',True);
      Result := False;
      Exit;
   end;

   FullName := FolderName + FileName;

//--- Save the current default printer then set the LPMS Default Printer as the
//--- default

   CurrDefault := GetDefaultPrinter();
   SetDefaultPrinter(DefPrinter);

//--- Print the generated document.

   ReturnValue := ShellExecute(0, 'print', PChar(FullName), nil, nil, SW_HIDE);

//--- Restore the previous default printer

   SetDefaultPrinter(CurrDefault);

//--- Indicate the result of the operation

   if (ReturnValue < 33) then begin
      LogMsg('Unable to print file - return code = ''' + IntToStr(ReturnValue) + '''',True);
      Result := False;
      Exit;
   end;

   Result := True;
end;

//---------------------------------------------------------------------------
// Procedure to print messages to the Running Log - Standard prolog/epilogue
//---------------------------------------------------------------------------
procedure TFldExcel.LogMsg(ThisType: integer; ThisInit: boolean; ShowOptions: boolean; ThisMsg: string);
var
   idx1       : integer;

begin

//--- Start the animation and clear the Running Log if ThisInit is set

   if (ThisInit = True) then begin
      lbProgress.Clear;

      if (RunType in [1,3..6,8,9,12,15,26]) then begin
         prbProgress.Show;
         stCount.Visible := True;
      end;

      FldExcel.Refresh;
   end;

//--- Processing if ThisType == PT_PROLOG

   if (ThisType = ord(PT_PROLOG)) then begin
      ShortDateFormat := 'yyyy/MM/dd';
      DateSeparator   := '/';

      LogMsg('*** LPMS Spreadsheet Generator (' + ParamStr(24) + ') - ' + ThisMsg + ' generation started on ' + FormatDateTime('yyyy/MM/dd', Now()) + ' at ' + FormatDateTime('HH:mm:ss.zzz',Now()) + ' ***',True);

//--- If LPMS is in Debug/Development mode then display all parameters

      if (VersionNum = 'DEBUG') then begin
         LogMsg(' ',True);

         LogMsg('    1 (' + ThisDebug[1] + ') : ' + ParamStr(1) + ' [' + ThisLabels[RunType + 1] + ']',True);

         for idx1 := 2 to 9 do
            LogMsg('    ' + IntToStr(idx1) + ' (' + ThisDebug[idx1] + ') : ' + ParamStr(idx1),True);

         for idx1 := 10 to 27 do
            LogMsg('   ' + IntToStr(idx1) + ' (' + ThisDebug[idx1] + ') : ' + ParamStr(idx1),True);
      end;

//--- If we are doing Periodic Billing then the options are passed via an XML
//--- file per instruction and are not available at the time that PT_PROLOG is
//--- invoked

      if (ShowOptions = True) then begin
         LogMsg(' ',True);

         if (CloseOnComplete = True) then
            LogMsg('  Close on completion: Yes',True)
         else
            LogMsg('  Close on completion: No',True);

         if (AutoOpen = True) then
            LogMsg('  Open generated ' + ThisMsg + ' after completion: Yes',True)
         else
            LogMsg('  Open generated ' + ThisMsg + ' after completion: No',True);

         if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then
            LogMsg('  Create PDF from generated Excel file: Yes',True);

         if ((DoPrint = True) and (DEFPrinter <> 'Not Found')) then
            LogMsg('  Print generated Excel file on LPMS Printer: Yes',True);

         if ((StoreInv = true) and (NilBalance = '0')) then
            LogMsg('  Store generated Invoice after completion: Yes',True);

         if (SendByEmail = '1') then
            LogMsg('  Email generated ' + ThisMsg + ' after completion: Yes',True);

         if (SendByEmail = '2') then begin
            LogMsg('  Send generated Invoice(s) by Email after completion: Yes',True);
            GroupAttach := true;
         end;
      end;
      LogMsg(' ',True);
   end;

//--- Processing if ThisType == PT_MIDLOG

   if (ThisType = ord(PT_MIDLOG)) then begin
      LogMsg('--- ' + ThisMsg + ' ---',True);
      LogMsg(' ',True);
   end;

//--- Processing if ThisType == PT_OPTLOG

   if (ThisType = ord(PT_OPTLOG))  then begin
      if (AutoOpen = True) then
         LogMsg('  Open generated ' + ThisMsg + ' after completion: Yes',True)
      else
         LogMsg('  Open generated ' + ThisMsg + ' after completion: No',True);

      if ((CreatePDF = True) and (PDFPrinter <> 'Not Found')) then
         LogMsg('  Create PDF from generated Excel file: Yes',True);

      if ((DoPrint = True) and (DEFPrinter <> 'Not Found')) then
         LogMsg('  Print generated Excel file on LPMS Printer: Yes',True);

      if ((StoreInv = true) and (NilBalance = '0')) then
         LogMsg('  Store generated Invoice after completion: Yes',True);

      if (SendByEmail = '1') then
         LogMsg('  Email generated ' + ThisMsg + ' after completion: Yes',True);

      LogMsg(' ', True);
   end;

//--- Processing if ThisType == PT_EPILOG

   if (ThisType = ord(PT_EPILOG)) then begin
      prbProgress.Hide;
      stCount.Visible := False;

      ShortDateFormat := 'yyyy/MM/dd';
      DateSeparator   := '/';

      LogMsg('*** LPMS Spreadsheet Generator (' + ParamStr(24) + ') - ' + ThisMsg + ' generation completed on ' + FormatDateTime('yyyy/MM/dd', Now()) + ' at ' + FormatDateTime('HH:mm:ss.zzz',Now()) + ' ***',True);
   end;

end;

//---------------------------------------------------------------------------
// Procedure to print a single line to the Running Log
//---------------------------------------------------------------------------
procedure TFldExcel.LogMsg(ThisMsg: string; DoAdjust: boolean);
begin
   lbProgress.Items.Add(ThisMsg);

   if (DoAdjust = True) then begin
      lbProgress.TopIndex := lbProgress.Items.Count - 1;
      lbProgress.Refresh;
   end;
end;

//---------------------------------------------------------------------------
// Function to Disassemble the Email Signature String
//---------------------------------------------------------------------------
function TFldExcel.Disassemble(Str: string; ThisDelim: char): TStringList;
var
   idx1, Len : integer;
   Delim     : char;
   ThisLine  : string;
   ThisList  : TStringList;

begin

   ThisList := TStringList.Create;
   Delim    := ThisDelim;
   Len      := Length(Str);
   ThisLine := '';

   for idx1 := 1 to Len do begin
      if ((Str[idx1] = Delim) or (idx1 = Len)) then begin
         if (ThisLine = '') then ThisLine := ' ';
         ThisList.Add(ThisLine);
         ThisLine := '';
      end else begin
         ThisLine := ThisLine + Str[idx1];
      end;
   end;
   Result := ThisList;
end;

//---------------------------------------------------------------------------
// Function to transform special characters
//---------------------------------------------------------------------------
function TFldExcel.ReplaceQuote(S1: string): string;
begin

   S1 := AnsiReplaceStr(S1,'&quot', '''');
   S1 := AnsiReplaceStr(S1,'&slash','\');
   S1 := AnsiReplaceStr(S1,'§',     '\');        // Retained for backwards compatibility

   Result := S1;
end;

//---------------------------------------------------------------------------
// Function to replace special characters in an XML line
//---------------------------------------------------------------------------
function TFldExcel.ReplaceXML(S1: string): string;
begin

   S1 := AnsiReplaceStr(S1,'$quote;','''');
   S1 := AnsiReplaceStr(S1,'&quot;','''');       // Stored in the DB like this
   S1 := AnsiReplaceStr(S1,'$slash;','\');
   S1 := AnsiReplaceStr(S1,'&slash','\');        // Stored in the DB like this
   S1 := AnsiReplaceStr(S1,'§','\');             // Retained for backwards compatibility
   S1 := AnsiReplaceStr(S1,'$amp;','&');
   S1 := AnsiReplaceStr(S1,'$LeftA;','<');
   S1 := AnsiReplaceStr(S1,'$RightA;','>');
   S1 := AnsiReplaceStr(S1,'##null##','');

   Result := S1;
end;

//---------------------------------------------------------------------------
// Function to Replace Symbolic Variables
//---------------------------------------------------------------------------
function TFldExcel.DoSymVars(ThisStr: AnsiString; FileStr: AnsiString): AnsiString;
var
   idx1               : integer;
   ThisDate, ThisTime : string;
   ThisDesc           : AnsiString;

begin
   ThisDesc := ThisStr;

//--- Set the Symbolic Variables specific to FldExcel

   ThisDate := FormatDateTime('yyyy/MM/dd',Now());
   ThisTime := FormatDateTime('HH:NN:SS',Now());

   SymVars_LPMS.SV[ord(SV_FROMDATE)].Value   := SDate;
   SymVars_LPMS.SV[ord(SV_ENDDATE)].Value    := EDate;
   SymVars_LPMS.SV[ord(SV_CURRDATE)].Value   := ThisDate;
   SymVars_LPMS.SV[ord(SV_FILE)].Value       := FileStr;
   SymVars_LPMS.SV[ord(SV_VATNUM)].Value     := VATNumber;
   SymVars_LPMS.SV[ord(SV_DATE)].Value       := ThisDate;
   SymVars_LPMS.SV[ord(SV_TIME)].Value       := ThisTime;
   SymVars_LPMS.SV[ord(SV_SHORTYEAR)].Value  := Copy(ThisDate,1,4);
   SymVars_LPMS.SV[ord(SV_SHORTTIME)].Value  := Copy(ThisTime,1,5);
   SymVars_LPMS.SV[ord(SV_USER)].Value       := UserName;

//--- We do Local and Global Symbolic Variables first as these take precendece

   for idx1 := 0 to Length(SymVars_Other.SV) - 1 do begin
      ThisDesc := AnsiReplaceStr(ThisDesc,'&' + SymVars_Other.SV[idx1].Variable,SymVars_Other.SV[idx1].Value);
   end;

//--- Then we do System Symbolic Variables

   for idx1 := 0 to ord(SV_COUNT) - 1 do begin
      ThisDesc := AnsiReplaceStr(ThisDesc,'&' + SymVars_LPMS.SV[idx1].Variable,SymVars_LPMS.SV[idx1].Value);
   end;
//   ThisDesc := AnsiReplaceStr(ThisDesc,'%',' ');

   Result := ReplaceQuote(ThisDesc);
end;

//---------------------------------------------------------------------------
// Function to Set up and intialise all Symbolic Varaibles
//---------------------------------------------------------------------------
procedure TFldExcel.SetUpSymVars();
var
   idx1, idx2                                     : integer;
   ThisDate, ThisTime, CpyFile, CpyName, HostName : AnsiString;
   RootFolder, Version, S1, ThisMonth             : Ansistring;
   RegIni                                         : TRegistry;

begin

//--- Initialise the System SymVar table with the built-in Symbolic Variables

   for idx1 := 0 to ord(SV_COUNT) - 1 do begin
      SymVars_LPMS.SV[idx1].Scope := 'LPMS';
      SymVars_LPMS.SV[idx1].Value := 'Not set';
   end;

{$INCLUDE 'SetUpSymVarsD.inc'}

//--- Load the values from the Registry that are known at this point

   RegIni := TRegistry.Create;
   RegIni.RootKey := HKEY_CURRENT_USER;
   RegIni.OpenKey(RegString,false);

   HostName   := RegIni.ReadString('HostName');
   RootFolder := RegIni.ReadString('Rootfolder');

   RegIni.CloseKey;
   RegIni.Free;

//--- Load the values from the Database that are known at this point

   FldExcel.Cursor := crHourGlass;

   S1 := 'SELECT CpyFile, CpyName, Version FROM lpms';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      On E : Exception do begin
         MessageDlg('Unexpected Data Base Error - "' + E.Message + '"', mtWarning, [mbOK], 0);
         FldExcel.Cursor := crDefault;
         Close;
      end;
   end;

   CpyFile := ReplaceQuote(Query1.FieldByName('CpyFile').AsString);
   CpyName := ReplaceQuote(Query1.FieldByName('CpyName').AsString);
   Version := ReplaceQuote(Query1.FieldByName('Version').AsString);

//--- Set all values that are known at this point

   ShortDateFormat := 'yyyy/MM/dd';
   DateSeparator   := '/';
   ThisDate        := FormatDateTime('yyyy/MM/dd',Now);
   ThisTime        := FormatDateTime('HH:nn:SS',Now);
   ThisMonth       := FormatDateTime('MMMM',Now);

   SymVars_LPMS.SV[ord(SV_CPYFILE)].Value   := CpyFile;
   SymVars_LPMS.SV[ord(SV_CPYNAME)].Value   := CpyName;
   SymVars_LPMS.SV[ord(SV_DATE)].Value      := ThisDate;
   SymVars_LPMS.SV[ord(SV_DAY)].Value       := Copy(ThisDate,9,2);
   SymVars_LPMS.SV[ord(SV_DBPREFIX)].Value  := DBPrefix;
   SymVars_LPMS.SV[ord(SV_HOSTNAME)].Value  := HostName;
   SymVars_LPMS.SV[ord(SV_LONGMONTH)].Value := ThisMonth;
   SymVars_LPMS.SV[ord(SV_MONTH)].Value     := Copy(ThisDate,6,2);
   SymVars_LPMS.SV[ord(SV_ROOTF)].Value     := RootFolder;
   SymVars_LPMS.SV[ord(SV_SHORTYEAR)].Value := Copy(ThisDate,3,2);
   SymVars_LPMS.SV[ord(SV_SHORTTIME)].Value := Copy(ThisTime,1,5);
   SymVars_LPMS.SV[ord(SV_TIME)].Value      := ThisTime;
   SymVars_LPMS.SV[ord(SV_VERSION)].Value   := 'Version ' + Version;
   SymVars_LPMS.SV[ord(SV_YEAR)].Value      := Copy(ThisDate,1,4);
   SymVars_LPMS.SV[ord(SV_AMP)].Value       := '&';

//--- Load the Global and Local Symbolic Variables. Global and Local Synboloc
//--- Variables come before System Symbolic Variables so that Global and Local
//--- Symbolic Variables can override System Symbolic Variables. For security
//--- reasons Global in turn overrides Local.

//--- Load Global Symbolic Variables

   S1 := 'SELECT Sym_Owner, Sym_Variable, Sym_Value FROM symvars WHERE Sym_Owner = ''Global'' ORDER BY Sym_Variable ASC';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      On E : Exception do begin
         MessageDlg('Unexpected Data Base Error - "' + E.Message + '"', mtWarning, [mbOK], 0);
         FldExcel.Cursor := crDefault;
         Close;
      end;
   end;

//--- Set the number of entries in the array to the number of records returned

   SetLength(SymVars_Other.SV,Query1.RecordCount);
   idx2 := Query1.RecordCount;

//--- Step through the Global Symbolic Variables

   Query1.First();

   for idx1 := 0 to Query1.RecordCount - 1 do begin
      SymVars_Other.SV[idx1].Scope    := ReplaceQuote(Query1.FieldByName('Sym_Owner').AsString);
      SymVars_Other.SV[idx1].Variable := ReplaceQuote(Query1.FieldByName('Sym_Variable').AsString);
      SymVars_Other.SV[idx1].Value    := ReplaceQuote(Query1.FieldByName('Sym_Value').AsString);

      Query1.Next;
   end;

//--- Load Local Symbolic Variables

   S1 := 'SELECT Sym_Owner, Sym_Variable, Sym_Value FROM symvars WHERE Sym_Owner = ''' + UserName + ''' ORDER BY Sym_Variable ASC';

   try
      Query1.Close;
      Query1.SQL.Text := S1;
      Query1.Open;
   except
      On E : Exception do begin
         MessageDlg('Unexpected Data Base Error - "' + E.Message + '"', mtWarning, [mbOK], 0);
         FldExcel.Cursor := crDefault;
         Close;
      end;
   end;

//--- Expand the number of entries in the array to accommodate the Local SymVars

   SetLength(SymVars_Other.SV,idx2 + Query1.RecordCount);

//--- Step through the Local Symbolic Variables

   Query1.First();

   for idx1 := 0 to Query1.RecordCount - 1 do begin
      SymVars_Other.SV[idx2 + idx1].Scope    := ReplaceQuote(Query1.FieldByName('Sym_Owner').AsString);
      SymVars_Other.SV[idx2 + idx1].Variable := ReplaceQuote(Query1.FieldByName('Sym_Variable').AsString);
      SymVars_Other.SV[idx2 + idx1].Value    := ReplaceQuote(Query1.FieldByName('Sym_Value').AsString);

      Query1.Next;
   end;

   FldExcel.Cursor := crDefault;
end;

//---------------------------------------------------------------------------
// Function to Round a Floating number to 'd' decimals
//---------------------------------------------------------------------------
function TFldExcel.RoundD(x: Extended; d: Integer): Extended;
var
  n: Extended;
begin
  n := IntPower(10, d);
  x := x * n;
  Result := (Int(x) + Int(Frac(x) * 2)) / n;
end;

//---------------------------------------------------------------------------
// Function to do a Vignere Cypher
//---------------------------------------------------------------------------
function TFldExcel.Vignere(ThisType: integer; Phrase: string; Key: string): string;
var
   idx1, idx2, PhraseLen, ThisKeyLen : integer;
   TempKey                           : array of char;
   NewKey                            : array of char;
   EncryptedMsg                      : array of char;
   ThisPhrase                        : string;

begin
   PhraseLen  := Length(Phrase);
   ThisKeyLen := Length(Key);

   SetLength(TempKey,ThisKeyLen);

//--- Remove spaces from the Key and translate to upper case

   idx1 := 0;
   idx2 := 1;

   while (idx2 <= ThisKeyLen) do begin
      if (Key[idx2] = ' ') then
         inc(idx2)
      else begin
         TempKey[idx1] := upcase(Key[idx2]);

         inc(idx1);
         inc(idx2);
      end;
   end;

   SetLength(TempKey,idx1);
   ThisKeyLen := Length(TempKey);

//--- Start by extending/limiting the Key to the same length as the Phrase

   SetLength(NewKey,PhraseLen);
   idx2 := 0;

   for idx1 := 0 to PhraseLen - 1 do begin
      if (idx2 = ThisKeyLen) then
         idx2 := 0;

      NewKey[idx1] := TempKey[idx2];
      inc(idx2);
   end;

//--- Do the Encryption or Decryption depending on the value of ThisType. Only
//--- characters in the range A-Z and a-z are transformed. The rest are
//--- preserved as is.

   SetLength(EncryptedMsg,PhraseLen);

   case ThisType of
      ord(CYPHER_ENC): begin
         for idx1 := 1 to PhraseLen do begin
            if (InRange(ord(Phrase[idx1]),ord('A'),ord('Z'))) then
               EncryptedMsg[idx1 - 1] := char((((ord(Phrase[idx1]) + ord(NewKey[idx1 - 1])) mod 26) + ord('A')))
            else if (InRange(ord(Phrase[idx1]),ord('a'),ord('z'))) then
               EncryptedMsg[idx1 - 1] := char((((ord(Phrase[idx1]) + ord(NewKey[idx1 - 1])) mod 26) + ord('a')))
            else
               EncryptedMsg[idx1 - 1] := char(Phrase[idx1]);
         end
      end;

      ord(CYPHER_DEC): begin
         for idx1 := 1 to PhraseLen do begin
            if (InRange(ord(Phrase[idx1]),ord('A'),ord('Z'))) then
               EncryptedMsg[idx1 - 1] := char(((((ord(Phrase[idx1]) - ord(NewKey[idx1 - 1])) + 26) mod 26) + ord('A')))
            else if (InRange(ord(Phrase[idx1]),ord('a'),ord('z'))) then
               EncryptedMsg[idx1 - 1] := char(((((ord(Phrase[idx1]) - ord(NewKey[idx1 - 1])) + 14) mod 26) + ord('a')))
            else
               EncryptedMsg[idx1 - 1] := char(Phrase[idx1]);
         end;
      end;
   end;

//--- Transfer the Encrypted/Decrypted phrase to a String so that it can be
//--- returned

   for idx1 := 0 to PhraseLen - 1 do
      ThisPhrase := ThisPhrase + EncryptedMsg[idx1];

   Result := ThisPhrase;
end;

//---------------------------------------------------------------------------
// Procedure to do a Bubble sort on the Satement records
//---------------------------------------------------------------------------
procedure TFldExcel.BubbleSort(var RecList: Array of LPMS_Statement);
var
   idx       : Integer;
   Changed   : Boolean;
   ThisRec   : LPMS_Statement;

begin
   Changed := True;

   while Changed do begin
      Changed := False;

      for idx := Low(RecList) to High(RecList) - 1 do begin

        if (RecList[idx].DateTime > RecList[idx + 1].DateTime) then begin
           ThisRec          := RecList[idx + 1];
           RecList[idx + 1] := RecList[idx];
           RecList[idx]     := ThisRec;
           Changed          := True;
        end;

      end;

   end;

end;

//---------------------------------------------------------------------------
end.
