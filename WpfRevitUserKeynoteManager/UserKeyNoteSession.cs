using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

namespace WpfRevitUserKeynoteManager
{
    public class UserKeyNoteSession : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        // Create the OnPropertyChanged method to raise the event
        protected void OnPropertyChanged(string propName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }

        private string knfolder;
        private string knfile = "<= Double click here for selecting a Revit Userkeynote file.";
        private DataTable kndataTable = new KNTable();
        private DataTable kndatacatState = new KNTableCatState();
        private string currentCatCode = "*";
        private string currentCatName = "*";
        private DataRowView currentDRV;
        private string findThisText = string.Empty;
        private string editKeyText;
        private string editCellText;
        private string editCatText;
        private int editRowIndex;
        private string notecount;
        private string replaceWithThisText = string.Empty;
        private bool replaceAsWord = false;
        private string statusMSG;
        private string statusMSGSettingsA;
        private string statusMSGSettingsB;
        private string categoryInReserve = string.Empty;
        private string categoryInReserveName = string.Empty;
        private bool categoryIsReserved = false;
        private bool reserveCategoryChangesNotCommitted;
        private DateTime lastModified;
        private bool replacementAllowed;
        private bool editcellisReadOnly;
        private string buttoncreatenewContent;
        private string lablecontentsContents;
        private bool tableisstale;
        private int pollmsec = 1000;
        private bool othersHaveReserved = false;
        private string histText = string.Empty;
        private string histTextFileName = string.Empty;
        private string histIndex = string.Empty;
        private string helpText = string.Empty;
        private bool readOnly = true;
        private List<string> newCategories = new List<string>();
        private bool seenNewCatMsg = false;
        private string keystatictext = string.Empty;
        private string keynumberformat = "00";
        private Single keyincval = 1F;
        private string historybuttontext = "History Files";

        public string Historybuttontext { get { return historybuttontext; } set { historybuttontext = value; OnPropertyChanged("Historybuttontext"); } }
        public string Keystatictext { get { return keystatictext; } set { keystatictext = value; OnPropertyChanged("Keystatictext"); } }
        public string Keynumberformat { get { return keynumberformat; } set { keynumberformat = value; OnPropertyChanged("Keynumberformat"); } }
        public Single Keyincval { get { return keyincval; } set { keyincval = value; OnPropertyChanged("Keyincval"); } }
        public bool SeenNewCatMsg { get { return seenNewCatMsg; } set { seenNewCatMsg = value; OnPropertyChanged("SeenNewCatMsg"); } }
        public List<string> NewCategories { get { return newCategories; } set { newCategories = value; OnPropertyChanged("NewCategories"); } }
        public string HelpText { get { return helpText; } set { helpText = value; OnPropertyChanged("HelpText"); } }
        public bool ReadOnly { get { return readOnly; } set { readOnly = value; OnPropertyChanged("ReadOnly"); } }
        public string HistTextFileName { get { return histTextFileName; } set { histTextFileName = value; OnPropertyChanged("HistTextFileName"); } }
        public string HistText { get { return histText; } set { histText = value; OnPropertyChanged("HistText"); } }
        public string HistIndex { get { return histIndex; } set { histIndex = value; OnPropertyChanged("HistIndex"); } }
        public DataRowView CurrentDRV  { get { return currentDRV; } set { currentDRV = value; OnPropertyChanged("CurrentDRV"); } }
        public bool OthersHaveReserved { get { return othersHaveReserved; } set { othersHaveReserved = value; OnPropertyChanged("OthersHaveReserved"); } }
        public int Pollmsec { get { return pollmsec; } set { pollmsec = value; OnPropertyChanged("Pollmsec"); } }
        public bool Tableisstale { get { return tableisstale; } set { tableisstale = value; OnPropertyChanged("Tableisstale"); } }
        public string LablecontentsContents { get { return lablecontentsContents; } set { lablecontentsContents = value; OnPropertyChanged("LablecontentsContents"); } }
        public string ButtoncreatenewContent { get { return buttoncreatenewContent; } set { buttoncreatenewContent = value; OnPropertyChanged("ButtoncreatenewContent"); } }
        public bool EditcellisReadOnly { get { return editcellisReadOnly; } set { editcellisReadOnly = value; OnPropertyChanged("EditcellisReadOnly"); } }
        public string Knfolder { get { return knfolder; } set { knfolder = value; OnPropertyChanged("Knfolder"); } }
        public string Knfile { get { return knfile; } set { knfile = value; OnPropertyChanged("Knfile"); } }
        public DataTable KndataTable { get { return kndataTable; } set { kndataTable = value; OnPropertyChanged("KndataTable"); } }
        public DataTable KndatacatState { get { return kndatacatState; } set { kndatacatState = value; OnPropertyChanged("KndatacatState"); } }
        public string CurrentCatCode { get { return currentCatCode; } set { currentCatCode = value.Trim(); OnPropertyChanged("CurrentCatCode"); } }
        public string CurrentCatName { get { return currentCatName; } set { currentCatName = value; OnPropertyChanged("CurrentCatName"); } }
        public string FindThisText { get { return findThisText; } set { findThisText = value; OnPropertyChanged("FindThisText"); } }
        public string EditKeyText { get { return editKeyText; } set { editKeyText = value; OnPropertyChanged("EditKeyText"); } }
        public string EditCellText { get { return editCellText; } set { editCellText = value; OnPropertyChanged("EditCellText"); } }
        public string EditCatText { get { return editCatText; } set { editCatText = value; OnPropertyChanged("EditCatText"); } }
        public int EditRowIndex { get { return editRowIndex; } set { editRowIndex = value; OnPropertyChanged("EditRowIndex"); } }
        public string Notecount { get { return notecount; } set { notecount = value; OnPropertyChanged("Notecount"); } }
        public string ReplaceWithThisText { get { return replaceWithThisText; } set { replaceWithThisText = value; OnPropertyChanged("ReplaceWithThisText"); } }
        public bool ReplaceAsWord { get { return replaceAsWord; } set { replaceAsWord = value; OnPropertyChanged("ReplaceAsWord"); } }
        public string StatusMSG { get { return statusMSG; } set { statusMSG = value; OnPropertyChanged("StatusMSG"); } }
        public string StatusMSGSettingsA { get { return statusMSGSettingsA; } set { statusMSGSettingsA = value; OnPropertyChanged("StatusMSGSettingsA"); } }
        public string StatusMSGSettingsB { get { return statusMSGSettingsB; } set { statusMSGSettingsB = value; OnPropertyChanged("StatusMSGSettingsB"); } }
        public string CategoryInReserve { get { return categoryInReserve; } set { categoryInReserve = value; OnPropertyChanged("CategoryInReserve"); } }
        public string CategoryInReserveName { get { return categoryInReserveName; } set { categoryInReserveName = value; OnPropertyChanged("CategoryInReserveName"); } }
        public bool CategoryIsReserved { get { return categoryIsReserved; } set { categoryIsReserved = value; OnPropertyChanged("CategoryIsReserved"); } }
        public bool ReserveCategoryChangesNotCommitted { get { return reserveCategoryChangesNotCommitted; } set { reserveCategoryChangesNotCommitted = value; OnPropertyChanged("ReserveCategoryChangesNotCommitted"); } }
        public DateTime LastModified { get { return lastModified; } set { lastModified = value; OnPropertyChanged("LastModified"); } }
        public bool ReplacementAllowed { get { return replacementAllowed; } set { replacementAllowed = value; OnPropertyChanged("ReplacementAllowed"); } }


        //private double[] scales = new double[7];
        //public double[] Scales { get { return scales; } set { scales = value; OnPropertyChanged("Scales"); } }

    }
    
}
