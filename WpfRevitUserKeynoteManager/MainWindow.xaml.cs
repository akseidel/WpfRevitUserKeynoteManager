using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Markup;
using System.Diagnostics;
using System.Reflection;

namespace WpfRevitUserKeynoteManager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool dbug = false;
        public UserKeyNoteSession KNS;
        public string Knfolder { get; set; }
        public string Knfile { get; set; }
        private string priorsfolder = @"History\PriorKeynoteFiles\";
        private string InfoFileName = "RevitKeyNotesExplained.rtf";
        private char pipechar = '|';
        private char space = ' ';
        private string extRKU = ".RKU";
        private string all;
        private string Available = "Available";
        private IBindingListView blvnotes;
        private IBindingListView blvcats;
        private string theLastStatusMSG = string.Empty;
        private int catcodepad = 3;
        private bool validNewCatCode = false;
        private bool validNewCatName = false;
        private readonly BackgroundWorker bwKNFFolderWatcher = new BackgroundWorker();
        public KNFFolderWatcherArgs KNFArgs = new KNFFolderWatcherArgs();
        private bool Inhibit = false;
        private readonly BackgroundWorker bkFileLoadingWorker = new BackgroundWorker();
        private string TempLastSelCat = "*";
        private string TempLastKey = "*";
        private Dictionary<string, List<string>> dictionaryPending = new Dictionary<string, List<string>>();
        private bool exitnormally = true;
        Instructions HowTo;

        public MainWindow()
        {
            InitializeComponent();
            PreviewKeyDown += new KeyEventHandler(CloseOnEscape);
            SetUpBackgroundWorkers();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            InitialzeTheKNS();
            RestoreLastSelections();
            if (dbug)
            {
                MessageBox.Show(TempLastSelCat + "  " + TempLastKey, "Temps: Window_Loaded after Restore");
                MessageBox.Show(KNS.CurrentCatCode + "  " + KNS.EditKeyText, "KNS: Window_Loaded after Restore");
            }
        }

        private void InitialzeTheKNS()
        {
            TempLastSelCat = Properties.Settings.Default.LastRestoreCat;
            TempLastKey = Properties.Settings.Default.LastSelKey;
            bwKNFFolderWatcher.CancelAsync();
            UnRegisterStatusAsKeyNoteUser();
            KNS = new UserKeyNoteSession();
            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
              new Action(() => DataContext = KNS));
            string thisKNF;
            thisKNF = EstablishKNFFileName(Knfolder, Knfile);
            if (IsThereExcelKeyedNoteSystem()) { this.Close(); return; };
            all = "*".PadRight(catcodepad) + space + pipechar + space + "*";
            RegisterStatusAsKeyNoteUser();
            ReadUpdateKNTables(thisKNF);
            ReadCurretnRKUStatus();
            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
              new Action(() => NotesGrid.DataContext = KNS.KndataTable));
            blvnotes = KNS.KndataTable.DefaultView;
            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
             new Action(() => CatStateGrid.DataContext = KNS.KndatacatState));
            blvcats = KNS.KndatacatState.DefaultView;
            KNS.HelpTextFileName = InfoFileName;
            SetCounts();
            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
            new Action(() => ComboCatCode.ItemsSource = KNS.KndatacatState.DefaultView));

            if (Application.Current.Dispatcher.CheckAccess())
            {
                ComboCatCode.SelectedIndex = 0;
            }
            else
            {
                Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
               new Action(() => ComboCatCode.SelectedIndex = 0));
            }
            //ComboCatCode.SelectedIndex = 0;
            SetEditingAndNewPrompts();
            SetReplacementButton(false);
            //MainStatusMsg("Pressing Escape key quits this window.");
            try
            {
                KNFArgs.Folderpath = KNS.Knfolder;
                KNFArgs.Knffilemarker = Path.GetFileNameWithoutExtension(KNS.Knfile);
                KNFArgs.Pollmsec = KNS.Pollmsec;
                KNFArgs.Usertag = Environment.UserName.ToUpper();

                if (bwKNFFolderWatcher.IsBusy != true)
                {
                    bwKNFFolderWatcher.RunWorkerAsync(KNFArgs);
                    MainStatusMsg("Watchdogging " + KNS.Knfolder);
                }
                if (KNS.ReadOnly)
                {
                    MainStatusMsg("Heck! This keynotefile is readonly.");
                }
            }
            catch (System.ArgumentException)
            {

            }
            // Load the help file at the very end on a thread
            //System.Threading.ThreadPool.QueueUserWorkItem(delegate { LoadUpHelpFile(); }, null);
        }

        private void CloseOnEscape(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                //Application.Current.Shutdown();
                Close();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            NormalClosingProceedure();
        }

        private void NormalClosingProceedure()
        {
            if (HowTo != null) { HowTo.Close(); };
            bwKNFFolderWatcher.CancelAsync();
            bwKNFFolderWatcher.Dispose();
            UnRegisterStatusAsKeyNoteUser();
            MainStatusMsg("Watchdogs called back.");
            if (exitnormally)
            {
                SaveSettingsState();
            }
            if (dbug)
            {
                MessageBox.Show(KNS.CurrentCatCode + "  " + KNS.EditKeyText, "KNS Window_Closing");
            }
        }

        private bool IsThereExcelKeyedNoteSystem()
        {
            if (KNS.Knfolder == null) { return false; }
            string searchPat = "*KeyNotes.xlsm";

            string[] files = Directory.GetFiles(@KNS.Knfolder, searchPat, SearchOption.TopDirectoryOnly);

            if (files.Count() > 0)
            {
                string msg = "This Excel file: \n\n" + files[0].ToString() + "\n\n";
                msg = msg + "is found in the vicinity. It looks like it might be an Excel RevitKeyNote file maker.";
                msg = msg + " Such a file typically stores the keynote information and then writes out the keynote table file.";
                msg = msg + " Using this keynote file manager will result in keynotes out of sync if the Excel";
                msg = msg + " is used later. If this file is actually Ok then rename or move the file to somewhere else so that it is not in the vicinity.";
                FormMsgWPF askthis = new FormMsgWPF("", 3, false);
                askthis.SetMsg("Aborting For Your Protection. This application will now quit.", msg);
                askthis.ShowDialog();
                exitnormally = false;
                return true;
            }
            return false;
        }

        private void SaveSettingsState()
        {
            Properties.Settings.Default.Manager_Top = Top;
            Properties.Settings.Default.Manager_Left = Left;
            Properties.Settings.Default.Manager_Height = Height;
            Properties.Settings.Default.Manager_Width = Width;
            Properties.Settings.Default.LastKNFile = KNS.Knfile;
            Properties.Settings.Default.LastRestoreCat = KNS.CurrentCatCode;
            Properties.Settings.Default.LastSelKey = KNS.EditKeyText;
            Properties.Settings.Default.Save();
            if (dbug)
            {
                MessageBox.Show(KNS.CurrentCatCode + "  " + KNS.EditKeyText, "KNS SaveSettingsState");
            }
        }

        public void DragWindow(object sender, MouseButtonEventArgs args)
        {
            // Watch out. Fatal error if not primary button!
            if (args.LeftButton == MouseButtonState.Pressed) { DragMove(); }
        }

        private string ThisSelectedFile(string thisfolder, string thisfile, string filter, string title)
        {
            string thisFilePath = thisfolder + thisfile;
            if (File.Exists(thisFilePath)) { return thisFilePath; }
            OpenFileDialog fd = new OpenFileDialog
            {
                Filter = filter,
                Title = title
            };


            if (Directory.Exists(thisfolder) && thisfolder.Length > 2)
            {
                fd.InitialDirectory = thisfolder;
            }

            var result = fd.ShowDialog();
            switch (result)
            {
                case true:
                    var file = fd.FileName;
                    return file.ToString();
                case false:
                default:
                    return null;
            }
        }

        private string EstablishKNFFileName(string atThisPath, string forThisFile)
        {   
            string thisKNF = ThisSelectedFile(atThisPath,
                forThisFile, @"Rvt UserKeynote files (*.txt)|*.txt",
                "Select the Revit UserKeynote File");

            if (thisKNF == null)
            {
                // assume user canceled, give opportunity to have the previous
                // file reopened.
                if (atThisPath != null) {
                    string orginalKNF = Path.Combine(atThisPath, forThisFile);
                    thisKNF = orginalKNF;
                } else {
                    return null;
                }
            }

            if (thisKNF.Equals(string.Empty)) { return null; }
            KNS.ReadOnly = IsFileReadOnly(thisKNF);
            KNS.Knfolder = Path.GetDirectoryName(thisKNF);
            KNS.Knfile = Path.GetFileName(thisKNF);
            Knfolder = Path.GetDirectoryName(thisKNF);
            Knfile = Path.GetFileName(thisKNF);
            return thisKNF;
        }

        // Returns wether a file is read-only.
        public static bool IsFileReadOnly(string FileName)
        {
            FileInfo fInfo = new FileInfo(FileName);
            return fInfo.IsReadOnly;
        }

        private void RemoveSuchStatusFromCategoryList(string status)
        {
            if (KNS.KndatacatState != null)
            {
                if (status != string.Empty)
                {
                    // find rows with matching status in the status column 
                    DataRow[] drs = KNS.KndatacatState.Select("Status = '" + status + "'");
                    if (drs != null)
                    {
                        foreach (DataRow d in drs)
                        {
                            KNS.KndatacatState.Rows.Remove(d);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Removes all notes from the notes datatable that are not currently reserved by thisuser
        /// </summary>
        /// <param name="thisUser"></param>
        private void RemoveAllNotesOfCategoryNotReservedBy(string thisUser)
        {
            if (KNS.KndataTable != null)
            {
                if (thisUser != null)
                {
                    // find rows NOT matching thisuser in the status column 
                    DataRow[] drcatstate = KNS.KndatacatState.Select("Status <> '" + thisUser + "'");
                    if (drcatstate != null)
                    {
                        foreach (DataRow d in drcatstate)
                        {
                            string notekey = d[0].ToString().Trim();
                            // find rows in notes table with notekeys matching the key
                            DataRow[] drnotes = KNS.KndataTable.Select("NoteKey = '" + notekey + "'");
                            if (drnotes != null)
                            {
                                foreach (DataRow noterow in drnotes)
                                {
                                    KNS.KndataTable.Rows.Remove(noterow);
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Adds unique to category datatable, returns false is table already contained
        /// the category code
        /// </summary>
        /// <param name="catcode"></param>
        /// <param name="catname"></param>
        /// <returns></returns>
        private bool AddedToCategoryList(string catcode, string catname, string status)
        {
            if (KNS.KndatacatState != null)
            {
                if (catcode != string.Empty)
                {
                    // find row with matching catcode in the key column to check if table already contains catcode
                    DataRow dr = KNS.KndatacatState.Select("Key = '" + catcode + "'").FirstOrDefault();
                    if (dr != null) { return false; }
                    DataRow drcatstate = KNS.KndatacatState.NewRow();
                    drcatstate[0] = catcode;
                    drcatstate[1] = catname;
                    drcatstate[2] = status;
                    KNS.KndatacatState.Rows.Add(drcatstate);
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Intended to be a universal update that loads or reloads the session while
        /// maintaining any user resevered category data
        /// </summary>
        /// <param name="thisKNF"></param>
        private void ReadUpdateKNTables(string thisKNF)
        {
            if (!File.Exists(thisKNF)) { return; }

            // Assuming this locks down the saved KNF so that others cannot mess with it while setups are first
            // made here prior to reading in the data.
            FileStream fileStream = null;
            try
            {
                fileStream = new FileStream(thisKNF, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (Exception er)
            {
                MainStatusMsg("Unable to refresh. Try again ... " + er.Message);
                return;
            }

            string thisUser = Environment.UserName.ToUpper();
            // All notes in the notes table that are not of a catefgory reserved by this user need to be removed
            // prior to loading their replacements from the KNF file about to be reloaded. Since this process relies
            // on the category table for the category keys to use, it must occur prior to cleaning the category list.
            RemoveAllNotesOfCategoryNotReservedBy(thisUser);
            // At this point the category table contains only the reserved catergories, by thisuser or by
            // others based on previous knowledge. The notes table contains only the user's reserved notes.
            // Remove all categories from status list that are currently available.
            RemoveSuchStatusFromCategoryList(Available);
            // At this point the category status list contains only those categories reserved by others or
            // reserved by thisuser.

            AddedToCategoryList("*", "ALL THE NOTES", "na");

            KNS.LastModified = File.GetLastWriteTime(thisKNF);
            using (var streamReader = new StreamReader(fileStream, Encoding.Default))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    string[] words = line.Split('\t');
                    if (words != null)
                    {
                        int qtyWords = words.Length;
                        // Use qtyWords to pick out the category headers. The datafile
                        // must be created this way to work right. 
                        //if (qtyWords == 2)  // assume this is a category header  // wrong assumption if there are extra tabs in data
                        if (qtyWords == 2 || (qtyWords > 1 && words[2].ToString().Equals(string.Empty))) // better
                        {
                            string catcode = string.Empty;
                            string catname = string.Empty;
                            if (words.Length > 0) // this would be the category code
                            {
                                catcode = words[0].ToString().Trim();
                            }
                            if (words.Length > 1) // this would be the category name
                            {
                                catname = words[1].ToString().Trim();
                            }
                            // The reserved categories are already in the list. This function
                            // only adds categories not in the list.
                            AddedToCategoryList(catcode, catname, Available);
                            continue;
                        }
                        // If we reach this point then words must be a keynote.
                        string notekey = string.Empty;
                        string notenote = string.Empty;
                        string notenotekeycode = string.Empty;
                        // read the note's notekey first, that is the category code
                        if (words.Length > 2)
                        {
                            notenotekeycode = words[2].ToString().Trim();
                        }
                        DataRow drcatstate = KNS.KndatacatState.Select("Key = '" + notenotekeycode + "'").FirstOrDefault();
                        if (drcatstate != null)
                        {
                            // Ignore all keynotes of category reserved by this session user. i.e. process
                            // i.e. Process only those not equal to thisuser
                            if (!drcatstate[2].Equals(thisUser))
                            {
                                if (words.Length > 0)
                                {
                                    notekey = words[0].ToString().Trim();
                                }
                                if (words.Length > 1)
                                {
                                    notenote = words[1].ToString().Trim();
                                }
                                DataRow dr = KNS.KndataTable.NewRow();
                                dr[0] = notekey;
                                dr[1] = notenote;
                                dr[2] = notenotekeycode;
                                KNS.KndataTable.Rows.Add(dr);
                            }
                        }
                    }
                }
            }
            KNS.KndataTable.DefaultView.Sort = "Key ASC";
            KNS.KndatacatState.DefaultView.Sort = "Key ASC";
            // reset the combobox selector to what is had been showing, but because this code
            // might be automatically run by the watcher outside of the ui thread we have to
            // use the dispacter invoker.
            if (Application.Current.Dispatcher.CheckAccess())
            {
                ComboCatCode.SelectedValue = KNS.CurrentCatCode;
            }
            else
            {
                Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
               new Action(() => ComboCatCode.SelectedValue = KNS.CurrentCatCode));
            }
            MainStatusMsg("Refreshed data from the master keynote file.");
            KNS.Tableisstale = false;
        }

        private void SetCounts()
        {
            if (blvnotes != null)
            {
                KNS.Notecount = "Showing " + blvnotes.Count.ToString() + " of " + KNS.KndataTable.Rows.Count.ToString();
            }
            else
            {
                KNS.Notecount = "Showing " + KNS.KndataTable.Rows.Count.ToString();
            }
        }

        private void NotesGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DataGrid datagrid = sender as DataGrid;
            AcknowledgeGridSelection(datagrid);
        }

        private void AcknowledgeGridSelection(DataGrid datagrid)
        {
            if (datagrid.SelectedItems.Count < 1) { return; }
            //Gets a list of the selected row indexes for the datatable bound to the datagrid.
            List<int> SelectedIndexes = datagrid
                              .SelectedItems
                              .Cast<DataRowView>()
                              .Select(view => KNS.KndataTable.Rows.IndexOf(view.Row))
                              .ToList();
            KNS.EditRowIndex = SelectedIndexes[0];

            if (datagrid.SelectedValue != null)
            {
                DataRowView dataRowView = (DataRowView)datagrid.SelectedItem;
                Inhibit = true;
                KNS.EditKeyText = dataRowView.Row.ItemArray[0].ToString();
                Inhibit = false;
                KNS.EditCellText = dataRowView.Row.ItemArray[1].ToString();
                KNS.EditCatText = dataRowView.Row.ItemArray[2].ToString();
                KNS.CurrentDRV = dataRowView;
            }
            else
            {
                KNS.EditKeyText = String.Empty;
                KNS.EditCellText = String.Empty;
                KNS.EditCatText = String.Empty;
                KNS.CurrentDRV = null;
            }
            SetEditingAndNewPrompts();
            Properties.Settings.Default.LastSelKey = KNS.EditKeyText;
            Properties.Settings.Default.Save();
        }

        private void SetEditingAndNewPrompts()
        {
            bool editingPossible = IsCellPossiblyEditableByMe(KNS.CurrentCatCode);
            if (editingPossible)
            {
                KNS.EditcellisReadOnly = false;
            }
            else
            {
                KNS.EditcellisReadOnly = true;
                MainStatusMsg(KNS.CurrentCatName + " is not editable by you.");
            }

            if (KNS.CurrentCatCode == "*")
            {
                if (KNS.EditRowIndex < 0)
                {
                    KNS.LablecontentsContents = "Edit Note Content Here (A Category Needs To Be Selected In Order To Add A Note.)";
                }
                else
                {
                    if (editingPossible)
                    { KNS.LablecontentsContents = "Edit Note " + KNS.EditKeyText + " Content Here"; }
                    else
                    { KNS.LablecontentsContents = KNS.EditKeyText + " Content"; }
                }
                KNS.ButtoncreatenewContent = "Add A New Note";
                if (Application.Current.Dispatcher.CheckAccess())
                {
                    ButtonCreateNew.IsEnabled = false;
                }
                else
                {
                    Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                   new Action(() => ButtonCreateNew.IsEnabled = false));
                }
            }
            else
            {
                if (KNS.EditRowIndex < 0)
                {
                    if (editingPossible)
                    { KNS.LablecontentsContents = "Add Content For A New " + KNS.CurrentCatName + " Note Here"; }
                    else
                    { KNS.LablecontentsContents = KNS.CurrentCatName + " is reserved. You cannot add."; }

                    KNS.ButtoncreatenewContent = "Add This New " + KNS.CurrentCatCode + " Note";
                }
                else
                {
                    if (editingPossible)
                    { KNS.LablecontentsContents = "Edit Note " + KNS.EditKeyText + " Content Here"; }
                    else
                    { KNS.LablecontentsContents = KNS.CurrentCatName + " is reserved. You cannot edit."; }

                    KNS.ButtoncreatenewContent = "Add " + KNS.CurrentCatCode + " Note";
                }
                // ButtonCreateNew.IsEnabled = editingPossible;
                if (Application.Current.Dispatcher.CheckAccess())
                {
                    ButtonCreateNew.IsEnabled = editingPossible;
                }
                else
                {
                    Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                   new Action(() => ButtonCreateNew.IsEnabled = editingPossible));
                }
            }
        }

        private void ComboCatCode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // seems odd but has evolved from large structural changes
            ComboBox cb = sender as ComboBox;
            if (cb.SelectedItem is DataRowView drv)
            {
                KNS.CurrentCatCode = drv[0].ToString().Trim();
                KNS.CurrentCatName = drv[1].ToString().Trim();
                ApplyCurrentSelectionFilters();
                SetReservedButtonContent();
                SetCommitButton(); // will also unset IsReserved but not unset InReserve
                SetReplacementButton(true);
                try
                {
                    // If no row has been selected yet, then set the selection to the first row in the 
                    // notes list for the newly changed category. But if there had been a selected row and
                    // that selection category is what has now been selected to be the category showing, then
                    // reselect that same row.


                    if (blvnotes.Count > 0)
                    {

                        if (KNS.CurrentDRV == null)
                        {
                            KNS.CurrentDRV = blvnotes[0] as DataRowView;
                        }
                        else if (KNS.CurrentDRV.Row.RowState == DataRowState.Detached)
                        {
                            KNS.CurrentDRV = blvnotes[0] as DataRowView;
                        }
                        else if (!KNS.CurrentDRV[2].ToString().Equals(((DataRowView)blvnotes[0])[2].ToString()))
                        {
                            KNS.CurrentDRV = blvnotes[0] as DataRowView;
                        }
                        SelectThisDataGridRow(NotesGrid, KNS.CurrentDRV);
                    }
                    else
                    {
                        MainStatusMsg("No " + KNS.CurrentCatName + " notes yet.");
                    }
                }
                catch (Exception erh)
                {
                    MainStatusMsg("Error caught at ComboCatCode_SelectionChanged: " + erh.Message);
                }
                ColorEditTextBoxAccordingToAccess();
            }
        }

        private void ApplyCurrentSelectionFilters()
        {
            if (KNS.CurrentCatCode == null || KNS.FindThisText == null) { return; }
            //http://www.csharp-examples.net/dataview-rowfilter/
            //blv.Filter = "Note LIKE '*" + KeyNoteSession.FindThisText + "*' AND NoteKey = '" + KeyNoteSession.CurrentCatCode + "'";
            string findthis = KNS.FindThisText;
            if (KNS.ReplaceAsWord)
            {
                findthis = space + findthis + space;
            }
            if (KNS.CurrentCatCode == "*")
            {
                if (KNS.FindThisText.Equals(string.Empty))
                {
                    blvnotes.RemoveFilter();
                }
                else
                {
                    blvnotes.Filter = "Note LIKE '*" + findthis + "*'";
                }
            }
            else
            {
                // Note: Filter supports "DataColumn.Expression" such as "NoteKey = 'M'"
                // or  //blv.Filter = "NoteKey LIKE  '*" + KeyNoteSession.CurrentCatCode + "*'"; 
                if (KNS.FindThisText.Equals(string.Empty))
                {
                    blvnotes.Filter = "NoteKey = '" + KNS.CurrentCatCode + "'";
                }
                else
                {
                    blvnotes.Filter = "Note LIKE '*" + findthis + "*' AND NoteKey = '" + KNS.CurrentCatCode + "'";
                }
            }
            // The edit area needs to be cleared but first set a -1 flag so that the cleared state doee not
            // overwrite whateber the text that was there.
            KNS.EditRowIndex = -1;
            KNS.EditKeyText = string.Empty;
            KNS.EditCellText = string.Empty;
            KNS.EditCatText = string.Empty;
            SetEditingAndNewPrompts();
            SetCounts();
            ReportOnSearch();
        }

        private void ReportOnSearch()
        {
            if (blvnotes != null)
            {
                if (KNS.FindThisText.Length > 0)
                {
                    if (blvnotes.Count == 0 && KNS.ReplaceAsWord)
                    {
                        KNS.StatusMSGSettingsA = "Nothing? Keep in mind the search is for 'whole word'.";
                        return;
                    }
                    if (blvnotes.Count != 0 && KNS.ReplaceAsWord)
                    {
                        KNS.StatusMSGSettingsA = "Thinking you got them all? 'Whole word' is not smart enough yet to pick up 'the_word,' or 'the_word' plural.";
                        return;
                    }
                }
                KNS.StatusMSGSettingsA = string.Empty;
            }
        }

        private void EditCell_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Using -1 as a flag that the datatable should not be updated with the edit contents. 
            if (KNS.EditRowIndex >= 0)
            {
                if (KNS.KndataTable.Rows.Count > 0)
                {
                    DataRow dr = KNS.KndataTable.Rows[KNS.EditRowIndex];
                    dr[1] = KNS.EditCellText;
                }
            }
        }

        private void AddToDictionaryPending(DataRow dr)
        {
            string catcode = dr[2].ToString();
            string note_numb = dr[0].ToString();
            List<string> note_numb_list;

            if (!dictionaryPending.TryGetValue(catcode, out note_numb_list))
            {
                dictionaryPending.Add(catcode, note_numb_list = new List<string>());
            }
            if (!note_numb_list.Contains(note_numb))
            {
                note_numb_list.Add(note_numb);
            }

            Reportdictpendingstatus("Added to dictionary, stat: ");
        }

        private void Reportdictpendingstatus(string mode)
        {
            string msg = string.Empty;
            foreach (KeyValuePair<string, List<string>> entry in dictionaryPending)
            {
                msg = msg + " " + entry.Key + "|" + entry.Value.Count().ToString();

            }
            //MainStatusMsg(mode  + msg);
        }

        private void RemoveFromDictionaryPending(string catcode)
        {
            dictionaryPending.Remove(catcode);
            Reportdictpendingstatus("Removed from dictionary, stat: ");
        }

        private void EditCell_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            ColorEditTextBoxAccordingToAccess();

            // Using -1 as a flag that the datatable should not be updated with the edit contents. 
            if (KNS.EditRowIndex >= 0)
            {
                if (KNS.KndataTable.Rows.Count > 0)
                {
                    DataRow dr = KNS.KndataTable.Rows[KNS.EditRowIndex];
                    AddToDictionaryPending(dr);
                }
            }
        }

        private void ColorEditTextBoxAccordingToAccess()
        {
            TextBox tb = EditCell;
            {
                if (!KNS.CurrentCatCode.Equals(KNS.CategoryInReserve, StringComparison.CurrentCultureIgnoreCase))
                {
                    if (Application.Current.Dispatcher.CheckAccess())
                    {
                        tb.Foreground = Brushes.Red;
                    }
                    else
                    {
                        Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                       new Action(() => tb.Foreground = Brushes.Red));
                    }
                }
                else
                {
                    if (Application.Current.Dispatcher.CheckAccess())
                    {
                        tb.Foreground = Brushes.DarkGreen;
                    }
                    else
                    {
                        Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                       new Action(() => tb.Foreground = Brushes.DarkGreen));
                    }
                }
            }
        }

        private void NotesGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (IsCellPossiblyEditableByMe(KNS.CurrentCatCode))
            {
                SendFocusToEditCell();
            }
            else
            {
                MainStatusMsg(KNS.CurrentCatName + " is not editable by you.");
            }
        }

        private void TboxFindThisText_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyCurrentSelectionFilters();
        }

        private void ButtonClearFindThisText_Click(object sender, RoutedEventArgs e)
        {
            KNS.FindThisText = string.Empty;
            ApplyCurrentSelectionFilters();
        }

        private void ButtonCreateNew_Click(object sender, RoutedEventArgs e)
        {
            AddANewNote();
        }

        private void AddANewNote()
        {
            // Divide the notekey into three parts: <catcode><static string><number>
            // For example: FP-1234-002
            // FP = catcode
            // static string = -1234-
            // number = 002
            string keycatcode = KNS.CurrentCatCode;
            string keystatictext = KNS.Keystatictext;
            string keystrnumpart = string.Empty;
            string keynumberformat = KNS.Keynumberformat;
            Single keyincval = KNS.Keyincval;

            DataView dv = new DataView(KNS.KndataTable)
            {
                RowFilter = "NoteKey =  '" + KNS.CurrentCatCode + "'",
                Sort = "Key ASC"
            };

            // Finding the first available new number using the scheme
            Single newnumberpart = 0;

            string newnotenumber;
            do
            {
                newnumberpart = newnumberpart + keyincval;
                newnotenumber = keycatcode + keystatictext + newnumberpart.ToString(keynumberformat);
            } while (dv.Find(newnotenumber) >= 0);

            DataRow dr = KNS.KndataTable.NewRow();
            dr[0] = newnotenumber;
            if (KNS.EditRowIndex == -1)
            {
                dr[1] = KNS.EditCellText;
            }
            else
            {
                dr[1] = string.Empty;
            }
            dr[2] = keycatcode;
            KNS.KndataTable.Rows.Add(dr);
            KNS.EditRowIndex = KNS.KndataTable.Rows.IndexOf(dr);
            KNS.EditKeyText = dr[0].ToString();
            KNS.EditCellText = dr[1].ToString();
            KNS.EditCatText = dr[2].ToString();
            NotesGrid.SelectedIndex = KNS.EditRowIndex;
            if (NotesGrid.SelectedItem != null) { NotesGrid.ScrollIntoView(NotesGrid.SelectedItem); }
            SendFocusToEditCell();
            SetEditingAndNewPrompts();
            SetCounts();
        }

        private void ButtonReplaceWithThisText_Click(object sender, RoutedEventArgs e)
        {
            string findthis = KNS.FindThisText;
            string replacewith = KNS.ReplaceWithThisText;
            string catcode = KNS.CurrentCatCode;

            if (findthis.Equals(string.Empty)) { return; }

            if (replacewith.Equals(string.Empty))
            {
                string msg = "Are you sure you want to replace the text '" + findthis + "' with nothing?";

                FormMsgWPF askthis = new FormMsgWPF("", 2, false);
                askthis.SetMsg("Replace with nothing?", msg, "");
                askthis.ShowDialog();
                if (askthis.TheResult != MessageBoxResult.OK)
                {
                    return;
                }
            }

            if (KNS.ReplaceAsWord)
            {
                findthis = space + findthis + space;
                if (!replacewith.Equals(string.Empty))
                {
                    replacewith = space + replacewith + space;
                }
                else
                {
                    replacewith = string.Empty;
                }
            }

            if (!findthis.Equals(string.Empty))
            {
                DataView dv = new DataView(KNS.KndataTable)
                {
                    RowFilter = "Note LIKE '*" + findthis + "*' AND NoteKey = '" + catcode + "'",
                    Sort = "Key ASC"
                };

                foreach (DataRowView rdv in dv)
                {
                    string revisednotepart = rdv[1].ToString().Replace(findthis, replacewith);
                    rdv[1] = revisednotepart.Trim();
                    AddToDictionaryPending(rdv.Row);
                }
                KNS.FindThisText = KNS.ReplaceWithThisText;
            }
        }

        private void ButtonCLRReplaceWithThisText_Click(object sender, RoutedEventArgs e)
        {
            KNS.ReplaceWithThisText = string.Empty;
        }

        private void ToggleWholeWord_Click(object sender, RoutedEventArgs e)
        {
            ApplyCurrentSelectionFilters();
            SendFocusToEditCell();
        }

        private void ButtonkeyNoteFile_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            SaveSettingsState();
            InitialzeTheKNS();
            RestoreLastSelections();
        }
        private void MainStatusMsg(string theMSG, bool wipeClean = false)
        {
            if (theLastStatusMSG.Equals(theMSG)) { return; }
            string thisUser = Environment.UserName.ToUpper();
            if (!theMSG.StartsWith(thisUser))
            {
                if (wipeClean)
                {
                    KNS.StatusMSG = "- " + theMSG + " " + DateTime.Now.ToLongTimeString();
                }
                else
                {
                    KNS.StatusMSG = "- " + theMSG + " " + DateTime.Now.ToLongTimeString() + "\n" + KNS.StatusMSG;
                }
            }
            theLastStatusMSG = theMSG;
        }

        private void ButtonCheckOutCategory_Click(object sender, RoutedEventArgs e)
        {
            // under construction

            // Don't allow if there are uncommited changes in another category.

            if (KNS.CategoryIsReserved == false && KNS.CategoryInReserve.Equals(string.Empty))  // nothing reserved, user wants to reserve
            {
                // Don't allow if category is not selected.
                if (KNS.CurrentCatCode == "*")
                {
                    MainStatusMsg("Not allowed to Bogart ALL the categories!");
                    return;
                }

                KNS.CategoryInReserve = KNS.CurrentCatCode;  // need to now make lock file
                KNS.CategoryInReserveName = KNS.CurrentCatName;
                RecordCategoryReserveStateChange(KNS.CurrentCatCode, Environment.UserName.ToUpper());

                KNS.CategoryIsReserved = true;  // commit button IsEnabled is bound to this.
                                                //ButtonWriteOutCategory.IsEnabled = true;
                MainStatusMsg(KNS.CategoryInReserveName
                    + " reserved. Others locked out until you commit or relinquish.");
            }
            else  // user wants to unreserve the category
            {
                RecordCategoryReserveStateChange(KNS.CategoryInReserve, Available);
                KNS.CategoryInReserve = string.Empty;
                KNS.CategoryIsReserved = false; // commit button IsEnabled is bound to this.
                MainStatusMsg(KNS.CategoryInReserveName + " given up. Others can make changes to this category.");
                KNS.CategoryInReserveName = string.Empty;
            }
            RegisterStatusAsKeyNoteUser();  // updates this user's intentions for others to see.
            SetReservedButtonContent();
            //SendFocusToEditCell();
            ColorEditTextBoxAccordingToAccess();
        }

        private void SetReservedButtonContent()
        {
            if (KNS.CategoryIsReserved || KNS.CategoryInReserve.Length > 0)
            {
                ButtonCheckOutCategory.Content = "Give up " + KNS.CategoryInReserveName + " Category";
                ButtonCheckOutCategory.IsEnabled = true;
            }
            else
            {
                if (KNS.CurrentCatCode == "*")
                {
                    ButtonCheckOutCategory.Content = "Reserve Category";
                    ButtonCheckOutCategory.IsEnabled = false;
                }
                else
                {
                    // Don't allow if someone else has reserved the category.
                    string proposedReservation = KNS.CurrentCatCode;
                    if (proposedReservation != "*")
                    {
                        string catStatus = GetCatReservationStatus(proposedReservation);
                        if (catStatus != Available)
                        {
                            MainStatusMsg(catStatus + " has " + KNS.CurrentCatName + " reserved. You cannot reserve.");
                            ButtonCheckOutCategory.IsEnabled = false;
                        }
                        else
                        {
                            ButtonCheckOutCategory.IsEnabled = true;
                        }
                    }
                    ButtonCheckOutCategory.Content = "Reserve " + KNS.CurrentCatName + " Category";
                }
            }
        }

        private void SetCommitButton()
        {
            // This is for the sole purpose to dim the commit button when the user has a category reserved
            // but is currently NOT looking at that category. This avoids making a nonreserved edit and thinking
            // the commit button would write that change out.
            if (KNS.CategoryInReserve != null && KNS.CurrentCatCode != null)
            {
                KNS.CategoryIsReserved = KNS.CategoryInReserve.Equals(KNS.CurrentCatCode);
            }
        }

        private void SendFocusToEditCell()
        {
            EditCell.Focus();
        }

        private void WriteCommitTheReserveCategory()
        {
            if (KNS.Knfolder == null) { return; }
            string thisKNF = Path.Combine(KNS.Knfolder, KNS.Knfile);
            if (!File.Exists(thisKNF)) { return; }

            // 1) opens and locks the main KNF file
            FileStream fsoriginalKNF = new FileStream(thisKNF, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

            // 2) create random name to be used temporarily for the the new KNF file
            string tempnewKNFname = Path.Combine(KNS.Knfolder, Path.GetRandomFileName());
            FileStream fsnewKNF = new FileStream(tempnewKNFname, FileMode.Create, FileAccess.ReadWrite, FileShare.Write);

            // 3) read original line by line 
            // 4) write the new line by line
            using (var streamWriterNew = new StreamWriter(fsnewKNF, Encoding.Default))
            {
                using (var streamReaderOriginal = new StreamReader(fsoriginalKNF, Encoding.Default))
                {
                    // first write header for the category in reserve
                    // The edited category header will be the top one.
                    string thisCategoryInReserveline = KNS.CategoryInReserve.Trim() + '\t' + KNS.CategoryInReserveName.Trim();
                    streamWriterNew.WriteLine(thisCategoryInReserveline);

                    // If there are other new categories, they should be written now regardless to avoid having to create
                    // & commit these one at a time, but not the one in reserve. It was just written.
                    foreach (string thisCat in KNS.NewCategories.Where(item => !thisCategoryInReserveline.Equals(item)))
                    {
                        streamWriterNew.WriteLine(thisCat);
                    }

                    // Now read the original file.
                    string line;
                    while ((line = streamReaderOriginal.ReadLine()) != null)
                    {
                        // MessageBox.Show(line);
                        string[] words = line.Split('\t');
                        if (words != null)
                        {
                            int qtyWords = words.Length;
                            // Use qtyWords to pick out the category headers. The datafile
                            // must be created this way to work right. 
                            if (qtyWords == 2)  // assume this is a category header
                            {
                                string preexistingcatcode = string.Empty;
                                string preexistingcatname = string.Empty;
                                if (words.Length > 0)
                                {
                                    preexistingcatcode = words[0].ToString().Trim();
                                }
                                if (words.Length > 1)
                                {
                                    preexistingcatname = words[1].ToString().Trim();
                                }
                                // for the time being we are reproducing all preexisting categories except
                                // for the categoryinresever
                                if (!preexistingcatcode.Equals(KNS.CategoryInReserve))
                                {
                                    // But also do not duplicate any new ones if somehow they sneaked into
                                    // the table.
                                    if (!KNS.NewCategories.Contains(line, StringComparer.CurrentCultureIgnoreCase))
                                    {
                                        streamWriterNew.WriteLine(line.Trim());
                                    }
                                }
                                continue;
                            }
                            // If we reach this point then words must be a keynote.
                            // Only its category code need to be examined.
                            string linesCatCode = string.Empty;
                            if (words.Length > 2)
                            {
                                linesCatCode = words[2].ToString();
                                // reproducing only what is not in reserve
                                if (!linesCatCode.Equals(KNS.CategoryInReserve))
                                {
                                    streamWriterNew.WriteLine(line.Trim());
                                }
                            }
                        }
                    }
                    // At this point all of the original knf that is to be retained has been reproduced.
                    // Now is time to write the catergoryinreserve material.
                    DataView dv = new DataView(KNS.KndataTable)
                    {
                        RowFilter = "NoteKey =  '" + KNS.CurrentCatCode + "'",
                        Sort = "Key ASC"
                    };
                    foreach (DataRowView rdv in dv)
                    {
                        string noteline = rdv[0].ToString().Trim() + '\t' + rdv[1].ToString().Trim() + '\t' + rdv[2].ToString().Trim();
                        //MessageBox.Show(noteline);
                        streamWriterNew.WriteLine(noteline);
                    }
                }
            }
            // 5) create timestamp
            string timestamp = DateTime.Now.Year.ToString("0000")
                             + DateTime.Now.Month.ToString("00")
                             + DateTime.Now.Day.ToString("00")
                             + DateTime.Now.Minute.ToString("00")
                             + DateTime.Now.Second.ToString("00");
            string thisTimeStampedMovedKNSName = Path.GetFileNameWithoutExtension(thisKNF) + "_" + timestamp + Path.GetExtension(thisKNF);
            // 6) create stash folder
            string priorsPath = Path.Combine(KNS.Knfolder, priorsfolder);
            DirectoryInfo di = Directory.CreateDirectory(priorsPath);
            // 7) old name with timestamp in name
            thisTimeStampedMovedKNSName = Path.Combine(priorsPath, thisTimeStampedMovedKNSName);
            //MessageBox.Show(thisTimeStampedMovedKNSName);
            // 8) move/rename old to stash folder
            File.Move(thisKNF, thisTimeStampedMovedKNSName);
            // 9) rename new as kn file
            File.Move(tempnewKNFname, thisKNF);
            fsoriginalKNF.Dispose();
            fsnewKNF.Dispose();
            MainStatusMsg(KNS.CategoryInReserveName + " changes commited. pwrball-" + DateTime.Now.Millisecond.ToString("000"));
            KNS.LastModified = DateTime.Now;
            RemoveFromDictionaryPending(KNS.CurrentCatCode);
            // also clear the NewCategories list
            KNS.NewCategories.Clear();
        }

        private void ButtonWriteOutCategory_Click(object sender, RoutedEventArgs e)
        {
            if (KNS.CurrentCatCode == "*")
            {
                string msg = "To allow simultaneous use, you are not permitted to reserve only one category at a time.";
                msg = msg + " You need to set and reserve that category before commiting changes you made.";
                MainStatusMsg(msg);
                return;
            }
            try
            {
                WriteCommitTheReserveCategory();
            }
            catch (Exception err)
            {
                MainStatusMsg("Unable to CommitTheReserveCategory: " + err.Message);
            }
        }

        private void RegisterStatusAsKeyNoteUser()
        {
            if (KNS == null) { return; }
            if (KNS.Knfolder == null) { return; }
            string username = Environment.UserName.ToUpper();
            string thisRKU = MakeRKUName(username);
            int tryThisManyTimes = 10;
            int attempts = 1;
            while (attempts <= tryThisManyTimes)
            {
                try
                {
                    FileStream fsthisRKU = new FileStream(thisRKU, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                    using (var streamWriterNew = new StreamWriter(fsthisRKU, Encoding.Default))
                    {
                        streamWriterNew.WriteLine("US:" + username);  // users login
                        streamWriterNew.WriteLine("TM:" + @DateTime.Now); // time now
                        streamWriterNew.WriteLine("KF:" + @KNS.Knfile); // knf name
                        streamWriterNew.WriteLine("RE:" + KNS.CategoryIsReserved.ToString()); // a category is resevered
                        if (KNS.CategoryIsReserved)  // category code if one is being reserved
                        {
                            streamWriterNew.WriteLine("CC:" + KNS.CategoryInReserve);
                            streamWriterNew.WriteLine("CN:" + KNS.CategoryInReserveName);
                        }
                    }
                    break;
                }
                catch (Exception e)
                {
                    MainStatusMsg("Unable to RegisterStatusAsKeyNoteUser on attempt " + attempts.ToString() + " : " + e.Message);
                    attempts++;
                }
            }
        }

        private void UnRegisterStatusAsKeyNoteUser()
        {
            if (KNS == null) { return; }
            if (KNS.Knfolder == null) { return; }
            string username = Environment.UserName.ToUpper();
            string thisRKU = MakeRKUName(username);
            int tryThisManyTimes = 10;
            int attempts = 1;
            while (attempts <= tryThisManyTimes)
            {
                try
                {
                    if (File.Exists(thisRKU))
                    {
                        File.Delete(thisRKU);
                    }
                    break;
                }
                catch (Exception err)
                {
                    MessageBox.Show("Not able to delete the Revit note user file " + thisRKU + " on attempt " + attempts.ToString() + "\n\n" + err.Message);
                    attempts++;
                }
            }
        }

        private string MakeRKUName(string username)
        {
            string rootKNSname = Path.GetFileNameWithoutExtension(KNS.Knfile) + "_";
            string thisRKU = Path.Combine(KNS.Knfolder, rootKNSname + username + extRKU);
            return thisRKU;
        }

        private void AWatchedFileChanged(string theFile)
        {
            string filename = theFile;
            string fext = Path.GetExtension(filename);
            // filter for desired file types 
            if (Regex.IsMatch(fext, @"\.txt|\" + extRKU, RegexOptions.IgnoreCase))
            {
                string fn = Path.GetFileNameWithoutExtension(filename);
                string rootKNSname = Path.GetFileNameWithoutExtension(KNS.Knfile) + "_";

                if (fext.Equals(extRKU) && (fn.Contains(rootKNSname)))
                {
                    UserStats userStats = new UserStats(Path.Combine(KNS.Knfolder, filename));
                    if (KNS.Knfile.Equals(userStats.Kf))
                    {
                        if (userStats.Re == "True")
                        {
                            RecordCategoryReserveStateChange(userStats.Cc, userStats.Us);
                            MainStatusMsg(userStats.Us + " reserved " + userStats.Cn);
                        }
                        else
                        {
                            RecordCategoryReserveStateChange(string.Empty, userStats.Us);
                            MainStatusMsg(userStats.Us + " relinquished and is here.");
                        }
                    }
                    SetEditingAndNewPrompts();
                    KNS.OthersHaveReserved = AreOthersReserved();
                    return;
                }
                if (fn.Equals(Path.GetFileNameWithoutExtension(KNS.Knfile), StringComparison.CurrentCultureIgnoreCase))
                {
                    MainStatusMsg(KNS.Knfile + " changed.");

                    if (KNS.OthersHaveReserved)
                    {
                        MainStatusMsg("The keynote file changed and others have parts reserved. You should refresh.");
                    }
                    KNS.Tableisstale = true;
                    ReloadAndRestore();
                    return;
                }
            }
        }

        private void AWatchedFileDeleted(string filename)
        {
            if (KNS.Knfile == null) { return; }
            string fext = Path.GetExtension(filename);
            // filter for desired file types - In this case RKU only
            if (Regex.IsMatch(fext, @"\" + extRKU, RegexOptions.IgnoreCase))
            {
                string fn = Path.GetFileNameWithoutExtension(filename);
                string rootKNSname = Path.GetFileNameWithoutExtension(KNS.Knfile) + "_";

                if (fn.Contains(rootKNSname))
                {
                    string user = fn.Substring(rootKNSname.Length);
                    RecordCategoryReserveStateChange(string.Empty, user);
                    MainStatusMsg(user + " exited.");
                    SetEditingAndNewPrompts();
                    KNS.OthersHaveReserved = AreOthersReserved();
                    return;
                }
            }
        }

        private void AWatchedFileCreated(string filename)
        {
            string fext = Path.GetExtension(filename);

            // filter for desired file types - In this case RKU only
            if (Regex.IsMatch(fext, @"\" + extRKU, RegexOptions.IgnoreCase))
            {
                string fn = Path.GetFileNameWithoutExtension(filename);
                string rootKNSname = Path.GetFileNameWithoutExtension(KNS.Knfile) + "_";

                if (fn.Contains(rootKNSname))
                {
                    string user = fn.Substring(rootKNSname.Length);
                    MainStatusMsg(user + " is here.");
                    return;
                }
            }
        }

        private void TboxReplaceWithThisText_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetReplacementButton(true);
        }

        private void SetReplacementButton(bool withMsg)
        {
            if (KNS.ReplaceWithThisText == null || KNS.CurrentCatCode == null) { return; }
            if (KNS.ReplaceWithThisText.Length > 0 && KNS.CurrentCatCode == "*")
            {
                if (withMsg)
                {
                    KNS.StatusMSGSettingsB = "For safety, replace works only for a selected category.";
                }
                KNS.ReplacementAllowed = false;
            }
            else
            {
                if (withMsg)
                {
                    KNS.StatusMSGSettingsB = String.Empty;
                }
                KNS.ReplacementAllowed = true;
            }
        }

        private bool AreOthersReserved()
        {
            if (KNS.KndatacatState != null)
            {
                string thisUser = Environment.UserName.ToUpper();
                int countThisUser = 0;
                int countAvailable = 0;
                int countOfAllCategories = KNS.KndatacatState.Rows.Count - 1;

                DataRow[] druser = KNS.KndatacatState.Select("Status = '" + thisUser + "'");
                if (druser != null)
                {
                    countThisUser = druser.Count();
                }

                DataRow[] dravail = KNS.KndatacatState.Select("Status = '" + Available + "'");
                if (dravail != null)
                {
                    countAvailable = dravail.Count();
                }

                return (countOfAllCategories - countAvailable - countThisUser > 0);
            }
            return false;
        }

        private void RecordCategoryReserveStateChange(string catcode, string who)
        {
            if (KNS.KndatacatState != null)
            {
                if (catcode != string.Empty)
                {
                    // find row with matching catcode in the key column and drop username into status column
                    DataRow dr = KNS.KndatacatState.Select("Key = '" + catcode + "'").FirstOrDefault();
                    if (dr != null) { dr[2] = who; }
                    // If this is the current category, then the reserve button must be disabled.
                    if ( //omg!
                        KNS.CurrentCatCode.Equals(catcode)
                        &&
                        (!who.Equals(Environment.UserName.ToUpper()))
                        &&
                        (!who.Equals(Available))
                        )
                    {
                        Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                       new Action(() => ButtonCheckOutCategory.IsEnabled = false));
                    }
                }
                else
                {
                    // find row with matching username in the status column and drop available in the status column.
                    // Taking advantage of user only able to check one category at a time.
                    DataRow dr = KNS.KndatacatState.Select("Status = '" + who + "'").FirstOrDefault();
                    if (dr != null)
                    {
                        dr[2] = Available;
                        if (KNS.CurrentCatCode.Equals(dr[0]))
                        {
                            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                           new Action(() => ButtonCheckOutCategory.IsEnabled = true));
                        }
                    }
                }
            }
        }

        private string GetCatReservationStatus(string catcode)
        {
            if (KNS.KndatacatState != null)
            {
                if (catcode != string.Empty)
                {
                    // find row with matching catcode in the key column and return status column value
                    DataRow dr = KNS.KndatacatState.Select("Key = '" + catcode + "'").FirstOrDefault();
                    if (dr != null) { return dr[2].ToString(); }
                }
            }
            return string.Empty;
        }

        private bool IsCellPossiblyEditableByMe(string catcode)
        {
            string status = string.Empty;
            if (KNS.KndatacatState != null)
            {
                if (catcode != string.Empty)
                {
                    // find row with matching catcode in the key column and return status column value
                    DataRow dr = KNS.KndatacatState.Select("Key = '" + catcode + "'").FirstOrDefault();
                    if (dr != null) { status = dr[2].ToString(); }
                }
            }
            if (status.Equals(Available) || status.Equals(Environment.UserName.ToUpper()))
            {
                KNS.EditcellisReadOnly = false;
                return true;
            }
            return false;
        }

        private void BwKNFFolderWatcher_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            KNFFolderWatcherArgs KNFArgs = e.Argument as KNFFolderWatcherArgs;

            string searchPat = "*";
            string rootKNSname = Path.GetFileNameWithoutExtension(KNFArgs.Knffilemarker);
            string foldertowatch = KNFArgs.Folderpath;
            string thisUserTag = KNFArgs.Usertag;
            Dictionary<string, DTI> ThisKNFActivity = new Dictionary<string, DTI>();

            while (worker.CancellationPending != true)
            //while (MessageBox.Show("Continue?", "In Loop", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                // collect all files for this keynote except for the user's RKU
                var files = from file in Directory.GetFiles(@foldertowatch, searchPat, SearchOption.TopDirectoryOnly)
                            let fileName = Path.GetFileNameWithoutExtension(file)
                            where (fileName.Contains(rootKNSname) && !fileName.Contains(thisUserTag))
                            select file;

                foreach (string f in files)
                {
                    string fname = Path.GetFileName(f);
                    if (fname.EndsWith(".txt", StringComparison.CurrentCultureIgnoreCase) || fname.EndsWith(extRKU, StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (!ThisKNFActivity.ContainsKey(fname))
                        {
                            // Add file as first seen
                            DTI fDTI = new DTI
                            {
                                DateTime = File.GetLastWriteTime(f),
                                Status = 3 // mark as file not seen before
                            };
                            ThisKNFActivity.Add(fname, fDTI);
                        }
                        else
                        {
                            // File already exists. Update record if need be.
                            if (ThisKNFActivity.TryGetValue(fname, out DTI lastdti))
                            {
                                DateTime thisfdatetime = File.GetLastWriteTime(f);
                                if (lastdti.DateTime < thisfdatetime)
                                {
                                    // File seems to be newer. Update status.
                                    lastdti.Status = 1;
                                }
                                else
                                {
                                    // File seems to be unchanged.
                                    lastdti.Status = 0;
                                }
                                lastdti.DateTime = thisfdatetime;
                            }
                        }
                    }
                }
                // At this point the dictionary holds record of the last polled files. Now report on any new files, any 
                // changed files, and any deleted files. Once reported change all entry status to 2. Therefore any entries
                // of value 2 had not been seen again and therefore were files deleted. Those same deleted entries are removed
                // from the dictionary onnce having been reported.
                Dictionary<string, string> KNStatChanges = new Dictionary<string, string>();
                foreach (KeyValuePair<string, DTI> entry in ThisKNFActivity)
                {
                    DTI thisentrydti = entry.Value;
                    if (thisentrydti.Status == 1)
                    {
                        // This one was newer. Therefore it had been updated.
                        KNStatChanges.Add(entry.Key, "Upd");
                        continue;
                    }
                    if (thisentrydti.Status == 2)
                    {
                        // This one was not seen. Therefore it was deleted.
                        KNStatChanges.Add(entry.Key, "Del");
                        continue;
                    }
                    if (thisentrydti.Status == 3)
                    {
                        // This one is brand new. Therefore it is new.
                        KNStatChanges.Add(entry.Key, "New");
                    }
                }

                // Delete entries not seen on this pass. Do this before marking the remainders as Status 2. 
                var itemsToRemove = ThisKNFActivity.Where(f => f.Value.Status == 2).ToArray();
                foreach (var item in itemsToRemove)
                {
                    ThisKNFActivity.Remove(item.Key);
                }

                // Now mark the remaining as status 2
                foreach (KeyValuePair<string, DTI> entry in ThisKNFActivity)
                {
                    entry.Value.Status = 2;
                }

                // report changes
                foreach (KeyValuePair<string, string> changed in KNStatChanges)
                {
                    int progcode;
                    switch (changed.Value)
                    {
                        case "Upd":
                            progcode = 10;
                            worker.ReportProgress(progcode, changed.Key);
                            break;
                        case "Del":
                            progcode = 20;
                            worker.ReportProgress(progcode, changed.Key);
                            break;
                        case "New":
                            progcode = 30;
                            worker.ReportProgress(progcode, changed.Key);
                            break;
                        default:
                            break;
                    }
                }
                System.Threading.Thread.Sleep(KNFArgs.Pollmsec);
            } // end while

            if ((worker.CancellationPending == true))
            {
                e.Cancel = true;
            }
            e.Result = KNFArgs;
        }

        private void BwKNFFolderWatcher_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e != null)
            {
                int stat = e.ProgressPercentage;
                string fn = e.UserState.ToString();
                if (stat == 10)
                {
                    // This one was newer. Therefore it had been updated.
                    // MainStatusMsg(fn + " => " + "Upd-P");
                    AWatchedFileChanged(fn);
                }
                if (stat == 20)
                {
                    // This one was not seen. Therefore it was deleted.
                    //MainStatusMsg(fn + " => " + "Del-P");
                    AWatchedFileDeleted(fn);
                }
                if (stat == 30)
                {
                    // This one is brand new. Therefore it is new.
                    //MainStatusMsg(fn + " => " + "New-P");
                    AWatchedFileCreated(fn);
                }
            }
        }

        private void SetUpBackgroundWorkers()
        {
            bwKNFFolderWatcher.WorkerSupportsCancellation = true;
            bwKNFFolderWatcher.WorkerReportsProgress = true;
            bwKNFFolderWatcher.DoWork += new DoWorkEventHandler(BwKNFFolderWatcher_DoWork);
            bwKNFFolderWatcher.ProgressChanged += new ProgressChangedEventHandler(BwKNFFolderWatcher_ProgressChanged);

            bkFileLoadingWorker.DoWork += BkFileLoadingWorker_DoWork;
            bkFileLoadingWorker.RunWorkerCompleted += BkFileLoadingWorker_RunWorkerCompleted;
        }

        public class DTI
        {
            public DateTime DateTime { get; set; }
            public int Status { get; set; }
        }

        public class KNFFolderWatcherArgs
        {
            public string Folderpath { get; set; }
            public string Knffilemarker { get; set; }
            public string Usertag { get; set; }
            public int Pollmsec { get; set; }
        }

        private void ReadCurretnRKUStatus()
        {
            if (KNS.Knfolder == null) { return; }
            KNS.OthersHaveReserved = false;
            string thisUserTag = "_" + Environment.UserName.ToUpper();
            string searchPat = "*" + extRKU;
            string rootKNSname = Path.GetFileNameWithoutExtension(KNS.Knfile) + "_";

            var files = from file in Directory.GetFiles(@KNS.Knfolder, searchPat, SearchOption.TopDirectoryOnly)
                        let fileName = Path.GetFileNameWithoutExtension(file)
                        where (fileName.Contains(rootKNSname) && !fileName.Contains(thisUserTag))
                        select file;

            foreach (string knuFile in files)
            {
                UserStats userStats = new UserStats(knuFile);
                if (KNS.Knfile.Equals(userStats.Kf))
                {
                    if (userStats.Re == "True")
                    {
                        RecordCategoryReserveStateChange(userStats.Cc, userStats.Us);
                        MainStatusMsg(userStats.Us + " is here and has " + userStats.Cn + " reserved.");
                        KNS.OthersHaveReserved = true;
                    }
                    else
                    {
                        RecordCategoryReserveStateChange(string.Empty, userStats.Us);
                        MainStatusMsg(userStats.Us + " is here with you.");
                    }
                }
            }
        }

        private void TextBoxesForNewCategory_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox tb = sender as TextBox;
            string txt = tb.Text;
            string tboxname = tb.Name;
            DataRow dr;
            switch (tboxname)
            {
                case "TextBoxNewCode":
                    dr = KNS.KndatacatState.Select("Key = '" + txt + "'").FirstOrDefault();
                    if (dr != null)
                    {
                        validNewCatCode = false;
                        tb.Foreground = Brushes.Red;
                        LableNewCatStatus.Content = "Code is not new.";
                    }
                    else
                    {
                        validNewCatCode = true;
                        tb.Foreground = Brushes.Green;
                        LableNewCatStatus.Content = string.Empty;
                    }
                    if (txt.Trim().Equals(string.Empty))
                    {
                        validNewCatCode = false;
                    }
                    break;
                case "TextBoxNewName":
                    dr = KNS.KndatacatState.Select("Name = '" + txt + "'").FirstOrDefault();
                    if (dr != null)
                    {
                        validNewCatName = false;
                        tb.Foreground = Brushes.Red;
                        LableNewCatStatus.Content = "Use a new Name.";
                    }
                    else
                    {
                        validNewCatName = true;
                        tb.Foreground = Brushes.Green;
                        LableNewCatStatus.Content = string.Empty;
                    }
                    if (txt.Trim().Equals(string.Empty))
                    {
                        validNewCatName = true;
                    }
                    break;
                default:
                    break;
            }
            ButtonAddNewCategory.IsEnabled = (validNewCatName && validNewCatCode);
        }

        private void ButtonAddNewCategory_Click(object sender, RoutedEventArgs e)
        {
            DataRow drcatstate = KNS.KndatacatState.NewRow();
            drcatstate[0] = TextBoxNewCode.Text;
            drcatstate[1] = TextBoxNewName.Text;
            drcatstate[2] = Available;
            KNS.NewCategories.Add(drcatstate[0].ToString().Trim() + '\t' + drcatstate[1].ToString().Trim());
            KNS.KndatacatState.Rows.Add(drcatstate);
            MainStatusMsg(TextBoxNewCode.Text + "  " + TextBoxNewName.Text + " category added.");
            TextBoxNewCode.Clear();
            TextBoxNewName.Clear();
            if (!KNS.SeenNewCatMsg) { PlayNewCatMsg(); }
        }

        private void PlayNewCatMsg()
        {
            string msg = "You should reserve and then commit as soon as possible after adding ";
            msg = msg + "a new category.\n\nYou don't need to do this right now if you plan to add ";
            msg = msg + "more new categories, but you should do it soon after adding the last one.\n\n";
            msg = msg + "Unlike reserved categories, all new categories are saved in the next single commit.\n\n";
            msg = msg + "Otherwise all the new categories will dissapear the next time the keynote table that was ";
            msg = msg + "changed by others is automatically added.";
            FormMsgWPF askthis = new FormMsgWPF("", 3, false);
            askthis.SetMsg("Capiche?", msg);
            askthis.ShowDialog();
            KNS.SeenNewCatMsg = true;
        }

        private void ButtonRefresh_Click(object sender, RoutedEventArgs e)
        {
            ReloadAndRestore();
        }

        private void ReloadAndRestore()
        {
            SaveSettingsState();
            TempLastSelCat = KNS.CurrentCatCode;
            TempLastKey = KNS.EditKeyText;
            ReadUpdateKNTables(Path.Combine(KNS.Knfolder, KNS.Knfile));
            ReadCurretnRKUStatus(); // Using this as a patch for out of whack reserve status
            RestoreLastSelections();
        }

        private void RestoreLastSelections()
        {
            if (KNS.Knfile.Equals(Properties.Settings.Default.LastKNFile))
            {
                int i = 0;
                if (dbug)
                {
                    MessageBox.Show(TempLastSelCat + "  " + TempLastKey, "RestoreLastSelections");
                }
                bool catisset = false;
                foreach (DataRowView cbi in ComboCatCode.Items)
                {
                    if (cbi[0].Equals(TempLastSelCat))
                    {
                        Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                        new Action(() => ComboCatCode.SelectedIndex = i));
                        KNS.EditKeyText = TempLastKey;
                        //MainStatusMsg("RestoreLastSelections " + TempLastSelCat + "  " + TempLastKey);
                        catisset = true;
                        break;
                    }
                    i = i + 1;
                }
                if (!catisset)
                {
                    Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                        new Action(() => ComboCatCode.SelectedIndex = 0));
                }
            }
        }

        private void ComboboxKey_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Inhibit) { return; }
            // The selected item in a datagrid bound to a table is a DataRowView
            ComboBox cb = sender as ComboBox;
            DataRowView drv = cb.SelectedItem as DataRowView;
            SelectThisDataGridRow(NotesGrid, drv);
        }

        public void SelectThisDataGridRow(DataGrid dataGrid, DataRowView drv)
        {
            if (drv != null && dataGrid != null)
            {
                if (Application.Current.Dispatcher.CheckAccess())
                {
                    dataGrid.SelectedItem = drv;
                    dataGrid.ScrollIntoView(drv);
                    dataGrid.Focus();
                }
                else
                {
                    Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                        new Action(() => dataGrid.SelectedItem = drv));
                    Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                        new Action(() => dataGrid.ScrollIntoView(drv)));
                    Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                        new Action(() => dataGrid.Focus()));
                }
                AcknowledgeGridSelection(dataGrid);
            }
        }

        private void LoadThisFileIntoHistViewer(string fn)
        {
            if (File.Exists(fn))
            {
                BkFileloadingArgs thisArgs = new BkFileloadingArgs
                {
                    FNAME = fn,
                    TARGETCODE = 0
                };
                if (!bkFileLoadingWorker.IsBusy)
                {
                    bkFileLoadingWorker.RunWorkerAsync(thisArgs);
                    KNS.HistTextFileName = Path.GetFileName(fn);
                }
            }
            else
            {
                KNS.HistText = "Did not find the file " + fn + ".";
                KNS.HistTextFileName = fn;
            }
        }

        private void BkFileLoadingWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BkFileloadingArgs theArgs = e.Argument as BkFileloadingArgs;
            string fname = theArgs.FNAME;
            if (File.Exists(fname))
            {
                theArgs.FILECONTENTS = File.ReadAllText(fname, Encoding.Default);
            }
            e.Result = theArgs;
        }

        private void BkFileLoadingWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BkFileloadingArgs theReturningArgs = e.Result as BkFileloadingArgs;
            string temp = string.Empty;
            if (e.Cancelled)
            {
                MainStatusMsg("Canceled reading the information document.");
            }
            else if (e.Error != null)
            {
                MainStatusMsg("Error. Details: " + (e.Error as Exception).ToString());
            }

            if (e.Result != null)
            {
                switch (theReturningArgs.TARGETCODE)
                {
                    case 0:
                        string fn = theReturningArgs.FNAME;
                        KNS.HistText = theReturningArgs.FILECONTENTS;
                        break;
                    //case 1:
                    //    string fnh = theReturningArgs.FNAME;
                    //    KNS.HelpText = theReturningArgs.FILECONTENTS;
                    //    break;
                    default:
                        break;
                }

            }
        }

        class BkFileloadingArgs
        {
            public string FNAME { get; set; }
            public string FILECONTENTS { get; set; }
            public int TARGETCODE { get; set; }
        }

        private void OnTabSelected(object sender, RoutedEventArgs e)
        {
            if (sender is TabItem tab)
            {
                TabItem thisTab = sender as TabItem;
                switch (thisTab.Name)
                {
                    case "TabItemResources":
                        if (Knfolder != null)
                        {
                            ShowHistoryFile(0);
                        }
                        break;
                    case "TabHelp":
                        if (HowTo == null) { HowTo = new Instructions(); }
                        HowTo.Show();
                        break;
                    default:
                        break;
                }
            }
        }

        private void ButtonQuit_Click(object sender, RoutedEventArgs e)
        {
            if (dictionaryPending.Count() > 0)
            {
                string msg = "\n";
                foreach (KeyValuePair<string, List<string>> entry in dictionaryPending)
                {
                    msg = msg + "\n" + entry.Key + " :: " + String.Join(", ", entry.Value.ToArray());
                    msg = msg + "\n";
                }
                FormMsgWPF askthis = new FormMsgWPF("", 2, false);
                askthis.SetMsg("Quit without saving anything.", "There are unsaved edits pending." + msg, "");
                askthis.ShowDialog();
                if (askthis.TheResult == MessageBoxResult.OK)
                {
                    Close();
                }
                else
                {
                    Dispatcher.BeginInvoke((Action)(() => TabControlMain.SelectedIndex = 0));
                }
            }
            else
            {
                Close();
            }

        }

        // home edit needs to be added
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
            }
        }

        // home edit needs to be added
        private void LoadThisHistoryFile(string thisKNF)
        {
            if (KNS.Knfolder != null && thisKNF != null)
            {
                string priorsPath = Path.Combine(KNS.Knfolder, priorsfolder);
                // Oddly, Path.Combine does not combine a path to a fully formed pathname.
                // So thisKNF remains the current KNF if that is what was passed into there.
                string thisFile = Path.Combine(priorsPath, thisKNF);
                if (File.Exists(thisFile))
                {
                    LoadThisFileIntoHistViewer(thisFile);
                    //MainStatusMsg("Displayed: " + Path.GetFileName(thisKNF));
                }
            }
        }

        // home edit needs to be added
        private void ShowHistoryFile(int direction)
        {
            if (Knfolder == null) { return; };
            // if direction 0 then we want to see the current file
            if (direction == 0)
            {
                string thisKNF = Path.Combine(Knfolder, Knfile);
                if (thisKNF != null)
                {
                    LoadThisHistoryFile(thisKNF);
                    KNS.HistIndex = "Current";
                }
                return;
            }
            string curfileshowing = string.Empty;
            string priorsPath = string.Empty;
            string rootKNSname = string.Empty;
            try
            {
                priorsPath = Path.Combine(KNS.Knfolder, priorsfolder);
                rootKNSname = Path.GetFileNameWithoutExtension(Knfile);
            }
            catch (Exception)
            {
                MainStatusMsg("No history found.");
                return;
            }
            if (!Directory.Exists(priorsPath))
            {
                MainStatusMsg("No history found.");
                return;
            }
            string searchPat = "*";
            // collect all history files for this keynote 
            var files = from file in Directory.GetFiles(@priorsPath, searchPat, SearchOption.TopDirectoryOnly)
                        let fileName = Path.GetFileNameWithoutExtension(file)
                        where (fileName.Contains(rootKNSname))
                        select Path.GetFileName(file);

            List<string> theHistoryList = files.ToList<string>();
            theHistoryList.Sort();
            KNS.Historybuttontext = "History Has " + theHistoryList.Count.ToString() + " Files";

            if (theHistoryList.Count == 0) { return; }

            // correction for when the current file showing has nothing to do with the keynote file
            if (KNS.HistTextFileName.Contains(rootKNSname))
            {
                curfileshowing = KNS.HistTextFileName;
            }
            else
            {
                curfileshowing = Knfile;
            }

            if (direction == -1 && curfileshowing.Equals(KNS.Knfile))
            {
                string histfile = theHistoryList.LastOrDefault();
                LoadThisHistoryFile(histfile);
                KNS.HistIndex = (theHistoryList.Count - 1).ToString() + " of " + theHistoryList.Count.ToString();
                return;
            }

            if (direction == 1 && curfileshowing.Equals(KNS.Knfile))
            {
                MainStatusMsg("You are looking at the current keynote file.");
                return;
            }

            if (!curfileshowing.Equals(KNS.Knfile))
            {
                int indexOfCurrentFileShowing = theHistoryList.FindIndex(x => x.Equals(curfileshowing));
                if (direction == -1)
                {
                    if (indexOfCurrentFileShowing > 0)
                    {
                        string histfile = theHistoryList[indexOfCurrentFileShowing - 1];
                        LoadThisHistoryFile(histfile);
                        KNS.HistIndex = (indexOfCurrentFileShowing).ToString() + " of " + theHistoryList.Count.ToString();
                    }
                    return;
                }
                if (direction == 1)
                {
                    if (indexOfCurrentFileShowing < theHistoryList.Count - 1)
                    {
                        string histfile = theHistoryList[indexOfCurrentFileShowing + 1];
                        LoadThisHistoryFile(histfile);
                        KNS.HistIndex = (indexOfCurrentFileShowing + 2).ToString() + " of " + theHistoryList.Count.ToString();
                    }
                    else
                    {
                        string thisKNF = Path.Combine(Knfolder, Knfile);
                        LoadThisHistoryFile(thisKNF);
                        KNS.HistIndex = "Current";
                    }
                    return;
                }
            }
        }
        // home edit needs to be added
        private void ButtonShowFileContent_Click(object sender, RoutedEventArgs e)
        {
            ShowHistoryFile(0);
        }
        // home edit needs to be added
        private void ButtonBackOneFileContent_Click(object sender, RoutedEventArgs e)
        {
            ShowHistoryFile(-1);
        }
        // home edit needs to be added
        private void ButtonForwardOneFileContent_Click(object sender, RoutedEventArgs e)
        {
            ShowHistoryFile(1);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            KNS.StatusMSG = "";
            ReadCurretnRKUStatus();
        }

        private void EditCell_DragOver(object sender, DragEventArgs e)
        {
            if (e.Effects.HasFlag(DragDropEffects.Copy))
            {
                Mouse.SetCursor(Cursors.Cross);
            }
            else if (e.Effects.HasFlag(DragDropEffects.Move))
            {
                Mouse.SetCursor(Cursors.Pen);
            }
            else
            {
                Mouse.SetCursor(Cursors.No);
            }
        }

        // We want this to be checked after the user has finished
        private void Textbox_numberformat_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            CheckIfNumberFormatIsValid();
            CheckIfIncrementIsValid();
        }

        private void CheckIfNumberFormatIsValid()
        {
            Single testval = 1f;
            if (!testval.IsFormatValid(KNS.Keynumberformat))
            {
                KNS.Keynumberformat = "00";
                string msg = "'" + KNS.Keynumberformat;
                msg = msg + " is not a valid 'Numeric Format String'.";
                FormMsgWPF askthis = new FormMsgWPF("", 3, false);
                askthis.SetMsg("It has been reset to '00'.", msg);
                askthis.ShowDialog();
            }
        }

        private void CheckIfIncrementIsValid()
        {
            Single A = 100f;
            Single B = A + KNS.Keyincval;
            string Aformatted = A.ToString(KNS.Keynumberformat, null);
            string BFormatted = B.ToString(KNS.Keynumberformat, null);

            if (KNS.Keyincval == 0)
            {
                KNS.Keyincval = 1;
                string msg = "The increment value cannot be zero.";
                FormMsgWPF askthis = new FormMsgWPF("", 3, false);
                askthis.SetMsg("It has been reset to 1.", msg);
                askthis.ShowDialog();
                return;
            }

            if (Aformatted.Equals(BFormatted))
            {
                string msg = "The numbering format is not coordinated with the numbering";
                msg = msg + " increment value. Added keynotes will probably not be";
                msg = msg + " numbered like you want.\n\n";
                msg = msg + "For example a numbering format of '00.00' requires an";
                msg = msg + " increment value of '0.01' in order to have numbering with";
                msg = msg + " two decimal places.";
                FormMsgWPF askthis = new FormMsgWPF("", 3, false);
                askthis.SetMsg("You should correct the settings.", msg);
                askthis.ShowDialog();
            }
        }

        private void Textbox_incval_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            CheckIfIncrementIsValid();
        }

        private void ButtonHistoryChannel_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            string priorsPath = string.Empty;
            try
            {
                priorsPath = Path.Combine(KNS.Knfolder, priorsfolder);
            }
            catch (Exception)
            {
                MainStatusMsg("No history path " + priorsPath + " found.");
                return;
            }
            if (!Directory.Exists(priorsPath))
            {
                MainStatusMsg("No history path " + priorsPath + " found.");
                return;
            }
            Process.Start(new ProcessStartInfo()
            {
                FileName = priorsPath,
                UseShellExecute = true,
                Verb = "open"
            });
        }

        private void TabHelp_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HowTo == null) { HowTo = new Instructions(); }
            HowTo.Show();
            e.Handled = true;
        }

        private void TabQuit_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            if (dictionaryPending.Count() > 0)
            {
                string msg = "\n";
                foreach (KeyValuePair<string, List<string>> entry in dictionaryPending)
                {
                    msg = msg + "\n" + entry.Key + " :: " + String.Join(", ", entry.Value.ToArray());
                    msg = msg + "\n";
                }
                FormMsgWPF askthis = new FormMsgWPF("", 2, false);
                askthis.SetMsg("Quit without saving anything.", "There are unsaved edits pending." + msg, "");
                askthis.ShowDialog();
                if (askthis.TheResult == MessageBoxResult.OK)
                {
                    Close();
                }
                else
                {
                    Dispatcher.BeginInvoke((Action)(() => TabControlMain.SelectedIndex = 0));
                }
            }
            else
            {
                Close();
            }
        }
    }

    public class InvertedBoolenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return !(bool)value;

        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return !(bool)value;
        }
    }

    internal class UserStats
    {
        private string rkucontent = string.Empty;
        private string us = string.Empty;
        private string tm = string.Empty;
        private string kf = string.Empty;
        private string re = string.Empty;
        private string cc = string.Empty;
        private string cn = string.Empty;

        public string Us { get { return us; } }
        public string Tm { get { return tm; } }
        public string Kf { get { return kf; } }
        public string Re { get { return re; } }
        public string Cc { get { return cc; } }
        public string Cn { get { return cn; } }

        public UserStats(string RKUfilename)
        {
            if (File.Exists(RKUfilename))
            {
                cc = string.Empty;
                cn = string.Empty;
                int tryThisManyTimes = 4;
                int attempts = 1;
                while (attempts <= tryThisManyTimes)
                {
                    try
                    {
                        var fileStream = new FileStream(RKUfilename, FileMode.Open, FileAccess.Read, FileShare.None);
                        using (var streamReader = new StreamReader(fileStream, Encoding.Default))
                        {
                            rkucontent = streamReader.ReadToEnd();
                        }
                        string[] lines = rkucontent.Split('\n');
                        foreach (string s in lines)
                        {
                            string lcode = s.Substring(0, 3);
                            string lvalue = s.Substring(3).Trim(); // remove trailing \r
                            if (lcode.Equals("US:")) { us = lvalue; continue; }
                            if (lcode.Equals("TM:")) { tm = lvalue; continue; }
                            if (lcode.Equals("KF:")) { kf = lvalue; continue; }
                            if (lcode.Equals("RE:")) { re = lvalue; continue; }
                            if (lcode.Equals("CC:")) { cc = lvalue; continue; }
                            if (lcode.Equals("CN:")) { cn = lvalue; continue; }
                        }
                        break;
                    }
                    catch (Exception)
                    {
                        attempts++;
                    }
                }
            }
        }
    }

    [ValueConversion(typeof(string), typeof(string))]
    public class RatioConverter : MarkupExtension, IValueConverter
    {
        private static RatioConverter _instance;

        public RatioConverter() { }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        { // do not let the culture default to local to prevent variable outcome re decimal syntax
            double size = System.Convert.ToDouble(value) * System.Convert.ToDouble(parameter, CultureInfo.InvariantCulture);
            return size.ToString("G0", CultureInfo.InvariantCulture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        { // read only converter...
            throw new NotImplementedException();
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return _instance ?? (_instance = new RatioConverter());
        }

    }

    public class RTFFile
    {
        public String File { get; set; }
    }

    public static class MyExtensions
    {

        public static bool IsFormatValid<T>(this T target, string Format)
            where T : IFormattable
        {
            try
            {
                target.ToString(Format, null);
            }
            catch
            {
                return false;
            }
            return true;
        }
    }

}



