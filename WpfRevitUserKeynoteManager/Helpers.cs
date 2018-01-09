using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WpfRevitUserKeynoteManager
{
    internal static class Helpers
    {
        internal static void RunExplorerHere(string path)
        {
            Process.Start("Explorer.exe", path);

            // to do
            //public bool ExploreFile(string filePath)
            //{
            //    if (!System.IO.File.Exists(filePath))
            //    {
            //        return false;
            //    }
            //    //Clean up file path so it can be navigated OK
            //    filePath = System.IO.Path.GetFullPath(filePath);
            //    System.Diagonstics.Process.Start("explorer.exe", string.Format("/select,\"{0}\"", filePath));
            //    return true;
            //}

        }

        internal static void WebBrowserToHere(string target)
        {
            try
            {
                Process.Start(target);
            }
            catch
                (
                 System.ComponentModel.Win32Exception noBrowser)
            {
                if (noBrowser.ErrorCode == -2147467259)
                    MessageBox.Show(noBrowser.Message);
            }
            catch (Exception other)
            {
                MessageBox.Show(other.Message);
            }


        }

        internal static string MakeFileNameDateTimeRegexPattern(List<string> fnameExtList)
        {
            /// regex pattern matches 012345.01.pic at string end
            // pat_general_datetime = @"\d{6}\.\d{2}\.pic\Z";
            string pat_general_datetime = @"\d{6}\.\d{2}\.(";
            // string resStr = @"^.+\.(";
            foreach (string s in fnameExtList)
            {
                pat_general_datetime = pat_general_datetime + s + '|';
            }
            pat_general_datetime = pat_general_datetime.TrimEnd('|');
            pat_general_datetime = pat_general_datetime + @")\Z";
            return pat_general_datetime;
        }

        internal static string GiveMeBestPathOutOfThisPath(string proposedPath)
        {
            string[] words = proposedPath.Split('\\');
            string bestPath = String.Empty;
            foreach (string token in words)
            {
                string attempt = string.Empty;
                if (token.Equals(words[0]))
                {
                    attempt = token;
                }
                else
                {
                    attempt = String.Concat(bestPath, '\\', token);
                }
                if (Directory.Exists(attempt))
                {
                    bestPath = attempt;
                }
                else
                {
                    break;
                }
            }
            return bestPath;
        }

        internal static void SetTextToASelectedFolder(object sender, string msg, bool dotmode, string basePath)
        {
            TextBox tb = sender as TextBox;
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog
            {
                SelectedPath = tb.Text
            };
            if (dotmode)
            {
                // if the dotmode folder does not exist then we want to drop back no
                // further than necessary woith the default folder. 
                string proposedPath = CombineIntoPath(basePath, tb.Text);
                folderDialog.SelectedPath = GiveMeBestPathOutOfThisPath(proposedPath);
                if (!Directory.Exists(folderDialog.SelectedPath))
                {
                    folderDialog.SelectedPath = basePath; // for the time being
                }
            }
            folderDialog.Description = msg;
            var result = folderDialog.ShowDialog();
            switch (result)
            {
                case System.Windows.Forms.DialogResult.OK:
                    var folder = folderDialog.SelectedPath;
                    string fp = EnsurePathStringEndsInBackSlash(folder);
                    if (!dotmode)
                    {
                        tb.Text = fp;
                    }
                    else
                    {
                        string t = fp.ReplaceString(basePath, ".\\", StringComparison.CurrentCultureIgnoreCase);
                        tb.Text = t;
                    }
                    break;
                case System.Windows.Forms.DialogResult.Cancel:
                default:

                    break;
            }
        }

        internal static void SetTextToASelectedFile(object sender, string msg, bool fullpath)
        {
            TextBox tb = sender as TextBox;
            var filepickerDialog = new System.Windows.Forms.OpenFileDialog
            {
                Title = msg
            };

            var result = filepickerDialog.ShowDialog();
            switch (result)
            {
                case System.Windows.Forms.DialogResult.OK:

                    string thepick = filepickerDialog.FileName;

                    if (!fullpath)
                    {
                        thepick = Path.GetFileName(thepick);
                    }

                    tb.Text = thepick;
                    break;
                case System.Windows.Forms.DialogResult.Cancel:
                default:
                    break;
            }
        }

        internal static string CombineIntoPath(string partA_RootPath, string partB_PartialPathWithDot, string partC_OptionalFileName = "")
        {
            try
            {
                string fullPath = Path.Combine(partA_RootPath, partB_PartialPathWithDot);
                fullPath = fullPath.Replace(".\\", "") + partC_OptionalFileName;
                return fullPath;
            }
            catch (ArgumentException e)
            {
                string msg = "Part of this path name " + partB_PartialPathWithDot + partC_OptionalFileName + " is illegal. You need to correct it";
                msg = msg + "\n\n" + e.Message;
                FormMsgWPF explain = new FormMsgWPF(null, 3);
                explain.SetMsg(msg, "The Path Has Illegal An Character");
                explain.ShowDialog();
            }
            return string.Empty;
        }

        internal static string EnsurePathStringEndsInBackSlash(string path)
        {
            if (!path.EndsWith("\\", StringComparison.CurrentCultureIgnoreCase))
            {
                path = path + "\\";
            }
            return path;
        }

        internal static void EndsInBackSlash(object sender)
        {
            TextBox tb = sender as TextBox;
            tb.Text = EnsurePathStringEndsInBackSlash(tb.Text);
            tb.CaretIndex = tb.Text.Length;
        }

        // colors a filepath textbox as to filepath's existance
        // Also returns as bool for path existance. False = path does not exist
        internal static bool MarkTextBoxForPath(TextBox theTextBox, string basePath, bool dotMode = false)
        {
            // if (path == null) { return; }
            if (theTextBox == null) { return false; }
            string path = theTextBox.Text;
            if (dotMode) { path = Helpers.CombineIntoPath(basePath, path); }
            if (Directory.Exists(path))
            {
                theTextBox.Foreground = Brushes.Black;
                return true;
            }
            else
            {
                theTextBox.Foreground = Brushes.Red;
                return false;
            }
        }

        // colors a filename textbox as to filename's existance or allowed characters
        internal static void MarkTextBoxForFile(TextBox theTextBox, string basePath, string subpath)
        {
            if (subpath == null) { return; }
            if (theTextBox == null) { return; }
            subpath = CombineIntoPath(basePath, subpath);
            subpath = CombineIntoPath(subpath, theTextBox.Text);
            if (File.Exists(subpath))
            {
                theTextBox.Foreground = Brushes.Black;
            }
            else
            {
                theTextBox.Foreground = Brushes.Red;
            }
        }

        // colors a filename textbox as to filename's existance or allowed characters
        internal static void MarkLabelForFile(Label theLabel, string fullPath, bool sense)
        {
            if (fullPath == null) { return; }
            if (theLabel == null) { return; }
            if (!IsValidWindowsFileName(Path.GetFileName(fullPath)))
            {
                theLabel.Foreground = Brushes.Red;
            }
            if (File.Exists(fullPath))
            {
                if (sense)
                {
                    theLabel.Foreground = Brushes.Black;
                }
                else
                {
                    theLabel.Foreground = Brushes.Red;
                }
            }
            else
            {
                if (sense)
                {
                    theLabel.Foreground = Brushes.Red;
                }
                else
                {
                    theLabel.Foreground = Brushes.Black;
                }
            }
        }

        // colors a filename datagridcell as to filename's existance
        internal static bool MarkDataGridCellForFile(DataGridCell theCell, string fullPath, bool sense)
        {
            if (fullPath == null) { return false; }
            if (theCell == null) { return false; }

            if (File.Exists(fullPath))
            {
                if (sense)
                {
                    theCell.Foreground = Brushes.Black;
                    return true;
                }
                else
                {
                    theCell.Foreground = Brushes.Red;
                    return false;
                }
            }
            else
            {
                if (sense)
                {
                    theCell.Foreground = Brushes.Red;
                    return false;
                }
                else
                {
                    theCell.Foreground = Brushes.Black;
                    return true;
                }
            }
        }

        internal static Dictionary<string, bool> BuildCheckListDictionary(string thePath, List<string> regexsearchPat)
        {
            string pat_toapply = MakeFilesListRegexString(regexsearchPat);
            Dictionary<string, bool> thisCheckListDict = new Dictionary<string, bool>();
            if (!Directory.Exists(thePath)) { return thisCheckListDict; }
            try
            {
                IEnumerable<String> ResultsList = Directory.GetFiles(thePath).Select(Path.GetFileName).Where(file => Regex.IsMatch(file, pat_toapply, RegexOptions.IgnoreCase));
                foreach (string f in ResultsList)
                {
                    thisCheckListDict.Add(Path.GetFileName(f), false);
                }
            }
            catch (ArgumentException e)
            {
                string msg = "Not Good! It looks like the regular expression \"" + pat_toapply + "\" has an invalid syntax. ";
                msg = msg + " This is happening in the BuildCheckListDictionary function.";
                msg = msg + "\n\n" + e.Message;
                FormMsgWPF explain = new FormMsgWPF(null, 3);
                explain.SetMsg(msg, "Up The Creek Without A Paddle");
                explain.ShowDialog();
            }
            return thisCheckListDict;
        }

        internal static string MakePcombRegexFiltetString(string strThis, List<string> extList)
        {
            string pat_toapply = strThis + @".*\.(";
            foreach (string s in extList)
            {
                pat_toapply = pat_toapply + s + '|';
            }
            pat_toapply = pat_toapply.TrimEnd('|');
            pat_toapply = pat_toapply + @")";
            return pat_toapply;
        }

        /// Given a list of file extensions, less any period, returns a regex expression
        /// that when used in context with file names, matches only those files that have
        /// the file extension text.
        internal static string MakeFilesListRegexString(List<string> extList)
        {
            string resStr = @"^.+\.(";
            foreach (string s in extList)
            {
                resStr = resStr + s + '|';
            }
            resStr = resStr.TrimEnd('|');
            resStr = resStr + ")$";
            return resStr;
        }
        
        internal static void CheckInThisListBox(ListBox thisListBox, bool beChecked)
        {
            foreach (var item in thisListBox.Items)
            {
                if (item.GetType() == typeof(CheckBox))
                {
                    CheckBox cb = item as CheckBox;
                    cb.IsChecked = beChecked;
                }
            }
        }
        
        internal static void CreateThisPath(string newPathToCreate)
        {
            try
            {
                Directory.CreateDirectory(@newPathToCreate);
            }
            catch (Exception er)
            {
                string ttl = "Create Directory Error";
                string msg = "Unable to create the paths in " + newPathToCreate + "\n\n" + er.Message;
                FormMsgWPF explain = new FormMsgWPF(null, 3);
                explain.SetMsg(msg, ttl);
                explain.ShowDialog();
            }
        }

        // reports if text in textbox would be ok for a windows filename
        internal static void SniffTextBoxToBeAValidFileName(TextBox theTextBox)
        {
            if (theTextBox == null) { return; }
            string fn = theTextBox.Text;
            if (fn.Trim() == string.Empty) { return; }
            if (!IsValidWindowsFileName(fn))
            {
                string msg = "Windows will not allow \n\n" + fn + "\n\n to be a file name.";
                FormMsgWPF explain = new FormMsgWPF(null, 3);
                explain.SetMsg(msg, "By The Way. Not A Valid Name");
                explain.ShowDialog();
            }
        }

        /// <summary>
        /// Only works properly on file names. Do not use for full path names.
        /// </summary>
        /// <param name="expression"></param>
        /// <returns></returns>
        internal static bool IsValidWindowsFileName(string expression)
        {
            // https://stackoverflow.com/questions/62771/how-do-i-check-if-a-given-string-is-a-legal-valid-file-name-under-windows
            string sPattern = @"^(?!^(PRN|AUX|CLOCK\$|NUL|CON|COM\d|LPT\d|\..*)(\..+)?$)[^\x00-\x1f\\?*:\"";|/]+$";
            return (Regex.IsMatch(expression, sPattern, RegexOptions.CultureInvariant));
        }

        /// <summary>
        /// Tries to find ghostscript via registry and typical install location.
        /// </summary>
        /// <returns></returns>
        internal static string FindGhostScript()
        {
            // 
            string registryRoot = @"SOFTWARE\\GPL Ghostscript";
            RegistryKey root;
            //     case "HKCU":
            //root = Registry.CurrentUser.OpenSubKey(registryRoot, false);
            //     case "HKLM":
            root = Registry.LocalMachine.OpenSubKey(registryRoot, false);

            if (root != null)
            {
                var subKeys = root.GetSubKeyNames();
                string gsExecPath = @"C:\Program Files\gs\gs" + subKeys.FirstOrDefault<string>().ToString() + @"\bin\gswin64.exe";
                //MessageBox.Show(gsExecPath);
                return gsExecPath;
            }
            else
            {
                return "Not Installed. Convert to PDF wil not be possible.";
            }
        }

        internal static string FindSomething(string something)
        {
            bool debug = false;
            string Sp = " ";
            String theArgs = string.Empty;
            Process p = new Process();
            ProcessStartInfo psi = new ProcessStartInfo();
            // Required when setting process EnvironmentVariables
            psi.UseShellExecute = false;
            // set the working directory
            // Tactic is for CMD.EXE and programname to be part of argument
            psi.FileName = "CMD.EXE";
            if (!debug)
            {
                psi.RedirectStandardError = true;
                psi.RedirectStandardOutput = true;
                // no window
                psi.CreateNoWindow = true;
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                theArgs = String.Concat(@"/c ", "where" + Sp + something);  // window closes
            }
            else
            {
                psi.RedirectStandardError = false;
                psi.RedirectStandardOutput = false;
                psi.CreateNoWindow = false;
                psi.WindowStyle = ProcessWindowStyle.Normal;
                theArgs = String.Concat(@"/k ", "where" + Sp + something);  // window stays open
            }
            // set the arguments
            psi.Arguments = theArgs;
            //// show window or not
            //if (!bedebug)
            //{
            //    psi.WindowStyle = ProcessWindowStyle.Hidden;
            //}
            //else
            //{
            //    psi.WindowStyle = ProcessWindowStyle.Normal;
            //}
            // Starts process
            string output = string.Empty;
            string error = string.Empty;
            p.StartInfo = psi;
            p.Start();
            // Do not wait for the child process to exit before
            // reading to the end of its redirected error stream.
            // p.WaitForExit();
            // Read the error stream first and then wait.
            if (!debug)
            {
                error = p.StandardError.ReadToEnd();
                output = p.StandardOutput.ReadToEnd();
            }
            p.WaitForExit();
            if (error.Equals(string.Empty))
            {
                return output;
            }
            else
            {
                //return error;
                return "Did not locate " + something;
            }
        }

        internal static bool RegistryValueExists(string _strHive_HKLM_HKCU, string _registryRoot, string _valueName)
        {
            RegistryKey root;
            switch (_strHive_HKLM_HKCU.ToUpper())
            {
                case "HKLM":
                    root = Registry.LocalMachine.OpenSubKey(_registryRoot, false);
                    break;
                case "HKCU":
                    root = Registry.CurrentUser.OpenSubKey(_registryRoot, false);
                    break;
                default:
                    throw new InvalidOperationException("parameter registryRoot must be either \"HKLM\" or \"HKCU\"");
            }

            return root.GetValue(_valueName) != null;
        }

        internal static void DeleteFileFromListBox(object sender, string rootPath, string contextSubPath)
        {
            ListBox lb = sender as ListBox;
            if (lb.SelectedItems.Count == 0) { return; }
            string msg = string.Empty;
            string ttl = "What? Delete This File?";
            if (lb.SelectedItems.Count > 1) { ttl = "What?, Delete These Files?"; }
            foreach (ListBoxItem lbi in lb.SelectedItems)
            {
                msg = msg + lbi.Content.ToString() + "\n";
            }
            FormMsgWPF explain = new FormMsgWPF(null, 2);
            explain.SetMsg(msg, ttl);
            explain.ShowDialog();
            if (explain.TheResult == MessageBoxResult.OK)
            {
                foreach (ListBoxItem lbi in lb.SelectedItems)
                {
                    String fname = Helpers.CombineIntoPath(rootPath, contextSubPath, lbi.Content.ToString());
                    if (fname != string.Empty)
                    {
                        if (File.Exists(fname))
                        {
                            try
                            {
                                File.Delete(fname);
                            }
                            catch (Exception err)
                            {
                                msg = "Cannot delete.";
                                msg = msg + "\n\n" + err.Message;
                                explain = new FormMsgWPF(null, 3);
                                explain.SetMsg(msg, "IO Error");
                                explain.ShowDialog();
                            }
                        }
                    }
                }
            }
        }

        internal static void ExploreListBoxSelection(ListBox listBox, string listPath)
        {
            var selected = listBox.SelectedItem;
            if (selected == null)
            {
                Process.Start("explorer.exe", listPath);
                return;
            }

            if (selected.GetType() == typeof(CheckBox))
            {
                CheckBox cb = selected as CheckBox;
                TextBlock tb = cb.Content as TextBlock;
                if (tb != null)
                {
                    string fullPathName = listPath + tb.Text;
                    ExploreFile(fullPathName);
                    return;
                }
            }

            if (selected.GetType() == typeof(ListBoxItem))
            {
                ListBoxItem lbi = selected as ListBoxItem;
                if (lbi != null)
                {
                    string fullPathName = listPath + lbi.Content.ToString();
                    ExploreFile(fullPathName);
                    return;
                }
            }
        }

        internal static bool ExploreFile(string filePath)
        {
            if (!System.IO.File.Exists(filePath))
            {
                return false;
            }
            //Clean up file path so it can be navigated OK
            filePath = Path.GetFullPath(filePath);
            Process.Start("explorer.exe", string.Format("/select,\"{0}\"", filePath));
            return true;
        }
        
    }

    internal static class StringExtensions
    {
        internal static string Left(this string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            maxLength = Math.Abs(maxLength);
            return (value.Length <= maxLength
                    ? value
                    : value.Substring(0, maxLength)
                    );
        }

        internal static string ReplaceString(this string str, string oldValue, string newValue, StringComparison comparison)
        {
            StringBuilder sb = new StringBuilder();

            int previousIndex = 0;
            int index = str.IndexOf(oldValue, comparison);
            while (index != -1)
            {
                sb.Append(str.Substring(previousIndex, index - previousIndex));
                sb.Append(newValue);
                index += oldValue.Length;

                previousIndex = index;
                index = str.IndexOf(oldValue, index, comparison);
            }
            sb.Append(str.Substring(previousIndex));

            return sb.ToString();
        }
    }

    internal static class BitmapExtensions
    {
        internal static BitmapImage GetBitmapImage(this Uri imageAbsolutePath, BitmapCacheOption bitmapCacheOption = BitmapCacheOption.Default)
        {
            BitmapImage bi = new BitmapImage();
            try
            {
                bi.BeginInit();
                bi.CacheOption = bitmapCacheOption;
                bi.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                bi.UriSource = imageAbsolutePath;
                bi.EndInit();
                return bi;
            }
            catch (Exception e)
            {
                string msg = "Trouble viewing \"" + imageAbsolutePath + "\".";
                msg = msg + "\n\n" + e.Message;
                FormMsgWPF explain = new FormMsgWPF(null, 3);
                explain.SetMsg(msg, "Sorry Charlie");
                explain.ShowDialog();
                bi = null;
            }
            return bi;
        }
    }

    /// <summary>
    /// Used to convert system drawing colors to WPF brush
    /// </summary>
    internal static class ColorExt
    {
        internal static System.Windows.Media.Brush ToBrush(System.Drawing.Color color)
        {
            {
                return new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B));
            }
        }
    }
}
