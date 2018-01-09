using System.Data;

namespace WpfRevitUserKeynoteManager
{
    internal class KNTable : DataTable
    {
        public KNTable()
        {
            this.Columns.Add("Key", typeof(string));
            this.Columns.Add("Note", typeof(string));
            this.Columns.Add("NoteKey", typeof(string));
        }
    }
   
    internal class KNTableCatState : DataTable
    {
        public KNTableCatState()
        {
            this.Columns.Add("Key", typeof(string));
            this.Columns.Add("Name", typeof(string));
            this.Columns.Add("Status", typeof(string));
        }
    }

}