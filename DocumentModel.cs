using System;

namespace Flexinets.Sharepoint
{
    public class DocumentModel
    {
        public Int32 Id
        {
            get;
            set;
        }
        public String Category
        {
            get;
            set;
        }
        public String Path
        {
            get;
            set;
        }
        public String Filename
        {
            get;
            set;
        }
        public DateTime Created
        {
            get;
            set;
        }
        public String HitHighlightedSummary
        {
            get;
            set;
        }
    }
}
