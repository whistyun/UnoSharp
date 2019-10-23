using System;
using System.IO;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.document;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.sheet;
using unoidl.com.sun.star.util;

namespace UnoSharp
{
    public class Workbook : IDisposable
    {
        private bool closed = false;

        private XStorable storePeer;
        private XCloseable closePeer;
        private XPropertySet propSetPeer;

        private WorksheetCollection sheets;

        private Uri savedFile;

        private Workbook(Uri uri)
        {
            if (uri.Scheme.ToLower() == "file")
            {
                if (!File.Exists(uri.LocalPath))
                    throw new FileNotFoundException(uri.LocalPath);
            }

            var aLoader = OfficeServiceManager.loader;

            Peer = (XSpreadsheetDocument)aLoader.loadComponentFromURL(
                uri.AbsoluteUri, "_blank", 0,
                new[] {
                    new PropertyValue(){
                        Name ="Hidden",
                        Value =new uno.Any(true)
                    },
                    new PropertyValue(){
                        Name ="MacroExecutionMode",
                        Value = new uno.Any(MacroExecMode.FROM_LIST_NO_WARN)
                    }
                }
            );

            this.storePeer = (XStorable)Peer;
            this.closePeer = (XCloseable)Peer;
            this.propSetPeer = (XPropertySet)Peer;
            this.FormatsSupplier = (XNumberFormatsSupplier)Peer;
        }

        /// <summary>
        /// Create workbook as new.
        /// </summary>
        public Workbook() : this(new Uri("private:factory/scalc")) { }

        /// <summary>
        /// Open workbook.
        /// </summary>
        /// <param name="filepath">The file's path to open</param>
        public Workbook(string filepath) : this(new Uri(Path.GetFullPath(filepath)))
        {
            savedFile = new Uri(Path.GetFullPath(filepath));
        }

        /// <summary>
        /// UNO Peer
        /// </summary>
        public XSpreadsheetDocument Peer { get; }

        public XNumberFormatsSupplier FormatsSupplier { get; }

        public System.DateTime NullDate
        {
            get
            {
                // Unbelievable, This Date's month start with 1.
                // Why, java.util.Calendar ...
                // https://www.openoffice.org/api/docs/common/ref/com/sun/star/util/Date.html
                var date = (Date)propSetPeer.getPropertyValue("NullDate").Value;
                return new System.DateTime(date.Year, date.Month, date.Day);
            }
        }

        public void SaveAs(string filepath)
        {
            savedFile = new Uri(Path.GetFullPath(filepath));
            Save();
        }

        public void Save()
        {
            CheckState();

            if (savedFile == null) throw new NullReferenceException(nameof(savedFile));

            storePeer.storeAsURL(savedFile.AbsoluteUri, new PropertyValue[] {
                new PropertyValue() { Name = "Overwrite", Value = new uno.Any(true)}
            });
        }

        public void Close()
        {
            if (!closed)
            {
                closePeer.close(false);
                closed = true;
            }
        }

        public WorksheetCollection Worksheets
        {
            get
            {
                CheckState();

                if (sheets == null)
                {
                    sheets = new WorksheetCollection(this);
                }
                return sheets;
            }
        }

        public void Dispose()
        {
            Close();
        }

        private void CheckState()
        {
            if (closed) throw new InvalidOperationException("Already closed");
        }
    }
}
