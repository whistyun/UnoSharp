using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unoidl.com.sun.star.sheet;
using unoidl.com.sun.star.container;
using System.Collections;

namespace UnoSharp
{
    public class WorksheetCollection : IEnumerable<Worksheet>
    {
        private Workbook owner;

        public WorksheetCollection(Workbook owner)
        {
            this.owner = owner;
            this.Peer = owner.Peer.getSheets();
        }

        public XSpreadsheets Peer { get; }

        public int Count
        {
            get { return ((XIndexAccess)Peer).getCount(); }
        }

        public Worksheet this[int idx]
        {
            get
            {
                XIndexAccess xsheetsIA = (XIndexAccess)Peer;

                if (idx >= xsheetsIA.getCount())
                    throw new IndexOutOfRangeException();

                return new Worksheet(owner, (XSpreadsheet)xsheetsIA.getByIndex(idx).Value);
            }
        }

        public Worksheet this[string name]
        {
            get
            {
                if (!Peer.hasByName(name))
                {
                    var msg = String.Format("'{0}' sheet is not found", name);
                    throw new KeyNotFoundException(msg);
                }

                return new Worksheet(owner, (XSpreadsheet)Peer.getByName(name).Value);
            }
        }

        public Worksheet Add(string sheetName)
        {
            return Add(sheetName, Count);
        }

        public Worksheet Add(string sheetName, int idx)
        {
            Peer.insertNewByName(sheetName, (short)idx);
            return this[idx];
        }

        public IEnumerator<Worksheet> GetEnumerator()
        {
            return new SheetEnumerator(this);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    class SheetEnumerator : IEnumerator<Worksheet>
    {
        private int index;
        private WorksheetCollection owner;

        public SheetEnumerator(WorksheetCollection owner)
        {
            this.index = -1;
            this.owner = owner;
        }

        public Worksheet Current => owner[index];

        object IEnumerator.Current => Current;

        public bool MoveNext() => ++index < owner.Count;

        public void Reset() => index = -1;

        public void Dispose() { }
    }
}
