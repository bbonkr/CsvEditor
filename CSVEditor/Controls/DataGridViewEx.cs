using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CSVEditor.Controls
{
    public class DataGridViewEx : DataGridView
    {
        public DataGridViewEx()
        {
            base.DoubleBuffered = true;
        }

        public string CurrentFilePath { get; set; }
    }
}