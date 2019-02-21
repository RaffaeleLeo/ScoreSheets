using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreSheets
{
    interface IView
    {
        /// <summary>
        /// fires if a user tries to register for some boggle
        /// </summary>
        event Action SelectFilePressed;

        event Action RegionalsCheckPressed;
    }
}
