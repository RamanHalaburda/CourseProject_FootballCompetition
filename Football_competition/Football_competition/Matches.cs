using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Football_competition
{
    class Matches
    {
        private int id_match;
	    private int id_firstteam;
	    private int id_secondteam;
	    private int id_stadium;
	    private int cost_ticket;
        private DateTime date;
        private int ball_first;
        private int ball_second;

        public Matches() { }

        public static void FirstPrintTitleDGV(System.Windows.Forms.DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "Соперник";
            dgv.Columns[1].HeaderCell.Value = "Дата матча";
            dgv.Columns[2].HeaderCell.Value = "Забито";
            dgv.Columns[3].HeaderCell.Value = "Пропущено";
        }

        public static void SecondPrintTitleDGV(System.Windows.Forms.DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "1 команда";
            dgv.Columns[1].HeaderCell.Value = "2 команда";
            dgv.Columns[2].HeaderCell.Value = "Дата матча";
            dgv.Columns[3].HeaderCell.Value = "Пропущено";
            dgv.Columns[4].HeaderCell.Value = "Забито";
        }
    }
}
