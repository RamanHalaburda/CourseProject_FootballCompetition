using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;

namespace Football_competition
{
    public partial class Form1 : Form
    {
        public Form1() { InitializeComponent(); }
/*
Разработка базы данных соревнований по футболу в рамках первенства страны.

Пусть требуется создать программную систему, предназначенную для организаторов соревнований по футболу в рамках первенства страны. 
Такая система должна обеспечивать хранение сведений о 
 * командах, участвующих в первенстве, 
 * об игроках команд, 
 * о расписании встреч и их результатах, 
 * о цене билетов на игры.
Сведения о команде - 
     * название команды, 
     * город, где она базируется, 
     * имя главного тренера, 
     * место в таблице прошлого сезона, 
     * расписание встреч. 
В один день команда может участвовать только в одной встрече. 
Сведения об игроке включают в себя 
 * фамилию и имя игрока, 
 * его возраст, 
 * номер и 
 * амплуа в команде. 
Сведения о стадионе, на котором происходит встреча содержат 
 * город, в котором он находится, 
 * название стадиона, и 
 * его вместимость . 
 Цена билета на матч зависит от вместимости стадиона и положения встречающихся команд в турнирной таблице прошлого сезона (наибольшая - при игре тройки призеров, наименьшая - при игре тройки аутсайдеров). 
Организаторы соревнований должны иметь возможность 
 * внести изменения в данные о составе команд, 
 * перенести встречу.
Им могут потребоваться следующие сведения:
    Даты встреч указанной команды, ее противники и счет?
+    Номера и фамилии игроков команд, участвовавших во встрече, которая проходила в указанный день в указанном городе? 
+    Цена, билета на матч между указанными командами?
    Игрок, забивший в турнире наибольшее количество мячей?
    Команды, имеющие наилучшую и наихудшую разницу забитых и пропущенных мячей?
+    Самый молодой участник турнира?
+    Команды, занявшие призовые места?
+    Расписание игр по стадионам?
    
По результатам турнира должен быть представлен отчет с результатами каждой игры. 
Для каждой игры указывается 
 * место и 
 * время ее проведения, 
 * команды – участницы, 
 * счет, 
 * игроки, выходившие на поле, и игроки, забившие мячи (указать на какой минуте). 
В отчете должны быть указаны призеры турнира и команда, занявшая последнее место.
*/
        OracleConnection oc;
        DataSet ds;
        OracleDataAdapter oda;

        public void AllMatch_PrintTitleDGV(DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД матча";
            dgv.Columns[1].HeaderCell.Value = "ИД команды";
            dgv.Columns[2].HeaderCell.Value = "ИД команды";
            dgv.Columns[3].HeaderCell.Value = "ИД стадиона";
            dgv.Columns[4].HeaderCell.Value = "Стоимость билета($)";
            dgv.Columns[5].HeaderCell.Value = "Дата";
            dgv.Columns[6].HeaderCell.Value = "Забито первой командой";
            dgv.Columns[7].HeaderCell.Value = "Забито второй командой";
        }

        public void Match4_PrintTitleDGV(DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД команды соперника";
            dgv.Columns[1].HeaderCell.Value = "Дата";
            dgv.Columns[2].HeaderCell.Value = "Забито";
            dgv.Columns[3].HeaderCell.Value = "Пропущено";
        }

        public void Match5_PrintTitleDGV(DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД команды";
            dgv.Columns[1].HeaderCell.Value = "ИД команды соперника";
            dgv.Columns[2].HeaderCell.Value = "Дата";
            dgv.Columns[3].HeaderCell.Value = "Забито";
            dgv.Columns[4].HeaderCell.Value = "Пропущено";
        }

        public void AllPlayer_PrintTitleDGV(DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД игрока";
            dgv.Columns[1].HeaderCell.Value = "ИД команды";
            dgv.Columns[2].HeaderCell.Value = "ФИО";
            dgv.Columns[3].HeaderCell.Value = "Возраст";
            dgv.Columns[4].HeaderCell.Value = "Номер в команде";
            dgv.Columns[5].HeaderCell.Value = "Роль";
        }

        public void Player4_PrintTitleDGV(DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ФИО";
            dgv.Columns[1].HeaderCell.Value = "Возраст";
            dgv.Columns[2].HeaderCell.Value = "Номер в команде";
            dgv.Columns[3].HeaderCell.Value = "Роль";
        }

        public void AllStadium_PrintTitleDGV(DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД стадиона";
            dgv.Columns[1].HeaderCell.Value = "Название";
            dgv.Columns[2].HeaderCell.Value = "Город";
            dgv.Columns[3].HeaderCell.Value = "Вместимость(чел.)";
        }

        public void AllTeam_PrintTitleDGV(DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД команды";
            dgv.Columns[1].HeaderCell.Value = "Название";
            dgv.Columns[2].HeaderCell.Value = "Город";
            dgv.Columns[3].HeaderCell.Value = "Тренер";
            dgv.Columns[4].HeaderCell.Value = "Место";
            dgv.Columns[5].HeaderCell.Value = "Побед";
            dgv.Columns[6].HeaderCell.Value = "Поражений";
            dgv.Columns[7].HeaderCell.Value = "Ничей";
            dgv.Columns[8].HeaderCell.Value = "Очки";
        }

        public void Team2_PrintTitleDGV(DataGridView dgv)
        {
            dgv.Columns[0].HeaderCell.Value = "ИД команды";
            dgv.Columns[1].HeaderCell.Value = "Название";
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            string oradb = "Data Source=localhost:1521/XE; User Id=SYSTEM;Password=1";
            oc = new OracleConnection(oradb);  
            oc.Open();
            ds = new DataSet();
            comboBox1.SelectedIndex = 0;

            // заполнение groupBox'а названиями стадионов
            ds = new DataSet();
            oda = new OracleDataAdapter("select title from football_stadiums order by id_stadium", oc);
            oda.Fill(ds);
            dataGridView6.DataSource = ds.Tables[0];

            for (int i = 0; i < dataGridView6.RowCount - 1; ++i)
            {
                comboBox2.Items.Add(dataGridView6[0, i].Value.ToString());
                comboBox6.Items.Add(dataGridView6[0, i].Value.ToString());
            }

            // заполнение groupBox'ов названиями команд
            ds = new DataSet();
            oda = new OracleDataAdapter("select name_team from football_teams order by id_team", oc);
            oda.Fill(ds);
            dataGridView6.DataSource = ds.Tables[0];

            for (int i = 0; i < dataGridView6.RowCount - 1; ++i)
            {
                comboBox3.Items.Add(dataGridView6[0, i].Value.ToString());
                comboBox4.Items.Add(dataGridView6[0, i].Value.ToString());
            } 
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 1;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                ds = new DataSet();
                oda = new OracleDataAdapter("select * from football_matches;", oc);
                oda.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                AllMatch_PrintTitleDGV(dataGridView1);
            }
            if (comboBox1.SelectedIndex == 1)
            {
                ds = new DataSet();
                oda = new OracleDataAdapter("select * from football_teams;", oc);
                oda.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                AllTeam_PrintTitleDGV(dataGridView1);
            }
            if (comboBox1.SelectedIndex == 2)
            {
                ds = new DataSet();
                oda = new OracleDataAdapter("select * from football_players;", oc);
                oda.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                AllPlayer_PrintTitleDGV(dataGridView1);
            }
            if (comboBox1.SelectedIndex == 3)
            {
                ds = new DataSet();
                oda = new OracleDataAdapter("select * from football_stadiums;", oc);
                oda.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                AllStadium_PrintTitleDGV(dataGridView1);
            }
        }
/************************************** tabPage 2 **************************************/
        private void tabPage2_Enter(object sender, EventArgs e)
        {
            ds = new DataSet();
            oda = new OracleDataAdapter("select id_team, name_team from football_teams;", oc);
            oda.Fill(ds);
            dataGridView2.DataSource = ds.Tables[0];
            Team2_PrintTitleDGV(dataGridView2);
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ds = new DataSet();
            oda = new OracleDataAdapter("select id_team_st,date_match,ball_first,ball_second from football_matches where id_team_ft = "
                + dataGridView2.Rows[e.RowIndex].Cells[0].Value + " and ball_first != ball_second", oc);
            oda.Fill(ds);
            dataGridView3.DataSource = ds.Tables[0];
            Match4_PrintTitleDGV(dataGridView3);

            ds = new DataSet();
            oda = new OracleDataAdapter("select id_team_ft,date_match,ball_second,ball_first from football_matches where id_team_st = "
                + dataGridView2.Rows[e.RowIndex].Cells[0].Value + " and ball_first != ball_second", oc);
            oda.Fill(ds);
            dataGridView4.DataSource = ds.Tables[0];
            Match4_PrintTitleDGV(dataGridView4);

            ds = new DataSet();
            oda = new OracleDataAdapter("select id_team_ft,id_team_st,date_match,ball_first,ball_second from football_matches where (id_team_ft = "
                + System.Convert.ToInt32(dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString()) 
                + " and ball_first = ball_second) or (id_team_st = "
                + System.Convert.ToInt32(dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString())
                + " and ball_first = ball_second)", oc);
            oda.Fill(ds);
            dataGridView5.DataSource = ds.Tables[0];
            Match5_PrintTitleDGV(dataGridView5);
        }
/************************************** tabPage 3 **************************************/
        private void tabPage3_Enter(object sender, EventArgs e)
        {
            ds.Tables[0].Clear();
            ds.Tables[0].Columns.Clear();
            dataGridView6.DataSource = ds.Tables[0];
            dataGridView7.DataSource = ds.Tables[0];

            label5.Text =  "________";
            label11.Text = "________";
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedValue != "")
            {
                DateTime picked_date = new DateTime(dateTimePicker1.Value.Year,
                                                    dateTimePicker1.Value.Month,
                                                    dateTimePicker1.Value.Day, 0, 0, 0);
                
// поиск ИД команд-участников
                string ft_id = "", st_id = "";
                try
                {
                    ds = new DataSet();
                    oda = new OracleDataAdapter("select id_team_ft, id_team_st from football_matches where id_stadium = "
                        + (comboBox2.SelectedIndex + 1) + " and date_match = TO_DATE('"
                        + picked_date.Year.ToString() + "."
                        + picked_date.Month.ToString() + "."
                        + picked_date.Day.ToString() + "' ,'yyyy/mm/dd hh24:mi:ss')", oc);

                    oda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                    ft_id = dataGridView6.Rows[0].Cells[0].Value.ToString();
                    st_id = dataGridView6.Rows[0].Cells[1].Value.ToString();
                }
                catch (Exception) { MessageBox.Show("Выбранный вами матч не существует"); }
// поиск названий команд-участников
                ds = new DataSet();
                oda = new OracleDataAdapter("select id_team,name_team from football_teams where id_team = "
                    + dataGridView6.Rows[0].Cells[1].Value, oc);
                oda.Fill(ds);
                dataGridView7.DataSource = ds.Tables[0];
                label11.Text = dataGridView7.Rows[0].Cells[1].Value.ToString();

                ds = new DataSet();
                oda = new OracleDataAdapter("select id_team,name_team from football_teams where id_team = "
                    + dataGridView6.Rows[0].Cells[0].Value, oc);
                oda.Fill(ds);
                dataGridView6.DataSource = ds.Tables[0];
                label5.Text = dataGridView6.Rows[0].Cells[1].Value.ToString();

// вывод игроков команд-участников
                ds = new DataSet();
                oda = new OracleDataAdapter("select fio,age,number_player,role_player from football_players where id_team_fk = "
                    + dataGridView6.Rows[0].Cells[0].Value, oc);
                oda.Fill(ds);
                dataGridView6.DataSource = ds.Tables[0];
                Player4_PrintTitleDGV(dataGridView6);

                ds = new DataSet();
                oda = new OracleDataAdapter("select fio,age,number_player,role_player from football_players where id_team_fk = "
                    + dataGridView7.Rows[0].Cells[0].Value, oc);
                oda.Fill(ds);
                dataGridView7.DataSource = ds.Tables[0];
                Player4_PrintTitleDGV(dataGridView7);
// печать отчета с игроками встречи
                if (checkBox2.Checked == true)
                {
                    Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook wb = application.Workbooks.Add(System.Reflection.Missing.Value);
                    Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)application.ActiveSheet;
                    string Name = "players(team " + ft_id + " vs team " + st_id + ")";
                    ws.Name = Name;

                    ws.Cells[1, 1] = label5.Text;
                    ws.Cells[3, 1] = "ФИО";
                    ws.Cells[3, 2] = "Возраст";
                    ws.Cells[3, 3] = "Номер игрока";
                    ws.Cells[3, 4] = "Роль";

                    ws.Cells[1, 6] = label11.Text;
                    ws.Cells[3, 6] = "ФИО";
                    ws.Cells[3, 7] = "Возраст";
                    ws.Cells[3, 8] = "Номер игрока";
                    ws.Cells[3, 9] = "Роль";

                    for (int i = 0; i < dataGridView6.Rows.Count; i++)
                        for (int j = 0; j < dataGridView6.ColumnCount; j++)
                            application.Cells[i + 4, j + 1] = dataGridView6.Rows[i].Cells[j].Value;

                    for (int i = 0; i < dataGridView7.Rows.Count; i++)
                        for (int j = 0; j < dataGridView7.ColumnCount; j++)
                            application.Cells[i + 4, j + 6] = dataGridView7.Rows[i].Cells[j].Value;

                    ws.Columns.AutoFit();

                    application.Visible = true;
                    application.UserControl = true;
                    try
                    {
                        string fileName = "e:\\study\\(1) Учёба\\БарГУ Инж.Фак\\Курсовое проектирование"
                                        + "\\Вика\\Football_competition\\Football_competition\\reports\\"
                                        + Name;
                        wb.SaveAs(fileName);
                    }
                    catch (Exception) { }
                }
            }
        }
/************************************** tabPage 4 **************************************/
        private void tabPage4_Enter(object sender, EventArgs e)
        {
// призеры первенства
            ds = new DataSet();
            oda = new OracleDataAdapter("select * from football_teams where position_team between 1 and 3 order by position_team", oc);
            oda.Fill(ds);
            dataGridView8.DataSource = ds.Tables[0];
            AllTeam_PrintTitleDGV(dataGridView8);

// последнее место
            int max_pos = 0;
            comboBox1.SelectedItem = "Команды";
            for (int i = 0; i < (dataGridView1.RowCount - 1); ++i)
                if (System.Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value) > max_pos)
                    max_pos = System.Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value);
                        
            ds = new DataSet();
            oda = new OracleDataAdapter("select * from football_teams where position_team = "
                + max_pos.ToString() + " order by position_team", oc);
            oda.Fill(ds);
            dataGridView9.DataSource = ds.Tables[0];
            AllTeam_PrintTitleDGV(dataGridView9);

// самый молодой и самый старший игроки
            int max_age = 0, min_age = 200;
            comboBox1.SelectedItem = "Игроки";
            for (int i = 0; i < (dataGridView1.RowCount - 1); ++i)
            {
                if (System.Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value) > max_age)
                    max_age = System.Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value);
                if (System.Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value) < min_age)
                    min_age = System.Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value);
            }

            ds = new DataSet();
            oda = new OracleDataAdapter("select * from football_players where age = "
                + max_age.ToString() + " or age = "
                + min_age.ToString() + " order by age", oc);
            oda.Fill(ds);
            dataGridView10.DataSource = ds.Tables[0];
            AllPlayer_PrintTitleDGV(dataGridView10);

// печать отчета с призерами
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = application.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)application.ActiveSheet;
            string Name = "winners";
            ws.Name = Name;

            ws.Cells[1, 1] = "ИД Команды";
            ws.Cells[1, 2] = "Название команды";
            ws.Cells[1, 3] = "Город базирования";
            ws.Cells[1, 4] = "Тренер";
            ws.Cells[1, 5] = "Позиция";
            ws.Cells[1, 6] = "Побед";
            ws.Cells[1, 7] = "Поражений";
            ws.Cells[1, 8] = "Ничей";
            ws.Cells[1, 9] = "Очков";


            for (int i = 0; i < dataGridView8.Rows.Count; i++)
                for (int j = 0; j < dataGridView8.ColumnCount; j++)
                    application.Cells[i + 2, j + 1] = dataGridView8.Rows[i].Cells[j].Value;

            ws.Columns.AutoFit();

            application.Visible = true;
            application.UserControl = true;
            try
            {
                string fileName = "e:\\study\\(1) Учёба\\БарГУ Инж.Фак\\Курсовое проектирование"
                                + "\\Вика\\Football_competition\\Football_competition\\reports\\"
                                + Name;
                wb.SaveAs(fileName);
            }
            catch (Exception) { }
        }
// цена билета на матч между указанными командами
        private void button2_Click(object sender, EventArgs e)
        {            
            if (comboBox3.SelectedValue != "" && comboBox4.SelectedValue != "")
            {
                if (comboBox3.SelectedIndex.ToString() != comboBox4.SelectedIndex.ToString())
                {
                    ds = new DataSet();
                    string query = "select * from football_matches where (id_team_ft like "
                        + (comboBox3.SelectedIndex + 1)
                        + " and id_team_st like "
                        + (comboBox4.SelectedIndex + 1)
                        + ") or (id_team_st like "
                        + (comboBox3.SelectedIndex + 1)
                        + " and id_team_ft like "
                        + (comboBox4.SelectedIndex + 1)
                        + ")";
                    oda = new OracleDataAdapter(query, oc);
                    oda.Fill(ds);
                    label10.Text = ds.Tables[0].Rows[0][4].ToString() + "$";
                    // печать билета
                    if (checkBox1.Checked == true)
                    {
                        int columns = 2;
                        int rows = 5;

                        Microsoft.Office.Interop.Word.Application applictaion = new Microsoft.Office.Interop.Word.Application();
                        Object missing = Type.Missing;
                        applictaion.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                        Microsoft.Office.Interop.Word.Document document = applictaion.ActiveDocument;

                        Paragraph p = document.Content.Paragraphs.Add(ref missing);
                        p.Range.Text = "Билет на матч: " + ds.Tables[0].Rows[0][0].ToString();
                        p.Range.InsertParagraphAfter();
                        Table table = document.Tables.Add(p.Range, rows, columns, ref missing, ref missing);

                        table.Borders.Enable = 1;
                        table.Cell(1, 1).Range.Text = "Команда";
                        table.Cell(2, 1).Range.Text = "Команда";
                        table.Cell(3, 1).Range.Text = "ИД Стадиона";
                        table.Cell(4, 1).Range.Text = "Стоимость билета";
                        table.Cell(5, 1).Range.Text = "Дата матча";
                        table.Cell(1, 2).Range.Text = comboBox3.Text;
                        table.Cell(2, 2).Range.Text = comboBox4.Text;
                        table.Cell(3, 2).Range.Text = ds.Tables[0].Rows[0][3].ToString();
                        table.Cell(4, 2).Range.Text = ds.Tables[0].Rows[0][4].ToString() + "$";
                        table.Cell(5, 2).Range.Text = ds.Tables[0].Rows[0][5].ToString();

                        applictaion.Visible = true;
                        try
                        {
                            object fileName = "e:\\study\\(1) Учёба\\БарГУ Инж.Фак\\Курсовое проектирование"
                                            + "\\Вика\\Football_competition\\Football_competition\\tickets\\ticket_match"
                                            + ds.Tables[0].Rows[0][0].ToString();
                            document.SaveAs2(fileName);
                        }
                        catch (Exception) { }
                    }
                }
                else
                    MessageBox.Show("Команды одинаковы!");
            }
            else
                MessageBox.Show("Команды не выбраны!");
        }
/************************************** tabPage 5 **************************************/
        private void tabPage5_Enter(object sender, EventArgs e)
        {
            comboBox5.Items.Clear();
            comboBox5.Items.Add("нападающий");
            comboBox5.Items.Add("защитник");
            comboBox5.Items.Add("вратарь");
            comboBox5.SelectedIndex = 0;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            ds = new DataSet();
            string query = "select * from football_players where role_player = '"
                + comboBox5.SelectedItem.ToString() + "' order by number_player";
            oda = new OracleDataAdapter(query, oc);
            oda.Fill(ds);
            dataGridView11.DataSource = ds.Tables[0];
            AllPlayer_PrintTitleDGV(dataGridView11);

// печать отчета со всем играками с заданной ролью
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = application.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)application.ActiveSheet;
            string Name = "allplayers(" + comboBox5.Text + ")";
            ws.Name = Name;
            
            ws.Cells[1, 1] = "ИД игрока";
            ws.Cells[1, 2] = "ИД команды";
            ws.Cells[1, 3] = "ФИО";
            ws.Cells[1, 4] = "Возраст";
            ws.Cells[1, 5] = "Номер игрока";
            ws.Cells[1, 6] = "Роль";


            for (int i = 0; i < dataGridView11.Rows.Count; i++)
                for (int j = 0; j < dataGridView11.ColumnCount; j++)
                    application.Cells[i + 2, j + 1] = dataGridView11.Rows[i].Cells[j].Value;

            ws.Columns.AutoFit();

            application.Visible = true;
            application.UserControl = true;
            try
            {
                string fileName = "e:\\study\\(1) Учёба\\БарГУ Инж.Фак\\Курсовое проектирование"
                                + "\\Вика\\Football_competition\\Football_competition\\reports\\"
                                + Name;
                wb.SaveAs(fileName);
            }
            catch (Exception) { }
        }
/************************************** tabPage 6 **************************************/
        private void tabPage6_Enter(object sender, EventArgs e)
        {
            comboBox6.SelectedIndex = 0;
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            ds = new DataSet();
            string query = "select * from football_matches where id_stadium = "
                + (comboBox6.SelectedIndex + 1) + " order by date_match";
            oda = new OracleDataAdapter(query, oc);
            oda.Fill(ds);
            dataGridView12.DataSource = ds.Tables[0];
            AllMatch_PrintTitleDGV(dataGridView12);
        }    
    }
}
