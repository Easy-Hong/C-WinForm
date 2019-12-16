using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSharpAnimeInfo
{
    public partial class Form1 : Form
    {
        MySqlConnection conn;
        MySqlDataAdapter dataAdaptera;
        MySqlDataAdapter dataAdapterb;
        MySqlDataAdapter dataAdapterc;
        DataSet dataSeta;
        DataSet dataSetb;
        DataSet dataSetc;
        int update_cnt = 0;
        string fsPath = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            pictureBox1.Load(@"C:\dbline.png");  // 이미지 변경 코드 _이미지 위치 수정

            string connStr = "server=localhost;port=3306;database=mydb;uid=root;pwd=1234";
            conn = new MySqlConnection(connStr);
            dataAdaptera = new MySqlDataAdapter("SELECT * FROM animation", conn);
            dataSeta = new DataSet();
            dataAdapterb = new MySqlDataAdapter("SELECT * FROM voice_actor", conn);
            dataSetb = new DataSet();
            dataAdapterc = new MySqlDataAdapter("SELECT * FROM studio", conn);
            dataSetc = new DataSet();

            dataAdaptera.Fill(dataSeta, "animation");
            dataGridView1.DataSource = dataSeta.Tables["animation"];
            dataAdapterb.Fill(dataSetb, "voice_actor");
            dataGridView2.DataSource = dataSetb.Tables["voice_actor"];
            dataAdapterc.Fill(dataSetc, "studio");
            dataGridView3.DataSource = dataSetc.Tables["studio"];
            //SetSearchComboBox();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataSeta.Clear();
            dataSetb.Clear();
            dataSetc.Clear();

            dataAdaptera = new MySqlDataAdapter("SELECT * FROM animation", conn);
            dataAdaptera.Fill(dataSeta, "animation");
            dataGridView1.DataSource = dataSeta.Tables["animation"];

            dataAdapterb = new MySqlDataAdapter("SELECT * FROM voice_actor", conn);
            dataAdapterb.Fill(dataSetb, "voice_actor");
            dataGridView2.DataSource = dataSetb.Tables["voice_actor"];

            dataAdapterc = new MySqlDataAdapter("SELECT * FROM studio", conn);
            dataAdapterc.Fill(dataSetc, "studio");
            dataGridView3.DataSource = dataSetc.Tables["studio"];
        }

        private void button1_Click(object sender, EventArgs e) //검색 버튼임
        {
            string queryStr;
            string condition_age;
            if (tabControl1.SelectedTab == tabControl1.TabPages[1]) //성우텝
            {
                string[] conditions = new string[4];
                conditions[0] = (textBox1.Text != "") ? "name=@name" : null;
                if (textBox3.Text != "" && textBox4.Text != "")
                {
                    condition_age = "age>=@min and age<=@max";
                }
                else if (textBox3.Text != "" || textBox4.Text != "")
                {
                    if (textBox3.Text != "")
                        condition_age = "age>=@min";
                    else
                        condition_age = "age <= @max";
                }
                else
                {
                    condition_age = null;
                }
                conditions[1] = condition_age;


                if (conditions[0] != null || conditions[1] != null || conditions[2] != null || conditions[3] != null)
                {
                    queryStr = $"SELECT * FROM voice_actor WHERE ";
                    bool firstCondition = true;
                    for (int i = 0; i < conditions.Length; i++)
                    {
                        if (conditions[i] != null)
                            if (firstCondition)
                            {
                                queryStr += conditions[i];
                                firstCondition = false;
                            }
                            else
                            {
                                queryStr += " and " + conditions[i];
                            }
                    }
                }
                else
                {
                    queryStr = "SELECT * FROM voice_actor";
                }

                dataAdapterb.SelectCommand = new MySqlCommand(queryStr, conn);
                dataAdapterb.SelectCommand.Parameters.AddWithValue("@name", textBox1.Text);
                dataAdapterb.SelectCommand.Parameters.AddWithValue("@min", textBox3.Text);
                dataAdapterb.SelectCommand.Parameters.AddWithValue("@max", textBox4.Text);

                try
                {
                    conn.Open();    
                    dataSetb.Clear();
                    if (dataAdapterb.Fill(dataSetb, "voice_actor") > 0)
                        dataGridView1.DataSource = dataSetb.Tables["voice_actor"];
                    else
                        MessageBox.Show("찾는 데이터가 없습니다.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    conn.Close();
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)  //검색 초기화 버튼
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages[0])
            {
                comboBox1.Text = "";
                comboBox2.Text = "";
                comboBox3.Text = "";
                textBox2.Clear();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages[1])
            {
                textBox1.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
                textBox7.Clear();
                textBox8.Clear();
                textBox9.Clear();
                //radioButton1.Checked = false;
                //radioButton2.Checked = false;
            }
            else
            {
                textBox6.Clear();
            }

            dataSeta.Clear();
            dataSetb.Clear();
            dataSetc.Clear();

            dataAdaptera = new MySqlDataAdapter("SELECT * FROM animation", conn);
            dataAdaptera.Fill(dataSeta, "animation");
            dataGridView1.DataSource = dataSeta.Tables["animation"];

            dataAdapterb = new MySqlDataAdapter("SELECT * FROM voice_actor", conn);
            dataAdapterb.Fill(dataSetb, "voice_actor");
            dataGridView2.DataSource = dataSetb.Tables["voice_actor"];

            dataAdapterc = new MySqlDataAdapter("SELECT * FROM studio", conn);
            dataAdapterc.Fill(dataSetc, "studio");
            dataGridView3.DataSource = dataSetc.Tables["studio"];
        }

        private void button2_Click(object sender, EventArgs e)//삭제 버튼
        {
            string target = textBox10.Text;

            string query = "delete from voice_actor where idname=@idname";
            dataAdapterb.DeleteCommand = new MySqlCommand(query, conn);
            dataAdapterb.DeleteCommand.Parameters.Add("@idname", MySqlDbType.Int32);
            dataAdapterb.DeleteCommand.Parameters["@idname"].Value = target;
            try
            {
                DataRow[] findRows = dataSetb.Tables["voice_actor"].Select($"idname={target}");
                findRows[0].Delete();
                dataAdapterb.Update(dataSetb, "voice_actor");
                MessageBox.Show("성우 사라짐", "삭제 완료");
            }
            catch
            {
                throw;
            }
        }

        private void button3_Click(object sender, EventArgs e)//수정 버튼
        {
            #region Update() 이용
            // Update를 호출하기 전에 명령을 명시적으로 설정해야 한다. 
            string sql = "UPDATE voice_actor SET agency=@agency WHERE gender=@gender";
            dataAdapterb.UpdateCommand = new MySqlCommand(sql, conn);
            dataAdapterb.UpdateCommand.Parameters.AddWithValue("@name", textBox1.Text);
            dataAdapterb.UpdateCommand.Parameters.AddWithValue("@age", textBox3.Text);
            dataAdapterb.UpdateCommand.Parameters.AddWithValue("@gender", textBox7.Text);
            dataAdapterb.UpdateCommand.Parameters.AddWithValue("@agency", textBox8.Text);
            dataAdapterb.UpdateCommand.Parameters.AddWithValue("@birth", textBox9.Text);

            //int id = (int)dataGridView1.SelectedRows[0].Cells["id"].Value;
            //string filter = "id=" + id;
            //DataRow[] findRows = dataSet.Tables["city"].Select(filter);
            //findRows[0]["id"] = id;
            //findRows[0]["name"] = txtName.Text;
            //findRows[0]["countrycode"] = txtCountryCode.Text;
            //findRows[0]["district"] = txtDistrict.Text;
            //findRows[0]["population"] = txtPopulation.Text;

            var selectedRows = dataGridView2.SelectedRows;
            int id;
            string filter;
            for (int i = 0; i < selectedRows.Count; i++)
            {
                id = (int)dataGridView2.SelectedRows[i].Cells["id"].Value;
                filter = "id=" + id;
                DataRow[] findRows = dataSetb.Tables["voice_actor"].Select(filter);
                findRows[0]["name"] = id;
                findRows[0]["age"] = (string)dataGridView2.SelectedRows[i].Cells["age"].Value;
                findRows[0]["gender"] = (string)dataGridView2.SelectedRows[i].Cells["gender"].Value;
                findRows[0]["agency"] = textBox8.Text;
                findRows[0]["birth"] = (string)dataGridView2.SelectedRows[i].Cells["birth"].Value;
            }

            dataAdapterb.Update(dataSetb, "voice_actor");
            //dataSet.Clear();
            //dataAdapter.Fill(dataSet, "city");
            //dataGridView1.DataSource = dataSet.Tables["city"];
            #endregion
        }

        private void button5_Click(object sender, EventArgs e)//외부 파일 저장 버튼
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages[0])
            {
                if (dataGridView1.RowCount == 0)
                {
                    MessageBox.Show("저장할 테이터가 없습니다. ", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // radio button가 체크 되어 있는가
                if (radioButton3.Checked)
                {
                    saveFileDialog1.Filter = "텍스트 파일(*.txt)|*.txt";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        fsPath = saveFileDialog1.FileName;  // SaveFileDialog에 지정한 파일경로를 전역 변수인 fsPath에 저장
                        TextFileSave();
                    }
                }
                else if (radioButton4.Checked)
                {
                    saveFileDialog1.Filter = "엑셀 파일(*.xlsx)|*.xlsx";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        fsPath = saveFileDialog1.FileName;
                        ExcelFileSave();
                    }
                }
                else
                {
                    MessageBox.Show("형식이 지정되어 있지 않습니다. ", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else if(tabControl1.SelectedTab == tabControl1.TabPages[1])
            {
                if (dataGridView2.RowCount == 0)
                {
                    MessageBox.Show("저장할 테이터가 없습니다. ", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // radio button가 체크 되어 있는가
                if (radioButton3.Checked)
                {
                    saveFileDialog1.Filter = "텍스트 파일(*.txt)|*.txt";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        fsPath = saveFileDialog1.FileName;  // SaveFileDialog에 지정한 파일경로를 전역 변수인 fsPath에 저장
                        TextFileSave();
                    }
                }
                else if (radioButton4.Checked)
                {
                    saveFileDialog1.Filter = "엑셀 파일(*.xlsx)|*.xlsx";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        fsPath = saveFileDialog1.FileName;
                        ExcelFileSave();
                    }
                }
                else
                {
                    MessageBox.Show("형식이 지정되어 있지 않습니다. ", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                if (dataGridView3.RowCount == 0)
                {
                    MessageBox.Show("저장할 테이터가 없습니다. ", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // radio button가 체크 되어 있는가
                if (radioButton3.Checked)
                {
                    saveFileDialog1.Filter = "텍스트 파일(*.txt)|*.txt";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        fsPath = saveFileDialog1.FileName;  // SaveFileDialog에 지정한 파일경로를 전역 변수인 fsPath에 저장
                        TextFileSave();
                    }
                }
                else if (radioButton4.Checked)
                {
                    saveFileDialog1.Filter = "엑셀 파일(*.xlsx)|*.xlsx";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        fsPath = saveFileDialog1.FileName;
                        ExcelFileSave();
                    }
                }
                else
                {
                    MessageBox.Show("형식이 지정되어 있지 않습니다. ", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void ExcelFileSave()//엑셀 저장
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages[0]) 
            {
                #region 1. 엑셀 사용에 필요한 객체 생성
                // 엑셀을 사용하기 위한 클래스 객체 생성
                Excel.Application eApp;     // 엑셀 프로그램 
                Excel.Workbook eWorkbook;   // 여러 WorkSheet 포함한 단위
                Excel.Worksheet eWorkSheet;

                string[,] data;     // 엑셀에 데이터를 저장하기 위한 2차원 배열

                eApp = new Excel.Application();         // 엑셀 프로그램 객체 생성
                eWorkbook = eApp.Workbooks.Add(true);   // 엑셀에 Workbook 추가, 초기화
                eWorkSheet = eWorkbook.Sheets[1] as Excel.Worksheet;    // WorkSheet 생성, Excel Sheet 배열은 1부터 시작한다.
                #endregion

                #region 2. 엑셀에 데이터를 저장할 2차원 데이터 배열(data[,]) 준비
                // 엑셀에 저장할 데이터 크기 지정
                int cnum = dataSeta.Tables["animation"].Columns.Count + 1;
                int rnum = dataSeta.Tables["animation"].Rows.Count + 1;
                data = new string[rnum + 1, cnum + 1];

                // 엑셀에 저장할 2차원 배열에 Column 이름 저장
                for (int i = 0; i < dataSeta.Tables["animation"].Columns.Count; i++)
                {
                    data[0, i] = dataSeta.Tables["animation"].Columns[i].ColumnName;
                }

                // 엑셀에 저장할 2차원 배열에 데이터 저장
                for (int i = 0; i < dataSeta.Tables["animation"].Rows.Count; i++)                    // 리스트뷰의 행수만큼 반복
                {
                    for (int j = 0; j < dataSeta.Tables["animation"].Columns.Count; j++)    // 한 행의 열수만큼 반복
                    {
                        data[i + 1, j] = dataSeta.Tables["animation"].Rows[i].ItemArray[j].ToString();    // data 배열에 데이터 저장
                    }
                }
                #endregion

                #region 3. 준비된 데이터를 엑셀에 저장
                //string EndStr = "F" + rnum.ToString();      // 8개의 파일을 선택한 경우 F9 => 마지막 데이터 저장셀 주소
                string EndStr = Convert.ToChar(cnum - 2 + 65) + rnum.ToString();      // 8개의 파일을 선택한 경우 F9 => 마지막 데이터 저장셀 주소
                eWorkSheet.get_Range("A1:" + EndStr).Value = data;     // 데이터 기록

                eWorkbook.SaveAs(fsPath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false,
                    Excel.XlSaveAsAccessMode.xlShared, false, false, Type.Missing, Type.Missing, Type.Missing);
                eWorkbook.Close(false, Type.Missing, Type.Missing);
                eApp.Quit();
                #endregion
            }
            else if(tabControl1.SelectedTab == tabControl1.TabPages[1])
            {
                #region 1. 엑셀 사용에 필요한 객체 생성
                // 엑셀을 사용하기 위한 클래스 객체 생성
                Excel.Application eApp;     // 엑셀 프로그램 
                Excel.Workbook eWorkbook;   // 여러 WorkSheet 포함한 단위
                Excel.Worksheet eWorkSheet;

                string[,] data;     // 엑셀에 데이터를 저장하기 위한 2차원 배열

                eApp = new Excel.Application();         // 엑셀 프로그램 객체 생성
                eWorkbook = eApp.Workbooks.Add(true);   // 엑셀에 Workbook 추가, 초기화
                eWorkSheet = eWorkbook.Sheets[1] as Excel.Worksheet;    // WorkSheet 생성, Excel Sheet 배열은 1부터 시작한다.
                #endregion

                #region 2. 엑셀에 데이터를 저장할 2차원 데이터 배열(data[,]) 준비
                // 엑셀에 저장할 데이터 크기 지정
                int cnum = dataSetb.Tables["voice_actor"].Columns.Count + 1;
                int rnum = dataSetb.Tables["voice_actor"].Rows.Count + 1;
                data = new string[rnum + 1, cnum + 1];

                // 엑셀에 저장할 2차원 배열에 Column 이름 저장
                for (int i = 0; i < dataSetb.Tables["voice_actor"].Columns.Count; i++)
                {
                    data[0, i] = dataSetb.Tables["voice_actor"].Columns[i].ColumnName;
                }

                // 엑셀에 저장할 2차원 배열에 데이터 저장
                for (int i = 0; i < dataSetb.Tables["voice_actor"].Rows.Count; i++)                    // 리스트뷰의 행수만큼 반복
                {
                    for (int j = 0; j < dataSetb.Tables["voice_actor"].Columns.Count; j++)    // 한 행의 열수만큼 반복
                    {
                        data[i + 1, j] = dataSetb.Tables["voice_actor"].Rows[i].ItemArray[j].ToString();    // data 배열에 데이터 저장
                    }
                }
                #endregion

                #region 3. 준비된 데이터를 엑셀에 저장
                //string EndStr = "F" + rnum.ToString();      // 8개의 파일을 선택한 경우 F9 => 마지막 데이터 저장셀 주소
                string EndStr = Convert.ToChar(cnum - 2 + 65) + rnum.ToString();      // 8개의 파일을 선택한 경우 F9 => 마지막 데이터 저장셀 주소
                eWorkSheet.get_Range("A1:" + EndStr).Value = data;     // 데이터 기록

                eWorkbook.SaveAs(fsPath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false,
                    Excel.XlSaveAsAccessMode.xlShared, false, false, Type.Missing, Type.Missing, Type.Missing);
                eWorkbook.Close(false, Type.Missing, Type.Missing);
                eApp.Quit();
                #endregion
            }
            else
            {
                #region 1. 엑셀 사용에 필요한 객체 생성
                // 엑셀을 사용하기 위한 클래스 객체 생성
                Excel.Application eApp;     // 엑셀 프로그램 
                Excel.Workbook eWorkbook;   // 여러 WorkSheet 포함한 단위
                Excel.Worksheet eWorkSheet;

                string[,] data;     // 엑셀에 데이터를 저장하기 위한 2차원 배열

                eApp = new Excel.Application();         // 엑셀 프로그램 객체 생성
                eWorkbook = eApp.Workbooks.Add(true);   // 엑셀에 Workbook 추가, 초기화
                eWorkSheet = eWorkbook.Sheets[1] as Excel.Worksheet;    // WorkSheet 생성, Excel Sheet 배열은 1부터 시작한다.
                #endregion

                #region 2. 엑셀에 데이터를 저장할 2차원 데이터 배열(data[,]) 준비
                // 엑셀에 저장할 데이터 크기 지정
                int cnum = dataSetc.Tables["studio"].Columns.Count + 1;
                int rnum = dataSetc.Tables["studio"].Rows.Count + 1;
                data = new string[rnum + 1, cnum + 1];

                // 엑셀에 저장할 2차원 배열에 Column 이름 저장
                for (int i = 0; i < dataSetc.Tables["studio"].Columns.Count; i++)
                {
                    data[0, i] = dataSetc.Tables["studio"].Columns[i].ColumnName;
                }

                // 엑셀에 저장할 2차원 배열에 데이터 저장
                for (int i = 0; i < dataSetc.Tables["studio"].Rows.Count; i++)                    // 리스트뷰의 행수만큼 반복
                {
                    for (int j = 0; j < dataSetc.Tables["studio"].Columns.Count; j++)    // 한 행의 열수만큼 반복
                    {
                        data[i + 1, j] = dataSetc.Tables["studio"].Rows[i].ItemArray[j].ToString();    // data 배열에 데이터 저장
                    }
                }
                #endregion

                #region 3. 준비된 데이터를 엑셀에 저장
                //string EndStr = "F" + rnum.ToString();      // 8개의 파일을 선택한 경우 F9 => 마지막 데이터 저장셀 주소
                string EndStr = Convert.ToChar(cnum - 2 + 65) + rnum.ToString();      // 8개의 파일을 선택한 경우 F9 => 마지막 데이터 저장셀 주소
                eWorkSheet.get_Range("A1:" + EndStr).Value = data;     // 데이터 기록

                eWorkbook.SaveAs(fsPath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false,
                    Excel.XlSaveAsAccessMode.xlShared, false, false, Type.Missing, Type.Missing, Type.Missing);
                eWorkbook.Close(false, Type.Missing, Type.Missing);
                eApp.Quit();
                #endregion
            }
        }

        private void TextFileSave()// 텍스트 파일 저장
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages[0])
            {
                using (System.IO.StreamWriter sw = new StreamWriter(fsPath))
                {
                    // Column Title 한번 출력
                    foreach (DataColumn col in dataSeta.Tables["animation"].Columns)
                    {
                        sw.Write($"{col.ColumnName}\t");
                    }
                    sw.WriteLine();

                    // DataGridView에 기록된 모든 파일 정보를 Text파일에 출력
                    foreach (DataRow row in dataSeta.Tables["animation"].Rows)
                    {
                        string rowString = "";
                        foreach (var item in row.ItemArray)
                        {
                            rowString += item.ToString() + "\t";
                        }
                        sw.WriteLine(rowString);
                    }
                    sw.Close();
                }
            }
            else if(tabControl1.SelectedTab == tabControl1.TabPages[0])
            {
                using (System.IO.StreamWriter sw = new StreamWriter(fsPath))
                {
                    // Column Title 한번 출력
                    foreach (DataColumn col in dataSetb.Tables["voice_actor"].Columns)
                    {
                        sw.Write($"{col.ColumnName}\t");
                    }
                    sw.WriteLine();

                    // DataGridView에 기록된 모든 파일 정보를 Text파일에 출력
                    foreach (DataRow row in dataSetb.Tables["voice_actor"].Rows)
                    {
                        string rowString = "";
                        foreach (var item in row.ItemArray)
                        {
                            rowString += item.ToString() + "\t";
                        }
                        sw.WriteLine(rowString);
                    }
                    sw.Close();
                }
            }
            else
            {
                using (System.IO.StreamWriter sw = new StreamWriter(fsPath))
                {
                    // Column Title 한번 출력
                    foreach (DataColumn col in dataSetc.Tables["studio"].Columns)
                    {
                        sw.Write($"{col.ColumnName}\t");
                    }
                    sw.WriteLine();

                    // DataGridView에 기록된 모든 파일 정보를 Text파일에 출력
                    foreach (DataRow row in dataSetc.Tables["studio"].Rows)
                    {
                        string rowString = "";
                        foreach (var item in row.ItemArray)
                        {
                            rowString += item.ToString() + "\t";
                        }
                        sw.WriteLine(rowString);
                    }
                    sw.Close();
                }
            }
        }

        private void GroupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void Button6_Click(object sender, EventArgs e)
        {
            string queryStr = "INSERT INTO voice_actor (name, age, gender,agency,birth) " +
    "VALUES(@name, @age, @gender, @agency, @birth)";

            dataAdapterb.InsertCommand = new MySqlCommand(queryStr, conn);
            dataAdapterb.InsertCommand.Parameters.Add("@name", MySqlDbType.VarChar);
            dataAdapterb.InsertCommand.Parameters.Add("@age", MySqlDbType.VarChar);
            dataAdapterb.InsertCommand.Parameters.Add("@gender", MySqlDbType.VarChar);
            dataAdapterb.InsertCommand.Parameters.Add("@agency", MySqlDbType.VarChar);
            dataAdapterb.InsertCommand.Parameters.Add("@birth", MySqlDbType.Date);
            dataAdapterb.InsertCommand.Parameters.Add("@idname", MySqlDbType.Int32);
            #region Parameter를 이용한 처리
            //dataAdapter.InsertCommand.Parameters["@name"].Value = txtName.Text;
            //dataAdapter.InsertCommand.Parameters["@countrycode"].Value = txtCountryCode.Text;
            //dataAdapter.InsertCommand.Parameters["@district"].Value = txtDistrict.Text;
            //dataAdapter.InsertCommand.Parameters["@population"].Value = txtPopulation.Text;

            //try
            //{
            //    conn.Open();
            //    dataAdapter.InsertCommand.ExecuteNonQuery();

            //    dataSet.Clear();                                        // 이전 데이터 지우기
            //    dataAdapter.Fill(dataSet, "city");                      // DB -> DataSet
            //    dataGridView1.DataSource = dataSet.Tables["city"];      // dataGridView에 테이블 표시
            //    ClearTextBoxes();                                       // 텍스트 박스 내용 지우기
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
            //finally
            //{
            //    conn.Close();
            //}
            #endregion

            #region MySqlDataAdapter.Update()를 이용한 처리
            DataRow newRow = dataSetb.Tables["voice_actor"].NewRow();
            newRow["name"] = textBox1.Text;
            newRow["age"] = textBox3.Text;
            newRow["gender"] = textBox7.Text;
            newRow["agency"] = textBox8.Text;
            newRow["birth"] = textBox9.Text;
            newRow["idname"] = textBox10.Text;
            dataSetb.Tables["voice_actor"].Rows.Add(newRow);

            dataAdapterb.InsertCommand.Parameters["@name"].Value = newRow["name"];
            dataAdapterb.InsertCommand.Parameters["@age"].Value = newRow["age"];
            dataAdapterb.InsertCommand.Parameters["@gender"].Value = newRow["gender"];
            dataAdapterb.InsertCommand.Parameters["@agency"].Value = newRow["agency"];
            dataAdapterb.InsertCommand.Parameters["@idname"].Value = newRow["idname"];
            dataAdapterb.InsertCommand.Parameters["@idname"].Value = newRow["idname"];
            dataAdapterb.Update(dataSetb, "voice_actor");

            dataSetb.Clear();
            dataAdapterb.Fill(dataSetb, "voice_actor");
            dataGridView2.DataSource = dataSetb.Tables["voice_actor"];
            #endregion
        }

        private void Label14_Click(object sender, EventArgs e)
        {

        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void GroupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void Button10_Click(object sender, EventArgs e)
        {
            if (textBox17.Text == "")
            {
                MessageBox.Show("ID를 입력하세요");
            }
            else
            {
                conn.Open();
                string query = "select * from voice_actor where idname = @idname";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.Add("@idname", MySqlDbType.Int32);
                cmd.Parameters["@idname"].Value = int.Parse(textBox17.Text);
                try
                {
                    MySqlDataReader reader = cmd.ExecuteReader();

                    if (reader.Read())
                    {
                        update_cnt = 1;
                        textBox14.Text = reader.GetString("agency");
                    }
                    else
                    {
                        MessageBox.Show("성우를 찾을 수 없음");
                    }

                    reader.Close();

                }
                catch (Exception)
                {
                    throw;
                }

                conn.Close();
            }
        }

        private void Button3_Click_1(object sender, EventArgs e)
        {
            if (update_cnt == 0)
            {
                MessageBox.Show("ID를 입력하고 조회누르세요.");
            }
            //name = @name, age = @age, agency = @agency,  gender = @gender, birth = @birth
            else
            {
                string sql = "UPDATE voice_actor SET  agency=@agency where idname=@idname";
                dataAdapterb.UpdateCommand = new MySqlCommand(sql, conn);

                dataAdapterb.UpdateCommand.Parameters.AddWithValue("@idname", textBox17.Text);

                if (textBox18.Text == "")
                {
                    dataAdapterb.UpdateCommand.Parameters.AddWithValue("@agency", textBox14.Text);
                }
                else
                {
                    dataAdapterb.UpdateCommand.Parameters.AddWithValue("@agency", textBox18.Text);
                }
                try
                {
                    conn.Open();

                    if (dataAdapterb.UpdateCommand.ExecuteNonQuery() > 0)
                    {
                        dataSetb.Clear();
                        dataAdapterb.Fill(dataSetb, "voice_actor");
                        dataGridView2.DataSource = dataSetb.Tables["voice_actor"];
                        MessageBox.Show("소속사 바뀜.", "수정 완료");
                        textBox14.Text = "";
                     
                        update_cnt = 0;
                    }
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    conn.Close();
                }
            }
        }
    }
}
