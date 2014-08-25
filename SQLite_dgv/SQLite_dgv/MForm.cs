using System;
using System.Windows.Forms;
using System.Drawing;
using System.Xml;
using System.Data;
using System.Data.SQLite;


class MForm : Form//, IDisposable
{

    private DataGridView dgv = null;      
        string cs = "URI=file:test.db";
	TextBox txtFilter = null;
	bool bFiltered=false;

	Timer timer;
//	public void Dispose (){
//		con.Close();
//	}

    public MForm()
    {

       this.Text = "DataGridView";
       this.Size = new Size(450, 350);
       
       this.InitUI();
       this.InitData();
       
       this.CenterToScreen();
		timer=new Timer();
		timer.Tick+=new EventHandler(timer_tick);
		timer.Interval=5000;
		timer.Start();
    }

	int count=0;
	public void timer_tick(object s, EventArgs a){
		count++;
		addRowSQL("test"+count.ToString(), "test");
	}
    
    void InitUI()
    {    
        dgv = new DataGridView();

        dgv.Location = new Point(8, 30);
        dgv.Size = new Size(350, 300);
        dgv.TabIndex = 0;
        dgv.Parent = this;        
		dgv.ReadOnly=true;
		dgv.AllowUserToAddRows=false;

		txtFilter=new TextBox();
		txtFilter.Location=new Point(8,8);
		txtFilter.Size=new System.Drawing.Size(200,80);
		txtFilter.Parent=this;

		txtFilter.TextChanged+=new EventHandler(txt_Changed);
    }

	void txt_Changed(object s, EventArgs a){
		if(txtFilter.Text.Length>0){
			filterData(txtFilter.Text);
		}
		else{
			DataTable dt = (DataTable)dgv.DataSource;
			dt.DefaultView.RowFilter = "";
			bFiltered=false;
		}
	}

	void filterData(string sID){
		DataTable dt = (DataTable)dgv.DataSource;
		dt.DefaultView.RowFilter = "Name like '%" + txtFilter.Text + "%'";
		bFiltered=true;		
	}

	void createData(){
        if (System.IO.File.Exists("test.db"))
            return;

		string createCmd = "CREATE TABLE IF NOT EXISTS Cars "+
			"( id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, "+
			" Name NVarChar, "+
			" Note NVarChar"+
				");";
//		string someDataInsert=
//			"INSERT INTO Cars (id, Name, Note) VALUES ("+
//				"NULL, 'BMW', 'Bayern')";
		using(SQLiteConnection connection=new SQLiteConnection(cs)){
            connection.Open();
			SQLiteCommand command = new SQLiteCommand(createCmd, connection);
			int iRes=command.ExecuteNonQuery();
		}

		//command.CommandText=someDataInsert;
		//iRes=command.ExecuteNonQuery();

	}

    void InitData()
    {
        createData();

        string stm = "SELECT * FROM Cars";
		using(SQLiteConnection connection=new SQLiteConnection(cs)){
			connection.Open();
            System.Data.DataSet ds = new DataSet();
			SQLiteDataAdapter da = new SQLiteDataAdapter(stm, connection);
			SQLiteCommandBuilder cmdBuilder=new SQLiteCommandBuilder(da);
				da.InsertCommand=
					new SQLiteCommand("INSERT INTO [Cars] ([id], [Name], [Note]) VALUES (NULL, @param1, @param2)");

                da.Fill(ds, "Cars");        
                dgv.DataSource = ds.Tables["Cars"];
		}                         
    }

	void doRefresh(){
		if(bFiltered)
			return;
		int iRow=0, iCol=0;
		if(dgv.Rows.Count>0){
			iRow=dgv.CurrentCell.RowIndex;
			iCol=dgv.CurrentCell.ColumnIndex;
		}
		dgv.ResetBindings();
		//dgv.DataSource=null;
		using(SQLiteConnection connection=new SQLiteConnection(cs)){
			connection.Open();
			System.Data.DataSet ds=new DataSet();
			SQLiteDataAdapter da = new SQLiteDataAdapter("SELECT * FROM Cars", connection);
			da.Fill(ds);
			dgv.DataSource=ds.Tables[0];
		}
		if(dgv.Rows.Count>0){
			dgv.CurrentCell=dgv.Rows[iRow].Cells[iCol];
		}
	}

	void addRowSQL(string sName, string sNote){
		//change sql data
		using (SQLiteConnection connection = new SQLiteConnection(cs)){
			connection.Open();
			SQLiteCommand cmd=new SQLiteCommand(
				string.Format("INSERT INTO Cars (id,name,note) VALUES(NULL,'{0}sql','{1}');",
			              sName, sNote),connection);
			cmd.ExecuteNonQuery();
			//get autoincrement value
			cmd=new SQLiteCommand("SELECT last_insert_rowid()",connection);
			long lastID = (long) cmd.ExecuteScalar();

			System.Console.WriteLine("added new row with "+lastID);
		}
		//change datagrid
//		DataTable dt = ds.Tables[0];
//		DataRow dr = dt.NewRow();
//		dr["name"]=sName;
//		dr["note"]=sNote;
//		dr["id"]=lastID;
//		dt.Rows.Add(dr);
//		dt.AcceptChanges();
//		ds.AcceptChanges();
//		da.Update(ds.Tables[0]);
		doRefresh();

	}

}

class MApplication 
{
    public static void Main() 
    {
        Application.Run(new MForm());
    }
}
