using System;
using System.Windows.Forms;
using System.Drawing;
using System.Data;
using Mono.Data.Sqlite;

class MForm : Form//, IDisposable
{

    private DataGridView dgv = null;      
    private DataSet ds = null;
	SqliteDataAdapter da=null;
	SqliteConnection con;

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

        dgv.Location = new Point(8, 0);
        dgv.Size = new Size(350, 300);
        dgv.TabIndex = 0;
        dgv.Parent = this;        
		dgv.ReadOnly=true;
		dgv.AllowUserToAddRows=false;
    }

	void createData(){
        string cs = "URI=file:test.db";
		string createCmd = "CREATE TABLE IF NOT EXISTS Cars "+
			"( id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, "+
			" Name NVarChar, "+
			" Note NVarChar"+
				");";
		string someDataInsert=
			"INSERT INTO Cars (id, Name, Note) VALUES ("+
				"NULL, 'BMW', 'Bayern')";

		SqliteCommand command = new SqliteCommand();
		command.CommandText=createCmd;
		command.Connection=con;
		int iRes=command.ExecuteNonQuery();


		command.CommandText=someDataInsert;
		iRes=command.ExecuteNonQuery();

	}

    void InitData()
    {    
        string cs = "URI=file:test.db";

//		createData();

        string stm = "SELECT * FROM Cars";

		con = new SqliteConnection(cs);
        
            con.Open();

            ds = new DataSet();
			SqliteCommandBuilder cmdBuilder;
			da = new SqliteDataAdapter(stm, con);

                da.Fill(ds, "Cars");        
				cmdBuilder = new SqliteCommandBuilder(da);

                dgv.DataSource = ds.Tables["Cars"];

				da.InsertCommand=
					new SqliteCommand("INSERT INTO [Cars] ([id], [Name], [Note]) VALUES (NULL, @param1, @param2)");

				addRowSQL("audi sql","another note");

				da.Update(ds.Tables[0]);
                         
    }

	void doRefresh(){
		dgv.ResetBindings();
		//dgv.DataSource=null;
		dgv.DataSource=ds.Tables[0];
	}

	void addRowSQL(string sName, string sNote){
		//change sql data
		SqliteCommand cmd=new SqliteCommand(
			string.Format("INSERT INTO Cars (id,name,note) VALUES(NULL,'{0}sql','{1}');",
		              sName, sNote),con);
		cmd.ExecuteNonQuery();
		//get autoincrement value
		cmd=new SqliteCommand("SELECT last_insert_rowid()",con);
		long lastID = (long) cmd.ExecuteScalar();

		System.Console.WriteLine("added new row with "+lastID);

		//change datagrid
		DataTable dt = ds.Tables[0];
		DataRow dr = dt.NewRow();
		dr["name"]=sName;
		dr["note"]=sNote;
		dr["id"]=lastID;
		dt.Rows.Add(dr);
		dt.AcceptChanges();
		ds.AcceptChanges();
		da.Update(ds.Tables[0]);
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
