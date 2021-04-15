using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ConsoleApp1.helpers
{
    class DataTableCases
    {
        public static void DatatableTestCopy()
        {

            // 创建一个空表
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable("Table_New");


            DataColumn title = new DataColumn("Title", typeof(String));
            dt.Columns.Add(title);

            DataColumn rank = new DataColumn("Rank", typeof(String));
            //rank.ColumnName
            dt.Columns.Add(rank);

            DataRow dr = dt.NewRow();

            dr[0] = "Amanda rocks";
            dr[1] = "345万";
            dt.Rows.Add(dr);
            List<Object> drArr = dr.ItemArray.ToList().GetRange(0, 2);

            Console.WriteLine(dt.Columns.Contains("Title")); // True
            
            dt.Rows.Add("Chen Wong", "233万");
            dt.Rows.Add("Lop Dio1", "275万");
            dt.Rows.Add("Lop Dio2", "295万");
            dt.Rows.Add("Lop Dio3", "235万");

            DataView dv = dt.DefaultView;
            dv.RowFilter = "`Title`like'Lop Dio%' and `Rank` > '255'";

            DataTable newdt = dt.DefaultView.ToTable(false, new string[] { "Title" });

            foreach (DataRow row in newdt.Rows)
            {
                Console.WriteLine(row["Title"]);
            }

            DataTable dt3 = dt.Copy();


            Console.WriteLine("++++" + dt.Rows[0][0]);
            dt.Rows[0][0] = "Amanda changed";
            Console.WriteLine("++++" + dt.Rows[0][0]);
            // var data = dt.AsEnumerable();
            foreach (DataRow row in dt.Rows)
            {
                Console.WriteLine(row["Title"]);
                Console.WriteLine(row["Rank"]);
            }

            Console.WriteLine("------------------Lop Dio--------------------");
            DataRow[] drs = dt.Select("Title like 'Lop Dio%'");
            foreach (DataRow row in drs)
            {
                Console.WriteLine(row["Title"]);
                Console.WriteLine(row["Rank"]);
            }


            Console.WriteLine("------------------only keep 0 and 1--------------------");
            for (int i = dt.Rows.Count - 1; i >= 2; i--)
            {
                dt.Rows.RemoveAt(i);
            }

            foreach (DataRow row in dt.Rows)
            {
                Console.WriteLine(row["Title"]);
                Console.WriteLine(row["Rank"]);
            }

           
        }


        static void DatatableTest()
        {

            // 创建一个空表
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable("Table_New");

            DataColumn Id1 = new DataColumn("Id", typeof(Int32));
            dt.Columns.Add(Id1);


            DataColumn Id = new DataColumn("Id", typeof(Int32));
            Id.AutoIncrement = true;

            dt2.Columns.Add(Id);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = dt2.Columns["Id"];
            dt2.PrimaryKey = keys;

            dt2.Merge(dt);
            DataRow dr = dt2.NewRow();
            dt2.Rows.Add(dr);

            Console.WriteLine(dt2);

            //DataColumn gender1 = new DataColumn("Gender", typeof(String));
            DataColumn gender1 = new DataColumn("isMale", typeof(Boolean));
            dt2.Columns.Add(gender1);
            gender1.DefaultValue = true;

            DataColumn name = new DataColumn("Name", typeof(String));
            dt.Columns.Add(name);
            int colIndex;
            int i = 0;
            foreach (DataColumn col in dt2.Columns)
            {
                if (col.ColumnName.Replace("\n", "").Replace(" ", "") == "最近6个月平均应还款")
                {
                    colIndex = i;
                    break;
                }
                i = i + 1;
            }

            //dt2.Rows.Remove(dt2.Rows.Find(new Object[] { 1}));
            //dt.Columns.Remove(dt2.Columns["Gender"]);


            //DataColumn id = new DataColumn("id", typeof(Int32));
            //dt.Columns.Add(id);



            //dt.Columns.Remove(gender2);

            //dt.Columns.RemoveAt(0);


            DataColumn 姓名2 = new DataColumn("姓名", typeof(String));

            //Bitmap bm = new Bitmap("D:/云扩科技/RPA/OEMLOGO.bmp");
            //Console.WriteLine(bm.Height);
            //DataRow dr = dt.NewRow();
            DataRow dr2 = dt2.NewRow();

            dr2[0] = 8;

            Int32 j = 0;
            String stra = "";
            while (j < 1)
            {
                stra = stra + "a";
                j = j + 1;
            }

            dr2[2] = stra;
            dt2.Rows.Add(dr2);

            //DataRow dr3 = dt2.NewRow();
            //dr3[0] = false;
            //dt2.Rows.Add(dr3);


            Console.WriteLine(dt2);


            //dt2.Merge(dt);
            Console.WriteLine(dt2.Rows.Count);
            //foreach (DataRow row in dt2.Rows)
            //{
            //    Console.WriteLine(row["姓名"]); 
            //}


        }

    }
}
